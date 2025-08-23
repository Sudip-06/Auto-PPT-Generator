#!/usr/bin/env python3
"""
Auto PPT Generator - Robust LLM parsing (Gemini JSON mode, no Schema hard-dep),
enhanced prompting, markdown-aware fallback, template-preserving append,
style & image reuse from uploaded deck, adaptive layout to keep text inside slide,
and in-memory download. Keeps index.html unchanged.

Behavior with uploads:
- If user uploads a .pptx/.potx, we PRESERVE its existing slides
  and APPEND LLM-generated slides to the end.
- We extract fonts/colors and pictures from the uploaded file and apply/reuse them.
"""

import os
import re
import json
import tempfile
import logging
from io import BytesIO
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple

from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS

# --- LLM SDKs ---
import openai
import anthropic
import google.generativeai as genai

# --- PowerPoint ---
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE

# ------------------------------------------------------------------------------
# Logging & App
# ------------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("app")

app = Flask(__name__)
CORS(app, origins=["*"])
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB uploads


# ------------------------------------------------------------------------------
# PPT Generator
# ------------------------------------------------------------------------------
class PPTGenerator:
    """PowerPoint generator with robust JSON/markdown parsing and safe template reuse."""

    def __init__(self):
        self.supported_providers = ["openai", "anthropic", "google"]
        self.max_slides = 12
        self.min_slides = 3
        self._style_ctx: Optional[dict] = None

    # ------------------------------ Public API ------------------------------

    def parse_text_to_slides(
        self, text: str, provider: str, api_key: str, guidance: str = ""
    ) -> List[Dict]:
        """Parse raw text into normalized slide dicts using an LLM with robust fallbacks."""
        word_count = len((text or "").split())
        estimated_slides = max(self.min_slides, min(self.max_slides, word_count // 120 + 1))
        prompt = self._create_enhanced_prompt(text, guidance, estimated_slides)

        try:
            logger.info(f"Parsing with {provider}, estimated slides: {estimated_slides}")
            response = self._call_llm_with_retry(provider, api_key, prompt)
            slides = self._robust_json_extraction(response)
            slides = self._validate_and_enhance_slides(slides)

            # Always try one refinement pass for denser, more accurate bullets
            try:
                improved = self._refine_with_provider(provider, api_key, response, estimated_slides + 2)
                slides2 = self._robust_json_extraction(improved)
                slides2 = self._validate_and_enhance_slides(slides2)
                if self._score_slides(slides2) >= self._score_slides(slides):
                    slides = slides2
                    logger.info(f"Refinement improved slides to {len(slides)}")
            except Exception:
                logger.warning("Refinement pass failed; keeping initial slides")

            # ensure a conclusion slide exists at the end
            if not slides or slides[-1].get("slide_type") != "conclusion_slide":
                slides.append(
                    {
                        "title": "Conclusion & Next Steps",
                        "content": ["Summary of key points", "Actionable next steps", "Q&A"],
                        "slide_type": "conclusion_slide",
                        "speaker_notes": "Wrap up",
                    }
                )
            logger.info(f"Generated {len(slides)} slides successfully")
            return slides[: self.max_slides]
        except Exception as e:
            logger.error(f"Error parsing text: {e}")
            logger.warning("Using emergency fallback parsing")
            slides = self._emergency_fallback_parsing(text)
            return slides[: self.max_slides]

    def create_presentation(self, slides_data: List[Dict], template_file: Optional[str] = None) -> Presentation:
        """Create a PPTX from normalized slide data.

        If a PPTX/POTX is uploaded:
        - Extract style & assets.
        - PRESERVE existing slides from the upload.
        - APPEND the generated slides after them.
        """
        try:
            if template_file:
                prs = Presentation(template_file)
                style = self._extract_style_and_assets(prs)
                logger.info("Using uploaded template/presentation (preserving existing slides, appending new)")
            else:
                prs = Presentation()
                style = {
                    "title_font": None, "body_font": None,
                    "title_color": RGBColor(31, 73, 125), "body_color": RGBColor(64, 64, 64),
                    "images": []
                }
                logger.info("Using default template")
        except Exception as e:
            logger.warning(f"Template error: {e}, using default")
            prs = Presentation()
            style = {
                "title_font": None, "body_font": None,
                "title_color": RGBColor(31, 73, 125), "body_color": RGBColor(64, 64, 64),
                "images": []
            }

        # Pagination + long-bullet splitting before rendering
        slides_data = self._split_long_bullets(slides_data)
        slides_data = self._paginate_content(slides_data)

        # keep style context for helpers
        self._style_ctx = style
        for i, slide_data in enumerate(slides_data):
            try:
                self._create_enhanced_slide(prs, slide_data, i)
                # opportunistic image reuse from template assets (position safely)
                if self._style_ctx["images"]:
                    img_path = self._style_ctx["images"][i % len(self._style_ctx["images"])]
                    try:
                        self._place_logo_safe(prs, prs.slides[-1], img_path, slide_data.get("slide_type") == "title_slide")
                    except Exception:
                        pass
            except Exception as e:
                logger.error(f"Error creating slide {i}: {e}")
                self._create_fallback_slide(prs, slide_data, i)
        self._style_ctx = None
        return prs

    # ------------------------------ Prompting ------------------------------

    def _create_enhanced_prompt(self, text: str, guidance: str, target_slides: int) -> str:
        """JSON-only prompt with tight schema and clear constraints."""
        style_guidance = guidance or "professional, clear, and engaging presentation"
        return f"""You are a presentation designer. Convert the CONTENT into a JSON object ONLY.
Rules:
- Return ONLY valid JSON (no markdown, code fences, or prose).
- JSON must follow the exact schema shown below.
- Style: {style_guidance}
- 1 title slide, 1+ content slides, and a conclusion slide.
- Each slide has 3–5 bullet points, each < 15 words, action oriented.
- Target slides: {target_slides} (±2 permitted).

JSON SCHEMA (example keys, replace placeholders with real content):
{{
  "presentation_title": "Clear, Compelling Title",
  "slides": [
    {{
      "title": "Welcome to [Topic]",
      "content": ["Opening statement", "Key agenda point", "What audience will learn"],
      "slide_type": "title_slide",
      "speaker_notes": "Short note"
    }},
    {{
      "title": "Main Point Title",
      "content": ["Key insight", "Supporting detail", "Example", "Action item"],
      "slide_type": "content_slide",
      "speaker_notes": "Short note"
    }},
    {{
      "title": "Conclusion & Next Steps",
      "content": ["Summary of key points", "Actionable next steps", "Q&A"],
      "slide_type": "conclusion_slide",
      "speaker_notes": "Short note"
    }}
  ]
}}

CONTENT:
{text[:8000]}
"""

    # ------------------------------ LLM Calls ------------------------------

    def _call_llm_with_retry(self, provider: str, api_key: str, prompt: str, max_retries: int = 2) -> str:
        last_err = None
        for attempt in range(max_retries + 1):
            try:
                if provider == "openai":
                    return self._call_openai_enhanced(prompt, api_key)
                if provider == "anthropic":
                    return self._call_anthropic_enhanced(prompt, api_key)
                if provider == "google":
                    return self._call_google_enhanced(prompt, api_key)
                raise ValueError(f"Unsupported provider: {provider}")
            except Exception as e:
                last_err = e
                if attempt < max_retries:
                    logger.warning(f"Attempt {attempt + 1} failed: {e}, retrying...")
                else:
                    break
        raise last_err or RuntimeError("LLM call failed")

    def _call_openai_enhanced(self, prompt: str, api_key: str) -> str:
        """OpenAI call tuned for JSON-only output. Uses your 'gpt-5' routing."""
        try:
            openai.api_key = api_key
            resp = openai.ChatCompletion.create(
                model="gpt-5",
                messages=[
                    {"role": "system", "content": "You are a presentation expert. Output valid JSON only."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.2,
                max_tokens=1800,
            )
            content = (resp.choices[0].message.content or "").strip()
            if not content:
                raise ValueError("Empty response from OpenAI")
            return content
        except Exception as e:
            msg = str(e).lower()
            if "api key" in msg or "auth" in msg:
                raise ValueError("Invalid OpenAI API key")
            if "rate" in msg or "quota" in msg or "limit" in msg:
                raise ValueError("OpenAI rate limit exceeded")
            raise ValueError(f"OpenAI error: {str(e)}")

    def _call_anthropic_enhanced(self, prompt: str, api_key: str) -> str:
        """Anthropic (Claude Opus 4.1) with safe text-part concatenation."""
        try:
            client = anthropic.Anthropic(api_key=api_key)
            resp = client.messages.create(
                model="claude-opus-4-1",
                max_tokens=1800,
                temperature=0.2,
                system="Return ONLY valid JSON (no markdown).",
                messages=[{"role": "user", "content": prompt}],
            )
            parts: List[str] = []
            for block in (resp.content or []):
                if getattr(block, "type", "") == "text" and getattr(block, "text", ""):
                    parts.append(block.text)
            text = ("\n".join(parts)).strip() if parts else ""
            if not text and resp.content and getattr(resp.content[0], "text", None):
                text = resp.content[0].text.strip()
            if not text:
                raise ValueError("Empty response from Claude")
            return text
        except anthropic.AuthenticationError:
            raise ValueError("Invalid Anthropic API key")
        except anthropic.RateLimitError:
            raise ValueError("Anthropic rate limit exceeded")
        except Exception as e:
            raise ValueError(f"Anthropic error: {str(e)}")

    def _call_google_enhanced(self, prompt: str, api_key: str) -> str:
        """
        Gemini 2.5 Pro with JSON mode.
        - Version-agnostic: If genai.types.Schema exists, use it; else fallback to MIME-only JSON mode.
        - Always aggregate from candidates[].content.parts[] (never rely on response.text quick accessor).
        - max_output_tokens set to 4000.
        """
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-2.5-pro")

            # Try to build a schema if this SDK version supports it
            response_schema = None
            try:
                if hasattr(genai, "types") and hasattr(genai.types, "Schema") and hasattr(genai.types, "Type"):
                    Schema, Type = genai.types.Schema, genai.types.Type
                    response_schema = Schema(
                        type=Type.OBJECT,
                        properties={
                            "presentation_title": Schema(type=Type.STRING),
                            "slides": Schema(
                                type=Type.ARRAY,
                                items=Schema(
                                    type=Type.OBJECT,
                                    properties={
                                        "title": Schema(type=Type.STRING),
                                        "content": Schema(type=Type.ARRAY, items=Schema(type=Type.STRING)),
                                        "slide_type": Schema(type=Type.STRING),
                                        "speaker_notes": Schema(type=Type.STRING),
                                    },
                                    required=["title", "content", "slide_type"],
                                ),
                            ),
                        },
                        required=["slides"],
                    )
            except Exception:
                response_schema = None  # fall back if building schema fails

            gen_cfg = {
                "temperature": 0.15,
                "max_output_tokens": 4000,   # as requested
                "candidate_count": 1,
                "response_mime_type": "application/json",
            }
            if response_schema is not None:
                gen_cfg["response_schema"] = response_schema

            resp = model.generate_content(prompt, generation_config=gen_cfg)

            # Aggregate JSON from parts
            texts = []
            for cand in getattr(resp, "candidates", []) or []:
                content = getattr(cand, "content", None)
                for part in (getattr(content, "parts", None) or []):
                    t = getattr(part, "text", None)
                    if t:
                        texts.append(t)

            # Do NOT call resp.text quick accessor (can raise when parts empty)
            out = ("\n".join(texts)).strip()
            if not out:
                raise ValueError("Empty response from Google")
            return out

        except Exception as e:
            msg = str(e)
            if "api key" in msg.lower():
                raise ValueError("Invalid Google API key")
            if "500" in msg:
                raise ValueError("Google error: transient server issue (500). Please retry.")
            if "has no attribute 'Schema'" in msg:
                # We already fell back to MIME mode; surface a cleaner hint
                raise ValueError("Google error: SDK lacks JSON schema helpers; JSON mode fallback used.")
            raise ValueError(f"Google error: {msg}")

    # ------------------------------ Parsing & Validation ------------------------------

    def _robust_json_extraction(self, response: str) -> List[Dict]:
        """
        Parse JSON from LLM output:
        - strips code fences if present,
        - extracts outermost JSON containing "slides",
        - validates & normalizes; otherwise falls back to markdown parsing.
        """
        cleaned = (response or "").strip()

        # Strip triple-backtick fences
        if cleaned.startswith("```"):
            cleaned = cleaned.strip("`").strip()
            if cleaned.lower().startswith("json"):
                cleaned = cleaned[4:].strip()

        # Extract outer JSON containing "slides"
        match = re.search(r"\{.*\"slides\".*\}[\s\S]*$", cleaned)
        candidate = match.group(0) if match else cleaned

        try:
            data = json.loads(candidate)
        except Exception:
            # Try extracting just the slides array
            arr = re.search(r"\"slides\"\s*:\s*(\[[\s\S]*?\])", cleaned)
            if arr:
                try:
                    slides = json.loads(arr.group(1))
                    return self._validate_and_enhance_slides(slides)
                except Exception:
                    pass
            # Markdown/prose fallback
            return self._md_to_slides(cleaned)

        if not self._validate_json_structure(data):
            return self._md_to_slides(cleaned)

        return self._validate_and_enhance_slides(data)

    def _validate_json_structure(self, data: Dict) -> bool:
        return isinstance(data, dict) and isinstance(data.get("slides"), list) and len(data["slides"]) > 0

    def _validate_and_enhance_slides(self, data_or_slides) -> List[Dict]:
        """Normalize to guaranteed non-empty, 3–5 concise bullets per slide."""
        slides = data_or_slides if isinstance(data_or_slides, list) else data_or_slides.get("slides", [])
        out: List[Dict] = []

        for i, s in enumerate(slides):
            title = self._clean_text(str(s.get("title", f"Slide {i+1}"))) or f"Slide {i+1}"
            notes = self._clean_text(str(s.get("speaker_notes", ""))) or f"Notes for {title}"
            stype = s.get("slide_type", "content_slide")

            content = s.get("content", [])
            if isinstance(content, str):
                content = [content]
            content = [self._clean_text(str(x))[:150] for x in content if str(x).strip()]

            # Guarantee 3–5 bullets where possible
            if len(content) < 3 and i > 0:
                content += [f"Key insight about {title.lower()}"]
            content = [c for c in content if c][:5]
            if not content and i > 0:
                content = [f"Overview of {title}", "Main takeaway", "Next step"]

            if stype not in {"title_slide", "content_slide", "conclusion_slide", "section_header"}:
                stype = "content_slide"
            if i == 0:
                stype = "title_slide"

            out.append(
                {"title": title, "content": content, "slide_type": stype, "speaker_notes": notes}
            )

        # enforce max slides
        return out[: self.max_slides]

    def _md_to_slides(self, text: str) -> List[Dict]:
        """
        Markdown/prose → slides:
        - # H1 → title slide
        - ##/### → section slides
        - Bullets/numbered → bullets
        - Guarantees 3–5 bullets where viable
        """
        lines = (text or "").splitlines()
        slides: List[Dict] = []
        cur: Optional[Dict] = None

        def push_slide():
            nonlocal cur
            if not cur:
                return
            bullets = [re.sub(r"\s+", " ", b).strip() for b in cur.get("content", []) if b.strip()]
            bullets = [b[:150] for b in bullets]
            if not bullets and cur.get("title"):
                bullets = [f"Overview of {cur['title']}"]
            # make 3–5 bullets when possible
            if len(bullets) < 3:
                bullets = bullets + [" "] * (3 - len(bullets))
            cur["content"] = bullets[:5]
            slides.append(cur)
            cur = None

        for raw in lines:
            line = raw.strip()
            if not line:
                continue
            if line.startswith("# "):
                if cur:
                    push_slide()
                t = line[2:].strip() or "Presentation"
                cur = {"title": t, "content": [], "slide_type": "title_slide", "speaker_notes": f"Intro: {t}"}
            elif line.startswith("## ") or line.startswith("### "):
                if cur:
                    push_slide()
                t = line.split(" ", 1)[1].strip() or "Section"
                cur = {"title": t, "content": [], "slide_type": "content_slide", "speaker_notes": f"Discuss: {t}"}
            elif re.match(r"^(\*|-|•|\d+\.)\s+", line):
                if not cur:
                    cur = {"title": "Key Points", "content": [], "slide_type": "content_slide", "speaker_notes": "Key points"}
                b = re.sub(r"^(\*|-|•|\d+\.)\s+", "", line).strip()
                if b:
                    cur["content"].append(b)
            else:
                if not cur:
                    cur = {"title": "Overview", "content": [], "slide_type": "content_slide", "speaker_notes": "Overview"}
                if len(line) > 40:
                    cur["content"].append(line)

        if cur:
            push_slide()

        if not slides:
            return self._create_default_slides(text)

        # ensure ending conclusion if not present
        if not any(s.get("slide_type") == "conclusion_slide" for s in slides[-2:]):
            slides.append(
                {
                    "title": "Conclusion & Next Steps",
                    "content": ["Summary of key points", "Actionable next steps", "Q&A"],
                    "slide_type": "conclusion_slide",
                    "speaker_notes": "Wrap up",
                }
            )
        return slides

    # ------------------------------ Utilities & Fallbacks ------------------------------

    def _score_slides(self, slides: List[Dict]) -> float:
        """Simple heuristic: average bullets per non-title slide; penalize empties."""
        if not slides:
            return 0.0
        counts, n = 0, 0
        for i, s in enumerate(slides):
            if s.get("slide_type") == "title_slide" or i == 0:
                continue
            b = [x for x in (s.get("content") or []) if x and x.strip()]
            counts += len(b)
            n += 1
        return (counts / max(n, 1)) if n else 0.0

    def _refine_with_provider(self, provider: str, api_key: str, prior_json: str, target_slides: int) -> str:
        """Ask the same provider to densify bullets and increase slide count if light."""
        refine_prompt = f"""
You previously returned this JSON presentation:

{prior_json}

Improve it:
- Keep JSON only (no markdown).
- Ensure 1 title slide, >= {max(4, target_slides)} total slides, and a clear conclusion.
- For each content slide, include 3–5 action-oriented bullets (<15 words each).
- Prefer concrete, specific bullets over generic statements.
Return ONLY the JSON object.
""".strip()
        if provider == "openai":
            return self._call_openai_enhanced(refine_prompt, api_key)
        if provider == "anthropic":
            return self._call_anthropic_enhanced(refine_prompt, api_key)
        if provider == "google":
            return self._call_google_enhanced(refine_prompt, api_key)
        raise ValueError(f"Unsupported provider for refine: {provider}")

    def _clean_text(self, text: str) -> str:
        if not text:
            return ""
        # Normalize Windows CRLF artifacts that sometimes leak from copied text/HTML
        text = text.replace("_x000D_", " ").replace("\r\n", " ").replace("\r", " ")
        cleaned = re.sub(r"\s+", " ", text.strip())
        # Keep common punctuation; strip control chars/odd glyphs
        cleaned = re.sub(r"[^\w\s\-.,!?()&%$#@:/]", "", cleaned)
        return cleaned

    def _create_default_slides(self, original_text: str) -> List[Dict]:
        words = (original_text or "").split()
        chunk_size = max(50, len(words) // 4 or 50)
        chunks = [" ".join(words[i: i + chunk_size]) for i in range(0, len(words), chunk_size)]
        slides = [
            {
                "title": "Presentation Overview",
                "content": ["Key insights from your content", "Structured information", "Professional presentation"],
                "slide_type": "title_slide",
                "speaker_notes": "Introduction to the presentation content",
            }
        ]
        for i, chunk in enumerate(chunks[:6]):
            sentences = re.split(r"[.!?]+", chunk)
            bullets = [s.strip() for s in sentences if len(s.strip()) > 20][:4]
            if bullets:
                slides.append(
                    {
                        "title": f"Key Point {i+1}",
                        "content": bullets,
                        "slide_type": "content_slide",
                        "speaker_notes": f"Detailed discussion of key point {i+1}",
                    }
                )
        slides.append(
            {
                "title": "Conclusion & Next Steps",
                "content": ["Summary of key points", "Actionable next steps", "Q&A"],
                "slide_type": "conclusion_slide",
                "speaker_notes": "Wrap up",
            }
        )
        return slides

    def _emergency_fallback_parsing(self, text: str) -> List[Dict]:
        """Simple paragraph-based split as last resort."""
        paragraphs = [p.strip() for p in (text or "").split("\n\n") if p.strip()]
        slides = [
            {
                "title": "Generated Presentation",
                "content": ["Content extracted from your text", "Organized professionally", "Ready for presentation"],
                "slide_type": "title_slide",
                "speaker_notes": "Opening slide introducing the presentation",
            }
        ]
        for i, para in enumerate(paragraphs[:8]):
            sentences = [s.strip() for s in re.split(r"[.!?]+", para) if len(s.strip()) > 10]
            if sentences:
                title = sentences[0][:50] + ("..." if len(sentences[0]) > 50 else "")
                content = sentences[1:5] if len(sentences) > 1 else [sentences[0]]
                slides.append(
                    {
                        "title": title,
                        "content": content,
                        "slide_type": "content_slide",
                        "speaker_notes": f"Discussion of paragraph {i+1} content",
                    }
                )
        slides.append(
            {
                "title": "Conclusion & Next Steps",
                "content": ["Summary of key points", "Actionable next steps", "Q&A"],
                "slide_type": "conclusion_slide",
                "speaker_notes": "Wrap up",
            }
        )
        return slides

    # ------------------------------ Layout Controls (keep text inside slide) ------------------------------

    def _split_long_bullets(self, slides: List[Dict], max_len: int = 140) -> List[Dict]:
        """Split very long bullets into two at natural breakpoints to avoid overflow."""
        out = []
        for s in slides:
            bullets = []
            for b in (s.get("content") or []):
                t = (b or "").strip()
                if len(t) <= max_len:
                    bullets.append(t)
                else:
                    # split at punctuation/comma/semicolon/ dash
                    split_pt = max(
                        t.rfind(". ", 0, max_len),
                        t.rfind("; ", 0, max_len),
                        t.rfind(", ", 0, max_len),
                        t.rfind(" - ", 0, max_len),
                    )
                    if split_pt < 60:  # if no good breakpoint, hard split
                        split_pt = max_len
                    bullets.append(t[:split_pt].strip())
                    rest = t[split_pt:].lstrip(".;,- ").strip()
                    if rest:
                        bullets.append(rest[:max_len].strip())
            s2 = dict(s)
            s2["content"] = bullets[:5] if bullets else s.get("content", [])
            out.append(s2)
        return out

    def _paginate_content(self, slides: List[Dict]) -> List[Dict]:
        """
        Split slides whose total character load is too high into continuation slides.
        Heuristic thresholds; avoids overflow on small placeholders in some templates.
        """
        paginated: List[Dict] = []
        per_slide_char_limit = 600  # rough capacity for 18–20pt with bullets
        for s in slides:
            bullets = s.get("content") or []
            total_len = sum(len(b or "") for b in bullets)
            if total_len <= per_slide_char_limit or s.get("slide_type") == "title_slide":
                paginated.append(s)
                continue

            # Split into chunks by cumulative length
            chunk: List[str] = []
            chunk_len = 0
            chunk_idx = 1
            for b in bullets:
                blen = len(b or "")
                if chunk and (chunk_len + blen) > per_slide_char_limit:
                    paginated.append({
                        "title": f"{s.get('title','Slide')} (cont. {chunk_idx})",
                        "content": chunk[:5],
                        "slide_type": "content_slide",
                        "speaker_notes": s.get("speaker_notes","")
                    })
                    chunk = []
                    chunk_len = 0
                    chunk_idx += 1
                chunk.append(b)
                chunk_len += blen

            if chunk:
                suffix = "" if chunk_idx == 1 else f" (cont. {chunk_idx})"
                paginated.append({
                    "title": f"{s.get('title','Slide')}{suffix}",
                    "content": chunk[:5],
                    "slide_type": "content_slide",
                    "speaker_notes": s.get("speaker_notes","")
                })
        return paginated

    def _safe_rect(self, prs: Presentation) -> Tuple[Emu, Emu, Emu, Emu]:
        """Return a safe content rectangle (left, top, width, height) with margins, using prs dims."""
        sw, sh = prs.slide_width, prs.slide_height  # EMU
        margin_x = Emu(Inches(0.8))
        top = Emu(Inches(1.6))
        bottom = Emu(Inches(0.8))
        left = margin_x
        width = sw - Emu(Inches(0.8)) - margin_x
        height = sh - top - bottom
        return left, top, width, height

    def _apply_textframe_presentation_defaults(self, tf, base_pt=20):
        """Word-wrap, auto-size to fit, margins, and adaptive font size."""
        try:
            tf.word_wrap = True
        except Exception:
            pass
        try:
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except Exception:
            pass
        # margins (EMU)
        try:
            tf.margin_left = Emu(Inches(0.1))
            tf.margin_right = Emu(Inches(0.1))
            tf.margin_top = Emu(Inches(0.05))
            tf.margin_bottom = Emu(Inches(0.05))
        except Exception:
            pass

        # Adaptive font size baseline
        size = base_pt
        joined = " ".join(p.text for p in tf.paragraphs if getattr(p, "text", ""))
        length = len(joined)
        if length > 500:
            size = 16
        elif length > 350:
            size = 18
        else:
            size = base_pt

        # Apply size + style to first paragraph; others set in add loop
        try:
            p0 = tf.paragraphs[0]
            p0.font.size = Pt(size)
            ctx = getattr(self, "_style_ctx", None) or {}
            if ctx.get("body_font"):
                p0.font.name = ctx["body_font"]
            p0.font.color.rgb = (ctx.get("body_color", RGBColor(64, 64, 64)))
        except Exception:
            pass

        return size

    def _ideal_bullet_font_size(self, bullets: List[str]) -> int:
        n = len(bullets or [])
        avg_len = sum(len(b or "") for b in bullets) / max(n, 1)
        size = 20
        if n >= 5 or avg_len > 90:
            size = 18
        if n >= 6 or avg_len > 120:
            size = 16
        return size

    def _place_logo_safe(self, prs: Presentation, slide, img_path: str, is_title: bool):
        """Place a small logo inside bounds at top-right (title) / bottom-right (others)."""
        sw, sh = prs.slide_width, prs.slide_height
        desired_h_in = 0.9 if is_title else 1.0
        pic = slide.shapes.add_picture(img_path, Emu(0), Emu(0), height=Emu(Inches(desired_h_in)))
        # place with margin
        margin = Emu(Inches(0.3 if is_title else 0.4))
        pic.left = sw - margin - pic.width
        pic.top = (Emu(Inches(0.3)) if is_title else (sh - margin - pic.height))

    # ------------------------------ Template Style/Assets ------------------------------

    def _extract_style_and_assets(self, prs: Presentation) -> dict:
        """
        Heuristically extract title/body font names, primary colors, and picture assets
        from the uploaded PPTX/POTX (without removing its slides).
        """
        style = {
            "title_font": None,
            "body_font": None,
            "title_color": RGBColor(31, 73, 125),  # default pro blue
            "body_color": RGBColor(64, 64, 64),
            "images": [],  # list of temp file paths
        }

        # Try reading first slide placeholders for font/color
        try:
            if len(prs.slides):
                s0 = prs.slides[0]
                if getattr(s0.shapes, "title", None) and s0.shapes.title.has_text_frame:
                    p = s0.shapes.title.text_frame.paragraphs[0].font
                    if getattr(p, "name", None):
                        style["title_font"] = p.name
                    if getattr(p.color, "rgb", None):
                        style["title_color"] = p.color.rgb
                # find any text placeholder for body font
                for shp in s0.placeholders:
                    if hasattr(shp, "text_frame") and shp is not getattr(s0.shapes, "title", None):
                        pf = shp.text_frame.paragraphs[0].font
                        if getattr(pf, "name", None):
                            style["body_font"] = style["body_font"] or pf.name
                        if getattr(pf.color, "rgb", None):
                            style["body_color"] = pf.color.rgb
                        break
        except Exception:
            pass

        # Harvest picture assets (from all existing slides)
        try:
            for s in prs.slides:
                for shp in s.shapes:
                    if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        try:
                            blob = shp.image.blob
                            ext = shp.image.ext or "png"
                            tf = tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}")
                            tf.write(blob)
                            tf.flush(); tf.close()
                            style["images"].append(tf.name)
                        except Exception:
                            continue
        except Exception:
            pass

        return style

    # ------------------------------ PPT Creation Helpers ------------------------------

    def _create_enhanced_slide(self, prs: Presentation, slide_data: Dict, index: int):
        # Choose layout
        if slide_data.get("slide_type") == "title_slide" and index == 0:
            layout_idx = 0 if len(prs.slide_layouts) > 0 else 1
        elif slide_data.get("content"):
            layout_idx = 1 if len(prs.slide_layouts) > 1 else 0
        else:
            layout_idx = 5 if len(prs.slide_layouts) > 5 else 1

        try:
            slide_layout = prs.slide_layouts[layout_idx]
        except Exception:
            slide_layout = prs.slide_layouts[0]

        slide = prs.slides.add_slide(slide_layout)

        # Title
        try:
            title_text = slide_data["title"]
            if getattr(slide.shapes, "title", None):
                slide.shapes.title.text = title_text
                self._style_title(slide.shapes.title, index == 0)
            else:
                # Safe title box near top, within margins
                sw = prs.slide_width
                left = Emu(Inches(0.8))
                top = Emu(Inches(0.5))
                width = sw - left - Emu(Inches(0.8))
                height = Emu(Inches(1.2))
                tb = slide.shapes.add_textbox(left, top, width, height)
                tf = tb.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                p.text = title_text
                p.font.bold = True
                p.font.size = Pt(36 if index == 0 else 30)
                ctx = getattr(self, "_style_ctx", None) or {}
                if ctx.get("title_font"):
                    p.font.name = ctx["title_font"]
                try:
                    p.font.color.rgb = ctx.get("title_color", RGBColor(31, 73, 125))
                except Exception:
                    pass
        except Exception as e:
            logger.debug(f"Title add failed: {e}")

        # Content
        try:
            bullets = slide_data.get("content") or []
            if bullets:
                # Try a content placeholder
                ph = None
                for ph_i in getattr(slide, "placeholders", []):
                    if hasattr(ph_i, "text_frame") and ph_i != getattr(slide.shapes, "title", None):
                        ph = ph_i
                        break

                use_ph = False
                if ph and hasattr(ph, "height") and hasattr(ph, "width"):
                    # If placeholder visibly large, use it; else create our own box
                    try:
                        if ph.height >= Emu(Inches(2.0)) and ph.width >= Emu(Inches(5.0)):
                            use_ph = True
                    except Exception:
                        use_ph = True

                if use_ph:
                    tf = ph.text_frame
                    tf.clear()
                    base_size = self._ideal_bullet_font_size(bullets)
                    self._apply_textframe_presentation_defaults(tf, base_pt=base_size)
                    for i, bullet_text in enumerate(bullets):
                        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                        p.text = bullet_text
                        p.level = 0
                        try:
                            p.font.size = Pt(base_size)
                            p.space_after = Pt(6)
                            ctx = getattr(self, "_style_ctx", None) or {}
                            if ctx.get("body_font"):
                                p.font.name = ctx["body_font"]
                            p.font.color.rgb = ctx.get("body_color", RGBColor(64, 64, 64))
                        except Exception:
                            pass
                else:
                    # Safe textbox within margins (use presentation dims)
                    left, top, width, height = self._safe_rect(prs)
                    tb = slide.shapes.add_textbox(left, top, width, height)
                    tf = tb.text_frame
                    tf.clear()
                    base_size = self._ideal_bullet_font_size(bullets)
                    self._apply_textframe_presentation_defaults(tf, base_pt=base_size)
                    for i, bullet_text in enumerate(bullets):
                        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                        p.text = f"{bullet_text}"
                        p.level = 0
                        try:
                            p.font.size = Pt(base_size)
                            p.space_after = Pt(6)
                            ctx = getattr(self, "_style_ctx", None) or {}
                            if ctx.get("body_font"):
                                p.font.name = ctx["body_font"]
                            p.font.color.rgb = ctx.get("body_color", RGBColor(64, 64, 64))
                        except Exception:
                            pass
        except Exception as e:
            logger.warning(f"Could not add content: {e}")
            self._add_basic_content(prs, slide, slide_data.get("content", []))

        # Speaker notes
        self._add_speaker_notes(slide, slide_data.get("speaker_notes", ""))

    def _style_title(self, title_shape, is_main_title: bool):
        try:
            if title_shape.text_frame:
                paragraph = title_shape.text_frame.paragraphs[0]
                font = paragraph.font
                font.size = Pt(36 if is_main_title else 30)
                font.bold = True
                ctx = getattr(self, "_style_ctx", None) or {}
                if ctx.get("title_font"):
                    font.name = ctx["title_font"]
                try:
                    font.color.rgb = ctx.get("title_color", RGBColor(31, 73, 125))
                except Exception:
                    pass
        except Exception as e:
            logger.debug(f"Title styling failed: {e}")

    def _add_basic_content(self, prs: Presentation, slide, content_list: List[str]):
        try:
            left, top, width, height = self._safe_rect(prs)
            textbox = slide.shapes.add_textbox(left, top, width, height)
            tf = textbox.text_frame
            tf.clear()
            base_size = self._ideal_bullet_font_size(content_list)
            self._apply_textframe_presentation_defaults(tf, base_pt=base_size)
            for i, bullet_text in enumerate(content_list):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = f"{bullet_text}"
                try:
                    p.font.size = Pt(base_size)
                    ctx = getattr(self, "_style_ctx", None) or {}
                    if ctx.get("body_font"):
                        p.font.name = ctx["body_font"]
                    p.font.color.rgb = ctx.get("body_color", RGBColor(64, 64, 64))
                except Exception:
                    pass
        except Exception as e:
            logger.warning(f"Basic content addition failed: {e}")

    def _add_speaker_notes(self, slide, notes: str):
        if not notes:
            return
        try:
            notes_slide = slide.notes_slide  # create if missing
            text_frame = notes_slide.notes_text_frame
            text_frame.text = notes
        except Exception as e:
            logger.debug(f"Could not add speaker notes: {e}")

    def _create_fallback_slide(self, prs: Presentation, slide_data: Dict, index: int):
        try:
            slide_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)
            # Title (use prs dims)
            sw = prs.slide_width
            left = Emu(Inches(0.8))
            top = Emu(Inches(0.5))
            width = sw - left - Emu(Inches(0.8))
            title_box = slide.shapes.add_textbox(left, top, width, Emu(Inches(1.2)))
            tf = title_box.text_frame
            tf.text = slide_data.get("title", f"Slide {index+1}")
            tf.paragraphs[0].font.size = Pt(32)
            tf.paragraphs[0].font.bold = True
            # Content
            self._add_basic_content(prs, slide, slide_data.get("content") or [])
        except Exception as e:
            logger.error(f"Even fallback slide creation failed: {e}")


# ------------------------------------------------------------------------------
# Initialize generator
# ------------------------------------------------------------------------------
ppt_generator = PPTGenerator()


# ------------------------------------------------------------------------------
# Routes
# ------------------------------------------------------------------------------
@app.route("/")
def index():
    try:
        # Keep index.html unchanged (must be served via templates/index.html)
        return render_template("index.html")
    except Exception:
        return "<h1>Auto PPT Generator</h1><p>Server running but template missing.</p>", 200


@app.route("/generate", methods=["POST"])
def generate_presentation():
    try:
        if "input_text" not in request.form:
            return jsonify({"error": "No input text provided"}), 400

        input_text = (request.form["input_text"] or "").strip()
        guidance = (request.form.get("guidance", "") or "").strip()
        provider = (request.form.get("provider", "openai") or "").strip().lower()
        api_key = (request.form.get("api_key", "") or "").strip()

        if len(input_text) < 20:
            return jsonify({"error": "Please provide at least 20 characters of content"}), 400
        if not api_key:
            return jsonify({"error": "API key is required"}), 400
        if provider not in ppt_generator.supported_providers:
            return jsonify({"error": f"Unsupported provider: {provider}"}), 400

        # Handle optional template file
        template_path = None
        if "template_file" in request.files:
            template_file = request.files["template_file"]
            if template_file and template_file.filename and template_file.filename.lower().endswith((".pptx", ".potx")):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                    template_file.save(tmp.name)
                    template_path = tmp.name

        try:
            slides_data = ppt_generator.parse_text_to_slides(
                text=input_text, provider=provider, api_key=api_key, guidance=guidance
            )
            if not slides_data:
                return jsonify({"error": "Could not generate slides from content"}), 400

            presentation = ppt_generator.create_presentation(slides_data, template_path)

            # In-memory output to avoid 200/0 length edge cases
            bio = BytesIO()
            presentation.save(bio)
            bio.seek(0)
            return send_file(
                bio,
                as_attachment=True,
                download_name=f'presentation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pptx',
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        except ValueError as e:
            return jsonify({"error": str(e)}), 400
        except Exception as e:
            logger.error(f"Generation error: {e}")
            return jsonify({"error": "Presentation generation failed. Please try again."}), 500
        finally:
            if template_path and os.path.exists(template_path):
                try:
                    os.unlink(template_path)
                except Exception:
                    pass
    except Exception as e:
        logger.error(f"Request error: {e}")
        return jsonify({"error": "Server error"}), 500


@app.route("/health")
def health():
    return jsonify({"status": "healthy", "timestamp": datetime.utcnow().isoformat()})


# ------------------------------------------------------------------------------
# Entrypoint
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
