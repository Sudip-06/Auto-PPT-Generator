#!/usr/bin/env python3
"""
Auto PPT Generator - Robust LLM parsing, enhanced prompting, improved content processing,
and safer slide creation with template reuse. (Keeps index.html unchanged)
"""

import os
import re
import json
import tempfile
import logging
from datetime import datetime
from typing import Dict, List, Any, Optional

from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS

# --- LLM SDKs (ensure they're installed in requirements) ---
import openai
import anthropic
import google.generativeai as genai

# --- PowerPoint ---
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

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

    # ------------------------------ Public API ------------------------------

    def parse_text_to_slides(
        self, text: str, provider: str, api_key: str, guidance: str = ""
    ) -> List[Dict]:
        """Parse raw text into normalized slide dicts using an LLM with robust fallbacks."""
        # Estimate slide count from length
        word_count = len((text or "").split())
        estimated_slides = max(self.min_slides, min(self.max_slides, word_count // 120 + 1))

        prompt = self._create_enhanced_prompt(text, guidance, estimated_slides)

        try:
            logger.info(f"Parsing with {provider}, estimated slides: {estimated_slides}")
            response = self._call_llm_with_retry(provider, api_key, prompt)
            slides = self._robust_json_extraction(response)
            slides = self._validate_and_enhance_slides(slides)
            logger.info(f"Generated {len(slides)} slides successfully")
            return slides
        except Exception as e:
            logger.error(f"Error parsing text: {e}")
            logger.warning("Using emergency fallback parsing")
            return self._emergency_fallback_parsing(text)

    def create_presentation(self, slides_data: List[Dict], template_file: Optional[str] = None) -> Presentation:
        """Create a PPTX from normalized slide data, safely reusing optional template."""
        try:
            if template_file:
                prs = Presentation(template_file)
                # Reuse theme/layouts; drop any existing slides to prevent duplicates
                if len(prs.slides) > 0:
                    self._safely_clear_slides(prs)
                logger.info("Using uploaded template")
            else:
                prs = Presentation()
                logger.info("Using default template")
        except Exception as e:
            logger.warning(f"Template error: {e}, using default")
            prs = Presentation()

        for i, slide_data in enumerate(slides_data):
            try:
                self._create_enhanced_slide(prs, slide_data, i)
            except Exception as e:
                logger.error(f"Error creating slide {i}: {e}")
                self._create_fallback_slide(prs, slide_data, i)

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
- Each slide has 3-5 bullet points, each < 15 words.
- Language must be concise and action oriented.
- Target slides: {target_slides} (±2 permitted).

JSON SCHEMA (example keys, real content must replace placeholders):
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

CONTENT (truncate on your own as needed to obey limits):
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
            if not text:
                # Fallback: sometimes content is present but not typed 'text'
                text = (resp.content[0].text.strip() if resp.content and getattr(resp.content[0], "text", None) else "")
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
        """Gemini 2.5 Pro with full parts aggregation (no response.text quick accessor)."""
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-2.5-pro")
            resp = model.generate_content(
                prompt,
                generation_config=genai.types.GenerationConfig(
                    temperature=0.2, max_output_tokens=1800, candidate_count=1
                ),
            )

            # Aggregate from candidates[].content.parts[]
            texts: List[str] = []
            for cand in getattr(resp, "candidates", []) or []:
                content = getattr(cand, "content", None)
                for part in getattr(content, "parts", []) or []:
                    t = getattr(part, "text", None)
                    if t:
                        texts.append(t)

            # Fallback to resp.text if SDK provided it as simple text
            if not texts and getattr(resp, "text", None):
                texts.append(resp.text)

            text = ("\n".join(texts)).strip()
            if not text:
                raise ValueError("Empty response from Google")
            return text
        except Exception as e:
            msg = str(e)
            if "api key" in msg.lower():
                raise ValueError("Invalid Google API key")
            if "500" in msg:
                raise ValueError("Google error: transient server issue (500). Please retry.")
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

    def _clean_text(self, text: str) -> str:
        if not text:
            return ""
        cleaned = re.sub(r"\s+", " ", text.strip())
        cleaned = re.sub(r"[^\w\s\-.,!?()&%$#@:/]", "", cleaned)
        return cleaned

    def _create_default_slides(self, original_text: str) -> List[Dict]:
        words = (original_text or "").split()
        chunk_size = max(50, len(words) // 4 or 50)
        chunks = [" ".join(words[i : i + chunk_size]) for i in range(0, len(words), chunk_size)]
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

    # ------------------------------ PPT Creation Helpers ------------------------------

    def _safely_clear_slides(self, prs: Presentation):
        """
        Properly remove existing slides from a template:
        - drop relationships and sldId entries
        - avoids duplicate slideX.xml warnings when saving
        """
        try:
            for idx in range(len(prs.slides) - 1, -1, -1):
                sldId = prs.slides._sldIdLst[idx]
                rId = sldId.rId
                prs.part.drop_rel(rId)
                prs.slides._sldIdLst.remove(sldId)
        except Exception as e:
            logger.warning(f"Could not clear template slides safely: {e}")

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
            if getattr(slide.shapes, "title", None):
                slide.shapes.title.text = slide_data["title"]
                self._style_title(slide.shapes.title, index == 0)
            else:
                tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8.4), Inches(1.0))
                tf = tb.text_frame
                tf.text = slide_data["title"]
                tf.paragraphs[0].font.size = Pt(32 if index == 0 else 28)
                tf.paragraphs[0].font.bold = True
        except Exception as e:
            logger.debug(f"Title add failed: {e}")

        # Content
        try:
            content = slide_data.get("content") or []
            if content:
                ph = None
                for ph_i in getattr(slide, "placeholders", []):
                    if hasattr(ph_i, "text_frame") and ph_i != getattr(slide.shapes, "title", None):
                        ph = ph_i
                        break
                if ph and hasattr(ph, "text_frame"):
                    tf = ph.text_frame
                    tf.clear()
                    for i, bullet_text in enumerate(content):
                        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                        p.text = bullet_text
                        p.level = 0
                        try:
                            p.font.size = Pt(20)
                            p.space_after = Pt(8)
                        except Exception:
                            pass
                else:
                    self._add_basic_content(slide, content)
        except Exception as e:
            logger.warning(f"Could not add content: {e}")
            self._add_basic_content(slide, slide_data.get("content", []))

        # Speaker notes
        self._add_speaker_notes(slide, slide_data.get("speaker_notes", ""))

    def _style_title(self, title_shape, is_main_title: bool):
        try:
            if title_shape.text_frame:
                paragraph = title_shape.text_frame.paragraphs[0]
                font = paragraph.font
                font.size = Pt(36 if is_main_title else 30)
                font.bold = True
                if is_main_title:
                    try:
                        font.color.rgb = RGBColor(31, 73, 125)
                    except Exception:
                        pass
        except Exception as e:
            logger.debug(f"Title styling failed: {e}")

    def _add_basic_content(self, slide, content_list: List[str]):
        try:
            left, top, width, height = Inches(1), Inches(2), Inches(8), Inches(4)
            textbox = slide.shapes.add_textbox(left, top, width, height)
            tf = textbox.text_frame
            for i, bullet_text in enumerate(content_list):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = f"• {bullet_text}"
                try:
                    p.font.size = Pt(18)
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
            # Title
            title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
            tf = title_box.text_frame
            tf.text = slide_data.get("title", f"Slide {index+1}")
            tf.paragraphs[0].font.size = Pt(32)
            tf.paragraphs[0].font.bold = True
            # Content
            content = slide_data.get("content") or []
            if content:
                content_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
                ctf = content_box.text_frame
                for i, item in enumerate(content[:5]):
                    p = ctf.paragraphs[0] if i == 0 else ctf.add_paragraph()
                    p.text = f"• {item}"
                    try:
                        p.font.size = Pt(18)
                    except Exception:
                        pass
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

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as out:
                presentation.save(out.name)
                return send_file(
                    out.name,
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
