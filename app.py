#!/usr/bin/env python3
"""
Auto PPT Generator - Improved Version with Better Content Processing
"""

import os
import json
import tempfile
import logging
import re
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple

from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
import openai
import anthropic
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import markdown

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app, origins=["*"])
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

class PPTGenerator:
    """Enhanced PowerPoint generator with improved content processing"""
    
    def __init__(self):
        self.supported_providers = ['openai', 'anthropic', 'google']
        self.max_slides = 12
        self.min_slides = 3
    
    def parse_text_to_slides(self, text: str, provider: str, api_key: str, guidance: str = "") -> List[Dict]:
        """Parse text into structured slide content with improved prompting"""
        
        # Better slide estimation
        word_count = len(text.split())
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
            logger.error(f"Error parsing text: {str(e)}")
            # Fallback to basic parsing if LLM fails
            return self._emergency_fallback_parsing(text)
    
    def _create_enhanced_prompt(self, text: str, guidance: str, target_slides: int) -> str:
        """Create a more robust prompt with better structure"""
        
        style_guidance = guidance or "professional, clear, and engaging presentation"
        
        prompt = f"""You are a professional presentation designer. Transform this content into a structured presentation.

CONTENT TO CONVERT:
{text[:3000]}...  # Truncate very long content

PRESENTATION REQUIREMENTS:
- Style: {style_guidance}
- Target slides: {target_slides} (can be ±2 slides based on content)
- Include one title slide + content slides + conclusion
- Each content slide should have 3-5 clear bullet points
- Use action-oriented, concise language

OUTPUT FORMAT - Return ONLY valid JSON:
{{
  "presentation_title": "Clear, Compelling Title",
  "slides": [
    {{
      "title": "Welcome to [Topic]",
      "content": ["Opening statement", "Key agenda point", "What audience will learn"],
      "slide_type": "title_slide",
      "speaker_notes": "Welcome everyone and set expectations"
    }},
    {{
      "title": "Main Point Title", 
      "content": ["Key insight #1", "Supporting detail", "Practical example", "Action item"],
      "slide_type": "content_slide",
      "speaker_notes": "Explain the main concept and give examples"
    }},
    {{
      "title": "Conclusion & Next Steps",
      "content": ["Summary of key points", "Actionable takeaways", "Questions welcome"],
      "slide_type": "conclusion_slide", 
      "speaker_notes": "Wrap up and engage audience"
    }}
  ]
}}

CRITICAL RULES:
1. Return ONLY the JSON object - no explanations or markdown
2. Each slide must have a clear title and 2-5 bullet points
3. Keep bullet points under 15 words each
4. Include practical, actionable content
5. Ensure logical flow between slides

Generate the JSON now:"""
        
        return prompt
    
    def _call_llm_with_retry(self, provider: str, api_key: str, prompt: str, max_retries: int = 2) -> str:
        """Call LLM with retry logic and better error handling"""
        
        for attempt in range(max_retries + 1):
            try:
                if provider == 'openai':
                    return self._call_openai_enhanced(prompt, api_key)
                elif provider == 'anthropic':
                    return self._call_anthropic_enhanced(prompt, api_key)
                elif provider == 'google':
                    return self._call_google_enhanced(prompt, api_key)
                else:
                    raise ValueError(f"Unsupported provider: {provider}")
                    
            except Exception as e:
                if attempt == max_retries:
                    raise e
                logger.warning(f"Attempt {attempt + 1} failed: {e}, retrying...")
                
    def _call_openai_enhanced(self, prompt: str, api_key: str) -> str:
        """Enhanced OpenAI call with better parameters"""
        try:
            openai.api_key = api_key
            
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo-1106",  # More recent model
                messages=[
                    {
                        "role": "system", 
                        "content": "You are a presentation expert. Always return valid JSON only. No markdown formatting or explanations."
                    },
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,  # Lower temperature for more consistent output
                max_tokens=1500,
                timeout=45
            )
            
            content = response.choices[0].message.content.strip()
            return content
            
        except Exception as e:
            if "invalid_api_key" in str(e) or "authentication" in str(e).lower():
                raise ValueError("Invalid OpenAI API key")
            elif "rate_limit" in str(e).lower():
                raise ValueError("OpenAI rate limit exceeded")
            else:
                raise ValueError(f"OpenAI error: {str(e)}")
    
    def _call_anthropic_enhanced(self, prompt: str, api_key: str) -> str:
        """Enhanced Anthropic call"""
        try:
            client = anthropic.Anthropic(api_key=api_key)
            
            response = client.messages.create(
                model="claude-3-haiku-20240307",
                max_tokens=1500,
                temperature=0.3,
                system="Return only valid JSON. No explanations or formatting.",
                messages=[{"role": "user", "content": prompt}]
            )
            
            return response.content[0].text.strip()
            
        except anthropic.AuthenticationError:
            raise ValueError("Invalid Anthropic API key")
        except anthropic.RateLimitError:
            raise ValueError("Anthropic rate limit exceeded")
        except Exception as e:
            raise ValueError(f"Anthropic error: {str(e)}")
    
    def _call_google_enhanced(self, prompt: str, api_key: str) -> str:
        """Enhanced Google call"""
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-pro')
            
            response = model.generate_content(
                prompt,
                generation_config=genai.types.GenerationConfig(
                    temperature=0.3,
                    max_output_tokens=1500,
                    candidate_count=1
                )
            )
            
            if not response.text:
                raise ValueError("Empty response from Google")
                
            return response.text.strip()
            
        except Exception as e:
            if 'api key' in str(e).lower():
                raise ValueError("Invalid Google API key")
            else:
                raise ValueError(f"Google error: {str(e)}")
    
    def _robust_json_extraction(self, response: str) -> List[Dict]:
        """More robust JSON extraction with multiple fallback methods"""
        
        # Clean the response
        cleaned = response.strip()
        
        # Remove code block markers
        if cleaned.startswith('```'):
            lines = cleaned.split('\n')
            if lines[0].startswith('```'):
                lines = lines[1:]
            if lines and lines[-1].strip() == '```':
                lines = lines[:-1]
            cleaned = '\n'.join(lines)
        
        # Try to extract JSON
        try:
            # Method 1: Direct JSON parsing
            data = json.loads(cleaned)
            if self._validate_json_structure(data):
                return data.get('slides', [])
        except json.JSONDecodeError:
            pass
        
        try:
            # Method 2: Find JSON within text
            json_match = re.search(r'\{.*"slides".*\].*\}', cleaned, re.DOTALL)
            if json_match:
                json_str = json_match.group()
                data = json.loads(json_str)
                if self._validate_json_structure(data):
                    return data.get('slides', [])
        except (json.JSONDecodeError, AttributeError):
            pass
        
        # Method 3: Try to extract just the slides array
        try:
            slides_match = re.search(r'"slides"\s*:\s*\[(.*?)\]', cleaned, re.DOTALL)
            if slides_match:
                slides_str = '[' + slides_match.group(1) + ']'
                slides = json.loads(slides_str)
                return slides
        except (json.JSONDecodeError, AttributeError):
            pass
        
        # If all else fails, use text parsing fallback
        logger.warning("JSON extraction failed, using text parsing fallback")
        return self._text_parsing_fallback(cleaned)
    
    def _validate_json_structure(self, data: Dict) -> bool:
        """Validate that JSON has expected structure"""
        return (
            isinstance(data, dict) and 
            'slides' in data and 
            isinstance(data['slides'], list) and
            len(data['slides']) > 0
        )
    
    def _text_parsing_fallback(self, text: str) -> List[Dict]:
        """Parse text when JSON extraction fails"""
        slides = []
        lines = text.split('\n')
        current_slide = None
        slide_count = 0
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Check for slide titles (various formats)
            if (line.startswith('#') or 
                (len(line.split()) <= 8 and line.isupper()) or
                line.startswith('Title:') or
                line.startswith('Slide')):
                
                # Save previous slide
                if current_slide and current_slide.get('content'):
                    slides.append(current_slide)
                
                # Create new slide
                title = re.sub(r'^#+\s*|Title:\s*|Slide\s*\d+:\s*', '', line).strip().title()
                slide_count += 1
                
                slide_type = 'title_slide' if slide_count == 1 else 'content_slide'
                if 'conclusion' in title.lower() or 'summary' in title.lower():
                    slide_type = 'conclusion_slide'
                
                current_slide = {
                    'title': title or f'Slide {slide_count}',
                    'content': [],
                    'slide_type': slide_type,
                    'speaker_notes': f'Key points about {title.lower() if title else "this topic"}'
                }
            
            # Check for bullet points
            elif current_slide and (line.startswith('-') or line.startswith('*') or 
                                  line.startswith('•') or re.match(r'^\d+\.', line)):
                content = re.sub(r'^[-*•]\s*|\d+\.\s*', '', line).strip()
                if content and len(content) > 10:  # Ensure meaningful content
                    current_slide['content'].append(content[:100])  # Limit length
        
        # Add the last slide
        if current_slide and current_slide.get('content'):
            slides.append(current_slide)
        
        # Ensure we have at least one slide
        if not slides:
            slides = self._create_default_slides(text)
        
        return slides
    
    def _create_default_slides(self, original_text: str) -> List[Dict]:
        """Create basic slides when all parsing fails"""
        words = original_text.split()
        
        # Split content into chunks
        chunk_size = max(50, len(words) // 4)
        chunks = [' '.join(words[i:i+chunk_size]) for i in range(0, len(words), chunk_size)]
        
        slides = [
            {
                'title': 'Presentation Overview',
                'content': ['Key insights from your content', 'Structured information', 'Professional presentation'],
                'slide_type': 'title_slide',
                'speaker_notes': 'Introduction to the presentation content'
            }
        ]
        
        for i, chunk in enumerate(chunks[:6]):  # Max 6 content slides
            # Extract key sentences as bullet points
            sentences = re.split(r'[.!?]+', chunk)
            bullets = [s.strip() for s in sentences if len(s.strip()) > 20][:4]
            
            if bullets:
                slides.append({
                    'title': f'Key Point {i+1}',
                    'content': bullets,
                    'slide_type': 'content_slide',
                    'speaker_notes': f'Detailed discussion of key point {i+1}'
                })
        
        return slides
    
    def _validate_and_enhance_slides(self, slides: List[Dict]) -> List[Dict]:
        """Enhanced slide validation and improvement"""
        validated_slides = []
        
        for i, slide in enumerate(slides):
            # Ensure all required fields
            enhanced_slide = {
                'title': self._clean_text(str(slide.get('title', f'Slide {i+1}'))),
                'content': [],
                'slide_type': slide.get('slide_type', 'content_slide'),
                'speaker_notes': self._clean_text(str(slide.get('speaker_notes', '')))
            }
            
            # Process and validate content
            raw_content = slide.get('content', [])
            if isinstance(raw_content, str):
                raw_content = [raw_content]
            elif not isinstance(raw_content, list):
                raw_content = []
            
            for item in raw_content:
                clean_item = self._clean_text(str(item))
                if clean_item and len(clean_item) > 5:
                    enhanced_slide['content'].append(clean_item[:150])  # Reasonable length
            
            # Ensure minimum content
            if not enhanced_slide['content'] and i > 0:
                enhanced_slide['content'] = [f"Key insights about {enhanced_slide['title'].lower()}"]
            
            # Limit content items to 5
            enhanced_slide['content'] = enhanced_slide['content'][:5]
            
            # Validate slide type
            valid_types = ['title_slide', 'content_slide', 'conclusion_slide', 'section_header']
            if enhanced_slide['slide_type'] not in valid_types:
                enhanced_slide['slide_type'] = 'content_slide'
            
            validated_slides.append(enhanced_slide)
        
        # Ensure reasonable number of slides
        return validated_slides[:self.max_slides]
    
    def _clean_text(self, text: str) -> str:
        """Clean and normalize text content"""
        if not text:
            return ""
        
        # Remove extra whitespace and clean up
        cleaned = re.sub(r'\s+', ' ', text.strip())
        
        # Remove problematic characters
        cleaned = re.sub(r'[^\w\s\-.,!?()&%$#@]', '', cleaned)
        
        return cleaned
    
    def _emergency_fallback_parsing(self, text: str) -> List[Dict]:
        """Last resort parsing when everything fails"""
        logger.warning("Using emergency fallback parsing")
        
        # Simple text splitting approach
        paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
        
        slides = [
            {
                'title': 'Generated Presentation',
                'content': ['Content extracted from your text', 'Organized professionally', 'Ready for presentation'],
                'slide_type': 'title_slide',
                'speaker_notes': 'Opening slide introducing the presentation'
            }
        ]
        
        # Create slides from paragraphs
        for i, para in enumerate(paragraphs[:8]):
            sentences = [s.strip() for s in re.split(r'[.!?]+', para) if len(s.strip()) > 10]
            
            if sentences:
                title = sentences[0][:50] + "..." if len(sentences[0]) > 50 else sentences[0]
                content = sentences[1:5] if len(sentences) > 1 else [sentences[0]]
                
                slides.append({
                    'title': title,
                    'content': content,
                    'slide_type': 'content_slide',
                    'speaker_notes': f'Discussion of paragraph {i+1} content'
                })
        
        return slides
    
    def create_presentation(self, slides_data: List[Dict], template_file=None) -> Presentation:
        """Create PowerPoint with improved slide generation"""
        
        try:
            if template_file:
                prs = Presentation(template_file)
                # Clear existing slides but keep layouts
                self._safely_clear_slides(prs)
                logger.info("Using uploaded template")
            else:
                prs = Presentation()
                logger.info("Using default template")
        except Exception as e:
            logger.warning(f"Template error: {e}, using default")
            prs = Presentation()
        
        # Generate slides with better error handling
        for i, slide_data in enumerate(slides_data):
            try:
                self._create_enhanced_slide(prs, slide_data, i)
            except Exception as e:
                logger.error(f"Error creating slide {i}: {e}")
                # Create a basic slide as fallback
                self._create_fallback_slide(prs, slide_data, i)
        
        return prs
    
    def _safely_clear_slides(self, prs: Presentation):
        """Safely remove existing slides from template"""
        try:
            slide_ids = list(range(len(prs.slides)))
            for slide_id in reversed(slide_ids):  # Remove from end to avoid index issues
                xml_slides = prs.slides._sldIdLst
                if len(xml_slides) > slide_id:
                    xml_slides.remove(xml_slides[slide_id])
        except Exception as e:
            logger.warning(f"Could not clear all template slides: {e}")
    
    def _create_enhanced_slide(self, prs: Presentation, slide_data: Dict, index: int):
        """Create slide with improved content handling"""
        
        # Choose layout more intelligently
        if slide_data['slide_type'] == 'title_slide' and index == 0:
            layout_idx = 0  # Title slide layout
        elif slide_data.get('content') and len(slide_data['content']) > 0:
            layout_idx = 1  # Title and content layout
        else:
            layout_idx = 5 if len(prs.slide_layouts) > 5 else 1  # Blank or content layout
        
        try:
            slide_layout = prs.slide_layouts[layout_idx]
        except IndexError:
            slide_layout = prs.slide_layouts[0]  # Fallback to first available
        
        slide = prs.slides.add_slide(slide_layout)
        
        # Add title with error handling
        if hasattr(slide, 'shapes') and slide.shapes.title:
            slide.shapes.title.text = slide_data['title']
            self._style_title(slide.shapes.title, index == 0)
        
        # Add content with better formatting
        if slide_data.get('content') and len(slide.placeholders) > 1:
            try:
                self._add_content_to_slide(slide, slide_data['content'])
            except Exception as e:
                logger.warning(f"Could not add formatted content: {e}")
                # Fallback to basic text
                self._add_basic_content(slide, slide_data['content'])
        
        # Add speaker notes
        self._add_speaker_notes(slide, slide_data.get('speaker_notes', ''))
    
    def _style_title(self, title_shape, is_main_title: bool):
        """Apply better title styling"""
        try:
            if title_shape.text_frame:
                paragraph = title_shape.text_frame.paragraphs[0]
                font = paragraph.font
                font.size = Pt(36 if is_main_title else 30)
                font.bold = True
                
                if is_main_title:
                    try:
                        font.color.rgb = RGBColor(31, 73, 125)  # Professional blue
                    except:
                        pass
        except Exception as e:
            logger.debug(f"Title styling failed: {e}")
    
    def _add_content_to_slide(self, slide, content_list: List[str]):
        """Add formatted content to slide"""
        
        # Find content placeholder
        content_placeholder = None
        for placeholder in slide.placeholders:
            if placeholder.placeholder_format.idx == 1:  # Content placeholder
                content_placeholder = placeholder
                break
        
        if not content_placeholder and len(slide.placeholders) > 1:
            content_placeholder = slide.placeholders[1]
        
        if content_placeholder and hasattr(content_placeholder, 'text_frame'):
            text_frame = content_placeholder.text_frame
            text_frame.clear()
            
            # Add bullet points
            for i, bullet_text in enumerate(content_list):
                if i == 0:
                    p = text_frame.paragraphs[0]
                else:
                    p = text_frame.add_paragraph()
                
                p.text = bullet_text
                p.level = 0
                
                # Style the paragraph
                try:
                    p.font.size = Pt(20)
                    p.space_after = Pt(12)
                    p.font.color.rgb = RGBColor(64, 64, 64)  # Dark gray
                except Exception:
                    pass  # Ignore styling errors
    
    def _add_basic_content(self, slide, content_list: List[str]):
        """Fallback method to add content"""
        try:
            # Try to add as text box if placeholders don't work
            left = Inches(1)
            top = Inches(2)
            width = Inches(8)
            height = Inches(4)
            
            textbox = slide.shapes.add_textbox(left, top, width, height)
            text_frame = textbox.text_frame
            
            for i, bullet_text in enumerate(content_list):
                if i == 0:
                    p = text_frame.paragraphs[0]
                else:
                    p = text_frame.add_paragraph()
                
                p.text = f"• {bullet_text}"
                p.font.size = Pt(18)
                
        except Exception as e:
            logger.warning(f"Basic content addition failed: {e}")
    
    def _add_speaker_notes(self, slide, notes: str):
        """Add speaker notes to slide"""
        if notes and hasattr(slide, 'notes_slide'):
            try:
                notes_slide = slide.notes_slide
                text_frame = notes_slide.notes_text_frame
                text_frame.text = notes
            except Exception as e:
                logger.debug(f"Could not add speaker notes: {e}")
    
    def _create_fallback_slide(self, prs: Presentation, slide_data: Dict, index: int):
        """Create a basic slide when normal creation fails"""
        try:
            slide_layout = prs.slide_layouts[6]  # Blank layout
            slide = prs.slides.add_slide(slide_layout)
            
            # Add title as text box
            title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
            title_frame = title_box.text_frame
            title_frame.text = slide_data['title']
            title_frame.paragraphs[0].font.size = Pt(32)
            title_frame.paragraphs[0].font.bold = True
            
            # Add content as text box
            if slide_data.get('content'):
                content_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
                content_frame = content_box.text_frame
                
                for i, item in enumerate(slide_data['content'][:5]):
                    if i == 0:
                        p = content_frame.paragraphs[0]
                    else:
                        p = content_frame.add_paragraph()
                    p.text = f"• {item}"
                    p.font.size = Pt(18)
                    
        except Exception as e:
            logger.error(f"Even fallback slide creation failed: {e}")

# Initialize the generator
ppt_generator = PPTGenerator()

# Keep the same Flask routes but with improved error handling
@app.route('/')
def index():
    try:
        return render_template('index.html')
    except Exception:
        return """<h1>Auto PPT Generator</h1><p>Server running but template missing.</p>""", 200

@app.route('/generate', methods=['POST'])
def generate_presentation():
    try:
        # Validate inputs
        if 'input_text' not in request.form:
            return jsonify({'error': 'No input text provided'}), 400
        
        input_text = request.form['input_text'].strip()
        guidance = request.form.get('guidance', '').strip()
        provider = request.form.get('provider', 'openai').strip().lower()
        api_key = request.form.get('api_key', '').strip()
        
        if not input_text or len(input_text) < 20:
            return jsonify({'error': 'Please provide at least 20 characters of content'}), 400
        
        if not api_key:
            return jsonify({'error': 'API key is required'}), 400
        
        if provider not in ppt_generator.supported_providers:
            return jsonify({'error': f'Unsupported provider: {provider}'}), 400
        
        # Handle template file
        template_path = None
        if 'template_file' in request.files:
            template_file = request.files['template_file']
            if template_file.filename and template_file.filename.lower().endswith(('.pptx', '.potx')):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                    template_file.save(tmp_file.name)
                    template_path = tmp_file.name
        
        try:
            # Generate slides
            slides_data = ppt_generator.parse_text_to_slides(
                text=input_text,
                provider=provider, 
                api_key=api_key,
                guidance=guidance
            )
            
            if not slides_data:
                return jsonify({'error': 'Could not generate slides from content'}), 400
            
            # Create presentation
            presentation = ppt_generator.create_presentation(slides_data, template_path)
            
            # Save and return
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as output_file:
                presentation.save(output_file.name)
                
                return send_file(
                    output_file.name,
                    as_attachment=True,
                    download_name=f'presentation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pptx',
                    mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
                )
                
        except ValueError as e:
            return jsonify({'error': str(e)}), 400
        except Exception as e:
            logger.error(f"Generation error: {e}")
            return jsonify({'error': 'Presentation generation failed. Please try again.'}), 500
        finally:
            if template_path and os.path.exists(template_path):
                try:
                    os.unlink(template_path)
                except:
                    pass
    
    except Exception as e:
        logger.error(f"Request error: {e}")
        return jsonify({'error': 'Server error'}), 500

@app.route('/health')
def health():
    return jsonify({'status': 'healthy', 'timestamp': datetime.utcnow().isoformat()})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
