#!/usr/bin/env python3
"""
Auto PPT Generator - Complete Flask Backend
Transforms text into PowerPoint presentations using AI
"""

import os
import json
import tempfile
import logging
import re
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple

from flask import Flask, request, jsonify, send_file, render_template, abort
from flask_cors import CORS
import openai
import anthropic
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
import markdown

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app, origins=["*"])

# Configuration
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

class PPTGenerator:
    """Advanced PowerPoint generator with comprehensive features"""
    
    def __init__(self):
        self.supported_providers = ['openai', 'anthropic', 'google']
        self.max_slides = 15
        self.min_slides = 3
    
    def parse_text_to_slides(self, text: str, provider: str, api_key: str, guidance: str = "") -> List[Dict]:
        """Parse text into structured slide content using LLM"""
        
        # Estimate optimal slide count based on content length
        word_count = len(text.split())
        estimated_slides = max(self.min_slides, min(self.max_slides, word_count // 150))
        
        prompt = self._create_parsing_prompt(text, guidance, estimated_slides)
        
        try:
            logger.info(f"Parsing with {provider}, estimated slides: {estimated_slides}")
            
            if provider == 'openai':
                response = self._call_openai(prompt, api_key)
            elif provider == 'anthropic':
                response = self._call_anthropic(prompt, api_key)
            elif provider == 'google':
                response = self._call_google(prompt, api_key)
            else:
                raise ValueError(f"Unsupported provider: {provider}")
            
            slides = self._extract_slides_from_response(response)
            
            # Validate and clean slides
            slides = self._validate_slides(slides)
            
            logger.info(f"Successfully generated {len(slides)} slides")
            return slides
            
        except Exception as e:
            logger.error(f"Error parsing text with {provider}: {str(e)}")
            raise
    
    def _create_parsing_prompt(self, text: str, guidance: str, target_slides: int) -> str:
        """Create optimized prompt for LLM"""
        
        prompt = f"""You are an expert presentation designer. Convert the following text into a well-structured presentation.

INPUT TEXT:
{text}

PRESENTATION STYLE: {guidance or "professional, engaging presentation"}

REQUIREMENTS:
1. Create approximately {target_slides} slides (can vary based on content)
2. Include a compelling title slide
3. Break content into logical, digestible sections
4. Use clear, concise bullet points (3-6 per slide max)
5. Add relevant speaker notes
6. End with a strong conclusion or next steps

RETURN ONLY VALID JSON in this exact format:
{{
    "title": "Compelling Presentation Title",
    "slides": [
        {{
            "title": "Slide Title",
            "content": ["Concise bullet point 1", "Clear bullet point 2", "Action-oriented point 3"],
            "slide_type": "title_slide",
            "notes": "Helpful speaker notes for this slide"
        }},
        {{
            "title": "Main Content Slide",
            "content": ["Key insight", "Supporting detail", "Actionable takeaway"],
            "slide_type": "content_slide", 
            "notes": "Additional context for the presenter"
        }}
    ]
}}

SLIDE TYPES:
- "title_slide": Opening slide with main title
- "section_header": Major section dividers
- "content_slide": Main content slides
- "conclusion": Closing slide with key takeaways

Ensure each slide has:
- Clear, engaging title
- 2-6 bullet points maximum
- Actionable language
- Logical flow from previous slide
"""
        return prompt
    
    def _call_openai(self, prompt: str, api_key: str) -> str:
        """Call OpenAI API with error handling"""
        try:
            openai.api_key = api_key
            
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {
                        "role": "system", 
                        "content": "You are an expert presentation designer. Always respond with valid JSON only. No explanations or additional text."
                    },
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=2000,
                timeout=30
            )
            
            return response.choices[0].message.content.strip()
            
        except openai.error.AuthenticationError:
            raise ValueError("Invalid OpenAI API key")
        except openai.error.RateLimitError:
            raise ValueError("OpenAI rate limit exceeded. Please try again in a moment.")
        except openai.error.APIConnectionError:
            raise ValueError("Failed to connect to OpenAI. Please check your internet connection.")
        except Exception as e:
            raise ValueError(f"OpenAI API error: {str(e)}")
    
    def _call_anthropic(self, prompt: str, api_key: str) -> str:
        """Call Anthropic Claude API with error handling"""
        try:
            client = anthropic.Anthropic(api_key=api_key)
            
            response = client.messages.create(
                model="claude-3-haiku-20240307",
                max_tokens=2000,
                temperature=0.7,
                timeout=30,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            
            return response.content[0].text.strip()
            
        except anthropic.AuthenticationError:
            raise ValueError("Invalid Anthropic API key")
        except anthropic.RateLimitError:
            raise ValueError("Anthropic rate limit exceeded. Please try again in a moment.")
        except anthropic.APIConnectionError:
            raise ValueError("Failed to connect to Anthropic. Please check your internet connection.")
        except Exception as e:
            raise ValueError(f"Anthropic API error: {str(e)}")
    
    def _call_google(self, prompt: str, api_key: str) -> str:
        """Call Google Gemini API with error handling"""
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-pro')
            
            response = model.generate_content(
                prompt,
                generation_config={
                    'temperature': 0.7,
                    'max_output_tokens': 2000,
                }
            )
            
            if not response.text:
                raise ValueError("Empty response from Google Gemini")
                
            return response.text.strip()
            
        except Exception as e:
            error_msg = str(e).lower()
            if 'api key' in error_msg or 'authentication' in error_msg:
                raise ValueError("Invalid Google API key")
            elif 'quota' in error_msg or 'limit' in error_msg:
                raise ValueError("Google API quota exceeded. Please try again later.")
            else:
                raise ValueError(f"Google API error: {str(e)}")
    
    def _extract_slides_from_response(self, response: str) -> List[Dict]:
        """Extract and validate slide data from LLM response"""
        try:
            # Clean the response
            response = response.strip()
            
            # Remove code block markers if present
            if response.startswith('```json'):
                response = response[7:]
            if response.startswith('```'):
                response = response[3:]
            if response.endswith('```'):
                response = response[:-3]
            
            # Try to find JSON in the response
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                json_str = json_match.group()
                data = json.loads(json_str)
                
                if 'slides' in data and isinstance(data['slides'], list):
                    return data['slides']
            
            # If JSON parsing fails, try fallback parsing
            logger.warning("JSON parsing failed, using fallback method")
            return self._fallback_parse(response)
            
        except json.JSONDecodeError as e:
            logger.warning(f"JSON decode error: {e}, using fallback method")
            return self._fallback_parse(response)
        except Exception as e:
            logger.error(f"Error extracting slides: {e}")
            return self._fallback_parse(response)
    
    def _fallback_parse(self, text: str) -> List[Dict]:
        """Fallback method to parse text when JSON parsing fails"""
        slides = []
        lines = text.split('\n')
        current_slide = None
        slide_count = 0
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Check for headers (slide titles)
            if line.startswith('#') or (line.isupper() and len(line.split()) <= 8):
                if current_slide and current_slide.get('content'):
                    slides.append(current_slide)
                
                title = re.sub(r'^#+\s*', '', line).title()
                slide_count += 1
                
                current_slide = {
                    'title': title,
                    'content': [],
                    'slide_type': 'title_slide' if slide_count == 1 else 'content_slide',
                    'notes': f"Key points about {title.lower()}"
                }
                
            # Check for bullet points
            elif (line.startswith('-') or line.startswith('*') or line.startswith('•')) and current_slide:
                content = re.sub(r'^[-*•]\s*', '', line)
                if content and len(content) > 5:  # Filter out very short content
                    current_slide['content'].append(content)
            
            # Check for numbered lists
            elif re.match(r'^\d+[\.)]\s+', line) and current_slide:
                content = re.sub(r'^\d+[\.)]\s+', '', line)
                if content and len(content) > 5:
                    current_slide['content'].append(content)
        
        # Add the last slide
        if current_slide and current_slide.get('content'):
            slides.append(current_slide)
        
        # If no slides were parsed, create a basic structure
        if not slides:
            slides = [
                {
                    'title': 'Generated Presentation',
                    'content': ['Content extracted from your text', 'Organized into professional slides', 'Ready for presentation'],
                    'slide_type': 'title_slide',
                    'notes': 'This presentation was automatically generated from your content'
                }
            ]
        
        return slides
    
    def _validate_slides(self, slides: List[Dict]) -> List[Dict]:
        """Validate and clean slide data"""
        validated_slides = []
        
        for i, slide in enumerate(slides):
            # Ensure required fields exist
            validated_slide = {
                'title': str(slide.get('title', f'Slide {i+1}')).strip(),
                'content': [],
                'slide_type': slide.get('slide_type', 'content_slide'),
                'notes': str(slide.get('notes', '')).strip()
            }
            
            # Validate and clean content
            if isinstance(slide.get('content'), list):
                for item in slide['content']:
                    clean_item = str(item).strip()
                    if clean_item and len(clean_item) > 3:  # Filter very short items
                        validated_slide['content'].append(clean_item[:200])  # Limit length
            
            # Ensure at least some content
            if not validated_slide['content'] and i > 0:  # Skip empty check for title slide
                validated_slide['content'] = [f"Key points about {validated_slide['title']}"]
            
            # Limit content items
            validated_slide['content'] = validated_slide['content'][:6]
            
            # Validate slide type
            if validated_slide['slide_type'] not in ['title_slide', 'content_slide', 'section_header', 'conclusion']:
                validated_slide['slide_type'] = 'content_slide'
            
            validated_slides.append(validated_slide)
        
        # Ensure we have a reasonable number of slides
        if len(validated_slides) > self.max_slides:
            validated_slides = validated_slides[:self.max_slides]
        
        return validated_slides
    
    def create_presentation(self, slides_data: List[Dict], template_file=None) -> Presentation:
        """Create PowerPoint presentation with enhanced styling"""
        
        # Try to use template, fallback to default
        if template_file:
            try:
                prs = Presentation(template_file)
                self._clear_existing_slides(prs)
                logger.info("Using uploaded template")
            except Exception as e:
                logger.warning(f"Could not use template: {e}, using default")
                prs = Presentation()
        else:
            prs = Presentation()
        
        # Create slides
        for i, slide_data in enumerate(slides_data):
            self._create_slide(prs, slide_data, i)
        
        return prs
    
    def _clear_existing_slides(self, prs: Presentation):
        """Remove existing slides from template"""
        try:
            while len(prs.slides) > 0:
                xml_slides = prs.slides._sldIdLst
                xml_slides.remove(xml_slides[0])
        except Exception as e:
            logger.warning(f"Could not clear template slides: {e}")
    
    def _create_slide(self, prs: Presentation, slide_data: Dict, index: int):
        """Create individual slide with content"""
        
        # Choose appropriate layout
        if slide_data['slide_type'] == 'title_slide' and index == 0:
            slide_layout = prs.slide_layouts[0]  # Title slide
        else:
            slide_layout = prs.slide_layouts[1]  # Content slide
        
        slide = prs.slides.add_slide(slide_layout)
        
        # Add title
        if slide.shapes.title:
            slide.shapes.title.text = slide_data['title']
            
            # Style title if possible
            try:
                title_frame = slide.shapes.title.text_frame
                title_frame.paragraphs[0].font.size = Pt(32 if index == 0 else 28)
                title_frame.paragraphs[0].font.bold = True
            except:
                pass
        
        # Add content
        if len(slide.placeholders) > 1 and slide_data.get('content'):
            content_placeholder = slide.placeholders[1]
            text_frame = content_placeholder.text_frame
            text_frame.clear()  # Clear existing content
            
            # Add bullet points
            for i, bullet in enumerate(slide_data['content']):
                if i == 0:
                    p = text_frame.paragraphs[0]
                else:
                    p = text_frame.add_paragraph()
                
                p.text = bullet
                p.level = 0
                
                # Style bullet points
                try:
                    p.font.size = Pt(18)
                    p.space_after = Pt(12)
                except:
                    pass
        
        # Add speaker notes if supported
        if hasattr(slide, 'notes_slide') and slide_data.get('notes'):
            try:
                notes_slide = slide.notes_slide
                notes_text_frame = notes_slide.notes_text_frame
                notes_text_frame.text = slide_data['notes']
            except Exception as e:
                logger.debug(f"Could not add speaker notes: {e}")

# Initialize generator
ppt_generator = PPTGenerator()

@app.route('/')
def index():
    """Serve the main application"""
    try:
        return render_template('index.html')
    except Exception as e:
        logger.error(f"Error serving template: {e}")
        # Fallback if template file is missing
        return """
        <!DOCTYPE html>
        <html>
        <head>
            <title>Auto PPT Generator</title>
            <style>
                body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
                .error { color: #e53e3e; }
            </style>
        </head>
        <body>
            <h1>Auto PPT Generator</h1>
            <div class="error">
                <p>Template file is missing. Please ensure 'templates/index.html' exists.</p>
                <p>Check your deployment configuration.</p>
            </div>
        </body>
        </html>
        """, 500

@app.route('/health')
def health_check():
    """Health check endpoint for monitoring"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.utcnow().isoformat() + 'Z',
        'service': 'auto-ppt-generator',
        'version': '1.0.0'
    })

@app.route('/generate', methods=['POST'])
def generate_presentation():
    """Generate PowerPoint presentation from text"""
    try:
        # Validate request
        if 'input_text' not in request.form:
            return jsonify({'error': 'No input text provided'}), 400
        
        if 'api_key' not in request.form:
            return jsonify({'error': 'No API key provided'}), 400
        
        if 'provider' not in request.form:
            return jsonify({'error': 'No provider specified'}), 400
        
        # Extract form data
        input_text = request.form['input_text'].strip()
        guidance = request.form.get('guidance', '').strip()
        provider = request.form['provider'].strip().lower()
        api_key = request.form['api_key'].strip()
        
        # Validate inputs
        if not input_text or len(input_text) < 10:
            return jsonify({'error': 'Input text is too short. Please provide at least 10 characters.'}), 400
        
        if len(input_text) > 10000:
            return jsonify({'error': 'Input text is too long. Please keep it under 10,000 characters.'}), 400
        
        if provider not in ppt_generator.supported_providers:
            return jsonify({'error': f'Unsupported provider: {provider}'}), 400
        
        if not api_key:
            return jsonify({'error': 'API key is required'}), 400
        
        logger.info(f"Generating presentation with {provider}, text length: {len(input_text)}")
        
        # Handle optional template file
        template_path = None
        if 'template_file' in request.files:
            template_file = request.files['template_file']
            if template_file.filename and template_file.filename.endswith(('.pptx', '.potx')):
                # Save template temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                    template_file.save(tmp_file.name)
                    template_path = tmp_file.name
                    logger.info(f"Using template: {template_file.filename}")
        
        try:
            # Parse text into slides using LLM
            slides_data = ppt_generator.parse_text_to_slides(
                text=input_text,
                provider=provider,
                api_key=api_key,
                guidance=guidance
            )
            
            # Create PowerPoint presentation
            presentation = ppt_generator.create_presentation(slides_data, template_path)
            
            # Save presentation to temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
                presentation.save(tmp_file.name)
                
                # Return the file
                return send_file(
                    tmp_file.name,
                    as_attachment=True,
                    download_name=f'presentation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pptx',
                    mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
                )
                
        except ValueError as e:
            # These are user-facing errors (API key, rate limits, etc.)
            logger.warning(f"User error: {e}")
            return jsonify({'error': str(e)}), 400
            
        except Exception as e:
            logger.error(f"Unexpected error during generation: {e}")
            return jsonify({'error': 'An unexpected error occurred. Please try again.'}), 500
            
        finally:
            # Clean up template file if it exists
            if template_path and os.path.exists(template_path):
                try:
                    os.unlink(template_path)
                except:
                    pass
    
    except Exception as e:
        logger.error(f"Error in generate_presentation: {e}")
        return jsonify({'error': 'Server error occurred'}), 500

@app.route('/preview', methods=['POST'])
def preview_slides():
    """Preview slide structure before generation"""
    try:
        data = request.get_json()
        
        if not data or 'input_text' not in data:
            return jsonify({'error': 'No input text provided'}), 400
        
        input_text = data['input_text'].strip()
        guidance = data.get('guidance', '').strip()
        provider = data.get('provider', 'openai').strip().lower()
        api_key = data.get('api_key', '').strip()
        
        if not input_text or len(input_text) < 10:
            return jsonify({'error': 'Input text is too short'}), 400
        
        if not api_key:
            return jsonify({'error': 'API key is required'}), 400
        
        # Generate slide preview
        slides_data = ppt_generator.parse_text_to_slides(
            text=input_text,
            provider=provider,
            api_key=api_key,
            guidance=guidance
        )
        
        # Create preview response
        preview_slides = []
        for slide in slides_data[:5]:  # Limit preview to first 5 slides
            preview_slides.append({
                'title': slide['title'],
                'content_count': len(slide.get('content', [])),
                'slide_type': slide['slide_type']
            })
        
        return jsonify({
            'success': True,
            'total_slides': len(slides_data),
            'preview_slides': preview_slides,
            'estimated_duration': f"{len(slides_data) * 2}-{len(slides_data) * 3} minutes"
        })
        
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        logger.error(f"Error in preview: {e}")
        return jsonify({'error': 'Preview generation failed'}), 500

@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'File too large. Maximum size is 10MB.'}), 413

@app.errorhandler(500)
def internal_error(e):
    logger.error(f"Internal server error: {e}")
    return jsonify({'error': 'Internal server error'}), 500

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'Endpoint not found'}), 404

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') == 'development'
    
    logger.info(f"Starting Auto PPT Generator on port {port}")
    logger.info(f"Debug mode: {debug}")
    
    app.run(
        host='0.0.0.0',
        port=port,
        debug=debug,
        threaded=True
    )
