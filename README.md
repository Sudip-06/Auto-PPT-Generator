# 🎯 Auto PPT Generator

> Transform your text into stunning PowerPoint presentations instantly using AI

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com)

## ✨ Features

- **🤖 Multi-AI Support** - Works with OpenAI GPT-3.5, Anthropic Claude, and Google Gemini
- **🎨 Template-Based Styling** - Upload your PowerPoint template to maintain branding
- **📝 Intelligent Content Parsing** - AI automatically structures your text into logical slides
- **🔒 Privacy-First** - API keys never stored, processed securely in memory
- **📱 Responsive Design** - Beautiful interface that works on all devices
- **⚡ Fast Generation** - Create presentations in under 30 seconds
- **💾 No Signup Required** - Use immediately without registration

## 🚀 Live Demo

Try it now: [https://auto-ppt-generator.onrender.com/](https://auto-ppt-generator.onrender.com/)

## 📋 What You Need

- API key from at least one provider:
  - [OpenAI API Key](https://platform.openai.com/api-keys) (Recommended)
  - [Anthropic API Key](https://console.anthropic.com/)
  - [Google AI API Key](https://makersuite.google.com/app/apikey)

## 🔧 One-Click Deployment to Render

### Option 1: Direct Deploy (Fastest)

1. **Fork this repository** to your GitHub account
2. **Sign up for Render** at [render.com](https://render.com)
3. **Click "New +" → "Web Service"**
4. **Connect your GitHub repository**
5. **Use these settings:**
   - **Name:** `auto-ppt-generator`
   - **Environment:** `Python 3`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn --bind 0.0.0.0:$PORT app:app`
   - **Plan:** Free (for testing) or Starter ($7/month)

6. **Click "Create Web Service"**
7. **Wait 2-3 minutes for deployment**
8. **Your app will be live!** 🎉

### Option 2: Using render.yaml (Automated)

1. Fork this repository
2. In Render dashboard, click "New +" → "Blueprint"
3. Connect your repo - Render will automatically detect the `render.yaml` file
4. Deploy with one click!

## 💻 Local Development Setup

```bash
# Clone the repository
git clone https://github.com/yourusername/auto-ppt-generator.git
cd auto-ppt-generator

# Create virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py

# Open in browser
# http://localhost:5000
```

## 📖 How to Use

### Step 1: Prepare Your Content
```markdown
# Your Presentation Title

## Introduction
Your opening content here...

## Main Points
- Key insight 1
- Supporting detail
- Action item

## Conclusion
Summary and next steps
```

### Step 2: Choose AI Provider & Enter API Key
- Select OpenAI, Anthropic, or Google
- Enter your API key (kept secure, never stored)
- Add optional style guidance like "investor pitch" or "technical presentation"

### Step 3: Upload Template (Optional)
- Upload your branded PowerPoint template (.pptx or .potx)
- App will extract colors, fonts, and layouts
- Maintains consistent branding across slides

### Step 4: Generate & Download
- Click "Generate Presentation"
- Wait 10-30 seconds for AI processing
- Download your professional PowerPoint file

## 🛠️ Technical Architecture

### Frontend
- **Pure HTML/CSS/JavaScript** - No frameworks, fast loading
- **Responsive Design** - Works on mobile and desktop
- **Real-time Validation** - Immediate feedback on inputs
- **File Upload Handling** - Drag & drop template support

### Backend
- **Flask Python Server** - Lightweight and fast
- **Multi-AI Integration** - Unified interface for different providers
- **PowerPoint Generation** - Uses python-pptx for PPTX creation
- **Template Processing** - Extracts and applies styles from uploaded templates

### AI Text Processing Algorithm

1. **Content Analysis**: AI analyzes text structure and identifies main topics
2. **Slide Mapping**: Content organized into logical slides with titles and bullets
3. **Type Detection**: Determines slide types (title, content, section headers, conclusion)
4. **Speaker Notes**: Generates helpful presenter notes for each slide

### Template Style Application

1. **Style Extraction**: Reads colors, fonts, layouts from uploaded template
2. **Asset Preservation**: Maintains existing images and graphics
3. **Layout Application**: Applies slide layouts to generated content
4. **Brand Consistency**: Ensures visual cohesion across all slides

## 🔒 Privacy & Security

- **🔐 Zero Data Storage** - Content and API keys never saved permanently
- **⏱️ Temporary Processing** - Files processed in memory and auto-deleted
- **🔒 Secure Transmission** - All communications use HTTPS
- **🏠 Local Template Processing** - Style extraction happens on server, not cloud

## 📊 API Endpoints

### `GET /`
Serves the main application interface

### `GET /health`
Health check endpoint for monitoring
```json
{
  "status": "healthy",
  "timestamp": "2024-01-01T12:00:00Z",
  "service": "auto-ppt-generator"
}
```

### `POST /generate`
Generates PowerPoint presentation

**Form Data:**
- `input_text`: Text content to convert
- `guidance`: Optional style guidance
- `provider`: AI provider (openai/anthropic/google)
- `api_key`: User's API key
- `template_file`: PowerPoint template (optional)

**Response:** Binary PPTX file download

### `POST /preview`
Preview slide structure before generation

**JSON Request:**
```json
{
  "input_text": "Your content...",
  "guidance": "presentation style",
  "provider": "openai", 
  "api_key": "your-key"
}
```

**JSON Response:**
```json
{
  "success": true,
  "total_slides": 8,
  "preview_slides": [...],
  "estimated_duration": "16-24 minutes"
}
```

## 🔧 Environment Variables

For production deployment, you can optionally set:

```env
FLASK_ENV=production
PORT=5000
WEB_CONCURRENCY=2
PYTHONUNBUFFERED=1

# Optional: Default API keys (not recommended for security)
OPENAI_API_KEY=your_key_here
ANTHROPIC_API_KEY=your_key_here
GOOGLE_API_KEY=your_key_here
```

## 🚨 Troubleshooting

### Common Issues

**"Invalid API Key" Error:**
- Verify your API key is correct and active
- Check you have credits/quota with your AI provider
- Ensure the provider matches your API key type

**"Generation Failed" Error:**
- Try with shorter content (under 10,000 characters)
- Check your internet connection
- Verify your API key has sufficient quota

**Template Upload Issues:**
- Ensure file is valid .pptx or .potx format
- Check file size is under 10MB
- Try with a different template file

**Slow Generation:**
- Large content takes longer to process
- Try breaking very long text into sections
- Consider using OpenAI for faster processing

### Getting Help

1. **Check the error message** - Usually provides specific guidance
2. **Try the demo content** - Helps identify if it's a content issue
3. **Verify API credentials** - Test your key with the provider directly
4. **Open GitHub issue** - Include error details (no sensitive data)

## 🎯 Best Practices

### Content Preparation
- Use markdown headers (# ## ###) for slide structure
- Keep bullet points concise and actionable
- Include clear introductions and conclusions
- Aim for 100-1500 words for optimal results

### Template Design
- Use consistent fonts and colors
- Include your branding elements
- Keep layouts simple and clean
- Test template compatibility beforehand

### AI Provider Selection
- **OpenAI**: Best for creative and business content
- **Anthropic**: Excellent for technical and analytical content  
- **Google**: Strong general-purpose performance

## 🤝 Contributing

We welcome contributions! Here's how to help:

### Development Setup
```bash
git clone https://github.com/yourusername/auto-ppt-generator.git
cd auto-ppt-generator
python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate on Windows
pip install -r requirements.txt
python app.py
```

### Contribution Guidelines
- Fork the repository
- Create a feature branch (`git checkout -b feature/amazing-feature`)
- Write clear, documented code
- Test with multiple AI providers
- Ensure responsive design
- Commit changes (`git commit -m 'Add amazing feature'`)
- Push to branch (`git push origin feature/amazing-feature`)
- Open a Pull Request

### Areas for Contribution
- Additional AI provider integrations
- Advanced template style extraction
- UI/UX improvements
- Performance optimizations
- Documentation improvements
- Bug fixes and testing

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

```
MIT License

Copyright (c) 2025 Sudip-06

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```


**Made with ❤️ by developers, for presenters everywhere**

Transform your ideas into presentations instantly • No signup required • Privacy-first • Open source

[🚀 Try it now](https://auto-ppt-generator.onrender.com/) • [⭐ Star on GitHub](https://github.com/Sudip-06/auto-ppt-generator) • [📝 Report Issues](https://github.com/Sudip-06/auto-ppt-generator/issues)
