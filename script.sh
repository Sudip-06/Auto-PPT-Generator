#!/bin/bash

# Auto PPT Generator - Deployment Helper Script
# This script helps you deploy to various platforms

set -e

echo "🎯 Auto PPT Generator - Deployment Helper"
echo "========================================="

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Function to print colored output
print_status() {
    echo -e "${GREEN}✅ $1${NC}"
}

print_warning() {
    echo -e "${YELLOW}⚠️  $1${NC}"
}

print_error() {
    echo -e "${RED}❌ $1${NC}"
}

print_info() {
    echo -e "${BLUE}ℹ️  $1${NC}"
}

# Check if git is initialized
if [ ! -d ".git" ]; then
    print_error "Git repository not found. Initializing..."
    git init
    git add .
    git commit -m "Initial commit: Auto PPT Generator"
    print_status "Git repository initialized"
fi

# Function to deploy to different platforms
deploy_render() {
    print_info "Deploying to Render..."
    
    echo ""
    echo "To deploy to Render:"
    echo "1. Push your code to GitHub:"
    echo "   git add ."
    echo "   git commit -m 'Deploy to Render'"
    echo "   git push origin main"
    echo ""
    echo "2. Go to https://render.com and sign up/login"
    echo "3. Click 'New +' → 'Web Service'"
    echo "4. Connect your GitHub repository"
    echo "5. Use these settings:"
    echo "   - Name: auto-ppt-generator"
    echo "   - Environment: Python 3"
    echo "   - Build Command: pip install -r requirements.txt"
    echo "   - Start Command: gunicorn --bind 0.0.0.0:\$PORT app:app"
    echo "   - Plan: Free (for testing) or Starter (\$7/month)"
    echo ""
    echo "6. Click 'Create Web Service'"
    echo "7. Wait 2-3 minutes for deployment"
    echo ""
    print_status "Instructions displayed! Your app will be live at: https://your-app-name.onrender.com"
}

deploy_heroku() {
    print_info "Preparing for Heroku deployment..."
    
    # Create Procfile if it doesn't exist
    if [ ! -f "Procfile" ]; then
        echo "web: gunicorn --bind 0.0.0.0:\$PORT app:app" > Procfile
        print_status "Procfile created"
    fi
    
    # Create runtime.txt if it doesn't exist
    if [ ! -f "runtime.txt" ]; then
        echo "python-3.9.18" > runtime.txt
        print_status "runtime.txt created"
    fi
    
    echo ""
    echo "To deploy to Heroku:"
    echo "1. Install Heroku CLI: https://devcenter.heroku.com/articles/heroku-cli"
    echo "2. Login: heroku login"
    echo "3. Create app: heroku create your-app-name"
    echo "4. Deploy: git push heroku main"
    echo ""
    print_status "Heroku files prepared!"
}

deploy_vercel() {
    print_info "Preparing for Vercel deployment..."
    
    # Create vercel.json if it doesn't exist
    if [ ! -f "vercel.json" ]; then
        cat > vercel.json << EOL
{
  "version": 2,
  "builds": [
    {
      "src": "app.py",
      "use": "@vercel/python"
    }
  ],
  "routes": [
    {
      "src": "/(.*)",
      "dest": "app.py"
    }
  ]
}
EOL
        print_status "vercel.json created"
    fi
    
    echo ""
    echo "To deploy to Vercel:"
    echo "1. Install Vercel CLI: npm i -g vercel"
    echo "2. Login: vercel login"
    echo "3. Deploy: vercel --prod"
    echo ""
    print_status "Vercel configuration prepared!"
}

# Local development setup
setup_local() {
    print_info "Setting up local development environment..."
    
    # Check Python version
    if ! command -v python3 &> /dev/null; then
        print_error "Python 3 is required but not installed"
        exit 1
    fi
    
    python_version=$(python3 --version 2>&1 | cut -d' ' -f2 | cut -d'.' -f1,2)
    if [ "$(printf '%s\n' "3.8" "$python_version" | sort -V | head -n1)" != "3.8" ]; then
        print_error "Python 3.8+ required. Found: $python_version"
        exit 1
    fi
    
    print_status "Python version check passed: $python_version"
    
    # Create virtual environment
    if [ ! -d "venv" ]; then
        print_info "Creating virtual environment..."
        python3 -m venv venv
        print_status "Virtual environment created"
    fi
    
    # Activate virtual environment and install dependencies
    print_info "Installing dependencies..."
    source venv/bin/activate 2>/dev/null || source venv/Scripts/activate 2>/dev/null
    pip install --upgrade pip
    pip install -r requirements.txt
    
    print_status "Dependencies installed successfully"
    
    # Create .env file if it doesn't exist
    if [ ! -f ".env" ]; then
        cat > .env << EOL
# Flask Configuration
FLASK_ENV=development
PORT=5000

# Optional: Default API Keys (not recommended for production)
# OPENAI_API_KEY=your_openai_key_here
# ANTHROPIC_API_KEY=your_anthropic_key_here
# GOOGLE_API_KEY=your_google_key_here
EOL
        print_status ".env file created"
        print_warning "Please add your API keys to the .env file if desired"
    fi
    
    echo ""
    print_status "Local setup completed!"
    echo ""
    echo "To start the application:"
    echo "1. Activate virtual environment:"
    echo "   source venv/bin/activate  # Linux/macOS"
    echo "   venv\\Scripts\\activate     # Windows"
    echo "2. Run the server:"
    echo "   python app.py"
    echo "3. Open browser:"
    echo "   http://localhost:5000"
}

# Test the application
test_app() {
    print_info "Running basic tests..."
    
    # Check if all required files exist
    required_files=("app.py" "requirements.txt" "README.md")
    for file in "${required_files[@]}"; do
        if [ ! -f "$file" ]; then
            print_error "Required file missing: $file"
            exit 1
        fi
    done
    
    print_status "All required files present"
    
    # Check Python syntax
    if command -v python3 &> /dev/null; then
        python3 -m py_compile app.py
        print_status "Python syntax check passed"
    fi
    
    print_status "Basic tests completed"
}

# Main menu
show_menu() {
    echo ""
    echo "Choose deployment option:"
    echo "1) 🚀 Render (Recommended - Free tier available)"
    echo "2) 🟣 Heroku (Classic PaaS)"
    echo "3) ▲ Vercel (Serverless)"
    echo "4) 💻 Local Development Setup"
    echo "5) 🧪 Test Application"
    echo "6) 📋 Show All Instructions"
    echo "7) ❌ Exit"
    echo ""
    read -p "Enter your choice (1-7): " choice
}

show_all_instructions() {
    echo ""
    print_info "=== ALL DEPLOYMENT OPTIONS ==="
    echo
    
    echo "🚀 RENDER (Recommended)"
    echo "----------------------"
    echo "✅ Free tier available"
    echo "✅ Easy deployment from GitHub"
    echo "✅ Automatic SSL certificates"
    echo "✅ Built-in monitoring"
    echo ""
    echo "Steps:"
    echo "1. Push code to GitHub"
    echo "2. Connect repo to Render"
    echo "3. Deploy with one click"
    echo "4. Live at: https://your-app.onrender.com"
    echo ""
    
    echo "🟣 HEROKU"
    echo "----------"
    echo "✅ Mature platform"
    echo "✅ Extensive add-ons"
    echo "❌ No free tier (starts at \$7/month)"
    echo ""
    echo "Steps:"
    echo "1. Install Heroku CLI"
    echo "2. heroku create your-app"
    echo "3. git push heroku main"
    echo ""
    
    echo "▲ VERCEL"
    echo "---------"
    echo "✅ Serverless deployment"
    echo "✅ Global CDN"
    echo "⚠️  Better for static sites"
    echo ""
    echo "Steps:"
    echo "1. Install Vercel CLI"
    echo "2. vercel --prod"
    echo ""
    
    echo "💻 LOCAL DEVELOPMENT"
    echo "-------------------"
    echo "✅ Full control"
    echo "✅ Easy debugging"
    echo "✅ No hosting costs"
    echo ""
    echo "Steps:"
    echo "1. python -m venv venv"
    echo "2. source venv/bin/activate"
    echo "3. pip install -r requirements.txt"
    echo "4. python app.py"
    echo ""
    
    print_info "Choose the option that best fits your needs!"
}

# Check prerequisites
check_prerequisites() {
    print_info "Checking prerequisites..."
    
    # Check if we're in the right directory
    if [ ! -f "app.py" ]; then
        print_error "app.py not found. Please run this script from the project root directory."
        exit 1
    fi
    
    # Check Git
    if ! command -v git &> /dev/null; then
        print_warning "Git not found. Some deployment options require Git."
    else
        print_status "Git found"
    fi
    
    # Check Python
    if ! command -v python3 &> /dev/null; then
        print_warning "Python 3 not found. Required for local development."
    else
        python_version=$(python3 --version 2>&1 | cut -d' ' -f2)
        print_status "Python found: $python_version"
    fi
    
    # Check if requirements.txt exists
    if [ ! -f "requirements.txt" ]; then
        print_error "requirements.txt not found"
        exit 1
    fi
    
    print_status "Prerequisites check completed"
}

# Create GitHub repository
create_github_repo() {
    print_info "GitHub Repository Setup"
    echo ""
    echo "To create a GitHub repository:"
    echo ""
    echo "1. Go to https://github.com and login"
    echo "2. Click the '+' icon → 'New repository'"
    echo "3. Repository name: 'auto-ppt-generator'"
    echo "4. Description: 'AI-powered PowerPoint generator'"
    echo "5. Keep it Public (required for free deployments)"
    echo "6. Don't initialize with README (we already have files)"
    echo "7. Click 'Create repository'"
    echo ""
    echo "Then run these commands in your terminal:"
    echo ""
    echo "git remote add origin https://github.com/YOUR_USERNAME/auto-ppt-generator.git"
    echo "git branch -M main"
    echo "git push -u origin main"
    echo ""
    print_status "Follow these steps to set up your GitHub repository"
}

# Generate deployment summary
generate_summary() {
    echo ""
    print_info "=== DEPLOYMENT SUMMARY ==="
    echo ""
    echo "📁 Project Structure:"
    echo "   ✅ app.py (Main Flask application)"
    echo "   ✅ requirements.txt (Python dependencies)" 
    echo "   ✅ render.yaml (Render configuration)"
    echo "   ✅ README.md (Documentation)"
    echo "   ✅ LICENSE (MIT License)"
    echo ""
    echo "🚀 Recommended Deployment: Render"
    echo "   • Free tier available"
    echo "   • Automatic deployments from GitHub"
    echo "   • Built-in SSL and monitoring"
    echo "   • URL: https://your-app.onrender.com"
    echo ""
    echo "🔑 What you need:"
    echo "   • GitHub account (free)"
    echo "   • Render account (free)"
    echo "   • API key from OpenAI/Anthropic/Google"
    echo ""
    echo "⏱️ Deployment time: ~3 minutes"
    echo "💰 Cost: Free (Render free tier)"
    echo ""
    print_status "Ready to deploy!"
}

# Update project files
update_files() {
    print_info "Updating project files..."
    
    # Update .gitignore if it doesn't exist
    if [ ! -f ".gitignore" ]; then
        cat > .gitignore << EOL
# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.egg-info/
.installed.cfg
*.egg

# Virtual Environment
venv/
ENV/
env/

# IDE
.vscode/
.idea/
*.swp
*.swo
*~

# OS
.DS_Store
Thumbs.db

# Environment Variables
.env
.env.local
.env.production

# Logs
*.log

# Temporary files
*.tmp
temp/
EOL
        print_status ".gitignore created"
    fi
    
    # Create a simple health check script
    if [ ! -f "health_check.py" ]; then
        cat > health_check.py << EOL
#!/usr/bin/env python3
"""
Simple health check script for the Auto PPT Generator
"""
import requests
import sys

def check_health(url="http://localhost:5000"):
    try:
        response = requests.get(f"{url}/health", timeout=10)
        if response.status_code == 200:
            print("✅ Application is healthy")
            return True
        else:
            print(f"❌ Health check failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Health check failed: {e}")
        return False

if __name__ == "__main__":
    url = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:5000"
    success = check_health(url)
    sys.exit(0 if success else 1)
EOL
        print_status "health_check.py created"
    fi
}

# Main execution
main() {
    # Check prerequisites first
    check_prerequisites
    
    while true; do
        show_menu
        
        case $choice in
            1)
                deploy_render
                ;;
            2)
                deploy_heroku
                ;;
            3)
                deploy_vercel
                ;;
            4)
                setup_local
                ;;
            5)
                test_app
                ;;
            6)
                show_all_instructions
                ;;
            7)
                print_info "Goodbye! 👋"
                exit 0
                ;;
            *)
                print_error "Invalid option. Please choose 1-7."
                ;;
        esac
        
        echo ""
        read -p "Press Enter to continue or Ctrl+C to exit..."
    done
}

# Show initial information
echo ""
print_info "Auto PPT Generator Deployment Helper"
echo "This script will help you deploy your application to various platforms."
echo ""

# Check if user wants to see summary first
read -p "Would you like to see the deployment summary first? (y/n): " show_summary

if [[ $show_summary =~ ^[Yy]$ ]]; then
    generate_summary
    echo ""
    read -p "Press Enter to continue to deployment options..."
fi

# Update files
update_files

# Run main menu
main
