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

from flask import Flask, request, jsonify, send_file, render_template_string
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

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app, origins=["*"])

# HTML Template - Complete Frontend
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Auto PPT Generator</title>
    <link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>🎯</text></svg>">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        :root {
            --primary-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            --secondary-gradient: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            --success-color: #48bb78;
            --error-color: #e53e3e;
            --text-primary: #2d3748;
            --text-secondary: #4a5568;
            --bg-light: #f8fafc;
            --bg-white: #ffffff;
            --border-light: #e2e8f0;
            --shadow-card: 0 20px 60px rgba(0,0,0,0.15);
            --shadow-hover: 0 10px 30px rgba(102, 126, 234, 0.4);
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: var(--primary-gradient);
            min-height: 100vh;
            color: var(--text-primary);
            line-height: 1.6;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }

        .header {
            text-align: center;
            margin-bottom: 40px;
            color: white;
        }

        .header h1 {
            font-size: 3.5rem;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
            font-weight: 700;
        }

        .header p {
            font-size: 1.3rem;
            opacity: 0.95;
            margin-bottom: 10px;
        }

        .header .subtitle {
            font-size: 1rem;
            opacity: 0.8;
            font-style: italic;
        }

        .main-card {
            background: var(--bg-white);
            border-radius: 25px;
            padding: 50px;
            box-shadow: var(--shadow-card);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255,255,255,0.2);
        }

        .form-section {
            margin-bottom: 40px;
        }

        .form-section h3 {
            margin-bottom: 20px;
            color: var(--text-secondary);
            font-size: 1.4rem;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .input-group {
            margin-bottom: 25px;
        }

        label {
            display: block;
            margin-bottom: 10px;
            font-weight: 600;
            color: var(--text-primary);
            font-size: 1rem;
        }

        .required {
            color: var(--error-color);
        }

        textarea, input[type="text"], input[type="password"], select {
            width: 100%;
            padding: 16px 20px;
            border: 2px solid var(--border-light);
            border-radius: 12px;
            font-size: 16px;
            font-family: inherit;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            background: var(--bg-light);
            resize: none;
        }

        textarea:focus, input[type="text"]:focus, input[type="password"]:focus, select:focus {
            outline: none;
            border-color: #667eea;
            background: var(--bg-white);
            box-shadow: 0 0 0 4px rgba(102, 126, 234, 0.1);
            transform: translateY(-2px);
        }

        textarea {
            min-height: 220px;
            resize: vertical;
            font-family: 'SF Mono', Consolas, 'Liberation Mono', Menlo, monospace;
            line-height: 1.6;
        }

        .textarea-counter {
            font-size: 12px;
            color: var(--text-secondary);
            text-align: right;
            margin-top: 5px;
        }

        .file-upload {
            position: relative;
            display: block;
            width: 100%;
        }

        .file-upload input[type="file"] {
            position: absolute;
            opacity: 0;
            width: 100%;
            height: 100%;
            cursor: pointer;
            z-index: 2;
        }

        .file-upload-label {
            display: block;
            padding: 30px;
            border: 3px dashed var(--border-light);
            border-radius: 15px;
            text-align: center;
            background: var(--bg-light);
            cursor: pointer;
            transition: all 0.3s ease;
            position: relative;
        }

        .file-upload-label:hover {
            border-color: #667eea;
            background: #eef2ff;
            transform: translateY(-2px);
        }

        .file-upload.has-file .file-upload-label {
            border-color: var(--success-color);
            background: #f0fff4;
            color: #276749;
        }

        .file-icon {
            font-size: 3rem;
            margin-bottom: 10px;
            display: block;
        }

        .file-text {
            font-weight: 600;
            margin-bottom: 5px;
        }

        .file-subtext {
            font-size: 14px;
            color: var(--text-secondary);
        }

        .provider-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }

        .provider-option {
            padding: 25px 20px;
            border: 2px solid var(--border-light);
            border-radius: 15px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            background: var(--bg-light);
            position: relative;
            overflow: hidden;
        }

        .provider-option::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.4), transparent);
            transition: left 0.5s;
        }

        .provider-option:hover::before {
            left: 100%;
        }

        .provider-option:hover {
            border-color: #667eea;
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(102, 126, 234, 0.2);
        }

        .provider-option.selected {
            border-color: #667eea;
            background: linear-gradient(135deg, #eef2ff 0%, #e0e7ff 100%);
            color: #667eea;
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(102, 126, 234, 0.2);
        }

        .provider-option input[type="radio"] {
            display: none;
        }

        .provider-logo {
            font-size: 2rem;
            margin-bottom: 10px;
            display: block;
        }

        .provider-name {
            font-weight: 700;
            font-size: 1.1rem;
            margin-bottom: 5px;
        }

        .provider-model {
            font-size: 0.9rem;
            opacity: 0.8;
        }

        .api-key-input {
            position: relative;
        }

        .api-key-toggle {
            position: absolute;
            right: 15px;
            top: 50%;
            transform: translateY(-50%);
            background: none;
            border: none;
            cursor: pointer;
            font-size: 1.2rem;
            color: var(--text-secondary);
            z-index: 3;
        }

        .generate-btn {
            width: 100%;
            padding: 20px 40px;
            background: var(--primary-gradient);
            color: white;
            border: none;
            border-radius: 15px;
            font-size: 1.2rem;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            text-transform: uppercase;
            letter-spacing: 1px;
            position: relative;
            overflow: hidden;
        }

        .generate-btn::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
            transition: left 0.6s;
        }

        .generate-btn:hover:not(:disabled)::before {
            left: 100%;
        }

        .generate-btn:hover:not(:disabled) {
            transform: translateY(-3px);
            box-shadow: var(--shadow-hover);
        }

        .generate-btn:active:not(:disabled) {
            transform: translateY(-1px);
        }

        .generate-btn:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
        }

        .status-section {
            margin-top: 30px;
        }

        .loading, .result, .error {
            display: none;
            text-align: center;
            margin-top: 30px;
            padding: 30px;
            border-radius: 15px;
            animation: fadeIn 0.3s ease;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .loading {
            background: linear-gradient(135deg, #f0f8ff 0%, #e6f3ff 100%);
            border: 2px solid #3182ce;
        }

        .result {
            background: linear-gradient(135deg, #f0fff4 0%, #e6fffa 100%);
            border: 2px solid var(--success-color);
        }

        .error {
            background: linear-gradient(135deg, #fed7d7 0%, #fbb6ce 100%);
            border: 2px solid var(--error-color);
            color: #742a2a;
        }

        .spinner {
            width: 50px;
            height: 50px;
            border: 5px solid #e2e8f0;
            border-top: 5px solid #667eea;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .status-icon {
            font-size: 3rem;
            margin-bottom: 15px;
            display: block;
        }

        .status-title {
            font-size: 1.5rem;
            font-weight: 700;
            margin-bottom: 10px;
        }

        .status-message {
            font-size: 1.1rem;
            margin-bottom: 20px;
            line-height: 1.5;
        }

        .download-btn {
            background: var(--success-color);
            color: white;
            border: none;
            padding: 15px 30px;
            border-radius: 10px;
            font-size: 1.1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            text-decoration: none;
            display: inline-block;
            margin-top: 15px;
        }

        .download-btn:hover {
            background: #38a169;
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(72, 187, 120, 0.3);
        }

        .demo-section {
            margin-top: 50px;
            padding-top: 40px;
            border-top: 2px solid var(--border-light);
        }

        .demo-section h3 {
            margin-bottom: 25px;
            color: var(--text-secondary);
            font-size: 1.4rem;
        }

        .demo-content {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 25px;
        }

        .demo-card {
            background: var(--bg-light);
            padding: 25px;
            border-radius: 15px;
            border: 1px solid var(--border-light);
            transition: all 0.3s ease;
        }

        .demo-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }

        .demo-title {
            font-weight: 700;
            margin-bottom: 15px;
            color: var(--text-primary);
        }

        .demo-text {
            background: white;
            padding: 20px;
            border-radius: 10px;
            font-size: 14px;
            line-height: 1.6;
            max-height: 200px;
            overflow-y: auto;
            border: 1px solid var(--border-light);
            font-family: 'SF Mono', Consolas, 'Liberation Mono', Menlo, monospace;
        }

        .use-demo-btn {
            width: 100%;
            margin-top: 15px;
            padding: 12px 20px;
            background: #667eea;
            color: white;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s ease;
        }

        .use-demo-btn:hover {
            background: #5a6fd8;
            transform: translateY(-2px);
        }

        .footer {
            text-align: center;
            margin-top: 60px;
            padding-top: 30px;
            border-top: 1px solid rgba(255,255,255,0.2);
            color: rgba(255,255,255,0.8);
        }

        .footer a {
            color: rgba(255,255,255,0.9);
            text-decoration: none;
        }

        .footer a:hover {
            text-decoration: underline;
        }

        /* Responsive Design */
        @media (max-width: 768px) {
            .container {
                padding: 15px;
            }

            .main-card {
                padding: 30px 25px;
                border-radius: 20px;
            }

            .header h1 {
                font-size: 2.5rem;
            }

            .header p {
                font-size: 1.1rem;
            }

            .provider-grid {
                grid-template-columns: 1fr;
            }

            .demo-content {
                grid-template-columns: 1fr;
            }

            .generate-btn {
                padding: 18px 30px;
                font-size: 1.1rem;
            }
        }

        @media (max-width: 480px) {
            .header h1 {
                font-size: 2rem;
            }

            .main-card {
                padding: 25px 20px;
            }

            .form-section h3 {
                font-size: 1.2rem;
            }
        }

        /* Dark mode support */
        @media (prefers-color-scheme: dark) {
            :root {
                --bg-light: #2d3748;
                --bg-white: #1a202c;
                --text-primary: #f7fafc;
                --text-secondary: #e2e8f0;
                --border-light: #4a5568;
            }
        }

        /* Loading animation */
        .pulse {
            animation: pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
        }

        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.5; }
        }

        /* Success checkmark animation */
        .checkmark {
            display: inline-block;
            width: 60px;
            height: 60px;
            border-radius: 50%;
            background: var(--success-color);
            position: relative;
            margin: 0 auto 20px;
        }

        .checkmark::after {
            content: '';
            position: absolute;
            left: 22px;
            top: 18px;
            width: 12px;
            height: 20px;
            border: solid white;
            border-width: 0 3px 3px 0;
            transform: rotate(45deg);
            animation: checkmark 0.5s ease-in-out 0.2s both;
        }

        @keyframes checkmark {
            0% {
                width: 0;
                height: 0;
            }
            100% {
                width: 12px;
                height: 20px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🎯 Auto PPT Generator</h1>
            <p>Transform your text into stunning presentations instantly</p>
            <div class="subtitle">Powered by AI • No signup required • Privacy-first</div>
        </div>

        <div class="main-card">
            <form id="pptForm">
                <div class="form-section">
                    <h3><span>📝</span> Your Content</h3>
                    <div class="input-group">
                        <label for="inputText">Input Text <span class="required">*</span></label>
                        <textarea 
                            id="inputText" 
                            placeholder="Paste your markdown, prose, or any text content here. The more detailed, the better your presentation will be...

Examples:
• Business proposals and pitch decks
• Research findings and reports  
• Course materials and tutorials
• Meeting notes and summaries
• Product documentation
• Conference presentations

Tip: Use headers (# ## ###) to structure your content - they'll become slide titles!"
                            required
                            maxlength="10000"
                        ></textarea>
                        <div class="textarea-counter" id="textCounter">0 / 10,000 characters</div>
                    </div>
                    <div class="input-group">
                        <label for="guidance">Style Guidance (Optional)</label>
                        <input 
                            type="text" 
                            id="guidance" 
                            placeholder="e.g., 'investor pitch deck', 'technical presentation', 'sales pitch', 'academic report'"
                            maxlength="200"
                        >
                        <small style="color: #718096; margin-top: 5px; display: block;">
                            Help the AI understand your presentation style and audience
                        </small>
                    </div>
                </div>

                <div class="form-section">
                    <h3><span>🎨</span> Template Upload</h3>
                    <div class="input-group">
                        <label>PowerPoint Template (Optional but Recommended)</label>
                        <div class="file-upload" id="templateUpload">
                            <input type="file" id="templateFile" accept=".pptx,.potx">
                            <label for="templateFile" class="file-upload-label">
                                <span class="file-icon">📁</span>
                                <div class="file-text">Click to upload your PowerPoint template</div>
                                <div class="file-subtext">Supports .pptx and .potx files • Max 10MB</div>
                            </label>
                        </div>
                        <small style="color: #718096; margin-top: 10px; display: block;">
                            Upload your branded template to maintain consistent styling, colors, and fonts
                        </small>
                    </div>
                </div>

                <div class="form-section">
                    <h3><span>🤖</span> AI Provider</h3>
                    <div class="provider-grid">
                        <label class="provider-option selected">
                            <input type="radio" name="provider" value="openai" checked>
                            <span class="provider-logo">🧠</span>
                            <div class="provider-name">OpenAI</div>
                            <div class="provider-model">GPT-3.5 Turbo</div>
                        </label>
                        <label class="provider-option">
                            <input type="radio" name="provider" value="anthropic">
                            <span class="provider-logo">🔮</span>
                            <div class="provider-name">Anthropic</div>
                            <div class="provider-model">Claude 3 Haiku</div>
                        </label>
                        <label class="provider-option">
                            <input type="radio" name="provider" value="google">
                            <span class="provider-logo">✨</span>
                            <div class="provider-name">Google</div>
                            <div class="provider-model">Gemini Pro</div>
                        </label>
                    </div>
                    <div class="input-group">
                        <label for="apiKey">API Key <span class="required">*</span></label>
                        <div class="api-key-input">
                            <input type="password" id="apiKey" placeholder="Enter your API key" required>
                            <button type="button" class="api-key-toggle" id="toggleApiKey">👁️</button>
                        </div>
                        <small style="color: #718096; margin-top: 8px; display: block;">
                            🔒 Your API key is used only for this session and never stored on our servers
                        </small>
                    </div>
                </div>

                <button type="submit" class="generate-btn" id="generateBtn">
                    🚀 Generate Presentation
                </button>
            </form>

            <div class="status-section">
                <div class="loading" id="loading">
                    <div class="spinner"></div>
                    <div class="status-title">Generating Your Presentation</div>
                    <div class="status-message">
                        <span id="loadingStep">Analyzing your content...</span>
                    </div>
                </div>

                <div class="result" id="result">
                    <div class="checkmark"></div>
                    <div class="status-title">🎉 Presentation Generated Successfully!</div>
                    <div class="status-message">
                        Your professional presentation is ready for download
                    </div>
                    <button id="downloadBtn" class="download-btn">
                        📥 Download Presentation
                    </button>
                </div>

                <div class="error" id="error">
                    <span class="status-icon">❌</span>
                    <div class="status-title">Generation Failed</div>
                    <div class="status-message" id="errorMessage">
                        Something went wrong. Please check your inputs and try again.
                    </div>
                    <button class="use-demo-btn" onclick="showDebugInfo()">Show Details</button>
                </div>
            </div>

            <div class="demo-section">
                <h3>💡 Try with Demo Content</h3>
                <div class="demo-content">
                    <div class="demo-card">
                        <div class="demo-title">🚀 Startup Pitch Deck</div>
                        <div class="demo-text" id="startupDemo"># Revolutionary AI-Powered Marketing Platform

## Executive Summary
Our platform combines artificial intelligence with advanced analytics to transform how businesses approach digital marketing. We've identified a $50B market opportunity in automated marketing optimization.

## The Problem
- 73% of marketing budgets are wasted on ineffective campaigns
- Traditional A/B testing takes weeks to show results
- Small businesses lack access to enterprise-level marketing tools
- Manual campaign optimization is time-consuming and error-prone

## Our Solution
- AI-driven real-time campaign optimization
- Automated audience segmentation and targeting
- Predictive analytics for ROI forecasting
- One-click integration with major advertising platforms

## Market Opportunity
The digital advertising market is projected to reach $786B by 2026. Our target segment of SMBs represents $150B of this market.

## Business Model
- SaaS subscription: $99-$999/month based on ad spend
- Success-based pricing: 10% of improved ROI
- Enterprise licensing: $50K-$500K annually

## Financial Projections
Year 1: $2M revenue, 500 customers
Year 2: $8M revenue, 2,000 customers  
Year 3: $25M revenue, 5,000 customers

## Team & Expertise
- CEO: Former Google Ads executive, 15 years experience
- CTO: Ex-Facebook ML engineer, PhD in Computer Science
- CMO: Led growth at 3 successful B2B SaaS companies

## Funding Requirements
Seeking $5M Series A to:
- Scale engineering team (40% of funds)
- Expand sales and marketing (35% of funds)  
- Develop enterprise features (25% of funds)</div>
                        <button class="use-demo-btn" onclick="useDemoContent('startup')">Use This Demo</button>
                    </div>

                    <div class="demo-card">
                        <div class="demo-title">📊 Research Report</div>
                        <div class="demo-text" id="researchDemo"># The Future of Remote Work: A Comprehensive Study

## Introduction
This study examines the long-term implications of remote work adoption across industries, analyzing data from 10,000+ companies worldwide between 2020-2024.

## Methodology
- Survey of 10,000+ companies across 25 countries
- Analysis of productivity metrics and employee satisfaction
- Longitudinal study spanning 4 years
- Control groups in traditional office environments

## Key Findings

### Productivity Metrics
- Remote workers show 22% increase in productivity
- Reduced sick days by 35%
- 40% improvement in work-life balance scores
- 18% increase in employee retention rates

### Challenges Identified
- Communication barriers in 67% of teams
- Difficulty in maintaining company culture
- Technology adoption hurdles for 45% of employees
- Management oversight concerns

### Industry Variations
- Tech sector: 95% positive adaptation
- Finance: 78% successful implementation  
- Manufacturing: 34% remote-compatible roles
- Healthcare: 23% remote opportunities

## Recommendations

### For Organizations
1. Invest in robust communication platforms
2. Develop clear remote work policies
3. Provide technology training and support
4. Create virtual team-building initiatives

### For Employees
1. Establish dedicated workspace
2. Maintain regular communication schedules
3. Set clear boundaries between work and personal time
4. Participate actively in virtual meetings

## Future Predictions
- 60% of companies will adopt hybrid models by 2025
- Remote work will become standard for knowledge workers
- New tools and technologies will emerge to support distributed teams
- Geographic talent pools will expand significantly

## Conclusion
Remote work represents a fundamental shift in how we approach employment, with significant benefits for both employers and employees when implemented thoughtfully.</div>
                        <button class="use-demo-btn" onclick="useDemoContent('research')">Use This Demo</button>
                    </div>

                    <div class="demo-card">
                        <div class="demo-title">🎓 Educational Content</div>
                        <div class="demo-text" id="educationDemo"># Introduction to Machine Learning

## What is Machine Learning?
Machine Learning (ML) is a subset of artificial intelligence that enables computers to learn and improve from experience without being explicitly programmed.

## Types of Machine Learning

### Supervised Learning
- Uses labeled training data
- Predicts outcomes for new data
- Examples: Email spam detection, image classification
- Common algorithms: Linear regression, decision trees, neural networks

### Unsupervised Learning
- Works with unlabeled data
- Finds hidden patterns and structures
- Examples: Customer segmentation, recommendation systems
- Common algorithms: K-means clustering, PCA, association rules

### Reinforcement Learning
- Learns through interaction with environment
- Uses rewards and penalties
- Examples: Game playing, robotics, autonomous vehicles
- Key concepts: Agent, environment, actions, rewards

## Key Concepts

### Training Data
- Historical data used to teach the algorithm
- Quality and quantity both matter
- Should be representative of real-world scenarios

### Features
- Individual measurable properties of observations
- Feature selection is crucial for model performance
- Can be numerical, categorical, or text-based

### Model Evaluation
- Accuracy: Percentage of correct predictions
- Precision: True positives / (True positives + False positives)
- Recall: True positives / (True positives + False negatives)
- F1-Score: Harmonic mean of precision and recall

## Common Applications

### Business Applications
- Customer churn prediction
- Fraud detection
- Price optimization
- Inventory management

### Healthcare
- Medical diagnosis assistance
- Drug discovery
- Personalized treatment plans
- Medical image analysis

### Technology
- Search engines
- Recommendation systems
- Voice assistants
- Autonomous systems

## Getting Started

### Prerequisites
- Basic statistics and probability
- Programming skills (Python recommended)
- Understanding of data structures
- Mathematical foundations (linear algebra, calculus)

### Learning Path
1. Start with supervised learning concepts
2. Practice with simple datasets
3. Learn data preprocessing techniques
4. Experiment with different algorithms
5. Work on real-world projects

### Tools and Libraries
- Python: scikit-learn, pandas, numpy
- R: caret, randomForest, e1071
- Online platforms: Kaggle, Google Colab
- Visualization: matplotlib, seaborn, ggplot2

## Best Practices
- Always start with data exploration
- Clean and preprocess data thoroughly  
- Use cross-validation for model evaluation
- Avoid overfitting with proper validation
- Document your experiments and results

## Conclusion
Machine Learning is a powerful tool that's reshaping industries. Success requires combining domain knowledge, technical skills, and practical experience through hands-on projects.</div>
                        <button class="use-demo-btn" onclick="useDemoContent('education')">Use This Demo</button>
                    </div>
                </div>
            </div>
        </div>

        <div class="footer">
            <p>Made with ❤️ for better presentations • <a href="https://github.com/yourusername/auto-ppt-generator" target="_blank">View on GitHub</a></p>
        </div>
    </div>

    <script>
        // Global variables
        let generatedBlob = null;
        let isGenerating = false;

        // Initialize the application
        document.addEventListener('DOMContentLoaded', function() {
            initializeApp();
        });

        function initializeApp() {
            setupEventListeners();
            setupProviderSelection();
            setupTextCounter();
            setupApiKeyToggle();
            loadSavedData();
        }

        function setupEventListeners() {
            document.getElementById('pptForm').addEventListener('submit', handleFormSubmit);
            document.getElementById('templateFile').addEventListener('change', handleFileUpload);
            document.getElementById('downloadBtn').addEventListener('click', downloadPresentation);
            
            // Auto-save form data
            ['inputText', 'guidance', 'apiKey'].forEach(id => {
                document.getElementById(id).addEventListener('input', saveFormData);
            });

            // Form validation
            document.getElementById('inputText').addEventListener('input', validateForm);
            document.getElementById('apiKey').addEventListener('input', validateForm);
        }

        function setupProviderSelection() {
            const providerOptions = document.querySelectorAll('.provider-option');
            providerOptions.forEach(option => {
                option.addEventListener('click', function() {
                    providerOptions.forEach(opt => opt.classList.remove('selected'));
                    this.classList.add('selected');
                    this.querySelector('input[type="radio"]').checked = true;
                    saveFormData();
                    updateApiKeyPlaceholder();
                });
            });
        }

        function setupTextCounter() {
            const textArea = document.getElementById('inputText');
            const counter = document.getElementById('textCounter');
            
            textArea.addEventListener('input', function() {
                const length = this.value.length;
                counter.textContent = `${length.toLocaleString()} / 10,000 characters`;
                
                if (length > 9000) {
                    counter.style.color = '#e53e3e';
                } else if (length > 7000) {
                    counter.style.color = '#dd6b20';
                } else {
                    counter.style.color = '#718096';
                }
            });
        }

        function setupApiKeyToggle() {
            const toggle = document.getElementById('toggleApiKey');
            const input = document.getElementById('apiKey');
            
            toggle.addEventListener('click', function() {
                if (input.type === 'password') {
                    input.type = 'text';
                    toggle.textContent = '🙈';
                } else {
                    input.type = 'password';
                    toggle.textContent = '👁️';
                }
            });
        }

        function updateApiKeyPlaceholder() {
            const provider = document.querySelector('input[name="provider"]:checked').value;
            const apiKeyInput = document.getElementById('apiKey');
            
            const placeholders = {
                'openai': 'sk-...',
                'anthropic': 'sk-ant-...',
                'google': 'AIza...'
            };
            
            apiKeyInput.placeholder = `Enter your ${provider.charAt(0).toUpperCase() + provider.slice(1)} API key (${placeholders[provider]})`;
        }

        function validateForm() {
            const inputText = document.getElementById('inputText').value.trim();
            const apiKey = document.getElementById('apiKey').value.trim();
            const generateBtn = document.getElementById('generateBtn');
            
            if (inputText && apiKey && !isGenerating) {
                generateBtn.disabled = false;
                generateBtn.textContent = '🚀 Generate Presentation';
            } else {
                generateBtn.disabled = true;
                if (isGenerating) {
                    generateBtn.textContent = '⏳ Generating...';
                } else {
                    generateBtn.textContent = '🚀 Generate Presentation';
                }
            }
        }

        function saveFormData() {
            const formData = {
                inputText: document.getElementById('inputText').value,
                guidance: document.getElementById('guidance').value,
                provider: document.querySelector('input[name="provider"]:checked').value,
                // Note: We don't save API keys for security
            };
            
            try {
                localStorage.setItem('autoPPTFormData', JSON.stringify(formData));
            } catch (e) {
                // Handle localStorage errors silently
            }
        }

        function loadSavedData() {
            try {
                const saved = localStorage.getItem('autoPPTFormData');
                if (saved) {
                    const formData = JSON.parse(saved);
                    
                    if (formData.inputText) {
                        document.getElementById('inputText').value = formData.inputText;
                        document.getElementById('inputText').dispatchEvent(new Event('input'));
                    }
                    
                    if (formData.guidance) {
                        document.getElementById('guidance').value = formData.guidance;
                    }
                    
                    if (formData.provider) {
                        const providerOption = document.querySelector(`input[value="${formData.provider}"]`);
                        if (providerOption) {
                            providerOption.checked = true;
                            providerOption.closest('.provider-option').classList.add('selected');
                            document.querySelectorAll('.provider-option').forEach(opt => {
                                if (!opt.contains(providerOption)) {
                                    opt.classList.remove('selected');
                                }
                            });
                        }
                    }
                }
            } catch (e) {
                // Handle localStorage errors silently
            }
            
            updateApiKeyPlaceholder();
            validateForm();
        }

        function handleFileUpload(event) {
            const file = event.target.files[0];
            const uploadDiv = document.getElementById('templateUpload');
            const label = uploadDiv.querySelector('.file-upload-label');
            
            if (file) {
                // Validate file size (10MB max)
                if (file.size > 10 * 1024 * 1024) {
                    showError('File size must be less than 10MB');
                    event.target.value = '';
                    return;
                }
                
                uploadDiv.classList.add('has-file');
                label.innerHTML = `
                    <span class="file-icon">✅</span>
                    <div class="file-text">${file.name}</div>
                    <div class="file-subtext">${(file.size / 1024 / 1024).toFixed(1)}MB • Ready to use</div>
                `;
            } else {
                uploadDiv.classList.remove('has-file');
                label.innerHTML = `
                    <span class="file-icon">📁</span>
                    <div class="file-text">Click to upload your PowerPoint template</div>
                    <div class="file-subtext">Supports .pptx and .potx files • Max 10MB</div>
                `;
            }
        }

        async function handleFormSubmit(event) {
            event.preventDefault();
            
            if (isGenerating) return;
            
            try {
                const formData = collectFormData();
                if (!validateFormData(formData)) return;
                
                isGenerating = true;
                showLoading();
                
                const result = await generatePresentation(formData);
                
                if (result.success) {
                    generatedBlob = result.blob;
                    showSuccess();
                } else {
                    throw new Error(result.error || 'Generation failed');
                }
                
            } catch (error) {
                console.error('Generation error:', error);
                showError(error.message);
            } finally {
                isGenerating = false;
                validateForm();
            }
        }

        function collectFormData() {
            return {
                inputText: document.getElementById('inputText').value.trim(),
                guidance: document.getElementById('guidance').value.trim(),
                provider: document.querySelector('input[name="provider"]:checked').value,
                apiKey: document.getElementById('apiKey').value.trim(),
                templateFile: document.getElementById('templateFile').files[0]
            };
        }

        function validateFormData(data) {
            if (!data.inputText) {
                showError('Please enter some text content');
                return false;
            }
            
            if (data.inputText.length < 50) {
                showError('Please provide more content (at least 50 characters) for a meaningful presentation');
                return false;
            }
            
            if (!data.apiKey) {
                showError('Please enter your API key');
                return false;
            }
            
            // Basic API key format validation
            const apiKeyFormats = {
                'openai': /^sk-[a-zA-Z0-9]{48,}$/,
                'anthropic': /^sk-ant-[a-zA-Z0-9-_]{90,}$/,
                'google': /^AIza[a-zA-Z0-9-_]{35}$/
            };
            
            if (apiKeyFormats[data.provider] && !apiKeyFormats[data.provider].test(data.apiKey)) {
                showError(`Invalid ${data.provider.charAt(0).toUpperCase() + data.provider.slice(1)} API key format`);
                return false;
            }
            
            return true;
        }

        async function generatePresentation(formData) {
            const apiFormData = new FormData();
            apiFormData.append('input_text', formData.inputText);
            apiFormData.append('guidance', formData.guidance);
            apiFormData.append('provider', formData.provider);
            apiFormData.append('api_key', formData.apiKey);
            
            if (formData.templateFile) {
                apiFormData.append('template_file', formData.templateFile);
            }

            try {
                updateLoadingStep('Connecting to AI service...');
                
                const response = await fetch('/generate', {
                    method: 'POST',
                    body: apiFormData
                });

                if (!response.ok) {
                    let errorMessage = 'Unknown error occurred';
                    try {
                        const errorData = await response.json();
                        errorMessage = errorData.error || `HTTP ${response.status}: ${response.statusText}`;
                    } catch (e) {
                        errorMessage = `HTTP ${response.status}: ${response.statusText}`;
                    }
                    throw new Error(errorMessage);
                }

                updateLoadingStep('Processing your presentation...');
                const blob = await response.blob();
                
                updateLoadingStep('Finalizing download...');
                
                return {
                    success: true,
                    blob: blob
                };

            } catch (error) {
                console.error('API call failed:', error);
                
                // Provide user-friendly error messages
                let userMessage = error.message;
                
                if (error.message.includes('401') || error.message.includes('Invalid API key')) {
                    userMessage = 'Invalid API key. Please check your key and try again.';
                } else if (error.message.includes('429')) {
                    userMessage = 'API rate limit exceeded. Please wait a moment and try again.';
                } else if (error.message.includes('400')) {
                    userMessage = 'Invalid request. Please check your input and try again.';
                } else if (error.message.includes('Failed to fetch')) {
                    userMessage = 'Network error. Please check your internet connection.';
                }
                
                throw new Error(userMessage);
            }
        }

        function updateLoadingStep(step) {
            const loadingStep = document.getElementById('loadingStep');
            if (loadingStep) {
                loadingStep.textContent = step;
            }
        }

        function showLoading() {
            hideAllStatus();
            document.getElementById('loading').style.display = 'block';
            updateLoadingStep('Analyzing your content...');
        }

        function showSuccess() {
            hideAllStatus();
            document.getElementById('result').style.display = 'block';
        }

        function showError(message) {
            hideAllStatus();
            document.getElementById('error').style.display = 'block';
            document.getElementById('errorMessage').textContent = message;
        }

        function hideAllStatus() {
            document.getElementById('loading').style.display = 'none';
            document.getElementById('result').style.display = 'none';
            document.getElementById('error').style.display = 'none';
        }

        function downloadPresentation() {
            if (generatedBlob) {
                const url = URL.createObjectURL(generatedBlob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `presentation_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.pptx`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                
                // Show feedback
                const btn = document.getElementById('downloadBtn');
                const originalText = btn.textContent;
                btn.textContent = '✅ Downloaded!';
                btn.style.background = '#38a169';
                
                setTimeout(() => {
                    btn.textContent = originalText;
                    btn.style.background = '#48bb78';
                }, 2000);
            }
        }

        function useDemoContent(type) {
            const demoTexts = {
                'startup': document.getElementById('startupDemo').textContent,
                'research': document.getElementById('researchDemo').textContent,
                'education': document.getElementById('educationDemo').textContent
            };
            
            const guidanceTexts = {
                'startup': 'investor pitch deck',
                'research': 'academic research presentation',
                'education': 'educational tutorial presentation'
            };
            
            document.getElementById('inputText').value = demoTexts[type];
            document.getElementById('guidance').value = guidanceTexts[type];
            
            // Update character counter
            document.getElementById('inputText').dispatchEvent(new Event('input'));
            
            // Scroll to top
            document.getElementById('inputText').scrollIntoView({ 
                behavior: 'smooth', 
                block: 'start' 
            });
            
            // Highlight the filled content
            setTimeout(() => {
                document.getElementById('inputText').focus();
                document.getElementById('inputText').setSelectionRange(0, 0);
            }, 500);
            
            // Save the demo data
            saveFormData();
            validateForm();
        }

        function showDebugInfo() {
            const errorDetails = document.getElementById('errorMessage').textContent;
            navigator.clipboard.writeText(`Auto PPT Generator Error:\n${errorDetails}\n\nTimestamp: ${new Date().toISOString()}`).then(() => {
                alert('Error details copied to clipboard. You can share this with support if needed.');
            }).catch(() => {
                alert(`Error details:\n${errorDetails}\n\nTimestamp: ${new Date().toISOString()}`);
            });
        }

        // Keyboard shortcuts
        document.addEventListener('keydown', function(e) {
            // Ctrl/Cmd + Enter to generate
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                e.preventDefault();
                if (!isGenerating && !document.getElementById('generateBtn').disabled) {
                    document.getElementById('generateBtn').click();
                }
            }
        });

        // Auto-focus on page load
        setTimeout(() => {
            if (!document.getElementById('inputText').value) {
                document.getElementById('inputText').focus();
            }
        }, 500);
    </script>
</body>
</html>
'''

class PPTGenerator:
    """Advanced PowerPoint generator with comprehensive features"""
    
    def __init__(self):
        self.supported_providers = ['openai', 'anthropic', 'google']
        self.max_slides = 20
        self.min_slides = 1
    
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
            model = genai.GenerativeModel('gemini-2.5-pro')
            
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
        return render_template_string(HTML_TEMPLATE)
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
