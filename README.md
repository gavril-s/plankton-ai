# Plankton AI Word Add-in

A powerful Microsoft Word Add-in that brings advanced AI capabilities and document formatting tools directly into your Word interface. Powered by Plankton AI's language models, this add-in helps you write better, faster, and more efficiently.

## Table of Contents
- [Features](#features)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage Guide](#usage-guide)
- [Development](#development)
  - [Project Structure](#project-structure)
  - [Available Scripts](#available-scripts)
  - [Building for Production](#building-for-production)

## Features

### AI-Powered Writing Assistant
- **Custom AI Prompts**: Create custom instructions for the AI to help with specific writing tasks
- **Intelligent Text Rewriting**: Rephrase your text while preserving its original meaning
- **Grammar Correction**: Advanced grammar and style improvement suggestions
- **Model Selection**: Choose from various AI models to suit your needs
- **Context-Aware**: AI understands the context of your document for better suggestions

### Document Formatting Tools
- **Font Management**:
  - Multiple font family options (Times New Roman, Arial, Calibri)
  - Font size control (12, 14, 16 pt)
- **Layout Controls**:
  - Line spacing options (Single, 1.5, Double)
  - Text alignment (Left, Center, Right, Justify)
  - Margin settings in millimeters
- **Real-time Preview**: See your formatting changes instantly

## Installation

1. Clone the repository:
```bash
git clone https://github.com/planktonai/word-addin.git
cd word-addin
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm start
```

4. Sideload the add-in in Word:
   - [Windows sideloading instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
   - [Mac sideloading instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)

## Configuration

1. **API Key Setup**:
   - Get your API key from [openrouter](https://openrouter.ai/)
   - Enter the key in the add-in's settings panel
   - The key is securely stored in your browser's local storage

2. **Model Selection**:
   - Choose your preferred AI model from the dropdown
   - Models are sorted by capability and response time
   - Your selection is remembered between sessions

3. **Default Settings** (optional):
   - Configure default font and spacing in the settings
   - Set your preferred margin values
   - Customize the AI behavior for your needs

## Usage Guide

### Document Formatting
1. Open the Plankton AI panel in Word
2. Select your desired formatting options:
   - Choose font family and size
   - Set line spacing
   - Adjust text alignment
   - Configure margins
3. Click "Apply Document Settings" to format your document

### AI Features
1. **Custom Prompts**:
   - Select the text you want to work with
   - Enter your prompt in the custom prompt field
   - Click "Custom Prompt" to get AI assistance

2. **Text Rewriting**:
   - Select the text to rewrite
   - Click "Rewrite Text"
   - The AI will preserve meaning while changing the wording

3. **Grammar Correction**:
   - Select text to check
   - Click "Fix Grammar"
   - Review and apply the suggested corrections

## Development

### Project Structure
```
plankton-ai/
├── src/
│   ├── services/        # Core services (AI, Word, Logger)
│   └── taskpane/       # Main add-in UI components
├── assets/             # Static assets
├── manifest.xml        # Add-in manifest
├── package.json        # Project dependencies
└── webpack.config.js   # Build configuration
```

### Available Scripts
- `npm start` - Start development server
- `npm run build` - Build for production
- `npm run build:dev` - Build for development
- `npm run validate` - Validate the manifest file

### Building for Production
1. Update version in package.json and manifest.xml
2. Run `npm run build`
3. Test the production build thoroughly
4. Deploy the contents of the `dist` folder
