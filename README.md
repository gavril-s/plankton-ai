# Word AI Assistant

A Microsoft Word Add-in that integrates with OpenRouter to provide AI-powered writing assistance.

## Features

- Improve text using AI suggestions
- Fix grammar and style
- Generate creative ideas based on context
- Support for multiple AI models (GPT-3.5, GPT-4, Claude 2)

## Prerequisites

- [Node.js](https://nodejs.org) (version 14 or higher)
- Microsoft Word (Desktop or Online)
- OpenRouter API key (get one at [OpenRouter](https://openrouter.ai))

## Setup

1. Clone this repository:
```bash
git clone https://github.com/yourusername/word-ai.git
cd word-ai
```

2. Install dependencies:
```bash
npm install
```

3. Generate development certificates:
```bash
npm run dev-certs
```

4. Build the project:
```bash
npm run build
```

5. Start the development server:
```bash
npm start
```

## Development

- `npm run dev` - Start development server with hot reload
- `npm run build` - Build production version
- `npm run validate` - Validate the manifest file

## Loading the Add-in in Word

### Word Desktop
1. Open Word
2. Go to Insert > Get Add-ins
3. Choose "My Add-ins"
4. Select "Upload My Add-in"
5. Browse to the manifest file in your project
6. Click "Install"

### Word Online
1. Open Word Online
2. Go to Insert > Office Add-ins
3. Choose "Upload My Add-in"
4. Browse to the manifest file in your project
5. Click "Install"

## Usage

1. Open the add-in from the Home tab
2. Enter your OpenRouter API key
3. Select the AI model you want to use
4. Select text in your document
5. Click one of the available actions:
   - Improve Text
   - Fix Grammar
   - Generate Ideas

## Security

- Your OpenRouter API key is stored locally in your browser
- No data is stored on external servers
- All communication with OpenRouter is done via HTTPS

## License

MIT 