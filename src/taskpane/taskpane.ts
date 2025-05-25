import { OpenRouterService } from '../services/openrouter';
import { WordService } from '../services/word';

/* global document, Office */

let openRouterService: OpenRouterService;
let wordService: WordService;

Office.onReady(info => {
    if (info.host === Office.HostType.Word) {
        wordService = new WordService();
        
        // Initialize event listeners
        document.getElementById('apiKey')?.addEventListener('change', initializeOpenRouter);
        document.getElementById('improveText')?.addEventListener('click', handleImproveText);
        document.getElementById('fixGrammar')?.addEventListener('click', handleFixGrammar);
        document.getElementById('generateIdeas')?.addEventListener('click', handleGenerateIdeas);

        // Try to load API key from localStorage
        const savedApiKey = localStorage.getItem('openRouterApiKey');
        if (savedApiKey) {
            (document.getElementById('apiKey') as HTMLInputElement).value = savedApiKey;
            initializeOpenRouter();
        }
    }
});

function initializeOpenRouter() {
    const apiKeyInput = document.getElementById('apiKey') as HTMLInputElement;
    const apiKey = apiKeyInput.value;
    
    if (apiKey) {
        openRouterService = new OpenRouterService(apiKey);
        localStorage.setItem('openRouterApiKey', apiKey);
        showStatus('API key saved', 'success');
    } else {
        showStatus('Please enter an API key', 'error');
    }
}

async function handleImproveText() {
    if (!openRouterService) {
        showStatus('Please enter an API key first', 'error');
        return;
    }

    try {
        showStatus('Improving text...', 'loading');
        const selectedText = await wordService.getSelectedText();
        
        if (!selectedText) {
            showStatus('Please select some text first', 'error');
            return;
        }

        const model = (document.getElementById('model') as HTMLSelectElement).value;
        const improvedText = await openRouterService.improveText(selectedText);
        await wordService.replaceSelection(improvedText);
        showStatus('Text improved successfully', 'success');
    } catch (error) {
        console.error('Error improving text:', error);
        showStatus('Error improving text. Please try again.', 'error');
    }
}

async function handleFixGrammar() {
    if (!openRouterService) {
        showStatus('Please enter an API key first', 'error');
        return;
    }

    try {
        showStatus('Fixing grammar...', 'loading');
        const selectedText = await wordService.getSelectedText();
        
        if (!selectedText) {
            showStatus('Please select some text first', 'error');
            return;
        }

        const model = (document.getElementById('model') as HTMLSelectElement).value;
        const correctedText = await openRouterService.fixGrammar(selectedText);
        await wordService.replaceSelection(correctedText);
        showStatus('Grammar fixed successfully', 'success');
    } catch (error) {
        console.error('Error fixing grammar:', error);
        showStatus('Error fixing grammar. Please try again.', 'error');
    }
}

async function handleGenerateIdeas() {
    if (!openRouterService) {
        showStatus('Please enter an API key first', 'error');
        return;
    }

    try {
        showStatus('Generating ideas...', 'loading');
        const selectedText = await wordService.getSelectedText();
        
        if (!selectedText) {
            showStatus('Please select some text as context', 'error');
            return;
        }

        const model = (document.getElementById('model') as HTMLSelectElement).value;
        const ideas = await openRouterService.generateIdeas(selectedText);
        await wordService.insertText(ideas);
        showStatus('Ideas generated successfully', 'success');
    } catch (error) {
        console.error('Error generating ideas:', error);
        showStatus('Error generating ideas. Please try again.', 'error');
    }
}

function showStatus(message: string, type: 'success' | 'error' | 'loading') {
    const statusElement = document.getElementById('status');
    if (statusElement) {
        statusElement.textContent = message;
        statusElement.className = `status-message ${type}`;
    }
} 