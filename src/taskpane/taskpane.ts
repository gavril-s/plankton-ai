import { OpenRouterService } from '../services/openrouter';
import { WordService } from '../services/word';
import { Logger } from '../services/logger';

/* global document, Office */

declare const Word: any;

interface OpenRouterModel {
    id: string;
    name: string;
    description?: string;
}

interface OpenRouterMessage {
    role: 'user' | 'assistant' | 'system';
    content: string;
}

interface WordRequestContext {
    document: WordDocument;
    sync(): Promise<void>;
    load(object: any, options?: string[]): void;
}

interface WordDocument {
    onSelectionChanged: {
        add: (handler: () => void) => void;
        off: (handler: () => void) => void;
    };
    body: WordBody;
    getSelection(): WordRange;
}

interface WordBody {
    font: WordFont;
    paragraphs: {
        items: WordParagraph[];
    };
    text: string;
}

interface WordFont {
    name: string;
    size: number;
}

interface WordParagraph {
    lineSpacing: number;
    alignment: number;
}

interface WordRange {
    text: string;
    insertText(text: string, insertLocation: string): void;
}

let openRouterService: OpenRouterService;
let wordService: WordService;
const logger = Logger.getInstance();
let availableModels: OpenRouterModel[] = []; // Store models in memory
let autocompleteTimeout: NodeJS.Timeout | null = null;
let isAutocompleteEnabled = false;
let selectionHandler: (() => void) | null = null;

function updateDebugStatus(message: string) {
    logger.log(message);
    const debugDiv = document.getElementById('debug');
    if (debugDiv) {
        debugDiv.textContent = `Status: ${message}`;
    }
}

Office.onReady(info => {
    logger.log('Office.onReady called with info: ' + JSON.stringify(info));
    updateDebugStatus('Office.onReady called');

    if (info.host === Office.HostType.Word) {
        logger.log('Word host detected');
        updateDebugStatus('Word host detected');
        
        try {
            wordService = new WordService();
            logger.log('WordService initialized');
            
            // Initialize event listeners
            const apiKeyElement = document.getElementById('apiKey');
            const rewriteTextElement = document.getElementById('rewriteText');
            const fixGrammarElement = document.getElementById('fixGrammar');
            const modelSearchElement = document.getElementById('modelSearch');
            const submitPromptElement = document.getElementById('submitPrompt');
            const applySettingsElement = document.getElementById('applySettings');
            const enableAutocompleteElement = document.getElementById('enableAutocomplete') as HTMLInputElement;
            const autocompleteDelayElement = document.getElementById('autocompleteDelay') as HTMLInputElement;
            const clearLogsElement = document.getElementById('clearLogs');

            if (!apiKeyElement) logger.log('apiKey element not found', 'error');
            if (!rewriteTextElement) logger.log('rewriteText element not found', 'error');
            if (!fixGrammarElement) logger.log('fixGrammar element not found', 'error');
            if (!modelSearchElement) logger.log('modelSearch element not found', 'error');
            if (!submitPromptElement) logger.log('submitPrompt element not found', 'error');
            if (!applySettingsElement) logger.log('applySettings element not found', 'error');
            if (!enableAutocompleteElement) logger.log('enableAutocomplete element not found', 'error');
            if (!autocompleteDelayElement) logger.log('autocompleteDelay element not found', 'error');
            if (!clearLogsElement) logger.log('clearLogs element not found', 'error');

            apiKeyElement?.addEventListener('change', initializeOpenRouter);
            rewriteTextElement?.addEventListener('click', handleRewriteText);
            fixGrammarElement?.addEventListener('click', handleFixGrammar);
            modelSearchElement?.addEventListener('input', handleModelSearch);
            submitPromptElement?.addEventListener('click', handleSubmitPrompt);
            applySettingsElement?.addEventListener('click', handleApplySettings);
            clearLogsElement?.addEventListener('click', () => {
                logger.clearLogs();
                logger.log('Logs cleared');
            });
            
            // Initialize autocomplete settings from localStorage
            const savedAutocomplete = localStorage.getItem('enableAutocomplete');
            const savedDelay = localStorage.getItem('autocompleteDelay');
            
            if (savedAutocomplete) {
                enableAutocompleteElement.checked = savedAutocomplete === 'true';
            }
            
            if (savedDelay) {
                autocompleteDelayElement.value = savedDelay;
            }

            // Add autocomplete event listeners
            enableAutocompleteElement?.addEventListener('change', () => {
                localStorage.setItem('enableAutocomplete', enableAutocompleteElement.checked.toString());
                if (enableAutocompleteElement.checked) {
                    initializeAutocomplete();
                } else {
                    removeAutocomplete();
                }
            });

            autocompleteDelayElement?.addEventListener('change', () => {
                localStorage.setItem('autocompleteDelay', autocompleteDelayElement.value);
                if (enableAutocompleteElement.checked) {
                    removeAutocomplete();
                    initializeAutocomplete();
                }
            });

            // Initialize autocomplete if enabled
            if (enableAutocompleteElement?.checked) {
                initializeAutocomplete();
            }
            
            logger.log('Event listeners initialized');
            updateDebugStatus('Event listeners initialized');

            // Try to load API key from localStorage
            const savedApiKey = localStorage.getItem('openRouterApiKey');
            if (savedApiKey) {
                logger.log('Found saved API key');
                (document.getElementById('apiKey') as HTMLInputElement).value = savedApiKey;
                initializeOpenRouter();
            } else {
                logger.log('No saved API key found');
                updateDebugStatus('Ready - Please enter API key');
            }
        } catch (error: any) {
            const errorMessage = `Error during initialization: ${error.message}`;
            logger.log(errorMessage, 'error');
            logger.log('Stack trace: ' + error.stack, 'error');
            updateDebugStatus(errorMessage);
        }
    } else {
        const errorMessage = `Not in Word context: ${info.host}`;
        logger.log(errorMessage, 'error');
        updateDebugStatus('Error: Not in Word context');
    }
});

function handleModelSearch(event: Event) {
    const searchInput = event.target as HTMLInputElement;
    const searchTerm = searchInput.value.toLowerCase();
    logger.log(`Searching models with term: ${searchTerm}`);
    
    // Filter models based on search term
    const filteredModels = availableModels.filter(model => {
        const searchableText = [
            model.id.toLowerCase(),
            model.name.toLowerCase(),
            (model.description || '').toLowerCase()
        ].join(' ');
        
        return searchableText.includes(searchTerm);
    });

    // Update dropdown with filtered models
    populateModelDropdown(filteredModels, false); // false means don't update availableModels
    logger.log(`Found ${filteredModels.length} matching models`);
}

async function initializeOpenRouter() {
    logger.log('Initializing OpenRouter');
    const apiKeyInput = document.getElementById('apiKey') as HTMLInputElement;
    const apiKey = apiKeyInput.value;
    
    if (apiKey) {
        try {
            openRouterService = new OpenRouterService(apiKey);
            localStorage.setItem('openRouterApiKey', apiKey);
            logger.log('OpenRouter service initialized');
            
            // Fetch available models
            try {
                const models = await openRouterService.getAvailableModels();
                availableModels = models; // Store models in memory
                populateModelDropdown(models, true);
                
                // Restore previously selected model if it exists
                const savedModelId = localStorage.getItem('selectedModelId');
                if (savedModelId) {
                    const modelSelect = document.getElementById('model') as HTMLSelectElement;
                    if (modelSelect) {
                        modelSelect.value = savedModelId;
                    }
                }
                
                logger.log('Models fetched and populated');
            } catch (error: any) {
                logger.log('Error fetching models: ' + error.message, 'error');
            }
            
            showStatus('API key saved', 'success');
            updateDebugStatus('OpenRouter initialized');
        } catch (error: any) {
            const errorMessage = `Error initializing OpenRouter: ${error.message}`;
            logger.log(errorMessage, 'error');
            logger.log('Stack trace: ' + error.stack, 'error');
            showStatus('Error initializing OpenRouter', 'error');
            updateDebugStatus(errorMessage);
        }
    } else {
        logger.log('No API key provided');
        showStatus('Please enter an API key', 'error');
        updateDebugStatus('Waiting for API key');
    }
}

function populateModelDropdown(models: OpenRouterModel[], updateStoredModels: boolean = true) {
    const modelSelect = document.getElementById('model') as HTMLSelectElement;
    if (!modelSelect) {
        logger.log('Model select element not found', 'error');
        return;
    }

    // Update stored models if needed
    if (updateStoredModels) {
        availableModels = models;
    }

    // Clear existing options
    modelSelect.innerHTML = '';

    if (models.length === 0) {
        const option = document.createElement('option');
        option.value = '';
        option.text = 'No matching models found';
        modelSelect.appendChild(option);
        return;
    }

    // Sort models by name
    models.sort((a, b) => a.name.localeCompare(b.name));

    // Add each model as an option
    models.forEach(model => {
        const option = document.createElement('option');
        option.value = model.id;
        option.text = model.name;
        if (model.description) {
            option.title = model.description; // Add tooltip with description
        }
        modelSelect.appendChild(option);
    });

    // Add change event listener to save selected model
    modelSelect.addEventListener('change', (event) => {
        const selectedValue = (event.target as HTMLSelectElement).value;
        localStorage.setItem('selectedModelId', selectedValue);
        logger.log(`Selected model saved: ${selectedValue}`);
    });

    logger.log(`Populated ${models.length} models in dropdown`);
}

async function handleRewriteText() {
    if (!openRouterService) {
        showStatus('Please enter an API key first', 'error');
        return;
    }

    try {
        showStatus('Rewriting text...', 'loading');
        const selectedText = await wordService.getSelectedText();
        
        if (!selectedText) {
            showStatus('Please select some text first', 'error');
            return;
        }

        const model = (document.getElementById('model') as HTMLSelectElement).value;
        
        const messages: OpenRouterMessage[] = [
            {
                role: 'system',
                content: 'You are a writing assistant. Rewrite the following text to convey the same meaning but with different wording. Maintain the tone and style but use different sentence structures and synonyms where appropriate.'
            },
            {
                role: 'user',
                content: selectedText
            }
        ];

        const rewrittenText = await openRouterService.generateCompletion(messages, model);
        await wordService.replaceSelection(rewrittenText);
        showStatus('Text rewritten successfully', 'success');
    } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        console.error('Error rewriting text:', errorMessage);
        showStatus('Error rewriting text. Please try again.', 'error');
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
        const customPrompt = (document.getElementById('customPrompt') as HTMLTextAreaElement).value;
        
        const messages: OpenRouterMessage[] = [
            {
                role: 'system',
                content: customPrompt || 'You are a grammar correction assistant. Fix any grammatical errors in the following text while preserving its meaning.'
            },
            {
                role: 'user',
                content: selectedText
            }
        ];

        const correctedText = await openRouterService.generateCompletion(messages, model);
        await wordService.replaceSelection(correctedText);
        showStatus('Grammar fixed successfully', 'success');
    } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        console.error('Error fixing grammar:', errorMessage);
        showStatus('Error fixing grammar. Please try again.', 'error');
    }
}

async function handleSubmitPrompt() {
    if (!openRouterService) {
        showStatus('Please enter an API key first', 'error');
        return;
    }

    try {
        showStatus('Processing custom prompt...', 'loading');
        const selectedText = await wordService.getSelectedText();
        const customPrompt = (document.getElementById('customPrompt') as HTMLTextAreaElement).value;
        
        if (!customPrompt) {
            showStatus('Please enter a custom prompt', 'error');
            return;
        }

        const model = (document.getElementById('model') as HTMLSelectElement).value;
        
        const messages: OpenRouterMessage[] = [
            {
                role: 'system',
                content: 'You are a helpful AI assistant.'
            },
            {
                role: 'user',
                content: customPrompt + (selectedText ? `\n\nContext:\n${selectedText}` : '')
            }
        ];

        const response = await openRouterService.generateCompletion(messages, model);
        
        if (selectedText) {
            await wordService.replaceSelection(response);
        } else {
            await wordService.insertText(response);
        }
        
        showStatus('Custom prompt processed successfully', 'success');
    } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        console.error('Error processing custom prompt:', errorMessage);
        showStatus('Error processing custom prompt. Please try again.', 'error');
    }
}

async function handleApplySettings() {
    try {
        showStatus('Applying document settings...', 'loading');

        await Word.run(async (context: WordRequestContext) => {
            const document = context.document;
            const body = document.body;
            
            // Load required properties
            context.load(body, ['font', 'paragraphs']);
            await context.sync();

            // Get values from form elements using window.document
            const lineSpacingElement = window.document.getElementById('lineSpacing') as HTMLSelectElement;
            const fontElement = window.document.getElementById('font') as HTMLSelectElement;
            const fontSizeElement = window.document.getElementById('fontSize') as HTMLSelectElement;
            const alignmentElement = window.document.getElementById('textAlignment') as HTMLSelectElement;

            // Apply font settings to the whole document
            body.font.name = fontElement.value;
            body.font.size = parseInt(fontSizeElement.value);

            // Apply line spacing and alignment to all paragraphs
            const paragraphs = body.paragraphs;
            paragraphs.items.forEach((paragraph: WordParagraph) => {
                paragraph.lineSpacing = parseFloat(lineSpacingElement.value);
                switch (alignmentElement.value) {
                    case 'Left':
                        paragraph.alignment = Word.Alignment.left;
                        break;
                    case 'Center':
                        paragraph.alignment = Word.Alignment.centered;
                        break;
                    case 'Right':
                        paragraph.alignment = Word.Alignment.right;
                        break;
                    case 'Justify':
                        paragraph.alignment = Word.Alignment.justified;
                        break;
                }
            });

            await context.sync();
            showStatus('Document settings applied successfully (margins must be set manually)', 'success');
        });
    } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        console.error('Error applying document settings:', errorMessage);
        showStatus('Error applying document settings. Please try again.', 'error');
    }
}

function initializeAutocomplete() {
    if (!openRouterService) {
        showStatus('Please enter an API key first', 'error');
        return;
    }

    try {
        // Create the handler function
        selectionHandler = handleSelectionChanged;

        // Use the Office.js event binding
        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            selectionHandler,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    isAutocompleteEnabled = true;
                    logger.log('Autocomplete initialized');
                    showStatus('Autocomplete enabled', 'success');
                } else {
                    logger.log(`Failed to register selection handler: ${result.error.message}`, 'error');
                    showStatus('Failed to enable autocomplete', 'error');
                }
            }
        );
    } catch (error: any) {
        logger.log(`Error initializing autocomplete: ${error.message}`, 'error');
        showStatus('Error initializing autocomplete', 'error');
    }
}

function removeAutocomplete() {
    if (!selectionHandler) {
        logger.log('No active autocomplete handler to remove');
        return;
    }

    try {
        // Remove the Office.js event binding
        Office.context.document.removeHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            { handler: selectionHandler },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    isAutocompleteEnabled = false;
                    selectionHandler = null;
                    logger.log('Autocomplete removed');
                    showStatus('Autocomplete disabled', 'success');
                } else {
                    logger.log(`Failed to remove selection handler: ${result.error.message}`, 'error');
                    showStatus('Failed to disable autocomplete', 'error');
                }
            }
        );
    } catch (error: any) {
        logger.log(`Error removing autocomplete: ${error.message}`, 'error');
        showStatus('Error disabling autocomplete', 'error');
    }
}

function handleSelectionChanged() {
    const enableAutocompleteElement = document.getElementById('enableAutocomplete') as HTMLInputElement;
    const autocompleteDelayElement = document.getElementById('autocompleteDelay') as HTMLInputElement;
    const model = (document.getElementById('model') as HTMLSelectElement).value;

    if (!enableAutocompleteElement?.checked || !openRouterService || !isAutocompleteEnabled) {
        return;
    }

    // Clear existing timeout
    if (autocompleteTimeout) {
        clearTimeout(autocompleteTimeout);
    }

    // Set new timeout
    const delay = parseInt(autocompleteDelayElement?.value || '1000');
    autocompleteTimeout = setTimeout(async () => {
        try {
            const currentText = await wordService.getSurroundingText(1000); // Get surrounding context
            if (!currentText) return;

            const suggestion = await openRouterService.autocomplete(currentText, model);
            if (suggestion) {
                await wordService.insertText(suggestion);
                logger.log('Autocomplete suggestion inserted');
            }
        } catch (error: any) {
            logger.log(`Autocomplete error: ${error.message}`, 'error');
        }
    }, delay);
}

function showStatus(message: string, type: 'success' | 'error' | 'loading') {
    const statusElement = document.getElementById('status');
    if (statusElement) {
        statusElement.textContent = message;
        statusElement.className = `status-message ${type}`;
    }
} 