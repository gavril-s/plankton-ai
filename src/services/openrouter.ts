export interface OpenRouterMessage {
  role: 'user' | 'assistant' | 'system';
  content: string;
}

export interface OpenRouterResponse {
  choices: {
    message: {
      content: string;
      role: string;
    };
  }[];
}

export class OpenRouterService {
  private baseUrl = 'https://openrouter.ai/api/v1';
  private apiKey: string;

  constructor(apiKey: string) {
    this.apiKey = apiKey;
  }

  async getAvailableModels(): Promise<any> {
    try {
      const response = await fetch(`${this.baseUrl}/models`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.apiKey}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      return data.data; // OpenRouter returns models in data array
    } catch (error) {
      console.error('Error fetching models:', error);
      throw error;
    }
  }

  async generateCompletion(
    messages: OpenRouterMessage[],
    model: string = 'openai/gpt-3.5-turbo'
  ): Promise<string> {
    try {
      const response = await fetch(`${this.baseUrl}/chat/completions`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.apiKey}`,
          'Content-Type': 'application/json',
          'HTTP-Referer': window.location.href,
          'X-Title': 'Word AI Assistant'
        },
        body: JSON.stringify({
          model,
          messages,
        }),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data: OpenRouterResponse = await response.json();
      return data.choices[0].message.content;
    } catch (error) {
      console.error('Error calling OpenRouter API:', error);
      throw error;
    }
  }

  async improveText(text: string): Promise<string> {
    const messages: OpenRouterMessage[] = [
      {
        role: 'system',
        content: 'You are a helpful writing assistant. Improve the following text while maintaining its original meaning and tone.'
      },
      {
        role: 'user',
        content: text
      }
    ];

    return this.generateCompletion(messages);
  }

  async generateIdeas(prompt: string): Promise<string> {
    const messages: OpenRouterMessage[] = [
      {
        role: 'system',
        content: 'You are a creative writing assistant. Generate ideas based on the given prompt.'
      },
      {
        role: 'user',
        content: prompt
      }
    ];

    return this.generateCompletion(messages);
  }

  async fixGrammar(text: string): Promise<string> {
    const messages: OpenRouterMessage[] = [
      {
        role: 'system',
        content: 'You are a grammar correction assistant. Fix any grammatical errors in the following text while preserving its meaning.'
      },
      {
        role: 'user',
        content: text
      }
    ];

    return this.generateCompletion(messages);
  }

  async autocomplete(text: string, model: string = 'openai/gpt-3.5-turbo'): Promise<string> {
    const messages: OpenRouterMessage[] = [
      {
        role: 'system',
        content: 'You are an autocomplete assistant. Given the current text, provide a natural continuation that matches the style and context. Keep the continuation concise and relevant. Only provide the continuation text, do not repeat the input text.'
      },
      {
        role: 'user',
        content: text
      }
    ];

    return this.generateCompletion(messages, model);
  }
} 