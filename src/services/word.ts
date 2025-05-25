/* global Word */

declare const Word: any;

export class WordService {
  async getSelectedText(): Promise<string> {
    return new Promise((resolve, reject) => {
      Word.run(async (context: Word.RequestContext) => {
        const selection = context.document.getSelection();
        selection.load('text');
        
        try {
          await context.sync();
          resolve(selection.text);
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  async insertText(text: string): Promise<void> {
    return new Promise((resolve, reject) => {
      Word.run(async (context: Word.RequestContext) => {
        const selection = context.document.getSelection();
        selection.insertText(text, Word.InsertLocation.replace);
        
        try {
          await context.sync();
          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  async replaceSelection(text: string): Promise<void> {
    return new Promise((resolve, reject) => {
      Word.run(async (context: Word.RequestContext) => {
        const selection = context.document.getSelection();
        selection.insertText(text, Word.InsertLocation.replace);
        
        try {
          await context.sync();
          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  async getSurroundingText(characterCount: number = 500): Promise<string> {
    return new Promise((resolve, reject) => {
      Word.run(async (context: Word.RequestContext) => {
        const selection = context.document.getSelection();
        const range = selection.getRange('Whole');
        range.expandTo(context.document.body);
        range.load('text');
        
        try {
          await context.sync();
          const fullText = range.text;
          const selectionStart = fullText.indexOf(selection.text);
          
          let start = Math.max(0, selectionStart - characterCount);
          let end = Math.min(fullText.length, selectionStart + selection.text.length + characterCount);
          
          resolve(fullText.substring(start, end));
        } catch (error) {
          reject(error);
        }
      });
    });
  }
} 