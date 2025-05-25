export class Logger {
    private static instance: Logger;
    private debugDiv: HTMLElement | null;

    private constructor() {
        this.debugDiv = document.getElementById('debug');
    }

    public static getInstance(): Logger {
        if (!Logger.instance) {
            Logger.instance = new Logger();
        }
        return Logger.instance;
    }

    public log(message: string, level: 'info' | 'error' = 'info'): void {
        const timestamp = new Date().toISOString();
        const prefix = '[Plankton AI]';
        const formattedMessage = `${prefix} ${timestamp} - ${message}`;

        if (level === 'error') {
            console.error(formattedMessage);
        } else {
            console.log(formattedMessage);
        }

        // Update debug div if it exists
        if (this.debugDiv) {
            this.debugDiv.textContent = formattedMessage;
            this.debugDiv.className = level === 'error' ? 'debug-error' : 'debug-info';
        }
    }
} 