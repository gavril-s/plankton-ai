export class Logger {
    private static instance: Logger;
    private logsContainer: HTMLElement | null;
    private maxLogs: number = 100;  // Maximum number of logs to keep

    private constructor() {
        this.logsContainer = document.getElementById('logs');
    }

    public static getInstance(): Logger {
        if (!Logger.instance) {
            Logger.instance = new Logger();
        }
        return Logger.instance;
    }

    public log(message: string, level: 'info' | 'error' = 'info'): void {
        const timestamp = new Date().toLocaleTimeString();
        const prefix = '[Plankton AI]';
        const formattedMessage = `${timestamp} - ${message}`;

        // Log to console
        if (level === 'error') {
            console.error(`${prefix} ${formattedMessage}`);
        } else {
            console.log(`${prefix} ${formattedMessage}`);
        }

        // Add to UI
        this.addLogToUI(formattedMessage, level);
    }

    private addLogToUI(message: string, level: 'info' | 'error'): void {
        if (this.logsContainer) {
            // Create new log entry
            const logEntry = document.createElement('div');
            logEntry.className = `log-entry ${level}`;
            logEntry.textContent = message;

            // Add to container
            this.logsContainer.appendChild(logEntry);

            // Scroll to bottom
            this.logsContainer.scrollTop = this.logsContainer.scrollHeight;

            // Remove old logs if exceeding maximum
            while (this.logsContainer.children.length > this.maxLogs) {
                this.logsContainer.removeChild(this.logsContainer.firstChild!);
            }
        }
    }

    public clearLogs(): void {
        if (this.logsContainer) {
            this.logsContainer.innerHTML = '';
        }
    }
} 