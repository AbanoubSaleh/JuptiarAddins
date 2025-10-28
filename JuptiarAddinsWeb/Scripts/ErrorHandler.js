/**
 * Centralized Error Handler for Jupiter Add-in
 * Provides consistent error handling, logging, and user feedback across all components
 */
class ErrorHandler {
    constructor() {
        this.errorCounts = new Map();
        this.maxRetries = 3;
        this.retryDelay = 1000; // 1 second
    }

    /**
     * Handle and display error with consistent formatting
     * @param {Error|string} error - Error object or message
     * @param {Object} options - Error handling options
     * @param {string} options.context - Context where error occurred
     * @param {string} options.userMessage - User-friendly message
     * @param {boolean} options.showToUser - Whether to show error to user (default: true)
     * @param {boolean} options.logError - Whether to log error (default: true)
     * @param {string} options.severity - Error severity: 'error', 'warning', 'info'
     * @param {Function} options.onRetry - Callback for retry action
     */
    handle(error, options = {}) {
        const {
            context = 'Unknown',
            userMessage = null,
            showToUser = true,
            logError = true,
            severity = 'error',
            onRetry = null
        } = options;

        // Extract error details
        const errorDetails = this.extractErrorDetails(error);
        
        // Log error if requested
        if (logError) {
            this.logError(errorDetails, context, severity);
        }

        // Track error frequency
        this.trackError(context, errorDetails.type);

        // Show user-friendly message if requested
        if (showToUser) {
            const displayMessage = userMessage || this.getUserFriendlyMessage(errorDetails);
            this.showUserError(displayMessage, severity, onRetry);
        }

        return {
            handled: true,
            errorType: errorDetails.type,
            canRetry: this.canRetry(context, errorDetails.type),
            retryCallback: onRetry
        };
    }

    /**
     * Extract standardized error details from various error types
     */
    extractErrorDetails(error) {
        if (typeof error === 'string') {
            return {
                message: error,
                type: 'GENERIC_ERROR',
                code: null,
                stack: null
            };
        }

        if (error instanceof Error) {
            return {
                message: error.message,
                type: this.classifyError(error),
                code: error.code || error.status || null,
                stack: error.stack
            };
        }

        // Handle API error responses
        if (error && typeof error === 'object') {
            return {
                message: error.message || error.Message || 'Unknown error occurred',
                type: error.code || error.Code || 'API_ERROR',
                code: error.status || error.statusCode || null,
                stack: null,
                details: error.errors || error.Errors || null
            };
        }

        return {
            message: 'Unknown error occurred',
            type: 'UNKNOWN_ERROR',
            code: null,
            stack: null
        };
    }

    /**
     * Classify error type based on error properties
     */
    classifyError(error) {
        if (error.name === 'NetworkError' || error.message.includes('fetch')) {
            return 'NETWORK_ERROR';
        }
        if (error.name === 'ValidationError') {
            return 'VALIDATION_ERROR';
        }
        if (error.name === 'AuthenticationError' || error.message.includes('401')) {
            return 'AUTH_ERROR';
        }
        if (error.name === 'AuthorizationError' || error.message.includes('403')) {
            return 'PERMISSION_ERROR';
        }
        if (error.message.includes('404')) {
            return 'NOT_FOUND_ERROR';
        }
        if (error.name === 'TimeoutError') {
            return 'TIMEOUT_ERROR';
        }
        return 'GENERIC_ERROR';
    }

    /**
     * Get user-friendly error message based on error type
     */
    getUserFriendlyMessage(errorDetails) {
        const messages = {
            'NETWORK_ERROR': 'Unable to connect to the server. Please check your internet connection and try again.',
            'AUTH_ERROR': 'Your session has expired. Please log in again.',
            'PERMISSION_ERROR': 'You do not have permission to perform this action.',
            'NOT_FOUND_ERROR': 'The requested resource was not found.',
            'VALIDATION_ERROR': 'Please check your input and try again.',
            'TIMEOUT_ERROR': 'The operation timed out. Please try again.',
            'GENERIC_ERROR': 'An unexpected error occurred. Please try again.',
            'UNKNOWN_ERROR': 'An unknown error occurred. Please contact support if the problem persists.'
        };

        return messages[errorDetails.type] || messages['GENERIC_ERROR'];
    }

    /**
     * Log error with appropriate level and context
     */
    logError(errorDetails, context, severity) {
        const logMessage = `[${context}] ${errorDetails.type}: ${errorDetails.message}`;
        
        switch (severity) {
            case 'error':
                console.error(logMessage, errorDetails);
                break;
            case 'warning':
                console.warn(logMessage, errorDetails);
                break;
            case 'info':
                console.info(logMessage, errorDetails);
                break;
            default:
                console.log(logMessage, errorDetails);
        }
    }

    /**
     * Show error message to user with consistent styling
     */
    showUserError(message, severity = 'error', onRetry = null) {
        // Try to use existing message banner if available
        if (window.messageBanner) {
            this.showMessageBanner(message, severity, onRetry);
            return;
        }

        // Fallback to alert for critical errors
        if (severity === 'error') {
            alert(`Error: ${message}`);
        } else {
            console.warn(`User message (${severity}): ${message}`);
        }
    }

    /**
     * Show message using Office UI Fabric MessageBanner
     */
    showMessageBanner(message, severity, onRetry) {
        try {
            const banner = window.messageBanner;
            if (banner && banner.showBanner) {
                // Set message
                const messageElement = $.select('.MessageBanner-text');
                if (messageElement) {
                    $.text(messageElement, message);
                }

                // Add retry button if callback provided
                if (onRetry) {
                    const actionsElement = $.select('.MessageBanner-actions');
                    if (actionsElement) {
                        $.empty(actionsElement);
                        $.append(actionsElement, `
                            <button class="ms-Button ms-Button--small" onclick="(${onRetry.toString()})()">
                                <span class="ms-Button-label">Retry</span>
                            </button>
                        `);
                    }
                }

                // Set appropriate styling based on severity
                const bannerElement = $.select('.MessageBanner');
                if (bannerElement) {
                    $.removeClass(bannerElement, 'is-error is-warning is-info');
                    $.addClass(bannerElement, `is-${severity}`);
                }

                banner.showBanner();

                // Auto-hide after delay for non-error messages
                if (severity !== 'error') {
                    setTimeout(() => {
                        if (banner.hideBanner) {
                            banner.hideBanner();
                        }
                    }, 5000);
                }
            }
        } catch (bannerError) {
            console.warn('Could not show message banner:', bannerError);
            // Fallback to alert
            alert(`${severity.toUpperCase()}: ${message}`);
        }
    }

    /**
     * Track error frequency for analysis
     */
    trackError(context, errorType) {
        const key = `${context}:${errorType}`;
        const count = this.errorCounts.get(key) || 0;
        this.errorCounts.set(key, count + 1);

        // Log frequent errors
        if (count > 5) {
            console.warn(`Frequent error detected: ${key} (${count + 1} occurrences)`);
        }
    }

    /**
     * Check if operation can be retried
     */
    canRetry(context, errorType) {
        const key = `${context}:${errorType}`;
        const count = this.errorCounts.get(key) || 0;
        
        // Don't retry certain error types
        const nonRetryableErrors = ['VALIDATION_ERROR', 'PERMISSION_ERROR', 'NOT_FOUND_ERROR'];
        if (nonRetryableErrors.includes(errorType)) {
            return false;
        }

        return count < this.maxRetries;
    }

    /**
     * Execute operation with automatic retry on failure
     */
    async withRetry(operation, context, options = {}) {
        const maxAttempts = options.maxAttempts || this.maxRetries;
        const delay = options.delay || this.retryDelay;
        
        for (let attempt = 1; attempt <= maxAttempts; attempt++) {
            try {
                return await operation();
            } catch (error) {
                const errorDetails = this.extractErrorDetails(error);
                
                if (attempt === maxAttempts || !this.canRetry(context, errorDetails.type)) {
                    // Final attempt or non-retryable error
                    throw error;
                }

                console.warn(`${context} attempt ${attempt} failed, retrying in ${delay}ms:`, errorDetails.message);
                await this.sleep(delay);
            }
        }
    }

    /**
     * Sleep utility for retry delays
     */
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * Clear error tracking (useful for testing or reset)
     */
    clearErrorTracking() {
        this.errorCounts.clear();
    }
}

// Create global instance
window.ErrorHandler = new ErrorHandler();
