/**
 * Configuration file for Jupiter Addins
 * Contains server settings and other configuration options
 */

window.JupiterConfig = {
    // Server Configuration
    server: {
        // Base URL of the Jupiter DMS API server
        baseUrl: 'https://localhost:7001',

        // API endpoint path
        apiEndpoint: '',

        // Request timeout in milliseconds
        timeout: 30000,

        // Health check endpoint
        healthEndpoint: '/health'
    },
    
    // Authentication Configuration
    auth: {
        // Whether to remember credentials by default
        rememberCredentials: false,
        
        // Whether to enable auto-login
        autoLogin: false,
        
        // Token refresh threshold (minutes before expiry)
        refreshThreshold: 5
    },
    
    // UI Configuration
    ui: {
        // Default page size for document listings
        defaultPageSize: 50,
        
        // Maximum file size for uploads (in bytes)
        maxFileSize: 100 * 1024 * 1024, // 100MB
        
        // Supported file types for upload
        supportedFileTypes: [
            '.doc', '.docx', '.pdf', '.txt', '.rtf',
            '.xls', '.xlsx', '.ppt', '.pptx',
            '.jpg', '.jpeg', '.png', '.gif', '.bmp',
            '.zip', '.rar', '.7z'
        ],
        
        // Auto-refresh interval for document lists (milliseconds)
        autoRefreshInterval: 30000
    },
    
    // Feature Flags
    features: {
        // Enable document search functionality
        enableSearch: true,
        
        // Enable document versioning
        enableVersioning: true,
        
        // Enable document check-in/check-out
        enableCheckInOut: true,
        
        // Enable document metadata editing
        enableMetadataEdit: true,
        
        // Enable folder management
        enableFolderManagement: true
    },
    
    // Logging Configuration
    logging: {
        // Log level: 'debug', 'info', 'warn', 'error'
        level: 'info',
        
        // Whether to log to console
        console: true,
        
        // Whether to send logs to server
        remote: false
    },
    
    // Development/Debug Settings
    debug: {
        // Enable debug mode
        enabled: false,
        
        // Show detailed error messages
        verboseErrors: false,
        
        // Enable API request/response logging
        logApiCalls: false
    }
};

/**
 * Get configuration value by path
 * @param {string} path - Dot-separated path to config value (e.g., 'server.baseUrl')
 * @param {*} defaultValue - Default value if path not found
 * @returns {*} Configuration value
 */
window.JupiterConfig.get = function(path, defaultValue = null) {
    const keys = path.split('.');
    let current = this;
    
    for (const key of keys) {
        if (current && typeof current === 'object' && key in current) {
            current = current[key];
        } else {
            return defaultValue;
        }
    }
    
    return current;
};

/**
 * Set configuration value by path
 * @param {string} path - Dot-separated path to config value
 * @param {*} value - Value to set
 */
window.JupiterConfig.set = function(path, value) {
    const keys = path.split('.');
    const lastKey = keys.pop();
    let current = this;
    
    for (const key of keys) {
        if (!(key in current) || typeof current[key] !== 'object') {
            current[key] = {};
        }
        current = current[key];
    }
    
    current[lastKey] = value;
};

/**
 * Initialize configuration with environment-specific overrides
 */
window.JupiterConfig.init = function() {
    // Check for environment-specific configuration
    const hostname = window.location.hostname;
    
    // Development environment
    if (hostname === 'localhost' || hostname === '127.0.0.1') {
        this.set('server.baseUrl', 'https://localhost:7001');
        this.set('debug.enabled', true);
        this.set('logging.level', 'debug');
    }
    
    // Production environment adjustments can be added here
    // Example:
    // if (hostname === 'your-production-domain.com') {
    //     this.set('server.baseUrl', 'https://api.your-production-domain.com');
    //     this.set('debug.enabled', false);
    //     this.set('logging.level', 'warn');
    // }
    
    console.log('Jupiter Config initialized:', this);
};

// Auto-initialize when script loads
document.addEventListener('DOMContentLoaded', function() {
    window.JupiterConfig.init();
});
