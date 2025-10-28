/**
 * Authentication Manager - Handles user authentication and session management
 */
class AuthManager {
    constructor() {
        this._isAuthenticated = false;
        this.currentUser = null;
        this.authToken = null;
        this.isInitialized = false;
        // Initialize settings from configuration
        const configBaseUrl = window.JupiterConfig?.get('server.baseUrl');
        this.settings = {
            serverUrl: configBaseUrl || 'https://localhost:7001',
            apiEndpoint: window.JupiterConfig?.get('server.apiEndpoint') || '/api',
            timeout: window.JupiterConfig?.get('server.timeout') || 30000,
            rememberCredentials: false, // Disabled for security - only remember tokens
            autoLogin: false // Disabled - require explicit login
        };
        // Note: loadSettings and loadStoredCredentials are now async and called in initialize()
    }
    /**
     * Load settings from Office Runtime storage (more secure and persistent)
     */
    async loadSettings() {
        try {
            // Try to load from Office Runtime storage first (more secure)
            let savedSettings = null;
            try {
                const settingsJson = await OfficeRuntime.storage.getItem('juptiarSettings');
                if (settingsJson) {
                    savedSettings = JSON.parse(settingsJson);
                }
            } catch (runtimeError) {
                console.warn('Could not load from OfficeRuntime storage, trying Office settings:', runtimeError);
                // Fallback to Office.context.document.settings for backward compatibility
                await Office.context.document.settings.refreshAsync();
                savedSettings = Office.context.document.settings.get('juptiarSettings');
            }
            if (savedSettings) {
                this.settings = { ...this.settings, ...savedSettings };
                // Initialize service with settings only if it doesn't have a baseUrl
                if (window.jupiterService) {
                    // Only re-initialize if the service doesn't have a baseUrl or if our settings are different
                    if (!window.jupiterService.baseUrl || window.jupiterService.baseUrl === '') {
                        window.jupiterService.initialize(this.settings);
                    } else {
                    }
                }
            }
        } catch (error) {
            console.warn('Could not load settings:', error);
        }
    }
    /**
     * Save settings to Office Runtime storage (more secure and persistent)
     */
    async saveSettings(newSettings) {
        try {
            this.settings = { ...this.settings, ...newSettings };
            // Save to Office Runtime storage (primary method)
            try {
                await OfficeRuntime.storage.setItem('juptiarSettings', JSON.stringify(this.settings));
            } catch (runtimeError) {
                console.warn('Could not save to OfficeRuntime storage, using Office settings fallback:', runtimeError);
                // Fallback to Office.context.document.settings
                Office.context.document.settings.set('juptiarSettings', this.settings);
                await Office.context.document.settings.saveAsync();
            }
            // Update service configuration
            if (window.jupiterService) {
                window.jupiterService.initialize(this.settings);
            }
            return true;
        } catch (error) {
            console.error('Could not save settings:', error);
            throw new Error('Failed to save settings');
        }
    }
    /**
     * Load stored credentials if remember is enabled (using Office Runtime storage)
     * NOTE: Password storage has been disabled for security reasons.
     * Only authentication tokens are stored.
     */
    async loadStoredCredentials() {
        // Password storage disabled for security
        // Only tokens are stored and managed separately
    }
    /**
     * Store credentials securely using Office Runtime storage
     * NOTE: Password storage has been disabled for security reasons.
     */
    async storeCredentials(username, password) {
        // Password storage disabled for security
    }
    /**
     * Clear stored credentials from both storage methods
     * NOTE: Only clears legacy stored credentials if they exist.
     */
    async clearStoredCredentials() {
        try {
            // Clear legacy stored credentials if they exist
            try {
                await OfficeRuntime.storage.removeItem('juptiarCredentials');
            } catch (runtimeError) {
                // Ignore - likely doesn't exist
            }
            // Also clear from Office settings (fallback/legacy)
            try {
                Office.context.document.settings.remove('juptiarCredentials');
                await Office.context.document.settings.saveAsync();
            } catch (settingsError) {
                // Ignore - likely doesn't exist
            }
            this.storedUsername = null;
            this.storedPassword = null;
        } catch (error) {
            console.error('Could not clear legacy credentials:', error);
        }
    }
    /**
     * Store authentication token securely in Office Runtime storage
     */
    async storeAuthToken(token) {
        try {
            await OfficeRuntime.storage.setItem('authToken', token);
        } catch (error) {
            console.error('Could not store auth token:', error);
        }
    }
    /**
     * Load authentication token from Office Runtime storage
     */
    async loadAuthToken() {
        try {
            const token = await OfficeRuntime.storage.getItem('authToken');
            if (token) {
                return token;
            }
        } catch (error) {
            console.warn('Could not load auth token:', error);
        }
        return null;
    }
    /**
     * Clear authentication token from Office Runtime storage
     */
    async clearAuthToken() {
        try {
            await OfficeRuntime.storage.removeItem('authToken');
        } catch (error) {
            console.warn('Could not clear auth token:', error);
        }
    }
    // Password encryption/decryption methods removed for security
    // Only authentication tokens are stored now
    /**
     * Attempt login with credentials
     * NOTE: rememberCredentials parameter is ignored - only tokens are stored
     */
    async login(username, password, rememberCredentials = false) {
        try {
            if (!window.jupiterService) {
                console.error('Jupiter service not initialized');
                throw new Error('Jupiter service not initialized');
            }
            // Validate inputs
            if (!username || !password) {
                throw new Error('Username and password are required');
            }
            if (!this.settings.serverUrl) {
                throw new Error('Server URL not configured. Please check settings.');
            }
            // Ensure JupiterService is properly configured before login
            if (!window.jupiterService.baseUrl || window.jupiterService.baseUrl === '') {
                // Force re-initialization with correct settings
                window.jupiterService.initialize({
                    serverUrl: this.settings.serverUrl,
                    apiEndpoint: this.settings.apiEndpoint,
                    timeout: this.settings.timeout
                });
            }
            // Attempt login
            const response = await window.jupiterService.login(username, password);
            if (response.success && response.token) {
                this._isAuthenticated = true;
                this.authToken = response.token;
                this.currentUser = response.user || { username: username };
                // Store authentication token securely in Office Runtime storage
                await this.storeAuthToken(response.token);
                // Set token in JupiterService
                if (window.jupiterService) {
                    window.jupiterService.setAuthToken(response.token);
                }
                // Password storage disabled for security - only tokens are stored
                // Tokens are automatically stored in storeAuthToken() above
                // Update settings (rememberCredentials always false for security)
                await this.saveSettings({ rememberCredentials: false });
                this.notifyAuthStateChange();
                return { success: true, user: this.currentUser };
            } else {
                throw new Error(response.message || 'Login failed');
            }
        } catch (error) {
            this._isAuthenticated = false;
            this.authToken = null;
            this.currentUser = null;
            if (window.jupiterService) {
                window.jupiterService.clearAuthToken();
            }
            // Use centralized error handling
            if (window.ErrorHandler) {
                window.ErrorHandler.handle(error, {
                    context: 'AuthManager.login',
                    userMessage: 'Login failed. Please check your credentials and try again.',
                    showToUser: true,
                    severity: 'error'
                });
            }
            throw error;
        }
    }
    /**
     * Logout current user
     */
    async logout() {
        try {
            if (window.jupiterService && this._isAuthenticated) {
                await window.jupiterService.logout();
            }
        } catch (error) {
            console.warn('Logout request failed:', error);
        } finally {
            this._isAuthenticated = false;
            this.authToken = null;
            this.currentUser = null;
            // Clear stored authentication token
            await this.clearAuthToken();
            if (window.jupiterService) {
                window.jupiterService.clearAuthToken();
            }
            this.notifyAuthStateChange();
        }
    }
    /**
     * Validate current session
     */
    async validateSession() {
        try {
            if (!this.authToken) {
                return false;
            }
            if (!window.jupiterService) {
                return false;
            }
            const response = await window.jupiterService.validateSession();
            if (response.valid) {
                // If we don't have user info but token is valid, try to get user info
                if (!this.currentUser && response.username) {
                    this.currentUser = {
                        username: response.username,
                        userId: response.userId
                    };
                }
                return true;
            } else {
                await this.logout();
                return false;
            }
        } catch (error) {
            console.warn('Session validation failed:', error);
            await this.logout();
            return false;
        }
    }
    /**
     * Auto-login disabled for security (no password storage)
     */
    async autoLogin() {
        return false;
    }
    /**
     * Test connection to server
     */
    async testConnection(serverUrl, apiEndpoint = '/api') {
        try {
            const tempService = new JupiterService();
            tempService.initialize({
                serverUrl: serverUrl,
                apiEndpoint: apiEndpoint,
                timeout: 10000 // Shorter timeout for testing
            });
            // Try to make a simple request to test connectivity
            // Use a POST to Auth/login with empty body to test if endpoint exists
            // This should return 400 (bad request) but not 404 (not found)
            try {
                await tempService.makeRequest('POST', '/Auth/login', {});
            } catch (error) {
                // If we get 400 (bad request), the endpoint exists - connection is good
                // If we get 404 (not found), the endpoint doesn't exist - connection failed
                if (error.message.includes('400')) {
                    return { success: true, message: 'Connection successful - endpoint found' };
                } else if (error.message.includes('404')) {
                    throw new Error('API endpoint not found (404)');
                } else {
                    throw error;
                }
            }
            return { success: true, response: response };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }
    /**
     * Check if user is currently authenticated
     * @returns {boolean} True if authenticated
     */
    isAuthenticated() {
        return this._isAuthenticated;
    }
    /**
     * Get current authentication status
     */
    getAuthStatus() {
        return {
            isAuthenticated: this._isAuthenticated,
            user: this.currentUser,
            serverUrl: this.settings.serverUrl
        };
    }
    /**
     * Get current authentication token
     */
    getToken() {
        return this.authToken;
    }
    /**
     * Get current settings
     */
    getSettings() {
        return { ...this.settings };
    }
    /**
     * Get stored credentials (for auto-fill)
     * NOTE: Password storage disabled for security
     */
    getStoredCredentials() {
        return {
            username: '',
            hasStoredPassword: false
        };
    }
    /**
     * Notify listeners of authentication state changes
     */
    notifyAuthStateChange() {
        const event = new CustomEvent('juptiarAuthStateChanged', {
            detail: this.getAuthStatus()
        });
        window.dispatchEvent(event);
    }
    /**
     * Initialize authentication manager
     */
    async initialize() {
        // Load settings first
        await this.loadSettings();
        // Load stored credentials
        await this.loadStoredCredentials();
        // Try to restore previous session from stored token
        const storedToken = await this.loadAuthToken();
        if (storedToken) {
            this.authToken = storedToken;
            // Set token in JupiterService
            if (window.jupiterService) {
                window.jupiterService.setAuthToken(storedToken);
            }
            // Validate the stored token
            try {
                const isValid = await this.validateSession();
                if (isValid) {
                    this._isAuthenticated = true;
                    this.notifyAuthStateChange();
                    return; // Session restored, no need for auto-login
                } else {
                    await this.clearAuthToken();
                }
            } catch (error) {
                console.warn('AuthManager: Token validation failed:', error);
                await this.clearAuthToken();
            }
        }
        // Auto-login disabled for security (no password storage)
        this.isInitialized = true;
    }
}
