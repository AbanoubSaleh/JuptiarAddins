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
            let savedSettings = null;

            // Prefer OfficeRuntime.storage when available
            if (window.OfficeRuntime && OfficeRuntime.storage && typeof OfficeRuntime.storage.getItem === 'function') {
                try {
                    const settingsJson = await OfficeRuntime.storage.getItem('juptiarSettings');
                    if (settingsJson) {
                        savedSettings = JSON.parse(settingsJson);
                    }
                } catch (runtimeError) {
                    if (window.JupiterConfig?.debug?.enabled) {
                        console.warn('Runtime storage load failed:', runtimeError);
                    }
                }
            }

            // Fallback to Office.context.document.settings when available
            if (!savedSettings && typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings && typeof Office.context.document.settings.get === 'function') {
                try {
                    savedSettings = Office.context.document.settings.get('juptiarSettings');
                } catch (settingsError) {
                    if (window.JupiterConfig?.debug?.enabled) {
                        console.warn('Office settings load failed:', settingsError);
                    }
                }
            }

            if (savedSettings) {
                this.settings = { ...this.settings, ...savedSettings };
                if (window.jupiterService) {
                    if (!window.jupiterService.baseUrl || window.jupiterService.baseUrl === '') {
                        window.jupiterService.initialize(this.settings);
                    }
                }
            }
        } catch (error) {
            if (window.JupiterConfig?.debug?.enabled) {
                console.warn('Could not load settings:', error);
            }
        }
    }
    /**
     * Save settings to Office Runtime storage (more secure and persistent)
     */
    async saveSettings(newSettings) {
        try {
            this.settings = { ...this.settings, ...newSettings };

            let saved = false;
            if (window.OfficeRuntime && OfficeRuntime.storage && typeof OfficeRuntime.storage.setItem === 'function') {
                try {
                    await OfficeRuntime.storage.setItem('juptiarSettings', JSON.stringify(this.settings));
                    saved = true;
                } catch (runtimeError) {
                    if (window.JupiterConfig?.debug?.enabled) {
                        console.warn('Runtime storage save failed:', runtimeError);
                    }
                }
            }

            if (!saved && typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings && typeof Office.context.document.settings.set === 'function' && typeof Office.context.document.settings.saveAsync === 'function') {
                try {
                    Office.context.document.settings.set('juptiarSettings', this.settings);
                    await Office.context.document.settings.saveAsync();
                    saved = true;
                } catch (settingsError) {
                    if (window.JupiterConfig?.debug?.enabled) {
                        console.warn('Office settings save failed:', settingsError);
                    }
                }
            }

            // Update service configuration
            if (window.jupiterService) {
                window.jupiterService.initialize(this.settings);
            }
            return saved;
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
            if (window.OfficeRuntime && OfficeRuntime.storage && typeof OfficeRuntime.storage.removeItem === 'function') {
                try {
                    await OfficeRuntime.storage.removeItem('juptiarCredentials');
                } catch (runtimeError) {
                    // Ignore - likely doesn't exist
                }
            }
            // Also clear from Office settings (fallback/legacy)
            if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings && typeof Office.context.document.settings.remove === 'function' && typeof Office.context.document.settings.saveAsync === 'function') {
                try {
                    Office.context.document.settings.remove('juptiarCredentials');
                    await Office.context.document.settings.saveAsync();
                } catch (settingsError) {
                    // Ignore - likely doesn't exist
                }
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
            if (window.OfficeRuntime && OfficeRuntime.storage && typeof OfficeRuntime.storage.setItem === 'function') {
                await OfficeRuntime.storage.setItem('authToken', token);
            }
        } catch (error) {
            console.error('Could not store auth token:', error);
        }
    }
    /**
     * Load authentication token from Office Runtime storage
     */
    async loadAuthToken() {
        try {
            if (window.OfficeRuntime && OfficeRuntime.storage && typeof OfficeRuntime.storage.getItem === 'function') {
                const token = await OfficeRuntime.storage.getItem('authToken');
                if (token) {
                    return token;
                }
            }
        } catch (error) {
            if (window.JupiterConfig?.debug?.enabled) {
                console.warn('Could not load auth token:', error);
            }
        }
        return null;
    }
    /**
     * Clear authentication token from Office Runtime storage
     */
    async clearAuthToken() {
        try {
            if (window.OfficeRuntime && OfficeRuntime.storage && typeof OfficeRuntime.storage.removeItem === 'function') {
                await OfficeRuntime.storage.removeItem('authToken');
            }
        } catch (error) {
            if (window.JupiterConfig?.debug?.enabled) {
                console.warn('Could not clear auth token:', error);
            }
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
     * Decode JWT token and extract payload
     */
    decodeJWT(token) {
        try {
            if (!token) return null;

            const parts = token.split('.');
            if (parts.length !== 3) return null;

            // Decode the payload (second part)
            const payload = parts[1];
            const decoded = atob(payload.replace(/-/g, '+').replace(/_/g, '/'));
            return JSON.parse(decoded);
        } catch (error) {
            console.error('Error decoding JWT:', error);
            return null;
        }
    }

    /**
     * Get user role from current auth token
     */
    getUserRole() {
        if (!this.authToken) return null;

        const payload = this.decodeJWT(this.authToken);
        if (!payload) return null;

        // The role claim is stored as "http://schemas.microsoft.com/ws/2008/06/identity/claims/role"
        return payload['http://schemas.microsoft.com/ws/2008/06/identity/claims/role'] || payload.role || null;
    }

    /**
     * Check if current user is Admin
     */
    isAdmin() {
        const role = this.getUserRole();
        return role === 'Admin';
    }
    /**
     * Test connection to server
     */
    async testConnection(serverUrl, apiEndpoint = '/api') {
        try {
            // Create a temporary service instance for testing
            const tempService = new JupiterService();
            tempService.initialize({
                serverUrl: serverUrl,
                apiEndpoint: apiEndpoint,
                timeout: 10000 // Shorter timeout for testing
            });

            // Try the health endpoint first (no authentication required)
            try {
                const healthUrl = `${serverUrl}/health`;
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), 10000);

                const response = await fetch(healthUrl, {
                    method: 'GET',
                    signal: controller.signal
                });

                clearTimeout(timeoutId);

                if (response.ok) {
                    return { success: true, message: 'Connection successful - Server is healthy' };
                }

                // If health endpoint returns non-200, try API endpoint
                throw new Error('Health endpoint returned non-200 status');

            } catch (healthError) {
                // If health endpoint fails, try the API base endpoint
                try {
                    await tempService.makeRequest('GET', '');
                    return { success: true, message: 'Connection successful - API endpoint accessible' };
                } catch (apiError) {
                    // Network errors
                    if (healthError.message.includes('Failed to fetch') ||
                        healthError.message.includes('NetworkError') ||
                        healthError.name === 'AbortError') {
                        throw new Error('Cannot connect to server. Please check the Server URL and your network connection.');
                    }

                    // If we get 404, the endpoint doesn't exist
                    if (apiError.message.includes('404')) {
                        throw new Error('API endpoint not found (404). Please check the Server URL and API Endpoint.');
                    }

                    // Any other error from API endpoint, but server is reachable
                    // This means the server exists but the endpoint might be wrong
                    throw new Error('Server is reachable but API endpoint may be incorrect. Please verify the API Endpoint path.');
                }
            }
        } catch (error) {
            console.error('Test connection error:', error);
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
