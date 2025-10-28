/**
 * Authentication Manager - Handles user authentication and session management
 */

class AuthManager {
    constructor() {
        console.log('AuthManager: Constructor called');
        this.isAuthenticated = false;
        this.currentUser = null;
        this.authToken = null;
        this.isInitialized = false;
        // Initialize settings from configuration
        const configBaseUrl = window.JupiterConfig?.get('server.baseUrl');
        console.log('AuthManager: Reading config baseUrl:', configBaseUrl);
        console.log('AuthManager: JupiterConfig object:', window.JupiterConfig);

        this.settings = {
            serverUrl: configBaseUrl || 'https://localhost:7001',
            apiEndpoint: window.JupiterConfig?.get('server.apiEndpoint') || '/api',
            timeout: window.JupiterConfig?.get('server.timeout') || 30000,
            rememberCredentials: window.JupiterConfig?.get('auth.rememberCredentials') || false,
            autoLogin: window.JupiterConfig?.get('auth.autoLogin') || false
        };

        console.log('AuthManager: Initial settings:', this.settings);

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
                    console.log('AuthManager: Loaded saved settings from OfficeRuntime storage:', savedSettings);
                }
            } catch (runtimeError) {
                console.warn('Could not load from OfficeRuntime storage, trying Office settings:', runtimeError);
                // Fallback to Office.context.document.settings for backward compatibility
                await Office.context.document.settings.refreshAsync();
                savedSettings = Office.context.document.settings.get('juptiarSettings');
                console.log('AuthManager: Loaded saved settings from Office settings (fallback):', savedSettings);
            }

            if (savedSettings) {
                this.settings = { ...this.settings, ...savedSettings };

                // Initialize service with settings only if it doesn't have a baseUrl
                if (window.jupiterService) {
                    console.log('AuthManager: Current JupiterService baseUrl:', window.jupiterService.baseUrl);
                    console.log('AuthManager: Saved settings:', this.settings);

                    // Only re-initialize if the service doesn't have a baseUrl or if our settings are different
                    if (!window.jupiterService.baseUrl || window.jupiterService.baseUrl === '') {
                        console.log('AuthManager: Re-initializing JupiterService with saved settings:', this.settings);
                        window.jupiterService.initialize(this.settings);
                        console.log('AuthManager: JupiterService baseUrl after re-init:', window.jupiterService.baseUrl);
                    } else {
                        console.log('AuthManager: JupiterService already has baseUrl, not re-initializing');
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
                console.log('AuthManager: Settings saved to OfficeRuntime storage');
            } catch (runtimeError) {
                console.warn('Could not save to OfficeRuntime storage, using Office settings fallback:', runtimeError);
                // Fallback to Office.context.document.settings
                Office.context.document.settings.set('juptiarSettings', this.settings);
                await Office.context.document.settings.saveAsync();
                console.log('AuthManager: Settings saved to Office settings (fallback)');
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
     */
    async loadStoredCredentials() {
        try {
            if (this.settings.rememberCredentials) {
                let storedCreds = null;

                // Try Office Runtime storage first (more secure)
                try {
                    const credsJson = await OfficeRuntime.storage.getItem('juptiarCredentials');
                    if (credsJson) {
                        storedCreds = JSON.parse(credsJson);
                        console.log('AuthManager: Loaded credentials from OfficeRuntime storage');
                    }
                } catch (runtimeError) {
                    console.warn('Could not load credentials from OfficeRuntime storage, trying Office settings:', runtimeError);
                    // Fallback to Office.context.document.settings
                    storedCreds = Office.context.document.settings.get('juptiarCredentials');
                    console.log('AuthManager: Loaded credentials from Office settings (fallback)');
                }

                if (storedCreds) {
                    this.storedUsername = storedCreds.username;
                    this.storedPassword = this.decryptPassword(storedCreds.password);
                }
            }
        } catch (error) {
            console.warn('Could not load stored credentials:', error);
        }
    }

    /**
     * Store credentials securely using Office Runtime storage
     */
    async storeCredentials(username, password) {
        try {
            if (this.settings.rememberCredentials) {
                const credentials = {
                    username: username,
                    password: this.encryptPassword(password)
                };

                // Store in Office Runtime storage (primary method - more secure)
                try {
                    await OfficeRuntime.storage.setItem('juptiarCredentials', JSON.stringify(credentials));
                    console.log('AuthManager: Credentials stored in OfficeRuntime storage');
                } catch (runtimeError) {
                    console.warn('Could not store credentials in OfficeRuntime storage, using Office settings fallback:', runtimeError);
                    // Fallback to Office.context.document.settings
                    Office.context.document.settings.set('juptiarCredentials', credentials);
                    await Office.context.document.settings.saveAsync();
                    console.log('AuthManager: Credentials stored in Office settings (fallback)');
                }
            }
        } catch (error) {
            console.error('Could not store credentials:', error);
        }
    }

    /**
     * Clear stored credentials from both storage methods
     */
    async clearStoredCredentials() {
        try {
            // Clear from Office Runtime storage
            try {
                await OfficeRuntime.storage.removeItem('juptiarCredentials');
                console.log('AuthManager: Credentials cleared from OfficeRuntime storage');
            } catch (runtimeError) {
                console.warn('Could not clear from OfficeRuntime storage:', runtimeError);
            }

            // Also clear from Office settings (fallback/legacy)
            try {
                Office.context.document.settings.remove('juptiarCredentials');
                await Office.context.document.settings.saveAsync();
                console.log('AuthManager: Credentials cleared from Office settings');
            } catch (settingsError) {
                console.warn('Could not clear from Office settings:', settingsError);
            }

            this.storedUsername = null;
            this.storedPassword = null;
        } catch (error) {
            console.error('Could not clear credentials:', error);
        }
    }

    /**
     * Store authentication token securely in Office Runtime storage
     */
    async storeAuthToken(token) {
        try {
            await OfficeRuntime.storage.setItem('authToken', token);
            console.log('AuthManager: Auth token stored securely in OfficeRuntime storage');
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
                console.log('AuthManager: Auth token loaded from OfficeRuntime storage');
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
            console.log('AuthManager: Auth token cleared from OfficeRuntime storage');
        } catch (error) {
            console.warn('Could not clear auth token:', error);
        }
    }

    /**
     * Simple encryption for stored passwords (base64 encoding)
     * Note: This is not secure encryption, just obfuscation for convenience
     * For production, consider using Web Crypto API or avoid storing passwords entirely
     */
    encryptPassword(password) {
        // Add a simple salt to make it slightly less obvious
        const salt = 'JupiterDMS2024';
        return btoa(salt + password + salt);
    }

    /**
     * Simple decryption for stored passwords
     */
    decryptPassword(encryptedPassword) {
        try {
            const salt = 'JupiterDMS2024';
            const decoded = atob(encryptedPassword);
            // Remove salt from both ends
            return decoded.substring(salt.length, decoded.length - salt.length);
        } catch (error) {
            return '';
        }
    }

    /**
     * Attempt login with credentials
     */
    async login(username, password, rememberCredentials = false) {
        try {
            console.log('AuthManager: Attempting login for user:', username);
            console.log('AuthManager: Using server URL:', this.settings.serverUrl);

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
                console.log('AuthManager: JupiterService baseUrl is empty, re-initializing...');

                // Force re-initialization with correct settings
                window.jupiterService.initialize({
                    serverUrl: this.settings.serverUrl,
                    apiEndpoint: this.settings.apiEndpoint,
                    timeout: this.settings.timeout
                });

                console.log('AuthManager: JupiterService re-initialized with baseUrl:', window.jupiterService.baseUrl);
            }

            // Attempt login
            console.log('AuthManager: Calling jupiterService.login...');
            const response = await window.jupiterService.login(username, password);
            console.log('AuthManager: Login response:', response);
            
            if (response.success && response.token) {
                this.isAuthenticated = true;
                this.authToken = response.token;
                this.currentUser = response.user || { username: username };

                // Store authentication token securely in Office Runtime storage
                await this.storeAuthToken(response.token);

                // Set token in JupiterService
                if (window.jupiterService) {
                    window.jupiterService.setAuthToken(response.token);
                    console.log('AuthManager: Token set in JupiterService');
                }

                // Store credentials if requested (for auto-login)
                // Note: In production, consider storing only tokens for better security
                if (rememberCredentials) {
                    await this.storeCredentials(username, password);
                }

                // Update settings
                await this.saveSettings({ rememberCredentials: rememberCredentials });

                this.notifyAuthStateChange();
                return { success: true, user: this.currentUser };
            } else {
                throw new Error(response.message || 'Login failed');
            }
        } catch (error) {
            this.isAuthenticated = false;
            this.authToken = null;
            this.currentUser = null;
            
            if (window.jupiterService) {
                window.jupiterService.clearAuthToken();
            }
            
            throw error;
        }
    }

    /**
     * Logout current user
     */
    async logout() {
        try {
            if (window.jupiterService && this.isAuthenticated) {
                await window.jupiterService.logout();
            }
        } catch (error) {
            console.warn('Logout request failed:', error);
        } finally {
            this.isAuthenticated = false;
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
     * Auto-login if enabled and credentials are stored
     */
    async autoLogin() {
        try {
            if (this.settings.autoLogin && this.storedUsername && this.storedPassword) {
                return await this.login(this.storedUsername, this.storedPassword, true);
            }
            return false;
        } catch (error) {
            console.warn('Auto-login failed:', error);
            return false;
        }
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
     * Get current authentication status
     */
    getAuthStatus() {
        return {
            isAuthenticated: this.isAuthenticated,
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
     */
    getStoredCredentials() {
        return {
            username: this.storedUsername || '',
            hasStoredPassword: !!this.storedPassword
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
        console.log('AuthManager: Initializing...');

        // Load settings first
        await this.loadSettings();

        // Load stored credentials
        await this.loadStoredCredentials();

        // Try to restore previous session from stored token
        const storedToken = await this.loadAuthToken();
        if (storedToken) {
            console.log('AuthManager: Found stored auth token, attempting to restore session');
            this.authToken = storedToken;

            // Set token in JupiterService
            if (window.jupiterService) {
                window.jupiterService.setAuthToken(storedToken);
            }

            // Validate the stored token
            try {
                const isValid = await this.validateSession();
                if (isValid) {
                    this.isAuthenticated = true;
                    console.log('AuthManager: Session restored successfully from stored token');
                    this.notifyAuthStateChange();
                    return; // Session restored, no need for auto-login
                } else {
                    console.log('AuthManager: Stored token is invalid, clearing it');
                    await this.clearAuthToken();
                }
            } catch (error) {
                console.warn('AuthManager: Token validation failed:', error);
                await this.clearAuthToken();
            }
        }

        // Try auto-login if enabled and no valid session was restored
        if (this.settings.autoLogin && !this.isAuthenticated) {
            try {
                console.log('AuthManager: Attempting auto-login');
                await this.autoLogin();
            } catch (error) {
                console.warn('Auto-login failed:', error);
            }
        }

        console.log('AuthManager: Initialization complete. Authenticated:', this.isAuthenticated);
        this.isInitialized = true;
    }
}


