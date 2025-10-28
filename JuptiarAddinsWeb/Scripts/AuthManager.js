/**
 * Authentication Manager - Handles user authentication and session management
 */

class AuthManager {
    constructor() {
        console.log('AuthManager: Constructor called');
        this.isAuthenticated = false;
        this.currentUser = null;
        this.authToken = null;
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
        
        this.loadSettings();
        this.loadStoredCredentials();
    }

    /**
     * Load settings from Office settings
     */
    async loadSettings() {
        try {
            await Office.context.document.settings.refreshAsync();

            const savedSettings = Office.context.document.settings.get('juptiarSettings');
            console.log('AuthManager: Loaded saved settings from Office:', savedSettings);
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
     * Save settings to Office settings
     */
    async saveSettings(newSettings) {
        try {
            this.settings = { ...this.settings, ...newSettings };
            
            Office.context.document.settings.set('juptiarSettings', this.settings);
            await Office.context.document.settings.saveAsync();
            
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
     * Load stored credentials if remember is enabled
     */
    loadStoredCredentials() {
        try {
            if (this.settings.rememberCredentials) {
                const storedCreds = Office.context.document.settings.get('juptiarCredentials');
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
     * Store credentials securely
     */
    async storeCredentials(username, password) {
        try {
            if (this.settings.rememberCredentials) {
                const credentials = {
                    username: username,
                    password: this.encryptPassword(password)
                };
                
                Office.context.document.settings.set('juptiarCredentials', credentials);
                await Office.context.document.settings.saveAsync();
            }
        } catch (error) {
            console.error('Could not store credentials:', error);
        }
    }

    /**
     * Clear stored credentials
     */
    async clearStoredCredentials() {
        try {
            Office.context.document.settings.remove('juptiarCredentials');
            await Office.context.document.settings.saveAsync();
            this.storedUsername = null;
            this.storedPassword = null;
        } catch (error) {
            console.error('Could not clear credentials:', error);
        }
    }

    /**
     * Simple encryption for stored passwords (base64 encoding)
     * Note: This is not secure encryption, just obfuscation
     */
    encryptPassword(password) {
        return btoa(password);
    }

    /**
     * Simple decryption for stored passwords
     */
    decryptPassword(encryptedPassword) {
        try {
            return atob(encryptedPassword);
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

                // Set token in JupiterService
                if (window.jupiterService) {
                    window.jupiterService.setAuthToken(response.token);
                    console.log('AuthManager: Token set in JupiterService');
                }

                // Store credentials if requested
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
            if (!this.isAuthenticated || !this.authToken) {
                return false;
            }

            if (!window.jupiterService) {
                return false;
            }

            const response = await window.jupiterService.validateSession();
            
            if (response.valid) {
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
        await this.loadSettings();
        
        // Try auto-login if enabled
        if (this.settings.autoLogin) {
            try {
                await this.autoLogin();
            } catch (error) {
                console.warn('Auto-login failed:', error);
            }
        }
        
        // Validate existing session
        if (this.isAuthenticated) {
            await this.validateSession();
        }
    }
}


