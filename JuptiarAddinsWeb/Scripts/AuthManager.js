/**
 * Authentication Manager - Handles user authentication and session management
 */

class AuthManager {
    constructor() {
        this.isAuthenticated = false;
        this.currentUser = null;
        this.authToken = null;
        this.settings = {
            serverUrl: '',
            apiEndpoint: '/api/v1',
            timeout: 30000,
            rememberCredentials: false,
            autoLogin: false
        };
        
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
            if (savedSettings) {
                this.settings = { ...this.settings, ...savedSettings };
                
                // Initialize service with settings
                if (window.juptiarService) {
                    window.juptiarService.initialize(this.settings);
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
            if (window.juptiarService) {
                window.juptiarService.initialize(this.settings);
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
            if (!window.juptiarService) {
                throw new Error('Juptiar service not initialized');
            }

            // Validate inputs
            if (!username || !password) {
                throw new Error('Username and password are required');
            }

            if (!this.settings.serverUrl) {
                throw new Error('Server URL not configured. Please check settings.');
            }

            // Attempt login
            const response = await window.juptiarService.login(username, password);
            
            if (response.success && response.token) {
                this.isAuthenticated = true;
                this.authToken = response.token;
                this.currentUser = response.user || { username: username };
                
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
            
            if (window.juptiarService) {
                window.juptiarService.clearAuthToken();
            }
            
            throw error;
        }
    }

    /**
     * Logout current user
     */
    async logout() {
        try {
            if (window.juptiarService && this.isAuthenticated) {
                await window.juptiarService.logout();
            }
        } catch (error) {
            console.warn('Logout request failed:', error);
        } finally {
            this.isAuthenticated = false;
            this.authToken = null;
            this.currentUser = null;
            
            if (window.juptiarService) {
                window.juptiarService.clearAuthToken();
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

            if (!window.juptiarService) {
                return false;
            }

            const response = await window.juptiarService.validateSession();
            
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
    async testConnection(serverUrl, apiEndpoint = '/api/v1') {
        try {
            const tempService = new JuptiarService();
            tempService.initialize({
                serverUrl: serverUrl,
                apiEndpoint: apiEndpoint,
                timeout: 10000 // Shorter timeout for testing
            });

            // Try to make a simple request (like getting server info)
            const response = await tempService.makeRequest('GET', '/health');
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

// Create global instance
window.authManager = new AuthManager();

// Initialize when Office is ready
Office.onReady(() => {
    window.authManager.initialize();
});
