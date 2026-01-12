/**
 * Settings Page - Handles configuration and authentication settings
 */

class SettingsPage {
    constructor() {
        this.currentSettings = {};
        this.isTestingConnection = false;
        
        this.initializeEventListeners();
        this.loadCurrentSettings();
    }

    /**
     * Initialize event listeners
     */
    initializeEventListeners() {
        // Save settings
        DOMUtils.on('#saveBtn', 'click', () => this.saveSettings());



        // Test connection, login, and logout
        DOMUtils.on('#testConnectionBtn', 'click', () => this.testConnection());
        DOMUtils.on('#loginBtn', 'click', () => this.performLogin());
        DOMUtils.on('#logoutBtn', 'click', () => this.performLogout());

        // Form validation (for authentication fields)
        DOMUtils.on('#username', 'blur', () => this.validateUsername());
        DOMUtils.on('#password', 'blur', () => this.validatePassword());


    }

    /**
     * Load current settings from AuthManager
     */
    async loadCurrentSettings() {
        try {
            this.showLoading('Loading settings...');
            
            // Get current settings
            this.currentSettings = window.authManager.getSettings();
            const credentials = window.authManager.getStoredCredentials();
            const authStatus = window.authManager.getAuthStatus();
            
            // Populate form fields
            this.populateForm(this.currentSettings, credentials);
            
            // Update connection status
            this.updateConnectionStatus(authStatus);
            
            // Load default library options
            await this.loadDefaultLibraryOptions();
            
            this.hideLoading();
            
        } catch (error) {
            console.error('Error loading settings:', error);
            this.showError('Failed to load settings: ' + error.message);
        }
    }

    /**
     * Populate form with current settings
     */
    populateForm(settings, credentials) {
        // Connection settings are now pre-configured in JupiterConfig

        // Authentication (only populate if fields exist)
        if (DOMUtils.select('#username')) {
            DOMUtils.val('#username', credentials.username || '');
            const stayLoggedIn = settings.rememberCredentials || settings.autoLogin || false;
            DOMUtils.prop('#stayLoggedIn', 'checked', stayLoggedIn);
        }

        // Advanced settings
        DOMUtils.val('#defaultLibrary', settings.defaultLibrary || '');
        DOMUtils.val('#documentsPerPage', settings.documentsPerPage || 50);

        DOMUtils.prop('#enableNotifications', 'checked', settings.enableNotifications !== false);
    }

    /**
     * Update connection status display
     */
    updateConnectionStatus(authStatus) {
        const indicator = DOMUtils.select('#statusIndicator');
        const statusText = DOMUtils.select('#statusText');
        const details = DOMUtils.select('#connectionDetails');

        if (authStatus.isAuthenticated) {
            DOMUtils.removeClass(indicator, 'offline');
            DOMUtils.removeClass(indicator, 'testing');
            DOMUtils.addClass(indicator, 'online');
            DOMUtils.text(statusText, 'Connected');

            DOMUtils.text('#connectedServer', authStatus.serverUrl || '-');
            DOMUtils.text('#connectedUser', authStatus.user?.username || '-');
            DOMUtils.text('#lastConnected', new Date().toLocaleString());

            DOMUtils.show(details);

            // Show logout button, hide login button
            DOMUtils.hide('#loginBtn');
            DOMUtils.show('#logoutBtn');

            // Hide authentication form when logged in
            this.hideAuthenticationForm();
        } else {
            DOMUtils.removeClass(indicator, 'online');
            DOMUtils.removeClass(indicator, 'testing');
            DOMUtils.addClass(indicator, 'offline');
            DOMUtils.text(statusText, 'Not connected');
            DOMUtils.hide(details);

            // Show login button, hide logout button
            DOMUtils.show('#loginBtn');
            DOMUtils.hide('#logoutBtn');

            // Show authentication form when not logged in
            this.showAuthenticationForm();
        }
    }

    /**
     * Load default library options
     */
    async loadDefaultLibraryOptions() {
        try {
            if (window.authManager.getAuthStatus().isAuthenticated) {
                const libraries = await window.jupiterService.getLibraryTree();
                this.populateDefaultLibrarySelect(libraries);
            }
        } catch (error) {
            console.warn('Could not load libraries for default selection:', error);
        }
    }

    /**
     * Populate default library select
     */
    populateDefaultLibrarySelect(libraries) {
        const select = DOMUtils.select('#defaultLibrary');
        const currentValue = DOMUtils.val('#defaultLibrary');

        DOMUtils.empty(select);
        DOMUtils.append(select, '<option value="">Select default library...</option>');

        if (Array.isArray(libraries)) {
            libraries.forEach(library => {
                DOMUtils.append(select, `<option value="${library.id}">${library.name}</option>`);
            });
        }

        // Restore previous selection
        if (currentValue) {
            DOMUtils.val('#defaultLibrary', currentValue);
        }
    }

    /**
     * Validate server URL (now pre-configured, always valid)
     */
    validateServerUrl() {
        // Server URL is pre-configured in JupiterConfig, so always return true
        return true;
    }

    /**
     * Validate username
     */
    validateUsername() {
        const usernameField = DOMUtils.select('#username');
        if (!usernameField) return true; // Field doesn't exist, skip validation

        const username = DOMUtils.val('#username').trim();

        if (!username) {
            DOMUtils.addClass(usernameField, 'error');
            this.showFieldError(usernameField, 'Username is required');
            return false;
        }

        DOMUtils.removeClass(usernameField, 'error');
        this.hideFieldError(usernameField);
        return true;
    }

    /**
     * Validate password
     */
    validatePassword() {
        const passwordField = DOMUtils.select('#password');
        if (!passwordField) return true; // Field doesn't exist, skip validation

        const password = DOMUtils.val('#password');

        if (!password) {
            DOMUtils.addClass(passwordField, 'error');
            this.showFieldError(passwordField, 'Password is required');
            return false;
        }

        DOMUtils.removeClass(passwordField, 'error');
        this.hideFieldError(passwordField);
        return true;
    }



    /**
     * Show field-specific error
     */
    showFieldError(field, message) {
        // Remove existing error messages
        const existingErrors = field.parentNode.querySelectorAll('.error-message');
        existingErrors.forEach(error => error.remove());

        // Add new error message
        const errorSpan = document.createElement('span');
        errorSpan.className = 'error-message';
        errorSpan.textContent = message;
        field.parentNode.insertBefore(errorSpan, field.nextSibling);
    }

    /**
     * Hide field-specific error
     */
    hideFieldError(field) {
        const existingErrors = field.parentNode.querySelectorAll('.error-message');
        existingErrors.forEach(error => error.remove());
    }





    /**
     * Test connection to server
     */
    async testConnection() {
        if (this.isTestingConnection) return;

        try {
            this.isTestingConnection = true;

            // Get connection settings from JupiterConfig instead of form fields
            const serverUrl = window.JupiterConfig.get('server.baseUrl');
            const apiEndpoint = window.JupiterConfig.get('server.apiEndpoint') || '/api';

            // Update UI
            DOMUtils.removeClass('#statusIndicator', 'online');
            DOMUtils.removeClass('#statusIndicator', 'offline');
            DOMUtils.addClass('#statusIndicator', 'testing');
            DOMUtils.text('#statusText', 'Testing connection...');
            DOMUtils.prop('#testConnectionBtn', 'disabled', true);
            this.showLoading('Testing connection...');

            // Test connection
            const result = await window.authManager.testConnection(serverUrl, apiEndpoint);

            if (result.success) {
                DOMUtils.removeClass('#statusIndicator', 'testing');
                DOMUtils.removeClass('#statusIndicator', 'offline');
                DOMUtils.addClass('#statusIndicator', 'online');
                DOMUtils.text('#statusText', 'Connection successful');
                this.showSuccess('Connection test successful');
            } else {
                DOMUtils.removeClass('#statusIndicator', 'testing');
                DOMUtils.removeClass('#statusIndicator', 'online');
                DOMUtils.addClass('#statusIndicator', 'offline');
                DOMUtils.text('#statusText', 'Connection failed');
                this.showError('Connection test failed: ' + result.error);
            }

        } catch (error) {
            console.error('Connection test error:', error);
            DOMUtils.removeClass('#statusIndicator', 'testing');
            DOMUtils.removeClass('#statusIndicator', 'online');
            DOMUtils.addClass('#statusIndicator', 'offline');
            DOMUtils.text('#statusText', 'Connection failed');
            this.showError('Connection test failed: ' + error.message);
        } finally {
            this.isTestingConnection = false;
            DOMUtils.prop('#testConnectionBtn', 'disabled', false);
            this.hideLoading();
        }
    }

    /**
     * Perform login with current credentials
     */
    async performLogin() {
        if (this.isLoggingIn) return;

        try {
            this.isLoggingIn = true;

            // Show authentication form if not visible
            this.showAuthenticationForm();

            // Check if we have credentials to validate
            const usernameField = DOMUtils.select('#username');
            const passwordField = DOMUtils.select('#password');

            if (!usernameField || !passwordField) {
                this.showError('Authentication form not available. Please refresh the page.');
                return;
            }

            // Validate credentials
            if (!this.validateUsername() || !this.validatePassword()) {
                this.showError('Please enter valid username and password');
                return;
            }

            const username = DOMUtils.val('#username').trim();
            const password = DOMUtils.val('#password');
            const stayLoggedIn = DOMUtils.prop('#stayLoggedIn', 'checked');

            // Update UI
            DOMUtils.prop('#loginBtn', 'disabled', true);
            this.showLoading('Logging in...');

            // Perform login
            const result = await window.authManager.login(username, password, stayLoggedIn);

            if (result.success) {
                this.showSuccess('Login successful! You are now authenticated.');

                // Update connection status
                const authStatus = window.authManager.getAuthStatus();
                this.updateConnectionStatus(authStatus);

                // Hide authentication form
                this.hideAuthenticationForm();

                // Load library options if authenticated
                await this.loadDefaultLibraryOptions();

                // Clear password field for security (unless stay logged in is checked)
                if (!stayLoggedIn) {
                    DOMUtils.val('#password', '');
                }
            } else {
                this.showError('Login failed: ' + (result.message || 'Invalid credentials'));
            }

        } catch (error) {
            console.error('Login error:', error);
            this.showError('Login failed: ' + error.message);
        } finally {
            this.isLoggingIn = false;
            DOMUtils.prop('#loginBtn', 'disabled', false);
            this.hideLoading();
        }
    }

    /**
     * Perform logout
     */
    async performLogout() {
        try {
            this.showLoading('Logging out...');

            // Perform logout
            await window.authManager.logout();

            // Update UI
            this.updateConnectionStatus({ isAuthenticated: false });
            DOMUtils.hide('#logoutBtn');
            DOMUtils.show('#loginBtn');

            // Show authentication form for re-login
            this.showAuthenticationForm();

            // Clear password field for security
            if (DOMUtils.select('#password')) {
                DOMUtils.val('#password', '');
            }

            this.showSuccess('Logged out successfully');

        } catch (error) {
            console.error('Logout error:', error);
            this.showError('Logout failed: ' + error.message);
        } finally {
            this.hideLoading();
        }
    }

    /**
     * Show authentication form when user needs to log in
     */
    showAuthenticationForm() {
        const authSection = DOMUtils.select('#authSection');
        if (authSection) {
            DOMUtils.show('#authSection');
        }
    }

    /**
     * Hide authentication form when user is logged in
     */
    hideAuthenticationForm() {
        const authSection = DOMUtils.select('#authSection');
        if (authSection) {
            DOMUtils.hide('#authSection');
        }
    }

    /**
     * Save settings and optionally authenticate
     */
    async saveSettings() {
        try {
            this.showLoading('Saving settings...');

            // Collect form data
            const stayLoggedIn = DOMUtils.select('#stayLoggedIn') ? DOMUtils.prop('#stayLoggedIn', 'checked') : false;
            const newSettings = {
                // Connection settings are now pre-configured in JupiterConfig
                serverUrl: window.JupiterConfig.get('server.baseUrl'),
                apiEndpoint: window.JupiterConfig.get('server.apiEndpoint') || '/api',
                timeout: window.JupiterConfig.get('server.timeout') || 30000,
                // Authentication settings
                rememberCredentials: stayLoggedIn,
                autoLogin: stayLoggedIn,
                defaultLibrary: DOMUtils.val('#defaultLibrary'),
                documentsPerPage: parseInt(DOMUtils.val('#documentsPerPage')),

                enableNotifications: DOMUtils.prop('#enableNotifications', 'checked')
            };

            // Save settings
            await window.authManager.saveSettings(newSettings);

            this.showSuccess('Settings saved successfully.');
            this.currentSettings = newSettings;

        } catch (error) {
            console.error('Error saving settings:', error);
            this.showError('Failed to save settings: ' + error.message);
        } finally {
            this.hideLoading();
        }
    }







    /**
     * Show loading indicator
     */
    showLoading(message = 'Loading...') {
        DOMUtils.text('#loadingText', message);
        DOMUtils.show('#loadingSection');
        DOMUtils.hide('#messageSection');
    }

    /**
     * Hide loading indicator
     */
    hideLoading() {
        DOMUtils.hide('#loadingSection');
    }

    /**
     * Show error message
     */
    showError(message) {
        const messageBar = DOMUtils.select('#messageBar');
        const messageIcon = DOMUtils.select('#messageIcon i');

        if (messageBar) {
            messageBar.className = 'ms-MessageBar ms-MessageBar--error';
        }
        if (messageIcon) {
            messageIcon.className = 'ms-Icon ms-Icon--ErrorBadge';
        }
        DOMUtils.text('#messageText', message);
        DOMUtils.show('#messageSection');
        DOMUtils.hide('#loadingSection');
    }

    /**
     * Show success message
     */
    showSuccess(message) {
        const messageBar = DOMUtils.select('#messageBar');
        const messageIcon = DOMUtils.select('#messageIcon i');

        if (messageBar) {
            messageBar.className = 'ms-MessageBar ms-MessageBar--success';
        }
        if (messageIcon) {
            messageIcon.className = 'ms-Icon ms-Icon--CheckMark';
        }
        DOMUtils.text('#messageText', message);
        DOMUtils.show('#messageSection');
        DOMUtils.hide('#loadingSection');

        // Auto-hide success message after 3 seconds
        setTimeout(() => {
            DOMUtils.hide('#messageSection');
        }, 3000);
    }

    /**
     * Show warning message
     */
    showWarning(message) {
        const messageBar = DOMUtils.select('#messageBar');
        if (messageBar) {
            messageBar.className = 'ms-MessageBar ms-MessageBar--warning';
        }
        const messageIcon = DOMUtils.select('#messageIcon i');
        if (messageIcon) {
            messageIcon.className = 'ms-Icon ms-Icon--Warning';
        }
        DOMUtils.text('#messageText', message);
        DOMUtils.show('#messageSection');
        DOMUtils.hide('#loadingSection');
    }

    /**
     * Show info message
     */
    showInfo(message) {
        const messageBar = DOMUtils.select('#messageBar');
        if (messageBar) {
            messageBar.className = 'ms-MessageBar ms-MessageBar--info';
        }
        const messageIcon = DOMUtils.select('#messageIcon i');
        if (messageIcon) {
            messageIcon.className = 'ms-Icon ms-Icon--Info';
        }
        DOMUtils.text('#messageText', message);
        DOMUtils.show('#messageSection');
        DOMUtils.hide('#loadingSection');
    }

    /**
     * Hide message
     */
    hideMessage() {
        DOMUtils.hide('#messageSection');
    }
}

// Initialize when Office is ready (only if Office is available)
if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(() => {
        // Initialize JupiterConfig first (this might not have been called)
        if (typeof window.JupiterConfig.init === 'function') {
            window.JupiterConfig.init();
        }

        // Initialize global services if not already done
        if (!window.jupiterService) {
            window.jupiterService = new JupiterService();
            window.jupiterService.initialize({
                serverUrl: window.JupiterConfig.get('server.baseUrl'),
                apiEndpoint: window.JupiterConfig.get('server.apiEndpoint') || '/api',
                timeout: window.JupiterConfig.get('server.timeout') || 30000
            });
        }

        if (!window.authManager) {
            window.authManager = new AuthManager();
            // Initialize AuthManager asynchronously and then create SettingsPage
            window.authManager.initialize().then(() => {
                console.log('AuthManager initialized, creating SettingsPage...');
                if (!window.settingsPage) {
                    window.settingsPage = new SettingsPage();
                }
            }).catch(error => {
                console.error('Failed to initialize AuthManager:', error);
                // Still create SettingsPage even if AuthManager fails
                if (!window.settingsPage) {
                    window.settingsPage = new SettingsPage();
                }
            });
        } else {
            // AuthManager already exists, create SettingsPage immediately
            if (!window.settingsPage) {
                window.settingsPage = new SettingsPage();
            }
        }
    });
} else {
    console.warn('Settings.js: Office.js not available, skipping Office.onReady initialization');
}
