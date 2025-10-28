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

        // Reset to defaults
        DOMUtils.on('#resetBtn', 'click', () => this.resetToDefaults());

        // Clear credentials
        DOMUtils.on('#clearCredentialsBtn', 'click', () => this.clearCredentials());

        // Test connection, login, and logout
        DOMUtils.on('#testConnectionBtn', 'click', () => this.testConnection());
        DOMUtils.on('#loginBtn', 'click', () => this.performLogin());
        DOMUtils.on('#logoutBtn', 'click', () => this.performLogout());

        // Form validation
        DOMUtils.on('#serverUrl', 'blur', () => this.validateServerUrl());
        DOMUtils.on('#username', 'blur', () => this.validateUsername());
        DOMUtils.on('#password', 'blur', () => this.validatePassword());

        // Auto-login checkbox dependency
        DOMUtils.on('#autoLogin', 'change', () => this.handleAutoLoginChange());
        DOMUtils.on('#rememberCredentials', 'change', () => this.handleRememberCredentialsChange());
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
        // Connection settings
        DOMUtils.val('#serverUrl', settings.serverUrl || '');
        DOMUtils.val('#apiEndpoint', settings.apiEndpoint || '/api/v1');
        DOMUtils.val('#connectionTimeout', settings.timeout / 1000 || 30);

        // Authentication
        DOMUtils.val('#username', credentials.username || '');
        // Map old settings to new "Stay Logged In" option
        const stayLoggedIn = settings.rememberCredentials || settings.autoLogin || false;
        DOMUtils.prop('#stayLoggedIn', 'checked', stayLoggedIn);

        // Advanced settings
        DOMUtils.val('#defaultLibrary', settings.defaultLibrary || '');
        DOMUtils.val('#documentsPerPage', settings.documentsPerPage || 50);

        DOMUtils.prop('#enableNotifications', 'checked', settings.enableNotifications !== false);
        
        // Handle auto-login dependency
        this.handleRememberCredentialsChange();
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
        } else {
            DOMUtils.removeClass(indicator, 'online');
            DOMUtils.removeClass(indicator, 'testing');
            DOMUtils.addClass(indicator, 'offline');
            DOMUtils.text(statusText, 'Not connected');
            DOMUtils.hide(details);

            // Show login button, hide logout button
            DOMUtils.show('#loginBtn');
            DOMUtils.hide('#logoutBtn');
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
        // Server URL is pre-configured, so always return true
        const field = DOMUtils.select('#serverUrl');
        DOMUtils.removeClass(field, 'error');
        this.hideFieldError(field);
        return true;
    }

    /**
     * Validate username
     */
    validateUsername() {
        const username = DOMUtils.val('#username').trim();
        const field = DOMUtils.select('#username');

        if (!username) {
            DOMUtils.addClass(field, 'error');
            this.showFieldError(field, 'Username is required');
            return false;
        }

        DOMUtils.removeClass(field, 'error');
        this.hideFieldError(field);
        return true;
    }

    /**
     * Validate password
     */
    validatePassword() {
        const password = DOMUtils.val('#password');
        const field = DOMUtils.select('#password');

        if (!password) {
            DOMUtils.addClass(field, 'error');
            this.showFieldError(field, 'Password is required');
            return false;
        }

        DOMUtils.removeClass(field, 'error');
        this.hideFieldError(field);
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
     * Handle remember credentials checkbox change
     */
    handleRememberCredentialsChange() {
        const rememberChecked = DOMUtils.prop('#rememberCredentials', 'checked');
        DOMUtils.prop('#autoLogin', 'disabled', !rememberChecked);

        if (!rememberChecked) {
            DOMUtils.prop('#autoLogin', 'checked', false);
        }
    }

    /**
     * Handle auto-login checkbox change
     */
    handleAutoLoginChange() {
        const autoLoginChecked = DOMUtils.prop('#autoLogin', 'checked');

        if (autoLoginChecked && !DOMUtils.prop('#rememberCredentials', 'checked')) {
            DOMUtils.prop('#rememberCredentials', 'checked', true);
        }
    }

    /**
     * Test connection to server
     */
    async testConnection() {
        if (this.isTestingConnection) return;
        
        try {
            this.isTestingConnection = true;
            
            // Validate required fields
            if (!this.validateServerUrl()) {
                return;
            }
            
            const serverUrl = DOMUtils.val('#serverUrl').trim();
            const apiEndpoint = DOMUtils.val('#apiEndpoint').trim();

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

            // Clear password field for security
            DOMUtils.val('#password', '');

            this.showSuccess('Logged out successfully');

        } catch (error) {
            console.error('Logout error:', error);
            this.showError('Logout failed: ' + error.message);
        } finally {
            this.hideLoading();
        }
    }

    /**
     * Save settings with credential validation (Best Practice)
     */
    async saveSettings() {
        try {
            // Validate form
            const isValid = this.validateServerUrl() && this.validateUsername() && this.validatePassword();
            if (!isValid) {
                this.showError('Please fix validation errors before saving');
                return;
            }

            this.showLoading('Validating credentials and saving settings...');

            // Collect form data
            const stayLoggedIn = DOMUtils.prop('#stayLoggedIn', 'checked');
            const newSettings = {
                serverUrl: DOMUtils.val('#serverUrl').trim(),
                apiEndpoint: DOMUtils.val('#apiEndpoint').trim(),
                timeout: parseInt(DOMUtils.val('#connectionTimeout')) * 1000,
                // Map new "Stay Logged In" to both old options for backward compatibility
                rememberCredentials: stayLoggedIn,
                autoLogin: stayLoggedIn,
                defaultLibrary: DOMUtils.val('#defaultLibrary'),
                documentsPerPage: parseInt(DOMUtils.val('#documentsPerPage')),

                enableNotifications: DOMUtils.prop('#enableNotifications', 'checked')
            };

            const username = DOMUtils.val('#username').trim();
            const password = DOMUtils.val('#password');

            // SECURITY BEST PRACTICE: Validate credentials before saving anything
            if (username && password) {
                console.log('Settings: Validating credentials before saving...');

                // First save settings so AuthManager can use the new server URL
                await window.authManager.saveSettings(newSettings);

                // Attempt login to validate credentials
                const loginResult = await window.authManager.login(username, password, stayLoggedIn);

                if (loginResult.success) {
                    // ✅ Credentials are valid
                    console.log('Settings: Credentials validated successfully');

                    // Update connection status
                    const authStatus = window.authManager.getAuthStatus();
                    this.updateConnectionStatus(authStatus);

                    // Load library options if authenticated
                    await this.loadDefaultLibraryOptions();

                    // Clear password field for security (token is now stored)
                    DOMUtils.val('#password', '');

                    this.showSuccess('✅ Settings saved and login successful! You are now authenticated.');

                } else {
                    // ❌ Invalid credentials
                    console.log('Settings: Invalid credentials provided');

                    // Still save settings (server URL, etc.) but don't store credentials
                    await window.authManager.clearStoredCredentials();

                    this.showError('❌ Settings saved, but login failed: ' + (loginResult.message || 'Invalid username or password'));
                }
            } else {
                // No credentials provided, just save settings
                await window.authManager.saveSettings(newSettings);

                // Clear any stored credentials if stay logged in is unchecked
                if (!stayLoggedIn) {
                    await window.authManager.clearStoredCredentials();
                }

                this.showSuccess('Settings saved successfully. Enter credentials to login.');
            }

            this.currentSettings = newSettings;

        } catch (error) {
            console.error('Error saving settings:', error);
            this.showError('Failed to save settings: ' + error.message);
        } finally {
            this.hideLoading();
        }
    }

    /**
     * Reset settings to defaults
     */
    resetToDefaults() {
        // Use a custom confirmation dialog since Office Add-ins don't support window.confirm()
        this.showConfirmation(
            'Reset Settings',
            'Are you sure you want to reset all settings to defaults? This action cannot be undone.',
            () => this.performReset()
        );
    }

    /**
     * Perform the actual reset after confirmation
     */
    performReset() {
        // Reset form to default values (but preserve server URL as it's pre-configured)
        // Get default values from config
        const defaultServerUrl = window.JupiterConfig?.get('server.baseUrl') || 'https://localhost:7001';
        const defaultApiEndpoint = window.JupiterConfig?.get('server.apiEndpoint') || '/api';

        // Only reset server URL if it's empty, otherwise keep the current value
        if (!DOMUtils.val('#serverUrl').trim()) {
            DOMUtils.val('#serverUrl', defaultServerUrl);
        }

        DOMUtils.val('#apiEndpoint', defaultApiEndpoint);
        DOMUtils.val('#connectionTimeout', '30');
        DOMUtils.val('#username', '');
        DOMUtils.val('#password', '');
        DOMUtils.prop('#stayLoggedIn', 'checked', false);
        DOMUtils.val('#defaultLibrary', '');
        DOMUtils.val('#documentsPerPage', '50');

        DOMUtils.prop('#enableNotifications', 'checked', true);

        // Clear validation errors
        DOMUtils.removeClass('.ms-TextField-field', 'error');
        const errorMessages = DOMUtils.selectAll('.error-message');
        errorMessages.forEach(el => el.remove());

        this.handleRememberCredentialsChange();
        this.showSuccess('Settings reset to defaults (server URL preserved)');
    }

    /**
     * Show a custom confirmation dialog (Office Add-ins don't support window.confirm)
     */
    showConfirmation(title, message, onConfirm) {
        // Create a simple confirmation using the existing message system
        const confirmHtml = `
            <div class="confirmation-dialog" style="background: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; margin: 10px 0; border-radius: 4px;">
                <h4 style="margin: 0 0 10px 0; color: #856404;">${title}</h4>
                <p style="margin: 0 0 15px 0; color: #856404;">${message}</p>
                <div style="text-align: right;">
                    <button id="confirmYes" class="ms-Button ms-Button--primary" style="margin-right: 10px;">
                        <span class="ms-Button-label">Yes, Reset</span>
                    </button>
                    <button id="confirmNo" class="ms-Button">
                        <span class="ms-Button-label">Cancel</span>
                    </button>
                </div>
            </div>
        `;

        // Show the confirmation
        DOMUtils.html('#messageSection', confirmHtml);
        DOMUtils.show('#messageSection');

        // Handle confirmation buttons
        DOMUtils.on('#confirmYes', 'click', () => {
            DOMUtils.hide('#messageSection');
            onConfirm();
        });

        DOMUtils.on('#confirmNo', 'click', () => {
            DOMUtils.hide('#messageSection');
        });
    }

    /**
     * Clear stored credentials
     */
    async clearCredentials() {
        const confirmed = confirm('Are you sure you want to clear all stored credentials?');
        if (!confirmed) return;
        
        try {
            await window.authManager.clearStoredCredentials();
            
            DOMUtils.val('#username', '');
            DOMUtils.val('#password', '');
            DOMUtils.prop('#rememberCredentials', 'checked', false);
            DOMUtils.prop('#autoLogin', 'checked', false);
            
            this.handleRememberCredentialsChange();
            this.showSuccess('Stored credentials cleared');
            
        } catch (error) {
            console.error('Error clearing credentials:', error);
            this.showError('Failed to clear credentials: ' + error.message);
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

// Initialize when Office is ready
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
