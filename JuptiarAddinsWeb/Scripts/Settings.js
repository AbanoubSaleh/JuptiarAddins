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
        $('#saveBtn').on('click', () => this.saveSettings());
        
        // Reset to defaults
        $('#resetBtn').on('click', () => this.resetToDefaults());
        
        // Clear credentials
        $('#clearCredentialsBtn').on('click', () => this.clearCredentials());
        
        // Test connection, login, and logout
        $('#testConnectionBtn').on('click', () => this.testConnection());
        $('#loginBtn').on('click', () => this.performLogin());
        $('#logoutBtn').on('click', () => this.performLogout());

        // Form validation
        $('#serverUrl').on('blur', () => this.validateServerUrl());
        $('#username').on('blur', () => this.validateUsername());
        $('#password').on('blur', () => this.validatePassword());
        
        // Auto-login checkbox dependency
        $('#autoLogin').on('change', () => this.handleAutoLoginChange());
        $('#rememberCredentials').on('change', () => this.handleRememberCredentialsChange());
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
        $('#serverUrl').val(settings.serverUrl || '');
        $('#apiEndpoint').val(settings.apiEndpoint || '/api/v1');
        $('#connectionTimeout').val(settings.timeout / 1000 || 30);
        
        // Authentication
        $('#username').val(credentials.username || '');
        $('#rememberCredentials').prop('checked', settings.rememberCredentials || false);
        $('#autoLogin').prop('checked', settings.autoLogin || false);
        
        // Advanced settings
        $('#defaultLibrary').val(settings.defaultLibrary || '');
        $('#documentsPerPage').val(settings.documentsPerPage || 50);
        $('#enableLogging').prop('checked', settings.enableLogging || false);
        $('#enableNotifications').prop('checked', settings.enableNotifications !== false);
        
        // Handle auto-login dependency
        this.handleRememberCredentialsChange();
    }

    /**
     * Update connection status display
     */
    updateConnectionStatus(authStatus) {
        const $indicator = $('#statusIndicator');
        const $statusText = $('#statusText');
        const $details = $('#connectionDetails');
        
        if (authStatus.isAuthenticated) {
            $indicator.removeClass('offline testing').addClass('online');
            $statusText.text('Connected');

            $('#connectedServer').text(authStatus.serverUrl || '-');
            $('#connectedUser').text(authStatus.user?.username || '-');
            $('#lastConnected').text(new Date().toLocaleString());

            $details.show();

            // Show logout button, hide login button
            $('#loginBtn').hide();
            $('#logoutBtn').show();
        } else {
            $indicator.removeClass('online testing').addClass('offline');
            $statusText.text('Not connected');
            $details.hide();

            // Show login button, hide logout button
            $('#loginBtn').show();
            $('#logoutBtn').hide();
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
        const $select = $('#defaultLibrary');
        const currentValue = $select.val();
        
        $select.empty().append('<option value="">Select default library...</option>');
        
        if (Array.isArray(libraries)) {
            libraries.forEach(library => {
                $select.append(`<option value="${library.id}">${library.name}</option>`);
            });
        }
        
        // Restore previous selection
        if (currentValue) {
            $select.val(currentValue);
        }
    }

    /**
     * Validate server URL (now pre-configured, always valid)
     */
    validateServerUrl() {
        // Server URL is pre-configured, so always return true
        const $field = $('#serverUrl');
        $field.removeClass('error');
        this.hideFieldError($field);
        return true;
    }

    /**
     * Validate username
     */
    validateUsername() {
        const username = $('#username').val().trim();
        const $field = $('#username');
        
        if (!username) {
            $field.addClass('error');
            this.showFieldError($field, 'Username is required');
            return false;
        }
        
        $field.removeClass('error');
        this.hideFieldError($field);
        return true;
    }

    /**
     * Validate password
     */
    validatePassword() {
        const password = $('#password').val();
        const $field = $('#password');
        
        if (!password) {
            $field.addClass('error');
            this.showFieldError($field, 'Password is required');
            return false;
        }
        
        $field.removeClass('error');
        this.hideFieldError($field);
        return true;
    }

    /**
     * Show field-specific error
     */
    showFieldError($field, message) {
        $field.siblings('.error-message').remove();
        $field.after(`<span class="error-message">${message}</span>`);
    }

    /**
     * Hide field-specific error
     */
    hideFieldError($field) {
        $field.siblings('.error-message').remove();
    }

    /**
     * Handle remember credentials checkbox change
     */
    handleRememberCredentialsChange() {
        const rememberChecked = $('#rememberCredentials').is(':checked');
        $('#autoLogin').prop('disabled', !rememberChecked);
        
        if (!rememberChecked) {
            $('#autoLogin').prop('checked', false);
        }
    }

    /**
     * Handle auto-login checkbox change
     */
    handleAutoLoginChange() {
        const autoLoginChecked = $('#autoLogin').is(':checked');
        
        if (autoLoginChecked && !$('#rememberCredentials').is(':checked')) {
            $('#rememberCredentials').prop('checked', true);
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
            
            const serverUrl = $('#serverUrl').val().trim();
            const apiEndpoint = $('#apiEndpoint').val().trim();
            
            // Update UI
            $('#statusIndicator').removeClass('online offline').addClass('testing');
            $('#statusText').text('Testing connection...');
            $('#testConnectionBtn').prop('disabled', true);
            this.showLoading('Testing connection...');
            
            // Test connection
            const result = await window.authManager.testConnection(serverUrl, apiEndpoint);
            
            if (result.success) {
                $('#statusIndicator').removeClass('testing offline').addClass('online');
                $('#statusText').text('Connection successful');
                this.showSuccess('Connection test successful');
            } else {
                $('#statusIndicator').removeClass('testing online').addClass('offline');
                $('#statusText').text('Connection failed');
                this.showError('Connection test failed: ' + result.error);
            }
            
        } catch (error) {
            console.error('Connection test error:', error);
            $('#statusIndicator').removeClass('testing online').addClass('offline');
            $('#statusText').text('Connection failed');
            this.showError('Connection test failed: ' + error.message);
        } finally {
            this.isTestingConnection = false;
            $('#testConnectionBtn').prop('disabled', false);
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

            const username = $('#username').val().trim();
            const password = $('#password').val();
            const rememberCredentials = $('#rememberCredentials').is(':checked');

            // Update UI
            $('#loginBtn').prop('disabled', true);
            this.showLoading('Logging in...');

            // Perform login
            const result = await window.authManager.login(username, password, rememberCredentials);

            if (result.success) {
                this.showSuccess('Login successful! You are now authenticated.');

                // Update connection status
                const authStatus = window.authManager.getAuthStatus();
                this.updateConnectionStatus(authStatus);

                // Load library options if authenticated
                await this.loadDefaultLibraryOptions();

                // Clear password field for security (unless remember is checked)
                if (!rememberCredentials) {
                    $('#password').val('');
                }
            } else {
                this.showError('Login failed: ' + (result.message || 'Invalid credentials'));
            }

        } catch (error) {
            console.error('Login error:', error);
            this.showError('Login failed: ' + error.message);
        } finally {
            this.isLoggingIn = false;
            $('#loginBtn').prop('disabled', false);
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
            $('#logoutBtn').hide();
            $('#loginBtn').show();

            // Clear password field for security
            $('#password').val('');

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
            const newSettings = {
                serverUrl: $('#serverUrl').val().trim(),
                apiEndpoint: $('#apiEndpoint').val().trim(),
                timeout: parseInt($('#connectionTimeout').val()) * 1000,
                rememberCredentials: $('#rememberCredentials').is(':checked'),
                autoLogin: $('#autoLogin').is(':checked'),
                defaultLibrary: $('#defaultLibrary').val(),
                documentsPerPage: parseInt($('#documentsPerPage').val()),
                enableLogging: $('#enableLogging').is(':checked'),
                enableNotifications: $('#enableNotifications').is(':checked')
            };

            const username = $('#username').val().trim();
            const password = $('#password').val();

            // SECURITY BEST PRACTICE: Validate credentials before saving anything
            if (username && password) {
                console.log('Settings: Validating credentials before saving...');

                // First save settings so AuthManager can use the new server URL
                await window.authManager.saveSettings(newSettings);

                // Attempt login to validate credentials
                const loginResult = await window.authManager.login(username, password, newSettings.rememberCredentials);

                if (loginResult.success) {
                    // ✅ Credentials are valid
                    console.log('Settings: Credentials validated successfully');

                    // Update connection status
                    const authStatus = window.authManager.getAuthStatus();
                    this.updateConnectionStatus(authStatus);

                    // Load library options if authenticated
                    await this.loadDefaultLibraryOptions();

                    // Clear password field for security (token is now stored)
                    $('#password').val('');

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

                // Clear any stored credentials if remember is unchecked
                if (!newSettings.rememberCredentials) {
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
        const confirmed = confirm('Are you sure you want to reset all settings to defaults?');
        if (!confirmed) return;
        
        // Reset form to default values
        $('#serverUrl').val('');
        $('#apiEndpoint').val('/api/v1');
        $('#connectionTimeout').val('30');
        $('#username').val('');
        $('#password').val('');
        $('#rememberCredentials').prop('checked', false);
        $('#autoLogin').prop('checked', false);
        $('#defaultLibrary').val('');
        $('#documentsPerPage').val('50');
        $('#enableLogging').prop('checked', false);
        $('#enableNotifications').prop('checked', true);
        
        // Clear validation errors
        $('.ms-TextField-field').removeClass('error');
        $('.error-message').remove();
        
        this.handleRememberCredentialsChange();
        this.showSuccess('Settings reset to defaults');
    }

    /**
     * Clear stored credentials
     */
    async clearCredentials() {
        const confirmed = confirm('Are you sure you want to clear all stored credentials?');
        if (!confirmed) return;
        
        try {
            await window.authManager.clearStoredCredentials();
            
            $('#username').val('');
            $('#password').val('');
            $('#rememberCredentials').prop('checked', false);
            $('#autoLogin').prop('checked', false);
            
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
        $('#loadingText').text(message);
        $('#loadingSection').show();
        $('#messageSection').hide();
    }

    /**
     * Hide loading indicator
     */
    hideLoading() {
        $('#loadingSection').hide();
    }

    /**
     * Show error message
     */
    showError(message) {
        $('#messageBar').removeClass().addClass('ms-MessageBar ms-MessageBar--error');
        $('#messageIcon i').removeClass().addClass('ms-Icon ms-Icon--ErrorBadge');
        $('#messageText').text(message);
        $('#messageSection').show();
        $('#loadingSection').hide();
    }

    /**
     * Show success message
     */
    showSuccess(message) {
        $('#messageBar').removeClass().addClass('ms-MessageBar ms-MessageBar--success');
        $('#messageIcon i').removeClass().addClass('ms-Icon ms-Icon--CheckMark');
        $('#messageText').text(message);
        $('#messageSection').show();
        $('#loadingSection').hide();

        // Auto-hide success message after 3 seconds
        setTimeout(() => {
            $('#messageSection').fadeOut();
        }, 3000);
    }

    /**
     * Show warning message
     */
    showWarning(message) {
        $('#messageBar').removeClass().addClass('ms-MessageBar ms-MessageBar--warning');
        $('#messageIcon i').removeClass().addClass('ms-Icon ms-Icon--Warning');
        $('#messageText').text(message);
        $('#messageSection').show();
        $('#loadingSection').hide();
    }

    /**
     * Show info message
     */
    showInfo(message) {
        $('#messageBar').removeClass().addClass('ms-MessageBar ms-MessageBar--info');
        $('#messageIcon i').removeClass().addClass('ms-Icon ms-Icon--Info');
        $('#messageText').text(message);
        $('#messageSection').show();
        $('#loadingSection').hide();
    }

    /**
     * Hide message
     */
    hideMessage() {
        $('#messageSection').hide();
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
