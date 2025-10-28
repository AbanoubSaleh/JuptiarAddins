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
        $.on('#saveBtn', 'click', () => this.saveSettings());

        // Reset to defaults
        $.on('#resetBtn', 'click', () => this.resetToDefaults());

        // Clear credentials
        $.on('#clearCredentialsBtn', 'click', () => this.clearCredentials());

        // Test connection, login, and logout
        $.on('#testConnectionBtn', 'click', () => this.testConnection());
        $.on('#loginBtn', 'click', () => this.performLogin());
        $.on('#logoutBtn', 'click', () => this.performLogout());

        // Form validation
        $.on('#serverUrl', 'blur', () => this.validateServerUrl());
        $.on('#username', 'blur', () => this.validateUsername());
        $.on('#password', 'blur', () => this.validatePassword());

        // Auto-login checkbox dependency
        $.on('#autoLogin', 'change', () => this.handleAutoLoginChange());
        $.on('#rememberCredentials', 'change', () => this.handleRememberCredentialsChange());
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
        $.val('#serverUrl', settings.serverUrl || '');
        $.val('#apiEndpoint', settings.apiEndpoint || '/api/v1');
        $.val('#connectionTimeout', settings.timeout / 1000 || 30);

        // Authentication
        $.val('#username', credentials.username || '');
        // Map old settings to new "Stay Logged In" option
        const stayLoggedIn = settings.rememberCredentials || settings.autoLogin || false;
        $.prop('#stayLoggedIn', 'checked', stayLoggedIn);

        // Advanced settings
        $.val('#defaultLibrary', settings.defaultLibrary || '');
        $.val('#documentsPerPage', settings.documentsPerPage || 50);

        $.prop('#enableNotifications', 'checked', settings.enableNotifications !== false);
        
        // Handle auto-login dependency
        this.handleRememberCredentialsChange();
    }

    /**
     * Update connection status display
     */
    updateConnectionStatus(authStatus) {
        const indicator = $.select('#statusIndicator');
        const statusText = $.select('#statusText');
        const details = $.select('#connectionDetails');

        if (authStatus.isAuthenticated) {
            $.removeClass(indicator, 'offline testing');
            $.addClass(indicator, 'online');
            $.text(statusText, 'Connected');

            $.text('#connectedServer', authStatus.serverUrl || '-');
            $.text('#connectedUser', authStatus.user?.username || '-');
            $.text('#lastConnected', new Date().toLocaleString());

            $.show(details);

            // Show logout button, hide login button
            $.hide('#loginBtn');
            $.show('#logoutBtn');
        } else {
            $.removeClass(indicator, 'online testing');
            $.addClass(indicator, 'offline');
            $.text(statusText, 'Not connected');
            $.hide(details);

            // Show login button, hide logout button
            $.show('#loginBtn');
            $.hide('#logoutBtn');
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
        const field = $.select('#serverUrl');
        $.removeClass(field, 'error');
        this.hideFieldError(field);
        return true;
    }

    /**
     * Validate username
     */
    validateUsername() {
        const username = $.val('#username').trim();
        const field = $.select('#username');
        
        if (!username) {
            $.addClass(field, 'error');
            this.showFieldError(field, 'Username is required');
            return false;
        }

        $.removeClass(field, 'error');
        this.hideFieldError(field);
        return true;
    }

    /**
     * Validate password
     */
    validatePassword() {
        const password = $.val('#password');
        const field = $.select('#password');

        if (!password) {
            $.addClass(field, 'error');
            this.showFieldError(field, 'Password is required');
            return false;
        }

        $.removeClass(field, 'error');
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
            const stayLoggedIn = $('#stayLoggedIn').is(':checked');

            // Update UI
            $('#loginBtn').prop('disabled', true);
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
            const stayLoggedIn = $.prop('#stayLoggedIn', 'checked');
            const newSettings = {
                serverUrl: $.val('#serverUrl').trim(),
                apiEndpoint: $.val('#apiEndpoint').trim(),
                timeout: parseInt($.val('#connectionTimeout')) * 1000,
                // Map new "Stay Logged In" to both old options for backward compatibility
                rememberCredentials: stayLoggedIn,
                autoLogin: stayLoggedIn,
                defaultLibrary: $.val('#defaultLibrary'),
                documentsPerPage: parseInt($.val('#documentsPerPage')),

                enableNotifications: $.prop('#enableNotifications', 'checked')
            };

            const username = $.val('#username').trim();
            const password = $.val('#password');

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
        if (!$('#serverUrl').val().trim()) {
            $('#serverUrl').val(defaultServerUrl);
        }

        $('#apiEndpoint').val(defaultApiEndpoint);
        $('#connectionTimeout').val('30');
        $('#username').val('');
        $('#password').val('');
        $('#stayLoggedIn').prop('checked', false);
        $('#defaultLibrary').val('');
        $('#documentsPerPage').val('50');

        $('#enableNotifications').prop('checked', true);

        // Clear validation errors
        $('.ms-TextField-field').removeClass('error');
        $('.error-message').remove();

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
        $('#messageSection').html(confirmHtml).show();

        // Handle confirmation buttons
        $('#confirmYes').on('click', () => {
            $('#messageSection').hide();
            onConfirm();
        });

        $('#confirmNo').on('click', () => {
            $('#messageSection').hide();
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
        $.text('#loadingText', message);
        $.show('#loadingSection');
        $.hide('#messageSection');
    }

    /**
     * Hide loading indicator
     */
    hideLoading() {
        $.hide('#loadingSection');
    }

    /**
     * Show error message
     */
    showError(message) {
        const messageBar = $.select('#messageBar');
        const messageIcon = $.select('#messageIcon i');

        messageBar.className = 'ms-MessageBar ms-MessageBar--error';
        messageIcon.className = 'ms-Icon ms-Icon--ErrorBadge';
        $.text('#messageText', message);
        $.show('#messageSection');
        $.hide('#loadingSection');
    }

    /**
     * Show success message
     */
    showSuccess(message) {
        const messageBar = $.select('#messageBar');
        const messageIcon = $.select('#messageIcon i');

        messageBar.className = 'ms-MessageBar ms-MessageBar--success';
        messageIcon.className = 'ms-Icon ms-Icon--CheckMark';
        $.text('#messageText', message);
        $.show('#messageSection');
        $.hide('#loadingSection');

        // Auto-hide success message after 3 seconds
        setTimeout(() => {
            $.hide('#messageSection');
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
