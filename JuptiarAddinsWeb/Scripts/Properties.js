/**
 * Properties Page - Main controller for the properties dialog
 */

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
        // Initialize AuthManager asynchronously and then initialize properties editor
        window.authManager.initialize().then(() => {
            console.log('AuthManager initialized, initializing properties editor...');
            // Initialize the properties editor when AuthManager is ready
            if (window.propertiesEditor) {
                window.propertiesEditor.initialize();
            }
        }).catch(error => {
            console.error('Failed to initialize AuthManager:', error);
            // Still initialize properties editor even if AuthManager fails
            if (window.propertiesEditor) {
                window.propertiesEditor.initialize();
            }
        });
    } else {
        // AuthManager already exists, initialize properties editor immediately
        if (window.propertiesEditor) {
            window.propertiesEditor.initialize();
        }
    }
    
    // Handle authentication button clicks
    $.on('#loginBtn', 'click', () => {
        // Redirect to settings page for authentication
        window.open('Settings.html', '_blank');
    });
    
    // Listen for authentication state changes
    window.addEventListener('juptiarAuthStateChanged', (e) => {
        if (e.detail.isAuthenticated && window.propertiesEditor) {
            window.propertiesEditor.initialize();
        }
    });
});
