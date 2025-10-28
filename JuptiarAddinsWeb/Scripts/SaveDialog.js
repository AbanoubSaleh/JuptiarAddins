/**
 * Save Dialog Page - Main controller for the save dialog
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
        // Initialize AuthManager asynchronously and then initialize document saver
        window.authManager.initialize().then(() => {
            console.log('AuthManager initialized, initializing document saver...');
            // Initialize the document saver when AuthManager is ready
            if (window.documentSaver) {
                window.documentSaver.initialize();
            }
        }).catch(error => {
            console.error('Failed to initialize AuthManager:', error);
            // Still initialize document saver even if AuthManager fails
            if (window.documentSaver) {
                window.documentSaver.initialize();
            }
        });
    } else {
        // AuthManager already exists, initialize document saver immediately
        if (window.documentSaver) {
            window.documentSaver.initialize();
        }
    }
    
    // Handle authentication button clicks
    $.on('#loginBtn', 'click', () => {
        // Redirect to settings page for authentication
        window.open('Settings.html', '_blank');
    });
    
    // Listen for authentication state changes
    window.addEventListener('juptiarAuthStateChanged', (e) => {
        if (e.detail.isAuthenticated && window.documentSaver) {
            window.documentSaver.initialize();
        }
    });
});
