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
    }

    // Initialize the document saver when the page loads
    if (window.documentSaver) {
        window.documentSaver.initialize();
    }
    
    // Handle authentication button clicks
    $('#loginBtn').on('click', () => {
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
