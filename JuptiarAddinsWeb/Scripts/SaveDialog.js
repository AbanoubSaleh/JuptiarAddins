/**
 * Save Dialog Page - Main controller for the save dialog
 */

// Initialize when Office is ready
Office.onReady(() => {
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
