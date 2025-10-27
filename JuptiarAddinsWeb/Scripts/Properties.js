/**
 * Properties Page - Main controller for the properties dialog
 */

// Initialize when Office is ready
Office.onReady(() => {
    // Initialize the properties editor when the page loads
    if (window.propertiesEditor) {
        window.propertiesEditor.initialize();
    }
    
    // Handle authentication button clicks
    $('#loginBtn').on('click', () => {
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
