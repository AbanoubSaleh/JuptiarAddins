/**
 * Properties Page - Main controller for the properties dialog
 */

// Initialize when Office is ready
Office.onReady(() => {
    // Force-hide legacy sections if cached HTML is still present
    try {
        const ds = document.getElementById('documentSection'); if (ds) ds.style.display = 'none';
        const ts = document.getElementById('tabsSection'); if (ts) ts.style.display = 'none';
    } catch (e) {}

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
            (async () => {
                try {
                    const dsm = new DocumentStateManager();
                    await dsm.initialize();
                    const isNew = await dsm.isNewDocument();
                    if (isNew) {
                        try {
                            const messageBar = document.getElementById('messageBar');
                            if (messageBar) {
                                messageBar.className = 'ms-MessageBar ms-MessageBar--warning';
                                const messageIcon = document.querySelector('#messageIcon i');
                                if (messageIcon) messageIcon.className = 'ms-Icon ms-Icon--Info';
                                const messageText = document.getElementById('messageText');
                                if (messageText) messageText.textContent = 'This dialog is for existing documents only. Use "Save to Jupiter" to save new documents first.';
                                DOMUtils.show('#messageSection');
                            }
                        } catch (e) { console.warn('Could not show warning banner:', e); }
                        DOMUtils.hide('#metaSection');
                        DOMUtils.hide('#actionSection');
                        DOMUtils.hide('#loadingSection');
                        return;
                    }
                    if (window.propertiesEditor) { window.propertiesEditor.initialize(); }
                } catch (e) {
                    console.error('Failed to evaluate document state:', e);
                    if (window.propertiesEditor) { window.propertiesEditor.initialize(); }
                }
            })();
        }).catch(error => {
            console.error('Failed to initialize AuthManager:', error);
            // Still initialize properties editor even if AuthManager fails
            (async () => {
                try {
                    const dsm = new DocumentStateManager();
                    await dsm.initialize();
                    const isNew = await dsm.isNewDocument();
                    if (isNew) {
                        try {
                            const messageBar = document.getElementById('messageBar');
                            if (messageBar) {
                                messageBar.className = 'ms-MessageBar ms-MessageBar--warning';
                                const messageIcon = document.querySelector('#messageIcon i');
                                if (messageIcon) messageIcon.className = 'ms-Icon ms-Icon--Info';
                                const messageText = document.getElementById('messageText');
                                if (messageText) messageText.textContent = 'This dialog is for existing documents only. Use \"Save to Jupiter\" to save new documents first.';
                                DOMUtils.show('#messageSection');
                            }
                        } catch (e) { console.warn('Could not show warning banner:', e); }
                        DOMUtils.hide('#metaSection');
                        DOMUtils.hide('#actionSection');
                        DOMUtils.hide('#loadingSection');
                        return;
                    }
                    if (window.propertiesEditor) { window.propertiesEditor.initialize(); }
                } catch (e) {
                    console.error('Failed to evaluate document state:', e);
                    if (window.propertiesEditor) { window.propertiesEditor.initialize(); }
                }
            })();
        });
    } else {
        // AuthManager already exists, initialize properties editor immediately
        (async () => {
            try {
                const dsm = new DocumentStateManager();
                await dsm.initialize();
                const isNew = await dsm.isNewDocument();
                if (isNew) {
                    try {
                        const messageBar = document.getElementById('messageBar');
                        if (messageBar) {
                            messageBar.className = 'ms-MessageBar ms-MessageBar--warning';
                            const messageIcon = document.querySelector('#messageIcon i');
                            if (messageIcon) messageIcon.className = 'ms-Icon ms-Icon--Info';
                            const messageText = document.getElementById('messageText');
                            if (messageText) messageText.textContent = 'This dialog is for existing documents only. Use "Save to Jupiter" to save new documents first.';
                            DOMUtils.show('#messageSection');
                        }
                    } catch (e) { console.warn('Could not show warning banner:', e); }
                    DOMUtils.hide('#metaSection');
                    DOMUtils.hide('#actionSection');
                    DOMUtils.hide('#loadingSection');
                    return;
                }
                if (window.propertiesEditor) { window.propertiesEditor.initialize(); }
            } catch (e) {
                console.error('Failed to evaluate document state:', e);
                if (window.propertiesEditor) { window.propertiesEditor.initialize(); }
            }
        })();
    }
    
    // Handle authentication button clicks
    $.on('#loginBtn', 'click', () => {
        // Redirect to settings page for authentication
        window.open('Settings.html', '_blank');
    });
    
    // Listen for authentication state changes
    window.addEventListener('juptiarAuthStateChanged', (e) => {
        if (e.detail.isAuthenticated) {
            (async () => {
                try {
                    const dsm = new DocumentStateManager();
                    await dsm.initialize();
                    const isNew = await dsm.isNewDocument();
                    if (isNew) {
                        try {
                            const messageBar = document.getElementById('messageBar');
                            if (messageBar) {
                                messageBar.className = 'ms-MessageBar ms-MessageBar--warning';
                                const messageIcon = document.querySelector('#messageIcon i');
                                if (messageIcon) messageIcon.className = 'ms-Icon ms-Icon--Info';
                                const messageText = document.getElementById('messageText');
                                if (messageText) messageText.textContent = 'This dialog is for existing documents only. Use "Save to Jupiter" to save new documents first.';
                                DOMUtils.show('#messageSection');
                            }
                        } catch (e) { console.warn('Could not show warning banner:', e); }
                        DOMUtils.hide('#metaSection');
                        DOMUtils.hide('#actionSection');
                        DOMUtils.hide('#loadingSection');
                        return;
                    }
                    if (window.propertiesEditor) { window.propertiesEditor.initialize(); }
                } catch (e) {
                    console.error('Failed to evaluate document state:', e);
                    if (window.propertiesEditor) { window.propertiesEditor.initialize(); }
                }
            })();
        }
    });
});
