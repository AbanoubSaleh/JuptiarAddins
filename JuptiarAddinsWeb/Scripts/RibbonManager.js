/**
 * RibbonManager - Manages ribbon state and visibility based on document state
 * Controls which tabs and buttons are visible based on whether document is new or existing
 */
class RibbonManager {
    constructor() {
        this.documentStateManager = null;
        this.isInitialized = false;
    }
    /**
     * Initialize the ribbon manager
     * @param {DocumentStateManager} documentStateManager - Document state manager instance
     */
    async initialize(documentStateManager) {
        try {
            this.documentStateManager = documentStateManager;
            await Office.onReady();
            // Set up event listeners for document state changes
            await this.setupEventListeners();
            // Update ribbon based on current document state
            await this.updateRibbonState();
            this.isInitialized = true;
        } catch (error) {
            console.error('Failed to initialize RibbonManager:', error);
            throw error;
        }
    }
    /**
     * Set up event listeners for document changes
     */
    async setupEventListeners() {
        try {
            // Listen for document selection changes
            Office.context.document.addHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                this.onDocumentSelectionChanged.bind(this)
            );
            // Listen for document saved events (if available)
            if (Office.context.document.addHandlerAsync) {
                try {
                    Office.context.document.addHandlerAsync(
                        Office.EventType.DocumentSaved,
                        this.onDocumentSaved.bind(this)
                    );
                } catch (e) {
                    // DocumentSaved event might not be available in all Office versions
                }
            }
        } catch (error) {
            console.error('Error setting up event listeners:', error);
        }
    }
    /**
     * Update ribbon state based on document state
     */
    async updateRibbonState() {
        try {
            if (!this.documentStateManager) {
                console.error('DocumentStateManager not initialized');
                return;
            }
            const isNew = await this.documentStateManager.isNewDocument();
            if (isNew) {
                await this.showNewDocumentRibbon();
            } else {
                await this.showExistingDocumentRibbon();
            }
        } catch (error) {
            console.error('Error updating ribbon state:', error);
        }
    }
    /**
     * Show ribbon for new documents
     * Uses Office.ribbon.requestUpdate to disable buttons not applicable to new documents
     */
    async showNewDocumentRibbon() {
        try {
            // Try to use Office.ribbon.requestUpdate to control button states
            try {
                if (Office.ribbon && Office.ribbon.requestUpdate) {
                    await Office.ribbon.requestUpdate({
                        tabs: [{
                            id: "Jupiter.Tab",
                            controls: [
                                {
                                    id: "Jupiter.SaveToJupiterButton",
                                    enabled: true,
                                    visible: true
                                },
                                {
                                    id: "Jupiter.PropertiesButton",
                                    enabled: false,
                                    visible: false  // Hide Properties button for new documents
                                },
                                {
                                    id: "Jupiter.CheckOutButton",
                                    enabled: false,
                                    visible: false  // Hide Check Out button for new documents
                                },
                                {
                                    id: "Jupiter.CheckInButton",
                                    enabled: false,
                                    visible: false  // Hide Check In button for new documents
                                }
                            ]
                        }]
                    });
                } else {
                    // Fallback: Use setButtonVisibility method
                    await this.setButtonVisibility('Jupiter.PropertiesButton', false);
                    await this.setButtonVisibility('Jupiter.CheckOutButton', false);
                    await this.setButtonVisibility('Jupiter.CheckInButton', false);
                    await this.setButtonVisibility('Jupiter.SaveToJupiterButton', true);
                }
            } catch (error) {
                // Fallback: Use setButtonVisibility method
                await this.setButtonVisibility('Jupiter.PropertiesButton', false);
                await this.setButtonVisibility('Jupiter.CheckOutButton', false);
                await this.setButtonVisibility('Jupiter.CheckInButton', false);
                await this.setButtonVisibility('Jupiter.SaveToJupiterButton', true);
            }
            // Store document state for button behavior
            if (this.documentStateManager) {
                await this.documentStateManager.setDocumentState({
                    isNew: true,
                    ribbonMode: 'new',
                    lastUpdated: new Date().toISOString()
                });
            }
        } catch (error) {
            console.error('Error showing new document ribbon:', error);
        }
    }
    /**
     * Show ribbon for existing documents
     * Uses Office.ribbon.requestUpdate to enable buttons applicable to existing documents
     */
    async showExistingDocumentRibbon() {
        try {
            // Try to use Office.ribbon.requestUpdate to control button states
            try {
                if (Office.ribbon && Office.ribbon.requestUpdate) {
                    await Office.ribbon.requestUpdate({
                        tabs: [{
                            id: "Jupiter.Tab",
                            controls: [
                                {
                                    id: "Jupiter.SaveToJupiterButton",
                                    enabled: false,
                                    visible: true  // Keep visible but disabled for existing documents
                                },
                                {
                                    id: "Jupiter.PropertiesButton",
                                    enabled: true,
                                    visible: true  // Show Properties button for existing documents
                                },
                                {
                                    id: "Jupiter.CheckOutButton",
                                    enabled: true,
                                    visible: true  // Show Check Out button for existing documents
                                },
                                {
                                    id: "Jupiter.CheckInButton",
                                    enabled: true,
                                    visible: true  // Show Check In button for existing documents
                                }
                            ]
                        }]
                    });
                } else {
                    // Fallback: Use setButtonVisibility method
                    await this.setButtonVisibility('Jupiter.PropertiesButton', true);
                    await this.setButtonVisibility('Jupiter.CheckOutButton', true);
                    await this.setButtonVisibility('Jupiter.CheckInButton', true);
                    await this.setButtonVisibility('Jupiter.SaveToJupiterButton', true);
                }
            } catch (error) {
                // Fallback: Use setButtonVisibility method
                await this.setButtonVisibility('Jupiter.PropertiesButton', true);
                await this.setButtonVisibility('Jupiter.CheckOutButton', true);
                await this.setButtonVisibility('Jupiter.CheckInButton', true);
                await this.setButtonVisibility('Jupiter.SaveToJupiterButton', true);
            }
            // Store document state for button behavior
            if (this.documentStateManager) {
                const currentState = await this.documentStateManager.getDocumentState();
                await this.documentStateManager.setDocumentState({
                    ...currentState,
                    isNew: false,
                    ribbonMode: 'existing',
                    lastUpdated: new Date().toISOString()
                });
            }
            // Update checkout status
            await this.updateCheckoutButtons();
        } catch (error) {
            console.error('Error showing existing document ribbon:', error);
        }
    }
    /**
     * Update checkout buttons based on document checkout status
     * Note: This method should only be called for existing documents
     */
    async updateCheckoutButtons() {
        try {
            // First check if this is a new document - if so, hide all checkout buttons
            const isNew = await this.documentStateManager.isNewDocument();
            if (isNew) {
                await this.setButtonVisibility('CheckOutButton', false);
                await this.setButtonVisibility('CheckInButton', false);
                return;
            }
            // For existing documents, show appropriate checkout buttons based on status
            const checkoutInfo = await this.documentStateManager.getCheckoutInfo();
            if (checkoutInfo && checkoutInfo.status === 'CheckedOut') {
                // Document is checked out
                await this.setButtonVisibility('CheckOutButton', false);
                await this.setButtonVisibility('CheckInButton', true);
                await this.updateButtonLabel('CheckInButton', 'Check In');
            } else {
                // Document is available
                await this.setButtonVisibility('CheckOutButton', true);
                await this.setButtonVisibility('CheckInButton', false);
                await this.updateButtonLabel('CheckOutButton', 'Check Out');
            }
        } catch (error) {
            console.error('Error updating checkout buttons:', error);
        }
    }
    /**
     * Set tab visibility
     * @param {string} tabId - Tab ID
     * @param {boolean} visible - Whether tab should be visible
     */
    async setTabVisibility(tabId, visible) {
        try {
            // Office.js doesn't have direct tab visibility control
            // This would need to be implemented through ribbon XML customization
            // For now, we'll use a workaround by enabling/disabling tab controls
            const tab = document.getElementById(tabId);
            if (tab) {
                tab.style.display = visible ? 'block' : 'none';
            }
        } catch (error) {
            console.error(`Error setting tab visibility for ${tabId}:`, error);
        }
    }
    /**
     * Set button visibility
     * @param {string} buttonId - Button ID
     * @param {boolean} visible - Whether button should be visible
     */
    async setButtonVisibility(buttonId, visible) {
        try {
            // In a real Office add-in, this would use Office.ribbon.requestUpdate
            // For now, we'll simulate with DOM manipulation
            const button = document.getElementById(buttonId);
            if (button) {
                button.style.display = visible ? 'inline-block' : 'none';
                button.disabled = !visible;
            }
        } catch (error) {
            console.error(`Error setting button visibility for ${buttonId}:`, error);
        }
    }
    /**
     * Update button label
     * @param {string} buttonId - Button ID
     * @param {string} label - New button label
     */
    async updateButtonLabel(buttonId, label) {
        try {
            const button = document.getElementById(buttonId);
            if (button) {
                button.textContent = label;
                button.title = label;
            }
        } catch (error) {
            console.error(`Error updating button label for ${buttonId}:`, error);
        }
    }
    /**
     * Enable or disable a button
     * @param {string} buttonId - Button ID
     * @param {boolean} enabled - Whether button should be enabled
     */
    async setButtonEnabled(buttonId, enabled) {
        try {
            const button = document.getElementById(buttonId);
            if (button) {
                button.disabled = !enabled;
                button.classList.toggle('disabled', !enabled);
            }
        } catch (error) {
            console.error(`Error setting button enabled state for ${buttonId}:`, error);
        }
    }
    /**
     * Show loading state on a button
     * @param {string} buttonId - Button ID
     * @param {boolean} loading - Whether to show loading state
     */
    async setButtonLoading(buttonId, loading) {
        try {
            const button = document.getElementById(buttonId);
            if (button) {
                if (loading) {
                    button.classList.add('loading');
                    button.disabled = true;
                    button.innerHTML = '<span class="spinner"></span> Processing...';
                } else {
                    button.classList.remove('loading');
                    button.disabled = false;
                    // Restore original button text - this would need to be stored
                    button.innerHTML = button.getAttribute('data-original-text') || 'Save';
                }
            }
        } catch (error) {
            console.error(`Error setting button loading state for ${buttonId}:`, error);
        }
    }
    /**
     * Handle document selection changed event
     */
    async onDocumentSelectionChanged(eventArgs) {
        try {
            // Update ribbon state when document selection changes
            // This might indicate document state has changed
            await this.updateRibbonState();
        } catch (error) {
            console.error('Error handling document selection changed:', error);
        }
    }
    /**
     * Handle document saved event
     */
    async onDocumentSaved(eventArgs) {
        try {
            // Document was saved - might need to update ribbon state
            await this.updateRibbonState();
        } catch (error) {
            console.error('Error handling document saved:', error);
        }
    }
    /**
     * Refresh ribbon state
     */
    async refresh() {
        await this.updateRibbonState();
    }
    /**
     * Force ribbon to new document mode (for testing)
     */
    async forceNewDocumentMode() {
        await this.showNewDocumentRibbon();
    }
    /**
     * Force ribbon to existing document mode (for testing)
     */
    async forceExistingDocumentMode() {
        await this.showExistingDocumentRibbon();
    }
    /**
     * Handle document opened from Jupiter DMS
     * @param {Object} documentInfo - Document information
     */
    async onDocumentOpenedFromDMS(documentInfo) {
        try {
            // Mark document as existing
            await this.documentStateManager.markAsExistingDocument(documentInfo);
            // Update ribbon to show existing document state
            await this.showExistingDocumentRibbon();
        } catch (error) {
            console.error('Error handling document opened from DMS:', error);
        }
    }
    /**
     * Handle document saved to Jupiter DMS
     * @param {Object} documentInfo - Document information
     */
    async onDocumentSavedToDMS(documentInfo) {
        try {
            // Mark document as existing
            await this.documentStateManager.markAsExistingDocument(documentInfo);
            // Update ribbon to show existing document state
            await this.showExistingDocumentRibbon();
        } catch (error) {
            console.error('Error handling document saved to DMS:', error);
        }
    }
    /**
     * Handle new document created
     */
    async onNewDocumentCreated() {
        try {
            // Mark document as new
            await this.documentStateManager.markAsNewDocument();
            // Update ribbon to show new document state
            await this.showNewDocumentRibbon();
        } catch (error) {
            console.error('Error handling new document created:', error);
        }
    }
}
// Export for use in other modules
window.RibbonManager = RibbonManager;
