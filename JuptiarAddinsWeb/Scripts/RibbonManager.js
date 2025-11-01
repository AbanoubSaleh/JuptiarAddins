/**
 * RibbonManager - Manages ribbon state and visibility based on document state
 * Controls which tabs and buttons are visible based on whether document is new or existing
 */
class RibbonManager {
    constructor() {
        this.documentStateManager = null;
        this.isInitialized = false;
        this.lastKnownDocumentType = null;
        this.stateCheckInterval = null;
        // Prevent recursive SettingsChanged loops triggered by our own saves
        this._suppressSettingsEvent = false;
        this._settingsChangedTimer = null;
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
                    console.warn('DocumentSaved event not available:', e.message);
                }
            }

            // Listen for settings changes (when document state is updated)
            try {
                Office.context.document.settings.addHandlerAsync(
                    Office.EventType.SettingsChanged,
                    this.onDocumentSettingsChanged.bind(this)
                );
            } catch (e) {
                console.warn('SettingsChanged event not available:', e.message);
            }

            // Set up periodic document state checking for external document detection
            this.setupPeriodicStateCheck();

        } catch (error) {
            console.error('Error setting up event listeners:', error);
        }
    }

    /**
     * Set up periodic document state checking
     * This helps detect when documents are opened externally
     */
    setupPeriodicStateCheck() {
        // Check document state every 30 seconds
        this.stateCheckInterval = setInterval(async () => {
            try {
                await this.checkForDocumentStateChanges();
            } catch (error) {
                console.error('Error during periodic state check:', error);
            }
        }, 30000); // 30 seconds

        console.log('✅ Periodic document state checking enabled');
    }

    /**
     * Check for document state changes (e.g., external document opened)
     */
    async checkForDocumentStateChanges() {
        try {
            if (!this.documentStateManager) return;

            // Get current document type
            const currentDocumentType = await this.documentStateManager.detectDocumentType();

            // Check if document type has changed since last check
            if (this.lastKnownDocumentType !== currentDocumentType) {
                console.log(`📄 Document type changed: ${this.lastKnownDocumentType} → ${currentDocumentType}`);

                // Update ribbon based on new document type
                switch (currentDocumentType) {
                    case 'new':
                        await this.showNewDocumentRibbon();
                        break;
                    case 'jupiter':
                        await this.showExistingDocumentRibbon();
                        break;
                    case 'external':
                        await this.showExternalDocumentRibbon();
                        break;
                }

                this.lastKnownDocumentType = currentDocumentType;
            }
        } catch (error) {
            console.error('Error checking document state changes:', error);
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
            // Note: do not write document state from ribbon; avoid SettingsChanged loops
            console.log('✅ New document ribbon state applied (Office 2019 compatible mode)');
            console.log('ℹ️  Note: In Office 2019, buttons remain visible but will show appropriate messages when clicked');
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
            // Note: do not write document state from ribbon; avoid SettingsChanged loops
            // Update checkout status
            await this.updateCheckoutButtons();
        } catch (error) {
            console.error('Error showing existing document ribbon:', error);
        }
    }

    /**
     * Show ribbon for external documents (opened via File > Open, not Jupiter-managed)
     * Hides Jupiter-specific buttons since they don't apply to external documents
     */
    async showExternalDocumentRibbon() {
        try {
            console.log('🔧 Applying external document ribbon state...');

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
                                    visible: true  // Allow saving external documents to Jupiter
                                },
                                {
                                    id: "Jupiter.PropertiesButton",
                                    enabled: false,
                                    visible: false  // Hide Properties - not applicable to external documents
                                },
                                {
                                    id: "Jupiter.CheckOutButton",
                                    enabled: false,
                                    visible: false  // Hide Check Out - not applicable to external documents
                                },
                                {
                                    id: "Jupiter.CheckInButton",
                                    enabled: false,
                                    visible: false  // Hide Check In - not applicable to external documents
                                }
                            ]
                        }]
                    });
                } else {
                    // Fallback: Use setButtonVisibility method
                    await this.setButtonVisibility('Jupiter.SaveToJupiterButton', true);
                    await this.setButtonVisibility('Jupiter.PropertiesButton', false);
                    await this.setButtonVisibility('Jupiter.CheckOutButton', false);
                    await this.setButtonVisibility('Jupiter.CheckInButton', false);
                }
            } catch (error) {
                // Fallback: Use setButtonVisibility method
                await this.setButtonVisibility('Jupiter.SaveToJupiterButton', true);
                await this.setButtonVisibility('Jupiter.PropertiesButton', false);
                await this.setButtonVisibility('Jupiter.CheckOutButton', false);
                await this.setButtonVisibility('Jupiter.CheckInButton', false);
            }

            // Note: do not write document state from ribbon; avoid SettingsChanged loops
            console.log('✅ External document ribbon state applied');
        } catch (error) {
            console.error('Error showing external document ribbon:', error);
        }
    }

    /**
     * Update checkout buttons based on document checkout status
     * Note: This method should only be called for existing documents
     */
    async updateCheckoutButtons() {
        try {
            // Mark the time of a UI update to suppress immediate SettingsChanged loops
            this._uiUpdateAt = Date.now();

            // First check if this is a new document - if so, hide all checkout buttons
            const isNew = await this.documentStateManager.isNewDocument();
            if (isNew) {
                await this.setButtonVisibility('Jupiter.CheckOutButton', false);
                await this.setButtonVisibility('Jupiter.CheckInButton', false);
                return;
            }
            // For existing documents, show appropriate checkout buttons based on status
            const checkoutInfo = await this.documentStateManager.getCheckoutInfo();
            if (checkoutInfo && checkoutInfo.status === 'CheckedOut') {
                // Document is checked out
                await this.setButtonVisibility('Jupiter.CheckOutButton', false);
                await this.setButtonVisibility('Jupiter.CheckInButton', true);
                await this.updateButtonLabel('Jupiter.CheckInButton', 'Check In');
            } else {
                // Document is available
                await this.setButtonVisibility('Jupiter.CheckOutButton', true);
                await this.setButtonVisibility('Jupiter.CheckInButton', false);
                await this.updateButtonLabel('Jupiter.CheckOutButton', 'Check Out');
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
     * Set button visibility for Office 2019 compatibility
     * @param {string} buttonId - Button ID
     * @param {boolean} visible - Whether button should be visible
     */
    async setButtonVisibility(buttonId, visible) {
        try {
            // Determine current value to avoid unnecessary writes that trigger SettingsChanged
            const key = buttonId;
            let current = (window.jupiterRibbonState && Object.prototype.hasOwnProperty.call(window.jupiterRibbonState, key))
                ? window.jupiterRibbonState[key]
                : undefined;

            // Try reading persisted value as a last resort (do not save back to settings here)
            if (typeof current === 'undefined' && Office.context?.document?.settings) {
                try {
                    const buttonStateKey = `ribbonButton_${key}_visible`;
                    current = Office.context.document.settings.get(buttonStateKey);
                } catch (_) {
                    // ignore
                }
            }

            if (current === visible) {
                // No change; avoid noisy logs/writes
                return;
            }

            console.log(`🔧 Setting button visibility: ${key} = ${visible}`);

            // Update only in-memory to avoid SettingsChanged loops in Office 2019 fallback
            if (!window.jupiterRibbonState) {
                window.jupiterRibbonState = {};
            }
            window.jupiterRibbonState[key] = visible;
        } catch (error) {
            console.error(`Error setting button visibility for ${buttonId}:`, error);
        }
    }

    /**
     * Check if button should be visible (Office 2019 compatibility)
     * @param {string} buttonId - Button ID
     * @returns {boolean} Whether button should be visible
     */
    async isButtonVisible(buttonId) {
        try {
            // Check global state first
            if (window.jupiterRibbonState && window.jupiterRibbonState.hasOwnProperty(buttonId)) {
                return window.jupiterRibbonState[buttonId];
            }

            // Check document settings
            if (Office.context && Office.context.document && Office.context.document.settings) {
                const buttonStateKey = `ribbonButton_${buttonId}_visible`;
                const visible = Office.context.document.settings.get(buttonStateKey);
                if (visible !== null && visible !== undefined) {
                    return visible;
                }
            }

            // Default: all buttons visible
            return true;

        } catch (error) {
            console.error(`Error checking button visibility for ${buttonId}:`, error);
            return true; // Default to visible on error
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
            console.log('📄 Document saved event detected');
            // Document was saved - might need to update ribbon state
            await this.updateRibbonState();
        } catch (error) {
            console.error('Error handling document saved:', error);
        }
    }

    /**
     * Handle document settings changed event
     * This fires when Jupiter document state is updated
     */
    async onDocumentSettingsChanged(eventArgs) {
        try {
            // Suppress ribbon churn while a Jupiter dialog is open
            if (window.jupiterDialogOpen) {
                console.log('⚙️ SettingsChanged ignored (dialog open)');
                return;
            }
            if (this._suppressSettingsEvent) {
                console.log('⚙️ SettingsChanged triggered by our own save — ignoring.');
                return;
            }
            // Ignore SettingsChanged events we explicitly caused while writing settings
            if (window._jupiterSuppressSettingsEvent) {
                console.log('⚙️ SettingsChanged ignored (suppressed write)');
                return;
            }
            // Ignore bursts of events immediately after UI updates
            if (this._uiUpdateAt && (Date.now() - this._uiUpdateAt) < 1000) {
                console.log('⚙️ SettingsChanged ignored (cooldown)');
                return;
            }
            // Debounce rapid successive events
            if (this._settingsChangedTimer) {
                clearTimeout(this._settingsChangedTimer);
            }
            this._settingsChangedTimer = setTimeout(async () => {
                console.log('⚙️ Document settings changed event detected');
                try {
                    // Check if Jupiter document state was updated
                    const documentState = await this.documentStateManager.getDocumentState();
                    if (documentState && documentState.documentId) {
                        const sig = `${documentState.documentId}|${documentState.checkoutStatus}|${documentState.version || ''}`;
                        if (this._lastProcessedStateSig === sig) {
                            console.log('ℹ️ Ribbon already up-to-date for current state; skipping update.');
                            return;
                        }
                        this._lastProcessedStateSig = sig;
                        console.log('🏷️ Jupiter document state detected, updating ribbon');
                        await this.showExistingDocumentRibbon();
                    }
                } catch (innerErr) {
                    console.error('Error handling (debounced) document settings changed:', innerErr);
                }
            }, 250);
        } catch (error) {
            console.error('Error scheduling document settings changed handling:', error);
        }
    }
    /**
     * Refresh ribbon state
     */
    async refresh() {
        await this.updateRibbonState();
    }

    /**
     * Cleanup resources when ribbon manager is destroyed
     */
    destroy() {
        try {
            // Clear periodic state check interval
            if (this.stateCheckInterval) {
                clearInterval(this.stateCheckInterval);
                this.stateCheckInterval = null;
                console.log('✅ Periodic state checking disabled');
            }

            // Remove event listeners (if Office.js supports it)
            // Note: Office.js doesn't always provide removeHandlerAsync

            this.isInitialized = false;
            console.log('✅ RibbonManager destroyed');
        } catch (error) {
            console.error('Error destroying RibbonManager:', error);
        }
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
