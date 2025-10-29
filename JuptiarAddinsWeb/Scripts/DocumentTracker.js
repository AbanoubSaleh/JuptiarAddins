/**
 * DocumentTracker - Core detection and validation system for Jupiter DMS documents
 * Implements the complete workflow as specified in the requirements
 */
class DocumentTracker {
    constructor() {
        this.isInitialized = false;
        this.isJupiterDocument = false;
        this.documentId = null;
        this.libraryId = null;
        this.currentStatus = null;
        this.validationInterval = null;
        this.validationIntervalMs = 20000; // 20 seconds
        this.isEditing = false;
        this.popupShown = false;
        
        // Bind methods to preserve 'this' context
        this.onDocumentChanged = this.onDocumentChanged.bind(this);
        this.onSelectionChanged = this.onSelectionChanged.bind(this);
        this.validateDocumentStatus = this.validateDocumentStatus.bind(this);
    }

    /**
     * Initialize document tracking on add-in load
     * This is the main entry point called from Office.onReady
     */
    async initializeDocumentTracking() {
        try {
            console.log('🔍 Initializing Jupiter Document Tracking...');

            // Step 1: Detect if document is managed by Jupiter
            const jupiterInfo = await this.getJupiterDocumentId();
            
            if (!jupiterInfo) {
                // Document is not managed by Jupiter
                this.isJupiterDocument = false;
                await this.showNotManagedPopup();
                console.log('ℹ️ Document is not managed by Jupiter DMS');
                return;
            }

            // Document is managed by Jupiter
            this.isJupiterDocument = true;
            this.documentId = jupiterInfo.documentId;
            this.libraryId = jupiterInfo.libraryId;
            
            console.log('✅ Jupiter document detected:', {
                documentId: this.documentId,
                libraryId: this.libraryId
            });

            // Step 2: Validate checkout status from backend
            await this.validateDocumentStatus();

            // Step 3: Set up edit detection
            await this.setupEditDetection();

            // Step 4: Set up alternative edit detection (polling-based fallback)
            this.setupAlternativeEditDetection();

            // Step 5: Start periodic validation
            this.startPeriodicValidation();

            this.isInitialized = true;
            console.log('🎉 Document tracking initialized successfully');

        } catch (error) {
            console.error('❌ Error initializing document tracking:', error);
            throw error;
        }
    }

    /**
     * Get Jupiter Document ID and Library ID from the document
     * - Supports both new (PascalCase) and legacy (camelCase) custom property keys
     * - Falls back to DocumentStateManager scanning
     * - If only documentId is found, attempts to fetch libraryId from backend
     */
    async getJupiterDocumentId() {
        try {
            let documentId = null;
            let libraryId = null;

            // 1) Read custom properties (both PascalCase and legacy camelCase)
            try {
                const result = await Word.run(async (context) => {
                    const props = context.document.properties.customProperties;
                    props.load('items');
                    await context.sync();

                    let docId = null;
                    let libId = null;

                    props.items.forEach(prop => {
                        const key = prop.key;
                        if (key === 'JupiterDocumentId' || key === 'jupiterDocumentId') {
                            docId = prop.value;
                        } else if (key === 'LibraryId' || key === 'jupiterLibraryId') {
                            libId = prop.value;
                        }
                    });

                    return { docId, libId };
                });

                documentId = result?.docId || null;
                libraryId = result?.libId || null;
            } catch (propError) {
                console.warn('Could not read custom properties:', propError.message);
            }

            // 2) Fallback: try DocumentStateManager scan (legacy docs opened outside the browser)
            if (!documentId && typeof DocumentStateManager !== 'undefined') {
                try {
                    if (!window.documentStateManager) {
                        window.documentStateManager = new DocumentStateManager();
                        if (typeof window.documentStateManager.initialize === 'function') {
                            await window.documentStateManager.initialize();
                        }
                    }
                    if (typeof window.documentStateManager.scanDocumentForJupiterMetadata === 'function') {
                        const meta = await window.documentStateManager.scanDocumentForJupiterMetadata();
                        if (meta && meta.documentId) {
                            documentId = meta.documentId;
                        }
                    }
                } catch (scanError) {
                    console.warn('Fallback scan for Jupiter metadata failed:', scanError.message);
                }
            }

            // 3) If we have a documentId but no libraryId, try to fetch from backend
            if (documentId && !libraryId && window.jupiterService && typeof window.jupiterService.getDocumentById === 'function') {
                try {
                    const info = await window.jupiterService.getDocumentById(documentId);
                    libraryId = info?.libraryId || null;
                } catch (infoError) {
                    console.warn('Could not fetch libraryId from backend:', infoError.message);
                }
            }

            // If we found a documentId, consider this a Jupiter document (libraryId optional)
            if (documentId) {
                // Backfill canonical custom properties to keep future detection consistent
                try {
                    await this.setJupiterCustomProperties(documentId, libraryId);
                } catch (setPropError) {
                    console.warn('Could not backfill canonical custom properties:', setPropError.message);
                }
                return { documentId, libraryId };
            }

            return null;
        } catch (error) {
            console.warn('getJupiterDocumentId failed:', error.message);
            return null;
        }
    }

    /**
     * Validate checkout status from backend API
     */
    async getDocumentStatus(documentId) {
        try {
            if (!window.jupiterService) {
                throw new Error('JupiterService not available');
            }

            // Call the backend API to get document status
            const response = await window.jupiterService.getDocumentById(documentId);
            
            if (!response) {
                throw new Error('No response from backend');
            }

            // Map the response to the expected format
            const currentUserEmail = await this.getCurrentUserEmail();
            
            return {
                isCheckedOut: response.checkoutStatus === 'CheckedOut',
                checkedOutBy: response.checkedOutBy || null,
                lockedByYou: response.checkoutStatus === 'CheckedOut' && 
                           response.checkedOutBy === currentUserEmail,
                checkoutStatus: response.checkoutStatus,
                documentInfo: response
            };

        } catch (error) {
            console.error('❌ Error getting document status:', error);
            throw error;
        }
    }

    /**
     * Get current user email for comparison
     */
    async getCurrentUserEmail() {
        try {
            if (window.authManager && window.authManager.getCurrentUser) {
                const user = await window.authManager.getCurrentUser();
                return user?.email || null;
            }
            return null;
        } catch (error) {
            console.warn('Could not get current user email:', error);
            return null;
        }
    }

    /**
     * Validate document status and show appropriate popups
     */
    async validateDocumentStatus() {
        try {
            if (!this.isJupiterDocument || !this.documentId) {
                return;
            }

            console.log('🔍 Validating document status...');
            const status = await this.getDocumentStatus(this.documentId);
            this.currentStatus = status;

            console.log('📄 Document status:', status);

            // Handle different scenarios
            if (!status.isCheckedOut) {
                // Document is available - user must check out before editing
                console.log('🔓 Document is available - checkout required for editing');
                // Don't show popup immediately, wait for edit attempt
                
            } else if (status.isCheckedOut && !status.lockedByYou) {
                // Document is locked by another user
                console.log('🔒 Document is locked by another user:', status.checkedOutBy);
                await this.showLockedByOtherUserPopup(status.checkedOutBy);
                
            } else if (status.isCheckedOut && status.lockedByYou) {
                // Document is checked out by current user - editing allowed
                console.log('✅ Document is checked out by you - editing allowed');
            }

        } catch (error) {
            console.error('❌ Error validating document status:', error);
        }
    }

    /**
     * Set up edit detection using Word events
     */
    async setupEditDetection() {
        try {
            console.log('🎯 Setting up edit detection...');

            await Word.run(async (context) => {
                console.log('📝 Adding Word event listeners...');

                // Listen for document content changes
                context.document.onContentChanged.add(this.onDocumentChanged);
                console.log('✅ onContentChanged listener added');

                // Listen for selection changes (indicates user interaction)
                context.document.onSelectionChanged.add(this.onSelectionChanged);
                console.log('✅ onSelectionChanged listener added');

                await context.sync();
                console.log('✅ Event listeners synced with Word');
            });

            console.log('🎉 Edit detection set up successfully - try typing now!');

        } catch (error) {
            console.error('❌ Error setting up edit detection:', error);
        }
    }

    /**
     * Handle document content changes - detect editing attempts
     */
    async onDocumentChanged(event) {
        try {
            console.log('🚨 EDIT DETECTED! onDocumentChanged triggered');
            console.log('📊 Current state:', {
                isJupiterDocument: this.isJupiterDocument,
                documentId: this.documentId,
                popupShown: this.popupShown,
                currentStatus: this.currentStatus
            });

            if (!this.isJupiterDocument) {
                console.log('⏭️ Not a Jupiter document, ignoring edit');
                return;
            }

            if (this.popupShown) {
                console.log('⏭️ Popup already shown, ignoring edit');
                return;
            }

            console.log('📝 Processing edit attempt...');
            this.isEditing = true;

            // Re-validate status to ensure we have the latest information
            console.log('🔍 Re-validating document status...');
            await this.validateDocumentStatus();

            // Check if user needs to check out
            if (this.currentStatus && !this.currentStatus.isCheckedOut) {
                console.log('⚠️ Edit attempt without checkout - showing warning popup');
                await this.showCheckoutRequiredPopup();
            } else if (this.currentStatus && this.currentStatus.isCheckedOut && !this.currentStatus.lockedByYou) {
                console.log('🚫 Edit attempt on document locked by another user');
                await this.showLockedByOtherUserPopup(this.currentStatus.checkedOutBy);
            } else {
                console.log('✅ Edit allowed - document is checked out by current user');
            }

        } catch (error) {
            console.error('❌ Error handling document change:', error);
        }
    }

    /**
     * Handle selection changes - detect user interaction
     */
    async onSelectionChanged(event) {
        console.log('👆 Selection changed - user interaction detected');
        // This can be used for additional interaction detection if needed
        // For now, we rely primarily on content changes
    }

    /**
     * Set up alternative edit detection using polling (fallback method)
     */
    setupAlternativeEditDetection() {
        try {
            console.log('🔄 Setting up alternative edit detection (polling-based)...');

            let lastContent = '';
            let isFirstCheck = true;

            // Poll document content every 2 seconds to detect changes
            this.editDetectionInterval = setInterval(async () => {
                try {
                    if (!this.isJupiterDocument || this.popupShown) {
                        return;
                    }

                    await Word.run(async (context) => {
                        const body = context.document.body;
                        body.load('text');
                        await context.sync();

                        const currentContent = body.text;

                        if (isFirstCheck) {
                            lastContent = currentContent;
                            isFirstCheck = false;
                            console.log('📄 Initial document content captured for edit detection');
                            return;
                        }

                        if (currentContent !== lastContent) {
                            console.log('🚨 EDIT DETECTED via polling! Content changed');
                            console.log('📊 Content length changed from', lastContent.length, 'to', currentContent.length);

                            lastContent = currentContent;

                            // Trigger the same edit handling logic
                            await this.onDocumentChanged({});
                        }
                    });

                } catch (error) {
                    console.warn('⚠️ Error in alternative edit detection:', error);
                }
            }, 2000); // Check every 2 seconds

            console.log('✅ Alternative edit detection started (polling every 2 seconds)');

        } catch (error) {
            console.error('❌ Error setting up alternative edit detection:', error);
        }
    }

    /**
     * Start periodic validation (every 20 seconds)
     */
    startPeriodicValidation() {
        if (this.validationInterval) {
            clearInterval(this.validationInterval);
        }

        this.validationInterval = setInterval(async () => {
            try {
                await this.validateDocumentStatus();
            } catch (error) {
                console.error('❌ Error in periodic validation:', error);
            }
        }, this.validationIntervalMs);

        console.log(`⏰ Periodic validation started (every ${this.validationIntervalMs/1000} seconds)`);
    }

    /**
     * Stop periodic validation
     */
    stopPeriodicValidation() {
        if (this.validationInterval) {
            clearInterval(this.validationInterval);
            this.validationInterval = null;
            console.log('⏹️ Periodic validation stopped');
        }
    }

    /**
     * Show popup for documents not managed by Jupiter
     */
    async showNotManagedPopup() {
        const message = "This document is not managed by Jupiter DMS.";
        await this.showPopup(message, [
            { text: 'OK', action: 'ok' }
        ]);
    }

    /**
     * Show popup requiring checkout before editing
     */
    async showCheckoutRequiredPopup() {
        if (this.popupShown) return;
        
        this.popupShown = true;
        const message = "You need to check out this document before editing.";
        
        const result = await this.showPopup(message, [
            { text: 'Check Out', action: 'checkout' },
            { text: 'Cancel', action: 'cancel' }
        ]);

        if (result === 'checkout') {
            await this.checkOutDocument(this.documentId);
        }
        
        this.popupShown = false;
    }

    /**
     * Show popup for documents locked by another user
     */
    async showLockedByOtherUserPopup(checkedOutBy) {
        const message = `This document is currently checked out by ${checkedOutBy || 'another user'}.`;
        await this.showPopup(message, [
            { text: 'OK', action: 'ok' }
        ]);
    }

    /**
     * Generic popup display method
     */
    async showPopup(message, options = []) {
        try {
            return new Promise((resolve) => {
                // Create popup URL with parameters
                const baseUrl = window.location.origin + window.location.pathname.replace(/\/[^\/]*$/, '');
                const popupUrl = `${baseUrl}/jupiter-popup.html?message=${encodeURIComponent(message)}&options=${encodeURIComponent(JSON.stringify(options))}`;

                Office.context.ui.displayDialogAsync(
                    popupUrl,
                    { height: 40, width: 60 },
                    (asyncResult) => {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            const dialog = asyncResult.value;

                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                                dialog.close();
                                resolve(arg.message);
                            });

                            dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                                dialog.close();
                                resolve('cancel');
                            });
                        } else {
                            console.error('Failed to show popup:', asyncResult.error);
                            resolve('cancel');
                        }
                    }
                );
            });
        } catch (error) {
            console.error('❌ Error showing popup:', error);
            return 'cancel';
        }
    }

    /**
     * Handle checkout action
     */
    async checkOutDocument(documentId) {
        try {
            console.log('🔒 Attempting to check out document:', documentId);

            if (!window.jupiterService) {
                throw new Error('JupiterService not available');
            }

            const response = await window.jupiterService.checkOutDocument(documentId);

            if (response.success) {
                console.log('✅ Document checked out successfully');

                // Update internal state
                this.currentStatus = {
                    isCheckedOut: true,
                    checkedOutBy: await this.getCurrentUserEmail(),
                    lockedByYou: true,
                    checkoutStatus: 'CheckedOut'
                };

                // Update document state manager if available
                if (window.documentStateManager) {
                    await window.documentStateManager.updateCheckoutStatus('CheckedOut');
                }

                // Show success message
                await this.showPopup('Document checked out successfully! You can now edit the document.', [
                    { text: 'OK', action: 'ok' }
                ]);

                return true;
            } else {
                throw new Error(response.error || 'Checkout failed');
            }

        } catch (error) {
            console.error('❌ Error checking out document:', error);
            await this.showPopup(`Failed to check out document: ${error.message}`, [
                { text: 'OK', action: 'ok' }
            ]);
            return false;
        }
    }

    /**
     * Set custom properties for Jupiter documents
     */
    async setJupiterCustomProperties(documentId, libraryId) {
        try {
            await Word.run(async (context) => {
                const properties = context.document.properties.customProperties;

                // Remove existing properties first
                properties.load('items');
                await context.sync();

                const existingProps = properties.items.filter(p =>
                    p.key === 'JupiterDocumentId' || p.key === 'LibraryId'
                );

                existingProps.forEach(prop => prop.delete());

                // Add new properties (only add LibraryId if available)
                properties.add('JupiterDocumentId', documentId);
                if (libraryId) {
                    properties.add('LibraryId', libraryId);
                }

                await context.sync();
                console.log('✅ Jupiter custom properties set:', { documentId, libraryId });
            });
        } catch (error) {
            console.error('❌ Error setting custom properties:', error);
            throw error;
        }
    }

    /**
     * Clean up resources
     */
    dispose() {
        try {
            this.stopPeriodicValidation();

            // Remove event listeners
            if (this.isInitialized) {
                Word.run(async (context) => {
                    context.document.onContentChanged.removeAll();
                    context.document.onSelectionChanged.removeAll();
                    await context.sync();
                }).catch(error => {
                    console.warn('Error removing event listeners:', error);
                });
            }

            console.log('🧹 DocumentTracker disposed');
        } catch (error) {
            console.error('❌ Error disposing DocumentTracker:', error);
        }
    }
}

// Export for global use
window.DocumentTracker = DocumentTracker;
