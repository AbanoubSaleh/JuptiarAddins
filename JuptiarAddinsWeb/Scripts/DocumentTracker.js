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

            // Prevent duplicate initialization
            if (this.isInitialized) {
                console.log('⚠️ DocumentTracker already initialized, skipping...');
                return;
            }

            // Clean up any existing event listeners before re-initializing
            await this.cleanupEventListeners();

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

            // Handle both numeric (1) and string ('CheckedOut') enum values
            // Backend may return: 0='Available', 1='CheckedOut', 2='CheckoutExpired', 3='Locked'
            const checkoutStatus = response.checkoutStatus;

            // Prefer backend's isCheckedOut if available, otherwise calculate it
            const isCheckedOut = typeof response.isCheckedOut === 'boolean'
                ? response.isCheckedOut
                : (checkoutStatus === 'CheckedOut' || checkoutStatus === 1);

            // Prefer backend's lockedByYou if available, otherwise calculate it
            const lockedByYou = typeof response.lockedByYou === 'boolean'
                ? response.lockedByYou
                : (isCheckedOut &&
                   response.checkedOutBy &&
                   currentUserEmail &&
                   response.checkedOutBy.toLowerCase() === currentUserEmail.toLowerCase());

            console.log('🔍 Backend response:', {
                checkoutStatus: checkoutStatus,
                checkoutStatusType: typeof checkoutStatus,
                checkedOutBy: response.checkedOutBy,
                currentUserEmail: currentUserEmail,
                backendIsCheckedOut: response.isCheckedOut,
                backendLockedByYou: response.lockedByYou,
                calculatedIsCheckedOut: isCheckedOut,
                calculatedLockedByYou: lockedByYou
            });

            return {
                isCheckedOut: isCheckedOut,
                checkedOutBy: response.checkedOutBy || null,
                lockedByYou: lockedByYou,
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

            // Sync local state so ribbon reflects accurate checkout status
            try {
                if (window.documentStateManager && typeof window.documentStateManager.getDocumentState === 'function') {
                    const localState = await window.documentStateManager.getDocumentState();
                    const desiredStatus = status.isCheckedOut ? 'CheckedOut' : 'Available';
                    const needsUpdate = !localState || localState.checkoutStatus !== desiredStatus
                        || (status.checkedOutBy && localState.checkedOutBy !== status.checkedOutBy);
                    if (needsUpdate) {
                        await window.documentStateManager.updateCheckoutStatus(desiredStatus, {
                            checkedOutBy: status.checkedOutBy,
                            lockedByYou: !!status.lockedByYou
                        });
                    }
                }
            } catch (syncErr) {
                console.warn('⚠️ Could not sync checkout status to document state:', syncErr);
            }

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

            // Fast-path: trust local document state if it clearly indicates CheckedOut BY CURRENT USER
            if (window.documentStateManager && typeof window.documentStateManager.getDocumentState === 'function') {
                try {
                    const localState = await window.documentStateManager.getDocumentState();
                    const currentUserEmail = await this.getCurrentUserEmail();

                    console.log('🔍 Local state check:', {
                        documentId: localState?.documentId,
                        checkoutStatus: localState?.checkoutStatus,
                        checkedOutBy: localState?.checkedOutBy,
                        currentUserEmail: currentUserEmail
                    });

                    // Only allow edit if document is checked out by the CURRENT user
                    if (localState &&
                        localState.documentId === this.documentId &&
                        localState.checkoutStatus === 'CheckedOut' &&
                        localState.checkedOutBy &&
                        currentUserEmail &&
                        localState.checkedOutBy.toLowerCase() === currentUserEmail.toLowerCase()) {
                        console.log('✅ Local state indicates document is checked out by YOU — allowing edit');
                        return; // Do not show any popup
                    }
                } catch (e) {
                    console.warn('Could not read local document state:', e && e.message ? e.message : e);
                }
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

            // Poll document content every 3 seconds to detect changes
            this.editDetectionInterval = setInterval(async () => {
                if (this._pollingBusy) {
                    return;
                }
                this._pollingBusy = true;
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
                } finally {
                    this._pollingBusy = false;
                }
            }, 3000); // Check every 3 seconds

            console.log('✅ Alternative edit detection started (polling every 3 seconds)');

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
                // Create popup URL with parameters - use origin only to avoid /Functions/ path issue
                const popupUrl = `${window.location.origin}/jupiter-popup.html?message=${encodeURIComponent(message)}&options=${encodeURIComponent(JSON.stringify(options))}`;

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

                // Get current user email
                const currentUserEmail = await this.getCurrentUserEmail();

                // Update internal state
                this.currentStatus = {
                    isCheckedOut: true,
                    checkedOutBy: currentUserEmail,
                    lockedByYou: true,
                    checkoutStatus: 'CheckedOut'
                };

                // Update document state manager if available - INCLUDE checkedOutBy
                if (window.documentStateManager) {
                    await window.documentStateManager.updateCheckoutStatus('CheckedOut', {
                        checkedOutBy: currentUserEmail,
                        lockedByYou: true
                    });
                    console.log('✅ Local state updated with checkedOutBy:', currentUserEmail);
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
     * Clean up event listeners to prevent duplicates
     */
    async cleanupEventListeners() {
        try {
            console.log('🧹 Cleaning up existing event listeners...');
            await Word.run(async (context) => {
                context.document.onContentChanged.removeAll();
                context.document.onSelectionChanged.removeAll();
                await context.sync();
                console.log('✅ Event listeners cleaned up');
            });
        } catch (error) {
            console.warn('⚠️ Error cleaning up event listeners (may not exist yet):', error);
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
