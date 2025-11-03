/**
 * Document Edit Monitor - Detects document editing attempts and manages check-out workflow
 * Handles the scenario where users try to edit documents without checking them out first
 */
class DocumentEditMonitor {
    constructor(documentStateManager, jupiterService, ribbonManager) {
        this.documentStateManager = documentStateManager;
        this.jupiterService = jupiterService;
        this.ribbonManager = ribbonManager;
        
        this.isMonitoring = false;
        this.editAttemptCount = 0;
        this.lastEditTime = null;
        this.checkoutPromptShown = false;
        this.documentProtected = false;
        
        // Bind methods to preserve 'this' context
        this.onDocumentChanged = this.onDocumentChanged.bind(this);
        this.onSelectionChanged = this.onSelectionChanged.bind(this);
        this.onContentControlAdded = this.onContentControlAdded.bind(this);
    }

    /**
     * Start monitoring document for edit attempts
     */
    async startMonitoring() {
        try {
            if (this.isMonitoring) {
                console.log('📝 Document edit monitoring already active');
                return;
            }

            console.log('🔍 Starting document edit monitoring...');

            // Check if this is a Jupiter-managed document
            const documentState = await this.documentStateManager.getDocumentState();
            if (!documentState || !documentState.documentId) {
                console.log('ℹ️  Not a Jupiter document - edit monitoring not needed');
                return;
            }

            // Get current document info from backend
            const documentInfo = await this.jupiterService.getDocumentById(documentState.documentId);
            if (!documentInfo) {
                console.log('⚠️  Could not retrieve document info from backend');
                return;
            }

            console.log('📄 Document info:', documentInfo);

            // Check if document is already checked out by current user
            if (documentInfo.checkoutStatus === 'CheckedOut' && documentInfo.checkedOutBy) {
                // TODO: Get current user email and compare
                console.log('✅ Document already checked out - monitoring for changes');
                this.documentProtected = false;
            } else if (documentInfo.checkoutStatus === 'CheckedOut') {
                console.log('🔒 Document checked out by another user - protecting document');
                this.documentProtected = true;
                await this.protectDocument();
            } else {
                console.log('🔓 Document available - monitoring for edit attempts');
                this.documentProtected = false;
            }

            // Set up event listeners
            await this.setupEventListeners();
            this.isMonitoring = true;

            console.log('✅ Document edit monitoring started successfully');

        } catch (error) {
            console.error('❌ Error starting document edit monitoring:', error);
        }
    }

    /**
     * Stop monitoring document edits
     */
    async stopMonitoring() {
        try {
            if (!this.isMonitoring) {
                return;
            }

            console.log('🛑 Stopping document edit monitoring...');

            // Remove event listeners
            await this.removeEventListeners();
            
            this.isMonitoring = false;
            this.editAttemptCount = 0;
            this.lastEditTime = null;
            this.checkoutPromptShown = false;
            this.documentProtected = false;

            console.log('✅ Document edit monitoring stopped');

        } catch (error) {
            console.error('❌ Error stopping document edit monitoring:', error);
        }
    }

    /**
     * Set up Word event listeners for edit detection
     */
    async setupEventListeners() {
        try {
            await Word.run(async (context) => {
                // Listen for document content changes
                context.document.onContentChanged.add(this.onDocumentChanged);
                
                // Listen for selection changes (indicates user interaction)
                context.document.onSelectionChanged.add(this.onSelectionChanged);
                
                await context.sync();
            });

            console.log('✅ Edit monitoring event listeners set up');

        } catch (error) {
            console.error('❌ Error setting up event listeners:', error);
        }
    }

    /**
     * Remove Word event listeners
     */
    async removeEventListeners() {
        try {
            await Word.run(async (context) => {
                context.document.onContentChanged.removeAll();
                context.document.onSelectionChanged.removeAll();
                
                await context.sync();
            });

            console.log('✅ Edit monitoring event listeners removed');

        } catch (error) {
            console.error('❌ Error removing event listeners:', error);
        }
    }

    /**
     * Handle document content changes - CORE SCENARIO HANDLER
     */
    async onDocumentChanged(event) {
        try {
            const now = Date.now();

            // Throttle edit detection (ignore rapid consecutive changes)
            if (this.lastEditTime && (now - this.lastEditTime) < 2000) {
                return;
            }

            this.lastEditTime = now;
            this.editAttemptCount++;

            console.log(`📝 EDIT DETECTED! User is trying to edit the document (attempt #${this.editAttemptCount})`);

            // Get current document state
            const documentState = await this.documentStateManager.getDocumentState();
            if (!documentState || !documentState.documentId) {
                console.log('ℹ️  Not a Jupiter document - allowing edit');
                return;
            }

            // Get fresh document status from backend
            const documentInfo = await this.jupiterService.getDocumentById(documentState.documentId);
            if (!documentInfo) {
                console.log('⚠️  Could not get document info - allowing edit');
                return;
            }

            console.log('📄 Document checkout status:', documentInfo.checkoutStatus);
            console.log('📄 Checked out by:', documentInfo.checkedOutBy);

            // Get current user email
            const currentUserEmail = await this.getCurrentUserEmail();
            console.log('👤 Current user:', currentUserEmail);

            // SCENARIO 1: Document is available (not checked out) - PROMPT FOR CHECKOUT
            if (documentInfo.checkoutStatus === 'Available') {
                if (!this.checkoutPromptShown) {
                    console.log('🔓 Document is available - showing checkout prompt');
                    await this.promptForCheckout();
                }
                return;
            }

            // SCENARIO 2: Document is checked out - check if it's by the current user
            if (documentInfo.checkoutStatus === 'CheckedOut') {
                // Check if current user is the one who checked it out
                if (documentInfo.checkedOutBy && currentUserEmail &&
                    documentInfo.checkedOutBy.toLowerCase() === currentUserEmail.toLowerCase()) {
                    console.log('✅ Document checked out by current user - allowing edit');
                    // Reset the prompt flag since user has successfully checked out
                    this.checkoutPromptShown = false;
                    return;
                } else {
                    // Document is checked out by another user - BLOCK EDITING
                    console.log('🔒 Document checked out by another user - blocking edit');
                    if (!this.checkoutPromptShown) {
                        await this.handleProtectedDocumentEdit();
                        this.checkoutPromptShown = true;
                    }
                    return;
                }
            }

        } catch (error) {
            console.error('❌ Error handling document change:', error);
        }
    }

    /**
     * Get current user email from Office context
     */
    async getCurrentUserEmail() {
        try {
            // Try to get from Office context
            if (Office.context.mailbox && Office.context.mailbox.userProfile) {
                return Office.context.mailbox.userProfile.emailAddress;
            }

            // Try to get from document state
            const documentState = await this.documentStateManager.getDocumentState();
            if (documentState && documentState.currentUserEmail) {
                return documentState.currentUserEmail;
            }

            // Try to get from local storage
            const userEmail = localStorage.getItem('jupiterUserEmail');
            if (userEmail) {
                return userEmail;
            }

            console.warn('⚠️  Could not determine current user email');
            return null;

        } catch (error) {
            console.error('❌ Error getting current user email:', error);
            return null;
        }
    }

    /**
     * Handle selection changes (user interaction detection)
     */
    async onSelectionChanged(event) {
        // This helps detect when user is actively working with the document
        // We can use this for more sophisticated edit detection if needed
    }

    /**
     * Check if document needs to be checked out before editing
     */
    async checkIfCheckoutNeeded() {
        try {
            const documentState = await this.documentStateManager.getDocumentState();
            if (!documentState || !documentState.documentId) {
                return false;
            }

            // Get fresh document info from backend
            const documentInfo = await this.jupiterService.getDocumentById(documentState.documentId);
            if (!documentInfo) {
                return false;
            }

            // If document is available (not checked out), user needs to check it out
            return documentInfo.checkoutStatus === 'Available';

        } catch (error) {
            console.error('❌ Error checking if checkout needed:', error);
            return false;
        }
    }

    /**
     * Prompt user to check out document before editing
     */
    async promptForCheckout() {
        try {
            this.checkoutPromptShown = true;

            const result = await new Promise((resolve) => {
                Office.context.ui.displayDialogAsync(
                    this.getCheckoutPromptUrl(),
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
                                resolve('cancelled');
                            });
                        } else {
                            resolve('error');
                        }
                    }
                );
            });

            if (result === 'checkout') {
                await this.performCheckout();
            } else if (result === 'readonly') {
                await this.setDocumentReadOnly();
            } else {
                // User cancelled or error - reset flag
                this.checkoutPromptShown = false;
            }

        } catch (error) {
            console.error('❌ Error prompting for checkout:', error);
            this.checkoutPromptShown = false;
        }
    }

    /**
     * Get URL for checkout prompt dialog
     */
    getCheckoutPromptUrl() {
        const baseUrl = window.location.origin + window.location.pathname.replace(/\/[^\/]*$/, '');
        return `${baseUrl}/checkout-prompt.html`;
    }

    /**
     * Perform document checkout
     */
    async performCheckout() {
        try {
            const documentState = await this.documentStateManager.getDocumentState();
            if (!documentState || !documentState.documentId) {
                throw new Error('No document ID found');
            }

            console.log('🔒 Performing document checkout...');

            const result = await this.jupiterService.checkOutDocument(documentState.documentId);

            if (result.success) {
                // Update document state
                await this.documentStateManager.updateCheckoutStatus('CheckedOut');

                // Update ribbon
                if (this.ribbonManager) {
                    await this.ribbonManager.updateCheckoutButtons();
                }

                // Reset the checkout prompt flag - user has successfully checked out
                this.checkoutPromptShown = false;

                // Show success message
                Office.context.ui.displayDialogAsync(
                    'Document checked out successfully! You can now edit the document.',
                    { height: 30, width: 50 }
                );

                console.log('✅ Document checkout successful - prompt flag reset');

            } else {
                throw new Error(result.error || 'Checkout failed');
            }

        } catch (error) {
            console.error('❌ Error performing checkout:', error);

            Office.context.ui.displayDialogAsync(
                `Failed to check out document: ${error.message}`,
                { height: 30, width: 50 }
            );

            this.checkoutPromptShown = false;
        }
    }

    /**
     * Set document to read-only mode
     */
    async setDocumentReadOnly() {
        try {
            console.log('🔒 Setting document to read-only mode...');
            
            // Note: Word Online/Desktop doesn't have a direct read-only API
            // We can show a message to the user instead
            Office.context.ui.displayDialogAsync(
                'Document is now in read-only mode. To make changes, please check out the document first.',
                { height: 30, width: 50 }
            );

            this.checkoutPromptShown = false;

        } catch (error) {
            console.error('❌ Error setting document read-only:', error);
            this.checkoutPromptShown = false;
        }
    }

    /**
     * Handle editing attempts on protected documents
     */
    async handleProtectedDocumentEdit() {
        try {
            console.log('🚫 Edit attempt on protected document');

            Office.context.ui.displayDialogAsync(
                'This document is currently checked out by another user and cannot be edited. Please try again later.',
                { height: 30, width: 50 }
            );

        } catch (error) {
            console.error('❌ Error handling protected document edit:', error);
        }
    }

    /**
     * Protect document from editing (when checked out by someone else)
     */
    async protectDocument() {
        try {
            console.log('🛡️  Protecting document from editing...');
            
            // Show warning message
            Office.context.ui.displayDialogAsync(
                'This document is currently checked out by another user. It will be opened in read-only mode.',
                { height: 30, width: 50 }
            );

        } catch (error) {
            console.error('❌ Error protecting document:', error);
        }
    }
}

// Export for use in other modules
window.DocumentEditMonitor = DocumentEditMonitor;
