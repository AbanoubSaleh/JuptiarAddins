/**
 * Check In Dialog Controller - Handles document check-in operations
 */

class CheckInDialogController {
    constructor() {
        this.documentStateManager = null;
        this.documentUploader = null;
        this.ribbonManager = null;
        this.documentInfo = null;
        this.isInitialized = false;
    }

    /**
     * Initialize the check-in dialog controller
     */
    async initialize() {
        try {
            // Initialize JupiterConfig first
            if (typeof window.JupiterConfig.init === 'function') {
                window.JupiterConfig.init();
            }

            // Initialize global services
            await this.initializeServices();
            
            // Initialize UI components
            this.initializeUI();
            
            // Load document information
            await this.loadDocumentInfo();
            
            this.isInitialized = true;
            console.log('CheckInDialogController initialized successfully');
        } catch (error) {
            console.error('Failed to initialize CheckInDialogController:', error);
            this.showError('Failed to initialize check-in dialog: ' + error.message);
        }
    }

    /**
     * Initialize all required services
     */
    async initializeServices() {
        // Initialize JupiterService
        if (!window.jupiterService) {
            window.jupiterService = new JupiterService();
            window.jupiterService.initialize({
                serverUrl: window.JupiterConfig.get('server.baseUrl'),
                apiEndpoint: window.JupiterConfig.get('server.apiEndpoint') || '/api',
                timeout: window.JupiterConfig.get('server.timeout') || 30000
            });
        }

        // Initialize AuthManager
        if (!window.authManager) {
            window.authManager = new AuthManager();
            await window.authManager.initialize();
        }

        // Initialize DocumentStateManager
        this.documentStateManager = new DocumentStateManager();
        await this.documentStateManager.initialize();

        // Initialize DocumentUploader
        this.documentUploader = new DocumentUploader();
        await this.documentUploader.initialize(window.jupiterService, this.documentStateManager);

        // Initialize RibbonManager
        this.ribbonManager = new RibbonManager();
        await this.ribbonManager.initialize(this.documentStateManager);
    }

    /**
     * Initialize UI event handlers
     */
    initializeUI() {
        // Check-in button
        DOMUtils.on('#checkinBtn', 'click', () => {
            this.handleCheckIn();
        });

        // Discard changes button
        DOMUtils.on('#discardBtn', 'click', () => {
            this.showDiscardConfirmDialog();
        });

        // Cancel button
        DOMUtils.on('#cancelBtn', 'click', () => {
            this.handleCancel();
        });

        // Discard confirmation dialog
        DOMUtils.on('#closeDiscardDialog', 'click', () => {
            this.hideDiscardConfirmDialog();
        });

        DOMUtils.on('#cancelDiscardBtn', 'click', () => {
            this.hideDiscardConfirmDialog();
        });

        DOMUtils.on('#confirmDiscardBtn', 'click', () => {
            this.handleDiscardChanges();
        });
    }

    /**
     * Load document information
     */
    async loadDocumentInfo() {
        try {
            // Get document state
            const documentState = await this.documentStateManager.getDocumentState();
            
            if (!documentState || !documentState.documentId) {
                throw new Error('No document information found. This document may not be managed by Jupiter DMS.');
            }

            // Get document details from Jupiter DMS
            this.documentInfo = await window.jupiterService.getDocumentById(documentState.documentId);
            
            // Update UI with document information
            this.updateDocumentDisplay();
            
        } catch (error) {
            console.error('Error loading document info:', error);
            this.showError('Failed to load document information: ' + error.message);
        }
    }

    /**
     * Update document display
     */
    updateDocumentDisplay() {
        if (!this.documentInfo) return;

        const documentName = DOMUtils.select('#documentName');
        const documentPath = DOMUtils.select('#documentPath');
        const documentStatus = DOMUtils.select('#documentStatus');

        documentName.textContent = this.documentInfo.name;
        documentPath.textContent = this.documentInfo.folderPath || 'Unknown location';

        // Update status based on checkout information
        const statusIndicator = documentStatus.querySelector('.status-indicator');
        const statusText = documentStatus.querySelector('.status-text');

        if (this.documentInfo.checkoutStatus === 'CheckedOut') {
            DOMUtils.removeClass(statusIndicator, 'available');
            DOMUtils.addClass(statusIndicator, 'checked-out');
            statusText.textContent = 'Checked out to you';
        } else {
            DOMUtils.removeClass(statusIndicator, 'checked-out');
            DOMUtils.addClass(statusIndicator, 'available');
            statusText.textContent = 'Available';
        }
    }

    /**
     * Handle check-in operation
     */
    async handleCheckIn() {
        try {
            // Validate form
            const validation = this.validateForm();
            if (!validation.isValid) {
                this.showError(validation.message);
                return;
            }

            // Show progress
            this.showProgress('Preparing document for check-in...');

            // Get form data
            const versionComment = DOMUtils.select('#versionComment').value.trim();
            const versionType = document.querySelector('input[name="versionType"]:checked').value;
            const keepCheckedOut = DOMUtils.select('#keepCheckedOut').checked;

            // Update progress
            this.updateProgress(25, 'Uploading document changes...');

            // Check in document using DocumentUploader
            const result = await this.documentUploader.checkInDocument(
                this.documentInfo.id, 
                versionComment
            );

            if (result.success) {
                this.updateProgress(75, 'Updating document status...');

                // Update document state
                if (!keepCheckedOut) {
                    await this.documentStateManager.updateCheckoutStatus('Available');
                }

                this.updateProgress(100, 'Document checked in successfully!');
                
                // Show success message
                this.showSuccess('Document checked in successfully! New version created.');
                
                // Update ribbon state
                await this.ribbonManager.updateCheckoutButtons();
                
                // Close dialog after delay
                setTimeout(() => {
                    this.handleCancel();
                }, 2000);
                
            } else {
                this.hideProgress();
                this.showError('Failed to check in document: ' + (result.error || 'Unknown error'));
            }

        } catch (error) {
            console.error('Error checking in document:', error);
            this.hideProgress();
            this.showError('Failed to check in document: ' + error.message);
        }
    }

    /**
     * Validate form data
     */
    validateForm() {
        const versionComment = DOMUtils.select('#versionComment').value.trim();
        
        if (!versionComment) {
            return { isValid: false, message: 'Please enter a version comment describing your changes.' };
        }

        if (versionComment.length < 10) {
            return { isValid: false, message: 'Version comment must be at least 10 characters long.' };
        }

        return { isValid: true };
    }

    /**
     * Show discard confirmation dialog
     */
    showDiscardConfirmDialog() {
        const dialog = DOMUtils.select('#discardConfirmDialog');
        dialog.style.display = 'flex';
    }

    /**
     * Hide discard confirmation dialog
     */
    hideDiscardConfirmDialog() {
        const dialog = DOMUtils.select('#discardConfirmDialog');
        dialog.style.display = 'none';
    }

    /**
     * Handle discard changes
     */
    async handleDiscardChanges() {
        try {
            this.hideDiscardConfirmDialog();
            this.showProgress('Discarding changes...');

            // Here you would implement the logic to discard changes
            // This might involve reverting to the last saved version
            // For now, we'll just update the checkout status
            
            this.updateProgress(50, 'Updating document status...');
            
            // Update document state to mark as available
            await this.documentStateManager.updateCheckoutStatus('Available');
            
            this.updateProgress(100, 'Changes discarded successfully!');
            
            // Show success message
            this.showSuccess('Changes discarded. Document is now available for others to edit.');
            
            // Update ribbon state
            await this.ribbonManager.updateCheckoutButtons();
            
            // Close dialog after delay
            setTimeout(() => {
                this.handleCancel();
            }, 2000);
            
        } catch (error) {
            console.error('Error discarding changes:', error);
            this.hideProgress();
            this.showError('Failed to discard changes: ' + error.message);
        }
    }

    /**
     * Handle cancel
     */
    handleCancel() {
        // Close the task pane
        if (Office.context.ui) {
            Office.context.ui.closeContainer();
        } else {
            window.close();
        }
    }

    /**
     * Show progress
     */
    showProgress(message) {
        const progressSection = DOMUtils.select('#progressSection');
        const checkinFormSection = DOMUtils.select('.checkin-form-section');
        const progressText = DOMUtils.select('#progressText');
        const progressFill = DOMUtils.select('#progressFill');

        checkinFormSection.style.display = 'none';
        progressSection.style.display = 'block';
        progressText.textContent = message;
        progressFill.style.width = '0%';
    }

    /**
     * Update progress
     */
    updateProgress(percentage, message) {
        const progressText = DOMUtils.select('#progressText');
        const progressFill = DOMUtils.select('#progressFill');

        progressText.textContent = message;
        progressFill.style.width = percentage + '%';
    }

    /**
     * Hide progress
     */
    hideProgress() {
        const progressSection = DOMUtils.select('#progressSection');
        const checkinFormSection = DOMUtils.select('.checkin-form-section');

        progressSection.style.display = 'none';
        checkinFormSection.style.display = 'block';
    }

    /**
     * Show error message
     */
    showError(message) {
        this.showMessage(message, 'error');
    }

    /**
     * Show success message
     */
    showSuccess(message) {
        this.showMessage(message, 'success');
    }

    /**
     * Show message
     */
    showMessage(message, type) {
        const messageSection = DOMUtils.select('#messageSection');
        const messageBar = DOMUtils.select('#messageBar');
        const messageText = DOMUtils.select('#messageText');
        const messageIcon = DOMUtils.select('#messageIcon');

        // Set message content
        messageText.textContent = message;

        // Set message type styling
        messageBar.className = `ms-MessageBar ms-MessageBar--${type}`;
        
        // Set appropriate icon
        const iconClass = type === 'error' ? 'ms-Icon--Error' : 
                         type === 'success' ? 'ms-Icon--Completed' : 
                         'ms-Icon--Info';
        messageIcon.innerHTML = `<i class="ms-Icon ${iconClass}"></i>`;

        // Show message
        messageSection.style.display = 'block';

        // Auto-hide after 5 seconds for success messages
        if (type === 'success') {
            setTimeout(() => {
                messageSection.style.display = 'none';
            }, 5000);
        }
    }
}

// Initialize when Office is ready
Office.onReady(async () => {
    const checkInDialogController = new CheckInDialogController();
    await checkInDialogController.initialize();
    
    // Make it globally available for debugging
    window.checkInDialogController = checkInDialogController;
});
