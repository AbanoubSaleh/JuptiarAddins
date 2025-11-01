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

            // Initialize UI components (dialog-only)
            this.initializeUI();

            // Set up lightweight parent<->dialog messaging bridge
            this.setupMessageBridge();

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
        // Dialog-safe no-op: Do not initialize heavy services in a dialog context.
        // Parent (taskpane/ribbon) handles auth, state, and network calls.
        console.log('[Dialog] initializeServices skipped');
        return;
    }

    /**
     * Initialize UI event handlers
     */
    initializeUI() {
        // Check-in button
        DOMUtils.on('#checkinBtn', 'click', () => {
            this.handleCheckIn();
        });

        // Cancel button
        DOMUtils.on('#cancelBtn', 'click', () => {
            this.handleCancel();
        });
    }

    /**
     * Load document information
     */

    /**
     * Receive initial data from parent and update the UI.
     * In dialog pages we cannot access Office document settings; we rely on the parent to send data.
     */
    setupMessageBridge() {
        window.addEventListener('message', (event) => {
            try {
                const data = typeof event.data === 'string' ? JSON.parse(event.data) : event.data;
                if (data && data.type === 'init' && data.document) {
                    this.documentInfo = {
                        name: data.document.name || 'This document',
                        folderPath: data.document.folderPath || '',
                        checkoutStatus: data.document.checkoutStatus || 'CheckedOut'
                    };
                    this.updateDocumentDisplay();
                }
            } catch (err) {
                console.warn('Ignored invalid init message from parent:', err);
            }
        });
    }

    async loadDocumentInfo() {
        // Dialog relies on parent to send initial document info via postMessage
        console.log('[Dialog] loadDocumentInfo skipped; waiting for parent init payload');
        return;
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

            // Collect form data
            const versionComment = DOMUtils.select('#versionComment').value.trim();
            const keepCheckedOut = false; // UI option removed; backend doesn’t support it

            // In dialog pages, Office doesn't allow accessing the host document content.
            // Send data back to the parent (taskpane/ribbon) to perform the actual check-in there.
            const payload = { cancelled: false, versionComment, keepCheckedOut };
            if (typeof Office !== 'undefined' && Office.context && Office.context.ui && typeof Office.context.ui.messageParent === 'function') {
                Office.context.ui.messageParent(JSON.stringify(payload));
            } else if (window.parent) {
                // Fallback for environments where messageParent isn't available
                window.parent.postMessage(JSON.stringify(payload), '*');
            }
        } catch (error) {
            console.error('Error preparing check-in data:', error);
            this.showError('Failed to prepare check-in: ' + error.message);
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
     * Handle cancel
     */
    handleCancel() {
        // Inform parent that dialog was cancelled
        const payload = { cancelled: true };
        if (typeof Office !== 'undefined' && Office.context && Office.context.ui && typeof Office.context.ui.messageParent === 'function') {
            Office.context.ui.messageParent(JSON.stringify(payload));
        } else if (window.parent) {
            window.parent.postMessage(JSON.stringify(payload), '*');
        } else {
            // Last resort
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

// Initialize when DOM is ready (avoid Office.onReady in dialog to prevent host registration issues)
document.addEventListener('DOMContentLoaded', async () => {
    const checkInDialogController = new CheckInDialogController();
    await checkInDialogController.initialize();
    // Make it globally available for debugging
    window.checkInDialogController = checkInDialogController;
});
