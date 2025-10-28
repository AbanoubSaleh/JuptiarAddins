/**
 * Save Dialog Page - Enhanced controller for the save dialog with new document upload functionality
 */

class SaveDialogController {
    constructor() {
        this.documentStateManager = null;
        this.documentUploader = null;
        this.ribbonManager = null;
        this.selectedFolderId = null;
        this.selectedFolderName = null;
        this.selectedFolderPath = null;
        this.isInitialized = false;
    }

    /**
     * Initialize the save dialog controller
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

            // Load initial data
            await this.loadInitialData();

            this.isInitialized = true;
            console.log('SaveDialogController initialized successfully');
        } catch (error) {
            console.error('Failed to initialize SaveDialogController:', error);
            this.showError('Failed to initialize save dialog: ' + error.message);
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
            console.log('Creating new AuthManager instance');
            window.authManager = new AuthManager();
            await window.authManager.initialize();
            console.log('AuthManager initialized, methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(window.authManager)));
        } else {
            console.log('Using existing AuthManager instance');
            console.log('Existing AuthManager methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(window.authManager)));
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
        // Authentication button
        DOMUtils.on('#loginBtn', 'click', () => {
            window.open('Settings.html', '_blank');
        });

        // Folder selection
        DOMUtils.on('#browseFolderBtn', 'click', () => {
            this.showFolderTreeModal();
        });

        // Folder tree modal
        DOMUtils.on('#closeFolderTreeBtn', 'click', () => {
            this.hideFolderTreeModal();
        });

        DOMUtils.on('#cancelFolderSelectionBtn', 'click', () => {
            this.hideFolderTreeModal();
        });

        DOMUtils.on('#confirmFolderSelectionBtn', 'click', () => {
            this.confirmFolderSelection();
        });

        // Save button
        DOMUtils.on('#saveBtn', 'click', () => {
            this.handleSaveDocument();
        });

        // Cancel button
        DOMUtils.on('#cancelBtn', 'click', () => {
            this.handleCancel();
        });

        // Preview button
        DOMUtils.on('#previewBtn', 'click', () => {
            this.showPreview();
        });

        // Listen for authentication state changes
        window.addEventListener('juptiarAuthStateChanged', (e) => {
            this.handleAuthStateChange(e.detail.isAuthenticated);
        });
    }

    /**
     * Load initial data for the dialog
     */
    async loadInitialData() {
        try {
            // First check if this is a new document
            if (this.documentStateManager) {
                const isNewDocument = await this.documentStateManager.isNewDocument();

                if (!isNewDocument) {
                    // Show message that this is for new documents only
                    this.showError('This dialog is for saving new documents only. Use the Properties button to edit existing document metadata.');
                    return;
                }

                console.log('✅ Document is new - Save Dialog is appropriate');
            }

            // Check authentication status
            console.log('AuthManager object:', window.authManager);
            console.log('AuthManager methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(window.authManager)));
            console.log('isAuthenticated method exists:', typeof window.authManager.isAuthenticated);

            let isAuthenticated = false;
            if (typeof window.authManager.isAuthenticated === 'function') {
                isAuthenticated = await window.authManager.isAuthenticated();
            } else if (window.authManager._isAuthenticated !== undefined) {
                // Fallback to direct property access
                isAuthenticated = window.authManager._isAuthenticated;
            } else if (window.authManager.getAuthStatus) {
                // Fallback to getAuthStatus method
                const authStatus = window.authManager.getAuthStatus();
                isAuthenticated = authStatus.isAuthenticated;
            }

            console.log('Final isAuthenticated value:', isAuthenticated);
            this.handleAuthStateChange(isAuthenticated);

            if (isAuthenticated) {
                // Get suggested document name
                const suggestedName = await this.documentUploader.getSuggestedDocumentName();
                DOMUtils.select('#fileName').value = suggestedName.replace(/\.[^/.]+$/, ""); // Remove extension

                // Try to get document metadata from Word
                await this.loadDocumentMetadata();
            }
        } catch (error) {
            console.error('Error loading initial data:', error);
            this.showError('Failed to load initial data: ' + error.message);
        }
    }

    /**
     * Load document metadata from Word
     */
    async loadDocumentMetadata() {
        try {
            // In Office.js, document properties are accessed differently
            // For now, we'll skip this and use default values
            console.log('Document metadata loading skipped - using default values');

            // Set default author to current user if available
            if (window.authManager && window.authManager.currentUser) {
                const authorField = DOMUtils.select('#documentAuthor');
                if (authorField) {
                    authorField.value = window.authManager.currentUser.username || '';
                }
            }
        } catch (error) {
            console.error('Error loading document metadata:', error);
        }
    }

    /**
     * Handle authentication state changes
     */
    handleAuthStateChange(isAuthenticated) {
        const authSection = DOMUtils.select('#authSection');
        const saveFormSection = DOMUtils.select('#saveFormSection');
        const authIndicator = DOMUtils.select('#authIndicator');
        const authStatusText = DOMUtils.select('#authStatusText');
        const loginBtn = DOMUtils.select('#loginBtn');

        if (isAuthenticated) {
            authSection.style.display = 'none';
            saveFormSection.style.display = 'block';

            DOMUtils.removeClass(authIndicator, 'offline');
            DOMUtils.addClass(authIndicator, 'online');
            authStatusText.textContent = 'Authenticated';
            loginBtn.style.display = 'none';
        } else {
            authSection.style.display = 'block';
            saveFormSection.style.display = 'none';

            DOMUtils.removeClass(authIndicator, 'online');
            DOMUtils.addClass(authIndicator, 'offline');
            authStatusText.textContent = 'Not authenticated';
            loginBtn.style.display = 'inline-block';
        }
    }

    /**
     * Show folder tree modal
     */
    async showFolderTreeModal() {
        try {
            const modal = DOMUtils.select('#folderTreeModal');
            const folderTree = DOMUtils.select('#folderTree');

            // Show modal
            modal.style.display = 'flex';
            setTimeout(() => {
                DOMUtils.addClass(modal, 'show');
            }, 10);

            // Load folder tree
            folderTree.innerHTML = '<div class="loading-spinner">Loading folders...</div>';

            const libraryTree = await window.jupiterService.getLibraryTree();
            this.renderFolderTree(libraryTree, folderTree);

        } catch (error) {
            console.error('Error showing folder tree modal:', error);
            this.showError('Failed to load folders: ' + error.message);
        }
    }

    /**
     * Hide folder tree modal
     */
    hideFolderTreeModal() {
        const modal = DOMUtils.select('#folderTreeModal');
        DOMUtils.removeClass(modal, 'show');

        setTimeout(() => {
            modal.style.display = 'none';
        }, 300);
    }

    /**
     * Render folder tree
     */
    renderFolderTree(folders, container) {
        if (!folders || folders.length === 0) {
            container.innerHTML = '<div class="no-folders">No folders available</div>';
            return;
        }

        const treeHtml = this.buildFolderTreeHtml(folders);
        container.innerHTML = treeHtml;

        // Add click handlers for folder selection
        const folderItems = container.querySelectorAll('.folder-item');
        folderItems.forEach(item => {
            item.addEventListener('click', (e) => {
                e.stopPropagation();

                // Remove previous selection
                DOMUtils.removeClass('.folder-item.selected', 'selected');

                // Add selection to clicked item
                DOMUtils.addClass(item, 'selected');

                // Store selection data
                this.selectedFolderId = item.getAttribute('data-folder-id');
                this.selectedFolderName = item.getAttribute('data-folder-name');
                this.selectedFolderPath = item.getAttribute('data-folder-path');

                // Enable confirm button
                DOMUtils.select('#confirmFolderSelectionBtn').disabled = false;
            });
        });
    }

    /**
     * Build folder tree HTML
     */
    buildFolderTreeHtml(folders, level = 0) {
        let html = '';

        folders.forEach(folder => {
            const indent = level * 20;
            html += `
                <div class="folder-item"
                     data-folder-id="${folder.id}"
                     data-folder-name="${folder.name}"
                     data-folder-path="${folder.path || folder.name}"
                     style="padding-left: ${indent}px;">
                    <div class="folder-content">
                        <span class="folder-icon">📁</span>
                        <span class="folder-name">${folder.name}</span>
                    </div>
                </div>
            `;

            if (folder.children && folder.children.length > 0) {
                html += this.buildFolderTreeHtml(folder.children, level + 1);
            }
        });

        return html;
    }

    /**
     * Confirm folder selection
     */
    confirmFolderSelection() {
        if (!this.selectedFolderId) {
            return;
        }

        // Update the folder display
        const folderDisplay = DOMUtils.select('#selectedFolderDisplay');
        const folderName = DOMUtils.select('#selectedFolderName');
        const folderPath = DOMUtils.select('#selectedFolderPath');

        folderName.textContent = this.selectedFolderName;
        folderPath.textContent = this.selectedFolderPath;

        DOMUtils.addClass(folderDisplay, 'has-selection');

        // Hide modal
        this.hideFolderTreeModal();
    }

    /**
     * Handle save document
     */
    async handleSaveDocument() {
        try {
            // Validate form
            const validation = this.validateForm();
            if (!validation.isValid) {
                this.showError(validation.message);
                return;
            }

            // Show progress
            this.showProgress('Preparing document for save...');

            // Prepare save options
            const saveOptions = {
                name: DOMUtils.select('#fileName').value.trim() + '.docx',
                folderId: this.selectedFolderId,
                title: DOMUtils.select('#documentTitle').value.trim(),
                description: DOMUtils.select('#documentDescription').value.trim(),
                tags: DOMUtils.select('#documentTags').value.trim(),
                duplicateAction: null // Will be determined by DocumentUploader
            };

            // Update progress
            this.updateProgress(25, 'Checking for duplicate files...');

            // Save document using DocumentUploader
            const result = await this.documentUploader.saveNewDocument(saveOptions);

            if (result.success) {
                this.updateProgress(100, 'Document saved successfully!');

                // Show success message
                this.showSuccess('Document saved successfully to Jupiter DMS!');

                // Update ribbon state
                await this.ribbonManager.onDocumentSavedToDMS(result.document);

                // Close dialog after delay
                setTimeout(() => {
                    this.handleCancel();
                }, 2000);

            } else if (result.cancelled) {
                this.hideProgress();
                this.showInfo('Save operation was cancelled.');
            } else {
                this.hideProgress();
                this.showError('Failed to save document: ' + (result.error || 'Unknown error'));
            }

        } catch (error) {
            console.error('Error saving document:', error);
            this.hideProgress();
            this.showError('Failed to save document: ' + error.message);
        }
    }

    /**
     * Validate form data
     */
    validateForm() {
        const fileName = DOMUtils.select('#fileName').value.trim();

        if (!fileName) {
            return { isValid: false, message: 'Please enter a file name.' };
        }

        if (!this.selectedFolderId) {
            return { isValid: false, message: 'Please select a destination folder.' };
        }

        // Validate file name characters
        const invalidChars = /[<>:"/\\|?*]/;
        if (invalidChars.test(fileName)) {
            return { isValid: false, message: 'File name contains invalid characters. Please remove: < > : " / \\ | ? *' };
        }

        return { isValid: true };
    }

    /**
     * Show preview modal
     */
    showPreview() {
        // Implementation for preview functionality
        console.log('Preview functionality not yet implemented');
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
        const saveFormSection = DOMUtils.select('#saveFormSection');
        const progressText = DOMUtils.select('#progressText');
        const progressFill = DOMUtils.select('#progressFill');

        saveFormSection.style.display = 'none';
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
        const saveFormSection = DOMUtils.select('#saveFormSection');

        progressSection.style.display = 'none';
        saveFormSection.style.display = 'block';
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
     * Show info message
     */
    showInfo(message) {
        this.showMessage(message, 'info');
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

        // Auto-hide after 5 seconds for success/info messages
        if (type !== 'error') {
            setTimeout(() => {
                messageSection.style.display = 'none';
            }, 5000);
        }
    }
}

// Initialize when Office is ready
Office.onReady(async () => {
    const saveDialogController = new SaveDialogController();
    await saveDialogController.initialize();

    // Make it globally available for debugging
    window.saveDialogController = saveDialogController;
});
