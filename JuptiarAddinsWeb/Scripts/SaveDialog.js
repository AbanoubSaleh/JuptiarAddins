/**
 * Save Dialog Page - Enhanced controller for the save dialog with new document upload functionality
 */

class SaveDialogController {
    constructor() {
        this.documentStateManager = null;
        this.documentUploader = null;
        this.ribbonManager = null;
        this.selectedLibraryId = null;
        this.selectedLibraryName = null;
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

        // Library selection
        DOMUtils.on('#librarySelect', 'change', (e) => {
            this.onLibraryChange(e.target.value);
        });

        // Folder selection
        DOMUtils.on('#folderSelect', 'change', (e) => {
            this.onFolderChange(e.target.value);
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
                // Load libraries for destination selection
                await this.loadLibraries();

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
     * Load libraries into the dropdown
     */
    async loadLibraries() {
        try {
            const librarySelect = DOMUtils.select('#librarySelect');
            const librarySpinner = DOMUtils.select('#libraryLoadingSpinner');

            // Show loading spinner
            librarySpinner.style.display = 'block';
            librarySelect.disabled = true;

            // Get libraries from API
            const libraries = await window.jupiterService.getLibraries();

            // Clear existing options (except the first placeholder)
            librarySelect.innerHTML = '<option value="">Select a library...</option>';

            // Add library options
            libraries.forEach(library => {
                const option = document.createElement('option');
                option.value = library.id;
                option.textContent = library.name;
                librarySelect.appendChild(option);
            });

            // Hide loading spinner and enable dropdown
            librarySpinner.style.display = 'none';
            librarySelect.disabled = false;

            console.log(`Loaded ${libraries.length} libraries`);
        } catch (error) {
            console.error('Error loading libraries:', error);
            this.showError('Failed to load libraries: ' + error.message);

            // Hide spinner and keep dropdown disabled
            const librarySpinner = DOMUtils.select('#libraryLoadingSpinner');
            librarySpinner.style.display = 'none';
        }
    }

    /**
     * Handle library selection change
     */
    async onLibraryChange(libraryId) {
        const folderSelect = DOMUtils.select('#folderSelect');
        const selectedPathDisplay = DOMUtils.select('#selectedPathDisplay');

        // Reset folder selection
        this.selectedFolderId = null;
        this.selectedFolderName = null;
        this.selectedFolderPath = null;
        selectedPathDisplay.style.display = 'none';

        if (!libraryId) {
            // No library selected
            this.selectedLibraryId = null;
            this.selectedLibraryName = null;
            folderSelect.innerHTML = '<option value="">Select a library first...</option>';
            folderSelect.disabled = true;
            return;
        }

        // Store selected library
        this.selectedLibraryId = libraryId;
        const librarySelect = DOMUtils.select('#librarySelect');
        this.selectedLibraryName = librarySelect.options[librarySelect.selectedIndex].text;

        // Load folders for the selected library
        await this.loadFolders(libraryId);
    }

    /**
     * Load folders for a specific library
     */
    async loadFolders(libraryId) {
        try {
            const folderSelect = DOMUtils.select('#folderSelect');
            const folderSpinner = DOMUtils.select('#folderLoadingSpinner');

            // Show loading spinner
            folderSpinner.style.display = 'block';
            folderSelect.disabled = true;
            folderSelect.innerHTML = '<option value="">Loading folders...</option>';

            // Get folders from API
            const folders = await window.jupiterService.getFolders(libraryId);

            // Clear existing options
            folderSelect.innerHTML = '<option value="">Select a folder...</option>';

            // Add folder options
            folders.forEach(folder => {
                const option = document.createElement('option');
                option.value = folder.id;
                option.textContent = folder.name;
                option.setAttribute('data-path', folder.path || folder.name);
                folderSelect.appendChild(option);
            });

            // Hide loading spinner and enable dropdown
            folderSpinner.style.display = 'none';
            folderSelect.disabled = false;

            console.log(`Loaded ${folders.length} folders for library ${libraryId}`);
        } catch (error) {
            console.error('Error loading folders:', error);
            this.showError('Failed to load folders: ' + error.message);

            // Hide spinner and show error state
            const folderSpinner = DOMUtils.select('#folderLoadingSpinner');
            const folderSelect = DOMUtils.select('#folderSelect');
            folderSpinner.style.display = 'none';
            folderSelect.innerHTML = '<option value="">Error loading folders</option>';
            folderSelect.disabled = true;
        }
    }

    /**
     * Handle folder selection change
     */
    onFolderChange(folderId) {
        const folderSelect = DOMUtils.select('#folderSelect');
        const selectedPathDisplay = DOMUtils.select('#selectedPathDisplay');
        const selectedPathText = DOMUtils.select('#selectedPathText');

        if (!folderId) {
            // No folder selected
            this.selectedFolderId = null;
            this.selectedFolderName = null;
            this.selectedFolderPath = null;
            selectedPathDisplay.style.display = 'none';
            return;
        }

        // Store selected folder
        this.selectedFolderId = folderId;
        const selectedOption = folderSelect.options[folderSelect.selectedIndex];
        this.selectedFolderName = selectedOption.text;
        this.selectedFolderPath = selectedOption.getAttribute('data-path') || selectedOption.text;

        // Update path display
        selectedPathText.textContent = `${this.selectedLibraryName} / ${this.selectedFolderName}`;
        selectedPathDisplay.style.display = 'block';

        console.log(`Selected folder: ${this.selectedFolderName} (ID: ${this.selectedFolderId})`);
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
