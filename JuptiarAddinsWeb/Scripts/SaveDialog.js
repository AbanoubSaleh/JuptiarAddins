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
        // Guard: prevent duplicate save submissions
        this._saving = false;
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
            // Remove any legacy UI sections if an older cached HTML is loaded
            this.scrubLegacySaveOptions();
            // Load initial data
            await this.loadInitialData();
            this.isInitialized = true;
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
        } else {
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
        DOMUtils.on('#saveBtn', 'click', (e) => {
            if (e) { e.preventDefault(); e.stopPropagation(); }
            this.handleSaveDocument(e);
        });
        // Cancel button
        DOMUtils.on('#cancelBtn', 'click', () => {
            this.handleCancel();
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
            // Start loading libraries immediately (optimistic loading)
            // This shows the loading spinner right away for better UX
            const librariesPromise = this.loadLibraries();

            // First check if this is a new document
            if (this.documentStateManager) {
                const isNewDocument = await this.documentStateManager.isNewDocument();
                if (!isNewDocument) {
                    // Show message that this is for new documents only
                    this.showError('This dialog is for saving new documents only. Use the Properties button to edit existing document metadata.');
                    return;
                }
            }
            // Check authentication status quickly first
            let isAuthenticated = false;

            // Try quick synchronous checks first
            if (window.authManager._isAuthenticated !== undefined) {
                isAuthenticated = window.authManager._isAuthenticated;
            } else if (window.authManager.getAuthStatus) {
                const authStatus = window.authManager.getAuthStatus();
                isAuthenticated = authStatus.isAuthenticated;
            }

            // If not authenticated via quick check, try async method
            if (!isAuthenticated && typeof window.authManager.isAuthenticated === 'function') {
                isAuthenticated = await window.authManager.isAuthenticated();
            }
            this.handleAuthStateChange(isAuthenticated);
            if (isAuthenticated) {
                // Wait for libraries to finish loading (started earlier)
                await librariesPromise;

                // Load other data in parallel
                const [suggestedName] = await Promise.all([
                    this.documentUploader.getSuggestedDocumentName(),
                    this.loadDocumentMetadata()
                ]);

                DOMUtils.select('#fileName').value = suggestedName.replace(/\.[^/.]+$/, ""); // Remove extension
            } else {
                // If not authenticated, cancel the libraries loading
                librariesPromise.catch(() => {
                    // Expected to fail if not authenticated
                });
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
        } catch (error) {
            console.error('Error loading libraries:', error);
            // Hide spinner and keep dropdown disabled (no error message shown to user)
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
    }
    /**
     * Handle save document
     */
    async handleSaveDocument(e) {
        const requestId = `save-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
        console.log(`[${requestId}] handleSaveDocument called`);
        if (this._saving) {
            console.warn(`[${requestId}] Save already in progress; ignoring duplicate click`);
            return;
        }
        this._saving = true;
        const btn = document.getElementById('saveBtn');
        if (btn) { btn.disabled = true; btn.classList.add('is-disabled'); }
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
            this.updateProgress(25, 'Saving document...');
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
                // no toast on cancel
                // this.showInfo('Save operation was cancelled.');
            } else {
                this.hideProgress();
                this.showError('Failed to save document: ' + (result.error || 'Unknown error'));
            }
        } catch (error) {
            console.error('Error saving document:', error);
            this.hideProgress();
            this.showError('Failed to save document: ' + error.message);
        } finally {
            this._saving = false;
            const btn = document.getElementById('saveBtn');
            if (btn) { btn.disabled = false; btn.classList.remove('is-disabled'); }
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
     * Hide/remove legacy Save Options and Preview elements if present (cached HTML)
     */
    scrubLegacySaveOptions() {
        try {
            // Remove checkboxes and info banners
            const ids = ['autoCheckDuplicates','notifyUsers','duplicateHandlingInfo','previewModal','confirmSaveBtn','editDetailsBtn'];
            ids.forEach(id => {
                const el = document.getElementById(id);
                if (!el) return;
                const section = el.closest('.form-section') || el.closest('.ms-Grid-col') || el.parentElement;
                if (section && section.querySelector('#autoCheckDuplicates')) {
                    section.remove();
                } else {
                    el.remove();
                }
            });
            // Remove any heading labeled "Save Options"
            const headings = Array.from(document.querySelectorAll('h3, .ms-font-l'));
            headings.forEach(h => {
                const txt = (h.textContent || '').trim().toLowerCase();
                if (txt === 'save options') {
                    const sec = h.closest('.form-section') || h.parentElement;
                    if (sec) sec.remove(); else h.remove();
                }
            });
            // Remove preview button if found by id or by label text
            let previewBtn = document.getElementById('previewBtn');
            if (!previewBtn) {
                previewBtn = Array.from(document.querySelectorAll('button')).find(b => (b.textContent||'').trim().toLowerCase() === 'preview');
            }
            if (previewBtn) {
                previewBtn.remove();
            }
        } catch (e) {
            console.warn('SaveDialog: unable to scrub legacy options:', e);
        }
    }
    /**
     * Handle cancel
     */
    handleCancel() {
        const fallbackClose = () => {
            try {
                if (Office.context && Office.context.ui && typeof Office.context.ui.messageParent === 'function') {
                    Office.context.ui.messageParent(JSON.stringify({ type: 'close' }));
                    return true;
                }
                if (typeof window.close === 'function') {
                    window.close();
                    return true;
                }
            } catch (err) {
                console.warn('Fallback close failed:', err && err.message ? err.message : err);
            }
            return false;
        };
        try {
            if (Office.addin && typeof Office.addin.hide === 'function') {
                const result = Office.addin.hide();
                if (result && typeof result.then === 'function') {
                    result.catch((err) => {
                        console.warn('Office.addin.hide() rejected, attempting fallback:', err && err.message ? err.message : err);
                        fallbackClose();
                    });
                }
                return;
            }
            // If hide API not available, use fallback
            if (!fallbackClose()) {
                console.warn('No supported method to close Save dialog in this host.');
            }
        } catch (e) {
            console.warn('Failed to close Save dialog (primary path):', e && e.message ? e.message : e);
            fallbackClose();
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
    if (window.saveDialogController) {
        console.warn('SaveDialogController already initialized; skipping duplicate init');
        return;
    }
    const saveDialogController = new SaveDialogController();
    await saveDialogController.initialize();
    // Make it globally available for debugging
    window.saveDialogController = saveDialogController;
});
