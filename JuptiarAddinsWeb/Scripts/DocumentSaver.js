/**
 * Document Saver - Handles saving Word documents to Juptiar
 */

class DocumentSaver {
    constructor() {
        this.libraries = [];
        this.folders = [];
        this.currentDocument = null;
        
        this.initializeEventListeners();
    }

    /**
     * Initialize event listeners
     */
    initializeEventListeners() {
        // Library selection change
        $('#librarySelect').on('change', () => this.onLibraryChange());
        
        // Auto-fill document name from Word document
        $('#fileName').on('blur', () => this.validateFileName());
        
        // Save button
        $('#saveBtn').on('click', () => this.saveDocument());
        
        // Preview button
        $('#previewBtn').on('click', () => this.showPreview());
        
        // Cancel button
        $('#cancelBtn').on('click', () => this.cancelSave());
        
        // Preview modal events
        $('#closePreview').on('click', () => this.hidePreview());
        $('#confirmSaveBtn').on('click', () => this.confirmSave());
        $('#editDetailsBtn').on('click', () => this.hidePreview());
    }

    /**
     * Initialize the save dialog
     */
    async initialize() {
        try {
            await this.checkAuthentication();
            await this.loadLibraries();
            await this.loadCurrentDocument();
            await this.populateDocumentInfo();
            
        } catch (error) {
            console.error('Error initializing save dialog:', error);
            this.showError('Failed to initialize save dialog: ' + error.message);
        }
    }

    /**
     * Check authentication status
     */
    async checkAuthentication() {
        const authStatus = window.authManager.getAuthStatus();
        
        if (authStatus.isAuthenticated) {
            this.showAuthenticatedState(authStatus.user);
        } else {
            this.showUnauthenticatedState();
            throw new Error('Authentication required');
        }
    }

    /**
     * Show authenticated state
     */
    showAuthenticatedState(user) {
        $('#authStatusText').text(`Logged in as: ${user.username || 'User'}`);
        $('#loginBtn').hide();
        $('.status-indicator').removeClass('offline').addClass('online');
        $('#saveFormSection').show();
        $('#authSection').hide();
    }

    /**
     * Show unauthenticated state
     */
    showUnauthenticatedState() {
        $('#authStatusText').text('Please login to save documents');
        $('#loginBtn').show();
        $('.status-indicator').removeClass('online').addClass('offline');
        $('#saveFormSection').hide();
        $('#authSection').show();
    }

    /**
     * Load available libraries
     */
    async loadLibraries() {
        try {
            this.showLoading('Loading libraries...');
            
            const treeData = await window.jupiterService.getLibraryTree();
            this.libraries = treeData;
            
            this.populateLibrarySelect(treeData);
            this.hideLoading();
            
        } catch (error) {
            console.error('Error loading libraries:', error);
            throw new Error('Failed to load libraries');
        }
    }

    /**
     * Populate library select dropdown
     */
    populateLibrarySelect(libraries) {
        const $select = $('#librarySelect');
        $select.empty().append('<option value="">Select a library...</option>');
        
        if (Array.isArray(libraries)) {
            libraries.forEach(library => {
                $select.append(`<option value="${library.id}">${library.name}</option>`);
            });
        }
    }

    /**
     * Handle library selection change
     */
    async onLibraryChange() {
        const libraryId = $('#librarySelect').val();
        const $folderSelect = $('#folderSelect');
        
        $folderSelect.empty().append('<option value="">Select a folder...</option>');
        
        if (!libraryId) {
            $folderSelect.prop('disabled', true);
            return;
        }

        try {
            // Find the selected library and populate folders
            const library = this.libraries.find(lib => lib.id === libraryId);
            if (library && library.children) {
                this.populateFolderSelect(library.children);
                $folderSelect.prop('disabled', false);
            } else {
                $folderSelect.prop('disabled', true);
            }
        } catch (error) {
            console.error('Error loading folders:', error);
            this.showError('Failed to load folders');
        }
    }

    /**
     * Populate folder select dropdown
     */
    populateFolderSelect(folders) {
        const $select = $('#folderSelect');
        
        const addFolders = (folderList, prefix = '') => {
            folderList.forEach(folder => {
                $select.append(`<option value="${folder.folderId || folder.id}">${prefix}${folder.name}</option>`);
                
                if (folder.children && folder.children.length > 0) {
                    addFolders(folder.children, prefix + '  ');
                }
            });
        };
        
        addFolders(folders);
    }

    /**
     * Load current Word document information
     */
    async loadCurrentDocument() {
        try {
            await Word.run(async (context) => {
                const document = context.document;
                const properties = document.properties;
                
                properties.load(['title', 'author', 'subject', 'keywords']);
                
                await context.sync();
                
                this.currentDocument = {
                    title: properties.title,
                    author: properties.author,
                    subject: properties.subject,
                    keywords: properties.keywords
                };
            });
        } catch (error) {
            console.error('Error loading document properties:', error);
            this.currentDocument = {};
        }
    }

    /**
     * Populate document information from Word document
     */
    populateDocumentInfo() {
        if (this.currentDocument) {
            $('#documentTitle').val(this.currentDocument.title || '');
            $('#documentAuthor').val(this.currentDocument.author || '');
            $('#documentDescription').val(this.currentDocument.subject || '');
            $('#documentTags').val(this.currentDocument.keywords || '');
            
            // Generate filename from title if available
            if (this.currentDocument.title && !$('#fileName').val()) {
                const fileName = this.sanitizeFileName(this.currentDocument.title);
                $('#fileName').val(fileName);
            }
        }
    }

    /**
     * Sanitize filename for safe storage
     */
    sanitizeFileName(name) {
        return name.replace(/[^a-zA-Z0-9\s\-_]/g, '').trim().replace(/\s+/g, '_');
    }

    /**
     * Validate filename input
     */
    validateFileName() {
        const fileName = $('#fileName').val().trim();
        const $field = $('#fileName');
        
        if (!fileName) {
            $field.addClass('error');
            return false;
        }
        
        // Check for invalid characters
        if (!/^[a-zA-Z0-9\s\-_]+$/.test(fileName)) {
            $field.addClass('error');
            this.showError('Filename contains invalid characters');
            return false;
        }
        
        $field.removeClass('error');
        return true;
    }

    /**
     * Show preview of document to be saved
     */
    showPreview() {
        if (!this.validateForm()) {
            return;
        }

        const formData = this.getFormData();
        
        // Populate preview modal
        $('#previewLibrary').text(this.getLibraryName(formData.libraryId));
        $('#previewFolder').text(this.getFolderName(formData.folderId) || 'Root');
        $('#previewFileName').text(formData.fileName + '.docx');
        $('#previewTitle').text(formData.title || '-');
        $('#previewDescription').text(formData.description || '-');
        $('#previewTags').text(formData.tags || '-');
        $('#previewAuthor').text(formData.author || '-');
        
        $('#previewModal').show();
    }

    /**
     * Hide preview modal
     */
    hidePreview() {
        $('#previewModal').hide();
    }

    /**
     * Get library name by ID
     */
    getLibraryName(libraryId) {
        const library = this.libraries.find(lib => lib.id === libraryId);
        return library ? library.name : 'Unknown';
    }

    /**
     * Get folder name by ID
     */
    getFolderName(folderId) {
        if (!folderId) return '';
        
        const findFolder = (folders) => {
            for (const folder of folders) {
                if (folder.folderId === folderId || folder.id === folderId) {
                    return folder.name;
                }
                if (folder.children) {
                    const found = findFolder(folder.children);
                    if (found) return found;
                }
            }
            return null;
        };
        
        for (const library of this.libraries) {
            if (library.children) {
                const folderName = findFolder(library.children);
                if (folderName) return folderName;
            }
        }
        
        return 'Unknown';
    }

    /**
     * Validate form data
     */
    validateForm() {
        let isValid = true;
        
        // Required fields
        const requiredFields = ['librarySelect', 'fileName'];
        requiredFields.forEach(fieldId => {
            const $field = $(`#${fieldId}`);
            if (!$field.val().trim()) {
                $field.addClass('error');
                isValid = false;
            } else {
                $field.removeClass('error');
            }
        });
        
        if (!isValid) {
            this.showError('Please fill in all required fields');
        }
        
        return isValid && this.validateFileName();
    }

    /**
     * Get form data
     */
    getFormData() {
        return {
            libraryId: $('#librarySelect').val(),
            folderId: $('#folderSelect').val() || null,
            fileName: $('#fileName').val().trim(),
            title: $('#documentTitle').val().trim(),
            description: $('#documentDescription').val().trim(),
            tags: $('#documentTags').val().trim(),
            author: $('#documentAuthor').val().trim(),
            language: $('#documentLanguage').val(),
            customId: $('#customId').val().trim(),
            source: $('#documentSource').val().trim(),
            overwriteExisting: $('#overwriteExisting').is(':checked'),
            createVersion: $('#createVersion').is(':checked'),
            notifyUsers: $('#notifyUsers').is(':checked')
        };
    }

    /**
     * Save document to Juptiar
     */
    async saveDocument() {
        if (!this.validateForm()) {
            return;
        }

        try {
            await this.performSave();
        } catch (error) {
            console.error('Save error:', error);
            this.showError('Failed to save document: ' + error.message);
        }
    }

    /**
     * Confirm save from preview
     */
    async confirmSave() {
        this.hidePreview();
        await this.performSave();
    }

    /**
     * Perform the actual save operation
     */
    async performSave() {
        const formData = this.getFormData();

        try {
            this.showProgress('Preparing document...', 10);

            // Get document content as blob
            const documentBlob = await this.getDocumentAsBlob();

            this.showProgress('Uploading to Juptiar...', 50);

            // Prepare metadata
            const metadata = {
                title: formData.title,
                description: formData.description,
                tags: formData.tags,
                author: formData.author,
                language: formData.language,
                customId: formData.customId,
                source: formData.source,
                overwriteExisting: formData.overwriteExisting,
                createVersion: formData.createVersion,
                notifyUsers: formData.notifyUsers
            };

            // Create file object
            const file = new File([documentBlob], formData.fileName + '.docx', {
                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });

            // Upload to Juptiar
            const response = await window.jupiterService.uploadDocument(
                formData.libraryId,
                formData.folderId,
                file,
                metadata
            );

            this.showProgress('Save completed!', 100);

            setTimeout(() => {
                this.hideProgress();
                this.showSuccess('Document saved successfully to Juptiar');

                // Store document ID for future reference
                if (response.documentId) {
                    Office.context.document.settings.set('currentJuptiarDocumentId', response.documentId);
                    Office.context.document.settings.saveAsync();
                }
            }, 1000);

        } catch (error) {
            this.hideProgress();
            throw error;
        }
    }

    /**
     * Get current Word document as blob
     */
    async getDocumentAsBlob() {
        return new Promise((resolve, reject) => {
            Word.run(async (context) => {
                try {
                    const document = context.document;
                    const documentData = document.getFileData();

                    await context.sync();

                    // Convert base64 to blob
                    const base64Data = documentData.value;
                    const binaryString = atob(base64Data);
                    const bytes = new Uint8Array(binaryString.length);

                    for (let i = 0; i < binaryString.length; i++) {
                        bytes[i] = binaryString.charCodeAt(i);
                    }

                    const blob = new Blob([bytes], {
                        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    });

                    resolve(blob);
                } catch (error) {
                    reject(error);
                }
            });
        });
    }

    /**
     * Show progress indicator
     */
    showProgress(message, percentage) {
        $('#progressText').text(message);
        $('#progressFill').css('width', percentage + '%');
        $('#progressSection').show();
        $('#saveFormSection').hide();
    }

    /**
     * Hide progress indicator
     */
    hideProgress() {
        $('#progressSection').hide();
        $('#saveFormSection').show();
    }

    /**
     * Cancel save operation
     */
    cancelSave() {
        // Close the task pane or reset form
        this.resetForm();
    }

    /**
     * Reset form to initial state
     */
    resetForm() {
        $('#librarySelect').val('');
        $('#folderSelect').val('').prop('disabled', true);
        $('#fileName').val('');
        $('#documentTitle').val('');
        $('#documentDescription').val('');
        $('#documentTags').val('');
        $('#documentAuthor').val('');
        $('#customId').val('');
        $('#documentSource').val('');
        $('#overwriteExisting').prop('checked', false);
        $('#createVersion').prop('checked', true);
        $('#notifyUsers').prop('checked', false);

        $('.ms-TextField-field').removeClass('error');
        this.hideMessage();
    }

    /**
     * Show loading indicator
     */
    showLoading(message = 'Loading...') {
        $('#loadingSection p').text(message);
        $('#loadingSection').show();
        $('#messageSection').hide();
    }

    /**
     * Hide loading indicator
     */
    hideLoading() {
        $('#loadingSection').hide();
    }

    /**
     * Show error message
     */
    showError(message) {
        $('#messageBar').removeClass().addClass('ms-MessageBar ms-MessageBar--error');
        $('#messageIcon i').removeClass().addClass('ms-Icon ms-Icon--ErrorBadge');
        $('#messageText').text(message);
        $('#messageSection').show();
        $('#loadingSection').hide();
    }

    /**
     * Show success message
     */
    showSuccess(message) {
        $('#messageBar').removeClass().addClass('ms-MessageBar ms-MessageBar--success');
        $('#messageIcon i').removeClass().addClass('ms-Icon ms-Icon--CheckMark');
        $('#messageText').text(message);
        $('#messageSection').show();
        $('#loadingSection').hide();
    }

    /**
     * Hide message
     */
    hideMessage() {
        $('#messageSection').hide();
    }
}

// Initialize when Office is ready
Office.onReady(() => {
    window.documentSaver = new DocumentSaver();
});
