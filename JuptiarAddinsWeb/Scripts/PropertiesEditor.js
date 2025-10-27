/**
 * Properties Editor - Handles document metadata editing
 */

class PropertiesEditor {
    constructor() {
        this.currentDocument = null;
        this.documentMetadata = null;
        this.userPermissions = null;
        this.availableDocuments = [];
        
        this.initializeEventListeners();
    }

    /**
     * Initialize event listeners
     */
    initializeEventListeners() {
        // Authentication
        $('#loginBtn').on('click', () => this.showLoginDialog());
        
        // Document selection
        $('#loadFromJuptiarBtn').on('click', () => this.showDocumentSelectionModal());
        $('#useCurrentDocBtn').on('click', () => this.useCurrentDocument());
        
        // Tab switching
        $('.tab-header').on('click', (e) => this.switchTab($(e.currentTarget)));
        
        // Document selection modal
        $('#closeDocumentModal').on('click', () => this.hideDocumentSelectionModal());
        $('#selectDocumentBtn').on('click', () => this.selectDocumentFromModal());
        $('#cancelDocumentBtn').on('click', () => this.hideDocumentSelectionModal());
        $('#documentSearch').on('input', () => this.filterDocuments());
        
        // Document list selection
        $(document).on('click', '.document-item', (e) => this.selectDocumentItem($(e.currentTarget)));
        
        // Action buttons
        $('#savePropertiesBtn').on('click', () => this.saveProperties());
        $('#resetPropertiesBtn').on('click', () => this.resetProperties());
        $('#cancelPropertiesBtn').on('click', () => this.cancelEditing());
        
        // Listen for authentication state changes
        window.addEventListener('juptiarAuthStateChanged', (e) => {
            this.handleAuthStateChange(e.detail);
        });
    }

    /**
     * Initialize the properties editor
     */
    async initialize() {
        try {
            await this.checkAuthentication();
            
            // Check if there's a specific document to edit
            const editDocumentId = Office.context.document.settings.get('editPropertiesDocumentId');
            if (editDocumentId) {
                await this.loadDocumentById(editDocumentId);
                Office.context.document.settings.remove('editPropertiesDocumentId');
                await Office.context.document.settings.saveAsync();
            } else {
                await this.useCurrentDocument();
            }
            
        } catch (error) {
            console.error('Error initializing properties editor:', error);
            this.showError('Failed to initialize properties editor: ' + error.message);
        }
    }

    /**
     * Check authentication status
     */
    async checkAuthentication() {
        const authStatus = window.authManager.getAuthStatus();
        this.handleAuthStateChange(authStatus);
        
        if (!authStatus.isAuthenticated) {
            throw new Error('Authentication required');
        }
    }

    /**
     * Handle authentication state changes
     */
    handleAuthStateChange(authStatus) {
        if (authStatus.isAuthenticated) {
            this.showAuthenticatedState(authStatus.user);
        } else {
            this.showUnauthenticatedState();
        }
    }

    /**
     * Show authenticated state
     */
    showAuthenticatedState(user) {
        $('#authStatusText').text(`Logged in as: ${user.username || 'User'}`);
        $('#loginBtn').hide();
        $('.status-indicator').removeClass('offline').addClass('online');
        $('#documentSection').show();
        $('#authSection').hide();
    }

    /**
     * Show unauthenticated state
     */
    showUnauthenticatedState() {
        $('#authStatusText').text('Please login to edit document properties');
        $('#loginBtn').show();
        $('.status-indicator').removeClass('online').addClass('offline');
        $('#documentSection').hide();
        $('#tabsSection').hide();
        $('#actionSection').hide();
        $('#authSection').show();
    }

    /**
     * Show login dialog
     */
    showLoginDialog() {
        // This would typically open the settings page or a login modal
        // For now, we'll just show a message
        this.showInfo('Please use the Settings page to login to Juptiar');
    }

    /**
     * Use current Word document
     */
    async useCurrentDocument() {
        try {
            this.showLoading('Loading current document properties...');
            
            await this.loadCurrentWordDocument();
            await this.loadDocumentMetadataFromWord();
            
            this.populatePropertiesForm();
            this.showPropertiesEditor();
            
            this.hideLoading();
            
        } catch (error) {
            console.error('Error using current document:', error);
            this.showError('Failed to load current document: ' + error.message);
        }
    }

    /**
     * Load current Word document information
     */
    async loadCurrentWordDocument() {
        return new Promise((resolve, reject) => {
            Word.run(async (context) => {
                try {
                    const document = context.document;
                    const properties = document.properties;
                    
                    properties.load(['title', 'author', 'subject', 'keywords', 'category', 'comments']);
                    
                    await context.sync();
                    
                    this.currentDocument = {
                        id: Office.context.document.settings.get('currentJuptiarDocumentId') || null,
                        title: properties.title,
                        author: properties.author,
                        subject: properties.subject,
                        keywords: properties.keywords,
                        category: properties.category,
                        comments: properties.comments,
                        isWordDocument: true
                    };
                    
                    resolve();
                } catch (error) {
                    reject(error);
                }
            });
        });
    }

    /**
     * Load document metadata from Word properties
     */
    async loadDocumentMetadataFromWord() {
        this.documentMetadata = {
            title: this.currentDocument.title || '',
            description: this.currentDocument.subject || '',
            tags: this.currentDocument.keywords || '',
            author: this.currentDocument.author || '',
            language: 'en',
            customId: '',
            source: '',
            originalId: this.currentDocument.id || '',
            recipient: '',
            object: this.currentDocument.category || '',
            coverage: '',
            type: '',
            securityLevel: 'internal',
            allowDownload: true,
            allowPrint: true,
            allowEdit: false
        };
        
        // If document is from Juptiar, load additional metadata
        if (this.currentDocument.id) {
            try {
                const jupiterMetadata = await window.jupiterService.getDocumentMetadata(this.currentDocument.id);
                this.documentMetadata = { ...this.documentMetadata, ...jupiterMetadata };

                // Load user permissions
                this.userPermissions = await window.jupiterService.getUserPermissions(this.currentDocument.id);
            } catch (error) {
                console.warn('Could not load Juptiar metadata:', error);
            }
        }
    }

    /**
     * Load document by ID from Juptiar
     */
    async loadDocumentById(documentId) {
        try {
            this.showLoading('Loading document from Juptiar...');
            
            const document = await window.jupiterService.getDocument(documentId);
            const metadata = await window.jupiterService.getDocumentMetadata(documentId);
            const permissions = await window.jupiterService.getUserPermissions(documentId);
            
            this.currentDocument = document;
            this.documentMetadata = metadata;
            this.userPermissions = permissions;
            
            this.populatePropertiesForm();
            this.showPropertiesEditor();
            
            this.hideLoading();
            
        } catch (error) {
            console.error('Error loading document:', error);
            this.showError('Failed to load document: ' + error.message);
        }
    }

    /**
     * Show document selection modal
     */
    async showDocumentSelectionModal() {
        try {
            this.showLoading('Loading documents...');
            
            // Load available documents
            const response = await window.jupiterService.searchDocuments('', 'all');
            this.availableDocuments = response.documents || [];
            
            this.populateDocumentList(this.availableDocuments);
            $('#documentSelectionModal').show();
            
            this.hideLoading();
            
        } catch (error) {
            console.error('Error loading documents:', error);
            this.showError('Failed to load documents: ' + error.message);
        }
    }

    /**
     * Hide document selection modal
     */
    hideDocumentSelectionModal() {
        $('#documentSelectionModal').hide();
        $('#documentSearch').val('');
        $('.document-item').removeClass('selected');
        $('#selectDocumentBtn').prop('disabled', true);
    }

    /**
     * Populate document list in modal
     */
    populateDocumentList(documents) {
        const $list = $('#modalDocumentList');
        $list.empty();
        
        if (!documents || documents.length === 0) {
            $list.append('<div class="no-documents">No documents found</div>');
            return;
        }
        
        documents.forEach(doc => {
            const $item = $(`
                <div class="document-item" data-document-id="${doc.id}">
                    <div class="document-name">${doc.fileName || doc.title || 'Untitled'}</div>
                    <div class="document-path">${doc.libraryName || ''} / ${doc.folderPath || ''}</div>
                </div>
            `);
            
            $list.append($item);
        });
    }

    /**
     * Filter documents in modal
     */
    filterDocuments() {
        const query = $('#documentSearch').val().toLowerCase();
        
        $('.document-item').each(function() {
            const $item = $(this);
            const name = $item.find('.document-name').text().toLowerCase();
            const path = $item.find('.document-path').text().toLowerCase();
            
            if (name.includes(query) || path.includes(query)) {
                $item.show();
            } else {
                $item.hide();
            }
        });
    }

    /**
     * Select document item in modal
     */
    selectDocumentItem($item) {
        $('.document-item').removeClass('selected');
        $item.addClass('selected');
        $('#selectDocumentBtn').prop('disabled', false);
    }

    /**
     * Select document from modal
     */
    async selectDocumentFromModal() {
        const selectedId = $('.document-item.selected').data('document-id');
        if (!selectedId) return;
        
        this.hideDocumentSelectionModal();
        await this.loadDocumentById(selectedId);
    }

    /**
     * Switch between tabs
     */
    switchTab($tabHeader) {
        const tabId = $tabHeader.data('tab');
        
        // Update tab headers
        $('.tab-header').removeClass('active');
        $tabHeader.addClass('active');
        
        // Update tab content
        $('.tab-content').removeClass('active');
        $(`#${tabId}Tab`).addClass('active');
        
        // Load tab-specific content
        if (tabId === 'permissions') {
            this.loadPermissionsTab();
        }
    }

    /**
     * Populate properties form with current data
     */
    populatePropertiesForm() {
        if (!this.documentMetadata) return;

        // General tab
        $('#docTitle').val(this.documentMetadata.title || '');
        $('#docDescription').val(this.documentMetadata.description || '');
        $('#docTags').val(this.documentMetadata.tags || '');
        $('#docAuthor').val(this.documentMetadata.author || '');
        $('#docLanguage').val(this.documentMetadata.language || 'en');

        // Extended tab
        $('#customId').val(this.documentMetadata.customId || '');
        $('#docSource').val(this.documentMetadata.source || '');
        $('#originalId').val(this.documentMetadata.originalId || '');
        $('#docRecipient').val(this.documentMetadata.recipient || '');
        $('#docObject').val(this.documentMetadata.object || '');
        $('#docCoverage').val(this.documentMetadata.coverage || '');
        $('#docType').val(this.documentMetadata.type || '');

        // Permissions tab
        $('#securityLevel').val(this.documentMetadata.securityLevel || 'internal');
        $('#allowDownload').prop('checked', this.documentMetadata.allowDownload !== false);
        $('#allowPrint').prop('checked', this.documentMetadata.allowPrint !== false);
        $('#allowEdit').prop('checked', this.documentMetadata.allowEdit === true);
    }

    /**
     * Show properties editor
     */
    showPropertiesEditor() {
        // Update document info
        const docName = this.currentDocument?.fileName || this.currentDocument?.title || 'Current Document';
        const status = this.currentDocument?.id ? 'Saved in Juptiar' : 'Not saved to Juptiar';

        $('#currentDocName').text(docName);
        $('#documentStatus').text(status);

        // Show editor sections
        $('#tabsSection').show();
        $('#actionSection').show();
        $('#loadingSection').hide();
    }

    /**
     * Load permissions tab content
     */
    loadPermissionsTab() {
        const $permissionsContainer = $('#userPermissions');
        $permissionsContainer.empty();

        if (this.userPermissions) {
            const permissions = [
                { name: 'Read', granted: this.userPermissions.canRead },
                { name: 'Write', granted: this.userPermissions.canWrite },
                { name: 'Delete', granted: this.userPermissions.canDelete },
                { name: 'Download', granted: this.userPermissions.canDownload },
                { name: 'Print', granted: this.userPermissions.canPrint },
                { name: 'Edit Properties', granted: this.userPermissions.canEditProperties }
            ];

            permissions.forEach(perm => {
                const iconClass = perm.granted ? 'granted' : 'denied';
                const $item = $(`
                    <div class="permission-item">
                        <div class="permission-icon ${iconClass}"></div>
                        <span>${perm.name}: ${perm.granted ? 'Granted' : 'Denied'}</span>
                    </div>
                `);
                $permissionsContainer.append($item);
            });
        } else {
            $permissionsContainer.append('<p>No permission information available</p>');
        }
    }

    /**
     * Save properties
     */
    async saveProperties() {
        try {
            this.showLoading('Saving properties...');

            // Collect form data
            const updatedMetadata = this.collectFormData();

            // Save to Juptiar if document exists there
            if (this.currentDocument?.id) {
                await window.jupiterService.updateDocumentMetadata(this.currentDocument.id, updatedMetadata);
            }

            // Update Word document properties
            await this.updateWordDocumentProperties(updatedMetadata);

            // Update local metadata
            this.documentMetadata = updatedMetadata;

            this.hideLoading();
            this.showSuccess('Properties saved successfully');

        } catch (error) {
            console.error('Error saving properties:', error);
            this.showError('Failed to save properties: ' + error.message);
        }
    }

    /**
     * Collect form data
     */
    collectFormData() {
        return {
            title: $('#docTitle').val().trim(),
            description: $('#docDescription').val().trim(),
            tags: $('#docTags').val().trim(),
            author: $('#docAuthor').val().trim(),
            language: $('#docLanguage').val(),
            customId: $('#customId').val().trim(),
            source: $('#docSource').val().trim(),
            originalId: $('#originalId').val().trim(),
            recipient: $('#docRecipient').val().trim(),
            object: $('#docObject').val().trim(),
            coverage: $('#docCoverage').val().trim(),
            type: $('#docType').val(),
            securityLevel: $('#securityLevel').val(),
            allowDownload: $('#allowDownload').is(':checked'),
            allowPrint: $('#allowPrint').is(':checked'),
            allowEdit: $('#allowEdit').is(':checked')
        };
    }

    /**
     * Update Word document properties
     */
    async updateWordDocumentProperties(metadata) {
        return new Promise((resolve, reject) => {
            Word.run(async (context) => {
                try {
                    const properties = context.document.properties;

                    properties.title = metadata.title || '';
                    properties.author = metadata.author || '';
                    properties.subject = metadata.description || '';
                    properties.keywords = metadata.tags || '';
                    properties.category = metadata.object || '';
                    properties.comments = `Updated via Juptiar on ${new Date().toISOString()}`;

                    await context.sync();
                    resolve();
                } catch (error) {
                    reject(error);
                }
            });
        });
    }

    /**
     * Reset properties to original values
     */
    resetProperties() {
        const confirmed = confirm('Are you sure you want to reset all changes?');
        if (!confirmed) return;

        this.populatePropertiesForm();
        this.showInfo('Properties reset to original values');
    }

    /**
     * Cancel editing
     */
    cancelEditing() {
        // Close the task pane or reset to initial state
        this.showInfo('Editing cancelled');
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

        // Auto-hide success message after 3 seconds
        setTimeout(() => {
            $('#messageSection').fadeOut();
        }, 3000);
    }

    /**
     * Show info message
     */
    showInfo(message) {
        $('#messageBar').removeClass().addClass('ms-MessageBar ms-MessageBar--info');
        $('#messageIcon i').removeClass().addClass('ms-Icon ms-Icon--Info');
        $('#messageText').text(message);
        $('#messageSection').show();
        $('#loadingSection').hide();
    }
}

// Initialize when Office is ready
Office.onReady(() => {
    window.propertiesEditor = new PropertiesEditor();
});
