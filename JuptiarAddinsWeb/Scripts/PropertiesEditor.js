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
        DOMUtils.on('#loginBtn', 'click', () => this.showLoginDialog());

        // Document selection
        DOMUtils.on('#loadFromJuptiarBtn', 'click', () => this.showDocumentSelectionModal());
        DOMUtils.on('#useCurrentDocBtn', 'click', () => this.useCurrentDocument());

        // Tab switching
        DOMUtils.on('.tab-header', 'click', (e) => this.switchTab(e.currentTarget));

        // Document selection modal
        DOMUtils.on('#closeDocumentModal', 'click', () => this.hideDocumentSelectionModal());
        DOMUtils.on('#selectDocumentBtn', 'click', () => this.selectDocumentFromModal());
        DOMUtils.on('#cancelDocumentBtn', 'click', () => this.hideDocumentSelectionModal());
        DOMUtils.on('#documentSearch', 'input', () => this.filterDocuments());

        // Document list selection
        DOMUtils.delegate(document, 'click', '.document-item', (e) => this.selectDocumentItem(e.currentTarget));

        // Action buttons
        DOMUtils.on('#savePropertiesBtn', 'click', () => this.saveProperties());
        DOMUtils.on('#resetPropertiesBtn', 'click', () => this.resetProperties());
        DOMUtils.on('#cancelPropertiesBtn', 'click', () => this.cancelEditing());

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
        DOMUtils.text('#authStatusText', `Logged in as: ${user.username || 'User'}`);
        DOMUtils.hide('#loginBtn');
        DOMUtils.removeClass('.status-indicator', 'offline');
        DOMUtils.addClass('.status-indicator', 'online');
        DOMUtils.show('#documentSection');
        DOMUtils.hide('#authSection');
    }

    /**
     * Show unauthenticated state
     */
    showUnauthenticatedState() {
        DOMUtils.text('#authStatusText', 'Please login to edit document properties');
        DOMUtils.show('#loginBtn');
        DOMUtils.removeClass('.status-indicator', 'online');
        DOMUtils.addClass('.status-indicator', 'offline');
        DOMUtils.hide('#documentSection');
        DOMUtils.hide('#tabsSection');
        DOMUtils.hide('#actionSection');
        DOMUtils.show('#authSection');
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
            DOMUtils.show('#documentSelectionModal');
            
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
        DOMUtils.hide('#documentSelectionModal');
        DOMUtils.val('#documentSearch', '');
        DOMUtils.removeClass('.document-item', 'selected');
        DOMUtils.prop('#selectDocumentBtn', 'disabled', true);
    }

    /**
     * Populate document list in modal
     */
    populateDocumentList(documents) {
        const list = DOMUtils.select('#modalDocumentList');
        DOMUtils.empty(list);

        if (!documents || documents.length === 0) {
            DOMUtils.append(list, '<div class="no-documents">No documents found</div>');
            return;
        }

        documents.forEach(doc => {
            const itemHtml = `
                <div class="document-item" data-document-id="${doc.id}">
                    <div class="document-name">${doc.fileName || doc.title || 'Untitled'}</div>
                    <div class="document-path">${doc.libraryName || ''} / ${doc.folderPath || ''}</div>
                </div>
            `;

            DOMUtils.append(list, itemHtml);
        });
    }

    /**
     * Filter documents in modal
     */
    filterDocuments() {
        const query = DOMUtils.val('#documentSearch').toLowerCase();

        const items = DOMUtils.selectAll('.document-item');
        items.forEach(item => {
            const name = DOMUtils.text(DOMUtils.select('.document-name', item)).toLowerCase();
            const path = DOMUtils.text(DOMUtils.select('.document-path', item)).toLowerCase();

            if (name.includes(query) || path.includes(query)) {
                DOMUtils.show(item);
            } else {
                DOMUtils.hide(item);
            }
        });
    }

    /**
     * Select document item in modal
     */
    selectDocumentItem(item) {
        DOMUtils.removeClass('.document-item', 'selected');
        DOMUtils.addClass(item, 'selected');
        DOMUtils.prop('#selectDocumentBtn', 'disabled', false);
    }

    /**
     * Select document from modal
     */
    async selectDocumentFromModal() {
        const selectedElement = DOMUtils.select('.document-item.selected');
        const selectedId = selectedElement ? selectedElement.dataset.documentId : null;
        if (!selectedId) return;
        
        this.hideDocumentSelectionModal();
        await this.loadDocumentById(selectedId);
    }

    /**
     * Switch between tabs
     */
    switchTab(tabHeader) {
        const tabId = tabHeader.dataset.tab;

        // Update tab headers
        DOMUtils.removeClass('.tab-header', 'active');
        DOMUtils.addClass(tabHeader, 'active');

        // Update tab content
        DOMUtils.removeClass('.tab-content', 'active');
        DOMUtils.addClass(`#${tabId}Tab`, 'active');
        
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
        DOMUtils.val('#docTitle', this.documentMetadata.title || '');
        DOMUtils.val('#docDescription', this.documentMetadata.description || '');
        DOMUtils.val('#docTags', this.documentMetadata.tags || '');
        DOMUtils.val('#docAuthor', this.documentMetadata.author || '');
        DOMUtils.val('#docLanguage', this.documentMetadata.language || 'en');

        // Extended tab
        DOMUtils.val('#customId', this.documentMetadata.customId || '');
        DOMUtils.val('#docSource', this.documentMetadata.source || '');
        DOMUtils.val('#originalId', this.documentMetadata.originalId || '');
        DOMUtils.val('#docRecipient', this.documentMetadata.recipient || '');
        DOMUtils.val('#docObject', this.documentMetadata.object || '');
        DOMUtils.val('#docCoverage', this.documentMetadata.coverage || '');
        DOMUtils.val('#docType', this.documentMetadata.type || '');

        // Permissions tab
        DOMUtils.val('#securityLevel', this.documentMetadata.securityLevel || 'internal');
        DOMUtils.prop('#allowDownload', 'checked', this.documentMetadata.allowDownload !== false);
        DOMUtils.prop('#allowPrint', 'checked', this.documentMetadata.allowPrint !== false);
        DOMUtils.prop('#allowEdit', 'checked', this.documentMetadata.allowEdit === true);
    }

    /**
     * Show properties editor
     */
    showPropertiesEditor() {
        // Update document info
        const docName = this.currentDocument?.fileName || this.currentDocument?.title || 'Current Document';
        const status = this.currentDocument?.id ? 'Saved in Juptiar' : 'Not saved to Juptiar';

        DOMUtils.text('#currentDocName', docName);
        DOMUtils.text('#documentStatus', status);

        // Show editor sections
        DOMUtils.show('#tabsSection');
        DOMUtils.show('#actionSection');
        DOMUtils.hide('#loadingSection');
    }

    /**
     * Load permissions tab content
     */
    loadPermissionsTab() {
        const permissionsContainer = DOMUtils.select('#userPermissions');
        DOMUtils.empty(permissionsContainer);

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
                const itemHtml = `
                    <div class="permission-item">
                        <div class="permission-icon ${iconClass}"></div>
                        <span>${perm.name}: ${perm.granted ? 'Granted' : 'Denied'}</span>
                    </div>
                `;
                DOMUtils.append(permissionsContainer, itemHtml);
            });
        } else {
            DOMUtils.append(permissionsContainer, '<p>No permission information available</p>');
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
            title: DOMUtils.val('#docTitle').trim(),
            description: DOMUtils.val('#docDescription').trim(),
            tags: DOMUtils.val('#docTags').trim(),
            author: DOMUtils.val('#docAuthor').trim(),
            language: DOMUtils.val('#docLanguage'),
            customId: DOMUtils.val('#customId').trim(),
            source: DOMUtils.val('#docSource').trim(),
            originalId: DOMUtils.val('#originalId').trim(),
            recipient: DOMUtils.val('#docRecipient').trim(),
            object: DOMUtils.val('#docObject').trim(),
            coverage: DOMUtils.val('#docCoverage').trim(),
            type: DOMUtils.val('#docType'),
            securityLevel: DOMUtils.val('#securityLevel'),
            allowDownload: DOMUtils.prop('#allowDownload', 'checked'),
            allowPrint: DOMUtils.prop('#allowPrint', 'checked'),
            allowEdit: DOMUtils.prop('#allowEdit', 'checked')
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
        DOMUtils.text('#loadingSection p', message);
        DOMUtils.show('#loadingSection');
        DOMUtils.hide('#messageSection');
    }

    /**
     * Hide loading indicator
     */
    hideLoading() {
        DOMUtils.hide('#loadingSection');
    }

    /**
     * Show error message
     */
    showError(message) {
        const messageBar = DOMUtils.select('#messageBar');
        messageBar.className = 'ms-MessageBar ms-MessageBar--error';
        const messageIcon = DOMUtils.select('#messageIcon i');
        messageIcon.className = 'ms-Icon ms-Icon--ErrorBadge';
        DOMUtils.text('#messageText', message);
        DOMUtils.show('#messageSection');
        DOMUtils.hide('#loadingSection');
    }

    /**
     * Show success message
     */
    showSuccess(message) {
        const messageBar = DOMUtils.select('#messageBar');
        messageBar.className = 'ms-MessageBar ms-MessageBar--success';
        const messageIcon = DOMUtils.select('#messageIcon i');
        messageIcon.className = 'ms-Icon ms-Icon--CheckMark';
        DOMUtils.text('#messageText', message);
        DOMUtils.show('#messageSection');
        DOMUtils.hide('#loadingSection');

        // Auto-hide success message after 3 seconds
        setTimeout(() => {
            DOMUtils.hide('#messageSection');
        }, 3000);
    }

    /**
     * Show info message
     */
    showInfo(message) {
        const messageBar = DOMUtils.select('#messageBar');
        messageBar.className = 'ms-MessageBar ms-MessageBar--info';
        const messageIcon = DOMUtils.select('#messageIcon i');
        messageIcon.className = 'ms-Icon ms-Icon--Info';
        DOMUtils.text('#messageText', message);
        DOMUtils.show('#messageSection');
        DOMUtils.hide('#loadingSection');
    }
}

// Initialize when Office is ready
Office.onReady(() => {
    window.propertiesEditor = new PropertiesEditor();
});
