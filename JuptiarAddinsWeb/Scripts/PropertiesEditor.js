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
            // Force-hide legacy sections if cached HTML is still present
            try {
                const ds = document.getElementById('documentSection'); if (ds) ds.style.display = 'none';
                const ts = document.getElementById('tabsSection'); if (ts) ts.style.display = 'none';
            } catch (e) {}

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
        DOMUtils.hide('#metaSection');
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
        // Resolve Jupiter-managed document ID from DocumentStateManager first
        let jupiterId = null;
        try {
            const dsm = new DocumentStateManager();
            await dsm.initialize();
            const state = await dsm.getDocumentState();
            jupiterId = (state && state.documentId) ? state.documentId : (Office.context.document.settings.get('currentJuptiarDocumentId') || null);
        } catch (e) {
            try { jupiterId = Office.context.document.settings.get('currentJuptiarDocumentId') || null; } catch (_) { jupiterId = null; }
        }
        return new Promise((resolve, reject) => {
            Word.run(async (context) => {
                try {
                    const document = context.document;
                    const properties = document.properties;
                    
                    properties.load(['title', 'author', 'subject', 'keywords', 'category', 'comments']);
                    
                    await context.sync();
                    
                    this.currentDocument = {
                        id: jupiterId,
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
        // Default empty metadata
        this.documentMetadata = {
            title: '',
            description: '',
            tags: ''
        };

        // If document is from Jupiter, load metadata from backend (by Id)
        if (this.currentDocument && this.currentDocument.id) {
            try {
                const doc = await window.jupiterService.getDocument(this.currentDocument.id);
                this.currentDocument = doc;
                this.documentMetadata = {
                    title: doc.title || '',
                    description: doc.description || '',
                    tags: doc.tags || ''
                };
            } catch (error) {
                console.warn('Could not load Jupiter document by id:', error);
                try {
                    const messageBar = document.getElementById('messageBar');
                    if (messageBar) {
                        messageBar.className = 'ms-MessageBar ms-MessageBar--warning';
                        const messageIcon = document.querySelector('#messageIcon i');
                        if (messageIcon) messageIcon.className = 'ms-Icon ms-Icon--Info';
                        const messageText = document.getElementById('messageText');
                        if (messageText) messageText.textContent = 'This dialog is for existing documents only. Use "Save to Jupiter" to save new documents first.';
                        DOMUtils.show('#messageSection');
                    }
                } catch (_) {}
                DOMUtils.hide('#metaSection');
                DOMUtils.hide('#actionSection');
            }
        }
    }

    /**
     * Load document by ID from Juptiar
     */
    async loadDocumentById(documentId) {
        try {
            this.showLoading('Loading document from Juptiar...');
            
            const doc = await window.jupiterService.getDocument(documentId);
            this.currentDocument = doc;
            this.documentMetadata = {
                title: doc.title || '',
                description: doc.description || '',
                tags: doc.tags || ''
            };

            this.populatePropertiesForm();
            this.showPropertiesEditor();
            
            this.hideLoading();
            
        } catch (error) {
            console.warn('Could not load document from Juptiar:', error);
            try {
                const messageBar = document.getElementById('messageBar');
                if (messageBar) {
                    messageBar.className = 'ms-MessageBar ms-MessageBar--warning';
                    const messageIcon = document.querySelector('#messageIcon i');
                    if (messageIcon) messageIcon.className = 'ms-Icon ms-Icon--Info';
                    const messageText = document.getElementById('messageText');
                    if (messageText) messageText.textContent = 'This dialog is for existing documents only. Use "Save to Jupiter" to save new documents first.';
                    DOMUtils.show('#messageSection');
                }
            } catch (_) {}
            DOMUtils.hide('#metaSection');
            DOMUtils.hide('#actionSection');
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
                    <div class="document-name">${doc.name || doc.fileName || doc.title || 'Untitled'}</div>
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

        // Minimal fields only
        DOMUtils.val('#docTitle', this.documentMetadata.title || '');
        DOMUtils.val('#docDescription', this.documentMetadata.description || '');
        DOMUtils.val('#docTags', this.documentMetadata.tags || '');
    }

    /**
     * Show properties editor
     */
    showPropertiesEditor() {
        // Show editor sections
        DOMUtils.show('#metaSection');
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
            tags: DOMUtils.val('#docTags').trim()
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
                    properties.subject = metadata.description || '';
                    properties.keywords = metadata.tags || '';

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
