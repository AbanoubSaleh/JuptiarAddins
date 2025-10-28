/**
 * Document Browser - Main functionality for browsing and opening documents
 * Integrates with Jupiter document management system
 */

class DocumentBrowser {
    constructor() {
        this.currentLibrary = null;
        this.currentFolder = null;
        this.documents = [];
        this.selectedDocument = null;
        this.folderTree = [];

        this.initializeEventListeners();

        // Listen for authentication state changes from AuthManager
        window.addEventListener('juptiarAuthStateChanged', (e) => {
            console.log('DocumentBrowser: Received auth state change:', e.detail);
            this.handleAuthStateChange(e.detail);
        });

        // Check authentication status (but don't rely on it being accurate yet)
        this.checkAuthenticationStatus();

        // Show the modal dialog
        this.showDialog();
    }

    /**
     * Initialize event listeners
     */
    initializeEventListeners() {
        // Authentication events
        $.on('#loginBtn', 'click', () => this.showLoginModal());
        $.on('#logoutBtn', 'click', () => this.logout());
        $.on('#loginSubmitBtn', 'click', () => this.handleLogin());
        $.on('#loginCancelBtn', 'click', () => this.hideLoginModal());
        $.on('#closeModal', 'click', () => this.hideLoginModal());

        // Dialog controls
        $.on('#closeDialogBtn', 'click', () => this.closeDialog());

        // Search and refresh
        $.on('#fileSearchInput', 'keypress', (e) => {
            if (e.which === 13 || e.keyCode === 13) this.performFileSearch();
        });
        $.on('#fullTextSearchInput', 'keypress', (e) => {
            if (e.which === 13 || e.keyCode === 13) this.performFullTextSearch();
        });
        $.on('#refreshBtn', 'click', () => this.refreshCurrentView());

        // Action buttons
        $.on('#openBtn', 'click', () => this.openSelectedDocument());
        $.on('#cancelBtn', 'click', () => this.closeDialog());

        // Close dialog button
        $.on('#closeDialogBtn', 'click', () => this.closeDialog());

        // Context menu
        $.delegate(document, 'contextmenu', '.document-row', (e) => {
            e.preventDefault();
            this.showContextMenu(e, e.currentTarget);
        });

        // Hide context menu on click elsewhere
        $.on(document, 'click', () => this.hideContextMenu());

        // Context menu actions
        $.on('#openDocument', 'click', async () => {
            try {
                await this.openSelectedDocument();
            } catch (error) {
                console.error('Error opening document:', error);
                this.showError('Failed to open document: ' + error.message);
            }
        });
        $.on('#editDocument', 'click', async () => {
            try {
                await this.editSelectedDocument();
            } catch (error) {
                console.error('Error editing document:', error);
                this.showError('Failed to edit document: ' + error.message);
            }
        });
        $.on('#deleteDocument', 'click', async () => {
            try {
                await this.deleteSelectedDocument();
            } catch (error) {
                console.error('Error deleting document:', error);
                this.showError('Failed to delete document: ' + error.message);
            }
        });
        $.on('#editProperties', 'click', async () => {
            try {
                await this.editDocumentProperties();
            } catch (error) {
                console.error('Error editing properties:', error);
                this.showError('Failed to edit properties: ' + error.message);
            }
        });

        // Document selection
        $.delegate(document, 'click', '.document-row', (e) => {
            this.selectDocument(e.currentTarget);
        });

        // Double-click to open document
        $.delegate(document, 'dblclick', '.document-row', (e) => {
            this.openSelectedDocument();
        });

        // Folder expand/collapse and selection
        $.delegate(document, 'click', '.expand-icon', (e) => {
            e.stopPropagation();
            this.toggleFolderExpansion(e.currentTarget.parentElement);
        });

        $.delegate(document, 'click', '.folder-item', (e) => {
            // If clicking on expand icon, don't select folder
            if (e.target.classList.contains('expand-icon')) {
                return;
            }

            const folderItem = e.currentTarget;
            const type = folderItem.getAttribute('data-type');

            if (type === 'library') {
                // For libraries, toggle expansion
                this.toggleLibraryExpansion(folderItem);
            } else {
                // For folders, select and load documents
                this.selectFolder(folderItem);
            }
        });

        // Listen for authentication state changes
        window.addEventListener('juptiarAuthStateChanged', (e) => {
            this.handleAuthStateChange(e.detail);
        });
    }

    /**
     * Check current authentication status
     */
    async checkAuthenticationStatus() {
        try {
            console.log('DocumentBrowser: Checking authentication status...');

            if (!window.authManager) {
                console.error('DocumentBrowser: AuthManager not available!');
                this.handleAuthStateChange({ isAuthenticated: false });
                return;
            }

            // Wait a bit for AuthManager to finish initialization if it's still initializing
            let retries = 0;
            while (retries < 10 && window.authManager && !window.authManager.isInitialized) {
                console.log('DocumentBrowser: Waiting for AuthManager to initialize...');
                await new Promise(resolve => setTimeout(resolve, 100));
                retries++;
            }

            const authStatus = window.authManager.getAuthStatus();
            console.log('DocumentBrowser: Auth status:', authStatus);

            // Also check if JupiterService has the token
            if (authStatus.isAuthenticated && window.jupiterService) {
                const token = window.authManager.getToken();
                if (token) {
                    window.jupiterService.setAuthToken(token);
                    console.log('DocumentBrowser: Token set in JupiterService');
                } else {
                    console.warn('DocumentBrowser: Auth status shows authenticated but no token found');
                    authStatus.isAuthenticated = false;
                }
            }

            this.handleAuthStateChange(authStatus);
        } catch (error) {
            console.error('DocumentBrowser: Error checking auth status:', error);
            this.handleAuthStateChange({ isAuthenticated: false });
        }
    }

    /**
     * Handle authentication state changes
     */
    handleAuthStateChange(authStatus) {
        if (authStatus.isAuthenticated) {
            this.showAuthenticatedState(authStatus.user);
            // Only load library tree if we have a valid token
            if (window.jupiterService && window.jupiterService.authToken) {
                this.loadLibraryTree();
            }
        } else {
            this.showUnauthenticatedState();
        }
    }

    /**
     * Show authenticated state
     */
    showAuthenticatedState(user) {
        $.text('#authStatusText', `Logged in as: ${user.username || 'User'}`);
        $.hide('#loginBtn');
        $.show('#logoutBtn');
        $.show('#searchSection');
        $.show('#mainContent');
        $.hide('#loadingSection');

        $.removeClass('.status-indicator', 'offline');
        $.addClass('.status-indicator', 'online');
    }

    /**
     * Show unauthenticated state
     */
    showUnauthenticatedState() {
        $.text('#authStatusText', 'Not authenticated - Please log in to access libraries');
        $.show('#loginBtn');
        $.hide('#logoutBtn');
        $.hide('#searchSection');
        $.hide('#mainContent');
        $.hide('#loadingSection');

        $.removeClass('.status-indicator', 'online');
        $.addClass('.status-indicator', 'offline');

        // Clear any existing tree data
        $.empty('#folderTree');
        $.empty('#documentTableBody');
    }

    /**
     * Close the document browser dialog
     */
    closeDialog() {
        const overlay = $.select('#dialogOverlay');
        if (overlay) {
            $.hide(overlay);
        }
    }

    /**
     * Show login modal
     */
    showLoginModal() {
        const credentials = window.authManager.getStoredCredentials();

        $.val('#username', credentials.username);
        $.show('#loginModal');
        $.focus('#username');
    }

    /**
     * Hide login modal
     */
    hideLoginModal() {
        $.hide('#loginModal');
        $.val('#username', '');
        $.val('#password', '');
    }

    /**
     * Handle login form submission
     */
    async handleLogin() {
        try {
            const username = $.val('#username').trim();
            const password = $.val('#password');

            if (!username || !password) {
                this.showError('Please enter username and password');
                return;
            }

            // Attempt login (server URL is pre-configured)
            $.prop('#loginSubmitBtn', 'disabled', true);
            $.text('#loginSubmitBtn', 'Logging in...');

            await window.authManager.login(username, password, true);

            // Ensure token is set in JupiterService
            const token = window.authManager.getToken();
            if (token && window.jupiterService) {
                window.jupiterService.setAuthToken(token);
                console.log('Token set in JupiterService after login');
            }

            this.hideLoginModal();
            this.showSuccess('Successfully logged in');

        } catch (error) {
            console.error('Login error:', error);
            this.showError(error.message || 'Login failed');
        } finally {
            $.prop('#loginSubmitBtn', 'disabled', false);
            $.text('#loginSubmitBtn', 'Login');
        }
    }

    /**
     * Logout current user
     */
    async logout() {
        try {
            await window.authManager.logout();
            this.showSuccess('Successfully logged out');
        } catch (error) {
            console.error('Logout error:', error);
            this.showError('Logout failed');
        }
    }

    /**
     * Load library tree structure
     */
    async loadLibraryTree() {
        try {
            // Check if we have authentication before making the request
            if (!window.jupiterService || !window.jupiterService.authToken) {
                console.log('No authentication token available, skipping library tree load');
                this.showError('Please log in to access libraries');
                return;
            }

            this.showLoading('Loading libraries...');

            const treeData = await window.jupiterService.getLibraryTree();
            this.folderTree = treeData;

            this.renderFolderTree(treeData);

            // Auto-select default library if configured
            await this.selectDefaultLibrary();

            this.hideLoading();

        } catch (error) {
            console.error('Error loading library tree:', error);
            if (error.message.includes('401') || error.message.includes('Unauthorized')) {
                this.showError('Authentication required. Please log in to access libraries.');
            } else {
                this.showError('Failed to load libraries: ' + error.message);
            }
        }
    }

    /**
     * Render folder tree with collapsible functionality
     */
    renderFolderTree(treeData) {
        const treeContainer = $.select('#folderTree');
        $.empty(treeContainer);

        if (Array.isArray(treeData)) {
            treeData.forEach(library => this.renderLibrary(library, treeContainer));
        }
    }

    /**
     * Render a library with collapsible folders
     */
    renderLibrary(library, container) {
        console.log('Rendering library with new code:', library.name);
        const hasChildren = library.children && library.children.length > 0;
        const expandIcon = hasChildren ? '▶' : '';

        // Use folder icon for libraries as requested
        const libraryItem = document.createElement('div');
        libraryItem.className = 'folder-item library-item';
        libraryItem.setAttribute('data-library-id', library.id);
        libraryItem.setAttribute('data-folder-id', '');
        libraryItem.setAttribute('data-type', 'library');
        libraryItem.setAttribute('data-node-id', library.id);
        libraryItem.setAttribute('data-has-children', hasChildren);
        libraryItem.setAttribute('data-expanded', 'false');

        libraryItem.innerHTML = `
            <span class="expand-icon">${expandIcon}</span>
            <span class="folder-icon">📁</span>
            <span class="folder-name">${library.name}</span>
        `;

        container.appendChild(libraryItem);

        // Create children container (initially hidden)
        if (hasChildren) {
            const childrenContainer = document.createElement('div');
            childrenContainer.className = 'children-container';
            childrenContainer.setAttribute('data-parent-id', library.id);
            childrenContainer.style.display = 'none';

            library.children.forEach(folder => {
                this.renderFolder(folder, childrenContainer, library.id, 1);
            });

            container.appendChild(childrenContainer);
        }
    }

    /**
     * Render a folder
     */
    renderFolder(folder, container, libraryId, level) {
        const hasChildren = folder.children && folder.children.length > 0;
        const expandIcon = hasChildren ? '▶' : '';

        const folderItem = document.createElement('div');
        folderItem.className = 'folder-item';
        folderItem.setAttribute('data-library-id', libraryId);
        folderItem.setAttribute('data-folder-id', folder.id);
        folderItem.setAttribute('data-type', 'folder');
        folderItem.setAttribute('data-node-id', folder.id);
        folderItem.setAttribute('data-has-children', hasChildren);
        folderItem.setAttribute('data-expanded', 'false');
        folderItem.style.marginLeft = (level * 16) + 'px';

        folderItem.innerHTML = `
            <span class="expand-icon">${expandIcon}</span>
            <span class="folder-icon">📁</span>
            <span class="folder-name">${folder.name}</span>
        `;

        container.appendChild(folderItem);

        // Add children if they exist
        if (hasChildren) {
            const childrenContainer = document.createElement('div');
            childrenContainer.className = 'children-container';
            childrenContainer.setAttribute('data-parent-id', folder.id);
            childrenContainer.style.display = 'none';

            folder.children.forEach(child => {
                this.renderFolder(child, childrenContainer, libraryId, level + 1);
            });

            container.appendChild(childrenContainer);
        }
    }

    /**
     * Toggle library expansion (show/hide folders)
     */
    toggleLibraryExpansion(libraryItem) {
        const nodeId = libraryItem.getAttribute('data-node-id');
        const hasChildren = libraryItem.getAttribute('data-has-children') === 'true';
        const isExpanded = libraryItem.getAttribute('data-expanded') === 'true';

        if (!hasChildren) {
            return; // No children to expand
        }

        const childrenContainer = $.select(`.children-container[data-parent-id="${nodeId}"]`);
        const expandIcon = $.find(libraryItem, '.expand-icon')[0];

        if (isExpanded) {
            // Collapse
            $.slideUp(childrenContainer, 200);
            $.text(expandIcon, '▶');
            libraryItem.setAttribute('data-expanded', 'false');
        } else {
            // Expand
            $.slideDown(childrenContainer, 200);
            $.text(expandIcon, '▼');
            libraryItem.setAttribute('data-expanded', 'true');
        }
    }

    /**
     * Toggle folder expansion (show/hide subfolders)
     */
    toggleFolderExpansion(folderItem) {
        const nodeId = folderItem.getAttribute('data-node-id');
        const hasChildren = folderItem.getAttribute('data-has-children') === 'true';
        const isExpanded = folderItem.getAttribute('data-expanded') === 'true';

        if (!hasChildren) {
            return; // No children to expand
        }

        const childrenContainer = $.select(`.children-container[data-parent-id="${nodeId}"]`);
        const expandIcon = $.find(folderItem, '.expand-icon')[0];

        if (isExpanded) {
            // Collapse
            $.slideUp(childrenContainer, 200);
            $.text(expandIcon, '▶');
            folderItem.setAttribute('data-expanded', 'false');
        } else {
            // Expand
            $.slideDown(childrenContainer, 200);
            $.text(expandIcon, '▼');
            folderItem.setAttribute('data-expanded', 'true');
        }
    }

    /**
     * Select a folder and load its documents
     */
    async selectFolder(folderItem) {
        try {
            console.log('Folder selected:', folderItem);

            // Update UI selection
            $.removeClass('.folder-item', 'selected');
            $.addClass(folderItem, 'selected');

            const libraryId = folderItem.getAttribute('data-library-id');
            const folderId = folderItem.getAttribute('data-folder-id');
            const type = folderItem.getAttribute('data-type');

            console.log('Selected folder data:', { libraryId, folderId, type });

            this.currentLibrary = libraryId;
            this.currentFolder = folderId || null;

            // Only load documents for folders (not libraries)
            if (type === 'folder' && folderId) {
                console.log('Loading documents for folder:', folderId);
                await this.loadDocumentsByFolder(folderId);
            } else {
                console.log('No documents to load - not a folder or no folder ID');
                // Clear document list for libraries
                this.renderDocumentList([]);
            }

        } catch (error) {
            console.error('Error selecting folder:', error);
            this.showError('Failed to load folder contents: ' + error.message);
        }
    }

    /**
     * Load documents for a specific folder using the folder endpoint
     */
    async loadDocumentsByFolder(folderId) {
        try {
            console.log('Loading documents for folder ID:', folderId);
            this.showLoading('Loading documents...');

            const documents = await window.jupiterService.getDocumentsByFolder(folderId);
            console.log('Documents response:', documents);

            this.documents = documents || [];
            console.log('Processed documents:', this.documents);

            this.renderDocumentList(this.documents);
            this.hideLoading();

        } catch (error) {
            console.error('Error loading documents by folder:', error);
            this.showError('Failed to load documents: ' + error.message);
        }
    }

    /**
     * Load documents for the selected library/folder
     */
    async loadDocuments(libraryId, folderId = null) {
        try {
            this.showLoading('Loading documents...');
            
            const response = await window.jupiterService.getDocuments(libraryId, folderId);
            this.documents = response.documents || [];
            
            this.renderDocumentList(this.documents);
            this.hideLoading();
            
        } catch (error) {
            console.error('Error loading documents:', error);
            this.showError('Failed to load documents: ' + error.message);
        }
    }

    /**
     * Render document list in the table
     */
    renderDocumentList(documents) {
        const tbody = $.select('#documentTableBody');
        $.empty(tbody);

        if (!documents || documents.length === 0) {
            $.append(tbody, '<tr><td colspan="4" class="text-center">No documents found</td></tr>');
            return;
        }

        documents.forEach(doc => {
            const row = document.createElement('tr');
            row.className = 'document-row';
            row.setAttribute('data-document-id', doc.id);

            row.innerHTML = `
                <td>
                    <span class="file-icon">${this.getFileTypeIcon(doc.fileName)}</span>
                </td>
                <td>${doc.fileName || 'Untitled'}</td>
                <td>${this.formatDate(doc.modifiedOn || doc.dateModified)}</td>
                <td>${this.formatFileSize(doc.size)}</td>
            `;

            tbody.appendChild(row);
        });

        // Clear selection and disable open button
        this.selectedDocument = null;
        $.prop('#openBtn', 'disabled', true);
    }

    /**
     * Get file type icon based on extension
     */
    getFileTypeIcon(fileName) {
        if (!fileName) return '📄';

        const ext = fileName.split('.').pop().toLowerCase();
        const icons = {
            'doc': '📝', 'docx': '📝',
            'pdf': '📕',
            'xls': '📊', 'xlsx': '📊',
            'ppt': '📊', 'pptx': '📊',
            'txt': '📄',
            'jpg': '🖼️', 'jpeg': '🖼️', 'png': '🖼️', 'gif': '🖼️'
        };

        return icons[ext] || '📄';
    }

    /**
     * Format date for display
     */
    formatDate(dateString) {
        if (!dateString) return '-';
        
        try {
            const date = new Date(dateString);
            return date.toLocaleDateString() + ' ' + date.toLocaleTimeString();
        } catch (error) {
            return dateString;
        }
    }

    /**
     * Format file size for display
     */
    formatFileSize(bytes) {
        if (!bytes || bytes === 0) return '-';
        
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(1024));
        return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
    }

    /**
     * Auto-select default library if configured
     */
    async selectDefaultLibrary() {
        try {
            if (!window.authManager) return;

            const settings = await window.authManager.getSettings();
            const defaultLibrary = settings.defaultLibrary;

            if (defaultLibrary && this.folderTree) {
                console.log('DocumentBrowser: Auto-selecting default library:', defaultLibrary);

                // Find the default library in the tree
                const libraryNode = this.findLibraryInTree(this.folderTree, defaultLibrary);
                if (libraryNode) {
                    // Simulate clicking on the library
                    this.selectFolder(libraryNode);

                    // Expand the library node
                    const libraryElement = $.select(`[data-folder-id="${libraryNode.id}"]`);
                    if (libraryElement) {
                        $.addClass(libraryElement, 'selected');
                        // Expand if it has children
                        if (libraryNode.children && libraryNode.children.length > 0) {
                            const folderToggle = $.find(libraryElement, '.folder-toggle')[0];
                            if (folderToggle) {
                                folderToggle.click();
                            }
                        }
                    }

                    console.log('DocumentBrowser: Default library selected successfully');
                }
            }
        } catch (error) {
            console.warn('DocumentBrowser: Could not select default library:', error);
        }
    }

    /**
     * Find library in tree by name or ID
     */
    findLibraryInTree(tree, libraryIdentifier) {
        if (!Array.isArray(tree)) return null;

        for (const node of tree) {
            // Check if this node matches (by name or ID)
            if (node.name === libraryIdentifier || node.id === libraryIdentifier) {
                return node;
            }

            // Search in children recursively
            if (node.children) {
                const found = this.findLibraryInTree(node.children, libraryIdentifier);
                if (found) return found;
            }
        }

        return null;
    }

    /**
     * Show loading indicator
     */
    showLoading(message = 'Loading...') {
        $.text('#loadingMessage', message);
        $.show('#loadingSection');
        $.hide('#errorSection');
    }

    /**
     * Hide loading indicator
     */
    hideLoading() {
        $.hide('#loadingSection');
    }

    /**
     * Show error message
     */
    showError(message) {
        $.text('#errorMessage', message);
        $.show('#errorSection');
        $.hide('#loadingSection');
    }

    /**
     * Show success message
     */
    showSuccess(message) {
        // For now, just log to console. Could add a success overlay later
        console.log('Success:', message);
        // Could show a temporary success message in the dialog
    }

    /**
     * Show success message
     */
    showSuccess(message) {
        // You can implement a success message display here
        console.log('Success:', message);
    }

    /**
     * Refresh current view
     */
    async refreshCurrentView() {
        // First re-check authentication status
        await this.checkAuthenticationStatus();

        // Then refresh the current view if authenticated
        if (this.currentLibrary) {
            await this.loadDocuments(this.currentLibrary, this.currentFolder);
        } else {
            await this.loadLibraryTree();
        }
    }

    /**
     * Perform search
     */
    async performSearch() {
        const query = $.val('#fileSearchInput').trim();
        if (!query) {
            this.refreshCurrentView();
            return;
        }

        try {
            this.showLoading('Searching...');

            const results = await window.jupiterService.searchDocuments(query);
            this.renderDocumentList(results.documents || []);

            this.hideLoading();
        } catch (error) {
            console.error('Search error:', error);
            this.showError('Search failed: ' + error.message);
        }
    }

    /**
     * Perform file name search
     */
    async performFileSearch() {
        const query = $.val('#fileSearchInput').trim();
        if (!query) {
            this.refreshCurrentView();
            return;
        }

        try {
            this.showLoading('Searching files...');

            const results = await window.jupiterService.searchDocuments(query, 'filename');
            this.renderDocuments(results);

            this.hideLoading();
        } catch (error) {
            console.error('File search failed:', error);
            this.showError('File search failed: ' + error.message);
        }
    }

    /**
     * Perform full-text search
     */
    async performFullTextSearch() {
        const query = $.val('#fullTextSearchInput').trim();
        if (!query) {
            this.showError('Please enter a search term for full-text search');
            return;
        }

        try {
            this.showLoading('Searching content...');

            const results = await window.jupiterService.searchDocuments(query, 'fulltext');
            this.renderDocuments(results);

            this.hideLoading();
        } catch (error) {
            console.error('Full-text search failed:', error);
            this.showError('Full-text search failed: ' + error.message);
        }
    }

    /**
     * Select a document row
     */
    selectDocument(row) {
        $.removeClass('.document-row', 'selected');
        $.addClass(row, 'selected');
        this.selectedDocument = row.getAttribute('data-document-id');

        // Enable open button when document is selected
        $.prop('#openBtn', 'disabled', false);
    }

    /**
     * Show the modal dialog
     */
    showDialog() {
        $.show('#dialogOverlay');
        $.show('#openDocumentDialog');
    }

    /**
     * Close the dialog/task pane
     */
    closeDialog() {
        $.hide('#dialogOverlay');
        $.hide('#openDocumentDialog');
        console.log('Dialog closed');
    }

    /**
     * Show context menu for document
     */
    showContextMenu(event, row) {
        this.selectDocument(row);

        const documentId = $.data(row, 'document-id');
        const document = this.documents.find(doc => doc.id === documentId);

        // Enable/disable menu items based on permissions
        $.toggle('#editDocument', document && document.canEdit);
        $.toggle('#deleteDocument', document && document.canDelete);

        const contextMenu = $.select('#contextMenu');
        if (contextMenu) {
            contextMenu.style.display = 'block';
            contextMenu.style.left = event.pageX + 'px';
            contextMenu.style.top = event.pageY + 'px';
        }
    }

    /**
     * Hide context menu
     */
    hideContextMenu() {
        $.hide('#contextMenu');
    }

    /**
     * Open selected document in Word
     */
    async openSelectedDocument() {
        if (!this.selectedDocument) return;
        await this.openDocument(this.selectedDocument);
    }

    /**
     * Open document by ID
     */
    async openDocument(documentId) {
        try {
            this.showLoading('Opening document...');

            // Download document content
            const blob = await window.jupiterService.downloadDocument(documentId);

            // Convert blob to base64
            const reader = new FileReader();
            reader.onload = async () => {
                try {
                    const base64Data = reader.result.split(',')[1];

                    // Insert document content into Word
                    await Word.run(async (context) => {
                        // Clear current document
                        context.document.body.clear();

                        // Insert the document content
                        context.document.body.insertFileFromBase64(base64Data, Word.InsertLocation.start);

                        await context.sync();
                    });

                    this.hideLoading();
                    this.showSuccess('Document opened successfully');

                } catch (error) {
                    console.error('Error inserting document:', error);
                    this.showError('Failed to open document in Word');
                }
            };

            reader.readAsDataURL(blob);

        } catch (error) {
            console.error('Error opening document:', error);
            this.showError('Failed to open document: ' + error.message);
        }
    }

    /**
     * Edit selected document
     */
    async editSelectedDocument() {
        if (!this.selectedDocument) return;
        await this.editDocument(this.selectedDocument);
    }

    /**
     * Edit document by ID
     */
    async editDocument(documentId) {
        try {
            // First open the document
            await this.openDocument(documentId);

            // Store document ID for saving later
            Office.context.document.settings.set('currentJuptiarDocumentId', documentId);
            await Office.context.document.settings.saveAsync();

            this.showSuccess('Document opened for editing. Use Save to update in Juptiar.');

        } catch (error) {
            console.error('Error editing document:', error);
            this.showError('Failed to edit document: ' + error.message);
        }
    }

    /**
     * Delete selected document
     */
    async deleteSelectedDocument() {
        if (!this.selectedDocument) return;
        await this.deleteDocument(this.selectedDocument);
    }

    /**
     * Delete document by ID
     */
    async deleteDocument(documentId) {
        try {
            const document = this.documents.find(doc => doc.id === documentId);
            if (!document) {
                this.showError('Document not found');
                return;
            }

            const confirmed = confirm(`Are you sure you want to delete "${document.fileName}"?`);
            if (!confirmed) return;

            this.showLoading('Deleting document...');

            await window.jupiterService.deleteDocument(documentId);

            // Refresh the document list
            await this.refreshCurrentView();

            this.showSuccess('Document deleted successfully');

        } catch (error) {
            console.error('Error deleting document:', error);
            this.showError('Failed to delete document: ' + error.message);
        }
    }

    /**
     * Edit document properties
     */
    async editDocumentProperties() {
        if (!this.selectedDocument) return;

        try {
            // Store the document ID and open properties dialog
            Office.context.document.settings.set('editPropertiesDocumentId', this.selectedDocument);
            await Office.context.document.settings.saveAsync();

            // Open properties task pane
            Office.ribbon.requestUpdate({
                tabs: [{
                    id: "Juptiar.Tab",
                    controls: [{
                        id: "Juptiar.PropertiesButton",
                        enabled: true
                    }]
                }]
            });

            this.showSuccess('Opening properties editor...');

        } catch (error) {
            console.error('Error opening properties:', error);
            this.showError('Failed to open properties editor');
        }
    }
}

// Initialize when Office is ready
Office.onReady(() => {
    console.log('Office is ready, initializing DocumentBrowser...');

    // Ensure all dependencies are loaded
    if (typeof JupiterConfig === 'undefined') {
        console.error('JupiterConfig not loaded!');
        return;
    }

    if (typeof JupiterService === 'undefined') {
        console.error('JupiterService not loaded!');
        return;
    }

    if (typeof AuthManager === 'undefined') {
        console.error('AuthManager not loaded!');
        return;
    }

    // Initialize global instances
    window.jupiterConfig = JupiterConfig; // JupiterConfig is an object, not a constructor

    // Initialize JupiterConfig first (this might not have been called)
    if (typeof window.JupiterConfig.init === 'function') {
        window.JupiterConfig.init();
    }

    // Initialize global services if not already done
    if (!window.jupiterService) {
        window.jupiterService = new JupiterService();

        // Debug configuration values
        const baseUrl = window.JupiterConfig.get('server.baseUrl');
        const apiEndpoint = window.JupiterConfig.get('server.apiEndpoint') || '/api';
        const timeout = window.JupiterConfig.get('server.timeout') || 30000;

        console.log('DocumentBrowser: Initializing JupiterService with config:', {
            baseUrl: baseUrl,
            apiEndpoint: apiEndpoint,
            timeout: timeout
        });

        window.jupiterService.initialize({
            serverUrl: baseUrl,
            apiEndpoint: apiEndpoint,
            timeout: timeout
        });

        console.log('DocumentBrowser: JupiterService initialized. Current baseUrl:', window.jupiterService.baseUrl);
    }

    if (!window.authManager) {
        window.authManager = new AuthManager();
        // Initialize AuthManager asynchronously and then create DocumentBrowser
        window.authManager.initialize().then(() => {
            console.log('AuthManager initialized, creating DocumentBrowser...');
            if (!window.documentBrowser) {
                window.documentBrowser = new DocumentBrowser();
            }
        }).catch(error => {
            console.error('Failed to initialize AuthManager:', error);
            // Still create DocumentBrowser even if AuthManager fails
            if (!window.documentBrowser) {
                window.documentBrowser = new DocumentBrowser();
            }
        });
    } else {
        // AuthManager already exists, create DocumentBrowser immediately
        console.log('Dependencies loaded, creating DocumentBrowser...');
        if (!window.documentBrowser) {
            window.documentBrowser = new DocumentBrowser();
        }
    }
});
