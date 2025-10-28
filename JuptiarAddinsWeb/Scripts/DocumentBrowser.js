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
        $('#loginBtn').on('click', () => this.showLoginModal());
        $('#logoutBtn').on('click', () => this.logout());
        $('#loginSubmitBtn').on('click', () => this.handleLogin());
        $('#loginCancelBtn').on('click', () => this.hideLoginModal());
        $('#closeModal').on('click', () => this.hideLoginModal());

        // Dialog controls
        $('#closeDialogBtn').on('click', () => this.closeDialog());

        // Search and refresh
        $('#fileSearchInput').on('keypress', (e) => {
            if (e.which === 13) this.performFileSearch();
        });
        $('#fullTextSearchInput').on('keypress', (e) => {
            if (e.which === 13) this.performFullTextSearch();
        });
        $('#refreshBtn').on('click', () => this.refreshCurrentView());

        // Action buttons
        $('#openBtn').on('click', () => this.openSelectedDocument());
        $('#cancelBtn').on('click', () => this.closeDialog());

        // Context menu
        $(document).on('contextmenu', '.document-row', (e) => {
            e.preventDefault();
            this.showContextMenu(e, $(e.currentTarget));
        });

        // Hide context menu on click elsewhere
        $(document).on('click', () => this.hideContextMenu());

        // Context menu actions
        $('#openDocument').on('click', () => this.openSelectedDocument());
        $('#editDocument').on('click', () => this.editSelectedDocument());
        $('#deleteDocument').on('click', () => this.deleteSelectedDocument());
        $('#editProperties').on('click', () => this.editDocumentProperties());

        // Document selection
        $(document).on('click', '.document-row', (e) => {
            this.selectDocument($(e.currentTarget));
        });

        // Double-click to open document
        $(document).on('dblclick', '.document-row', (e) => {
            this.openSelectedDocument();
        });

        // Folder expand/collapse and selection
        $(document).on('click', '.expand-icon', (e) => {
            e.stopPropagation();
            this.toggleFolderExpansion($(e.currentTarget).parent());
        });

        $(document).on('click', '.folder-item', (e) => {
            // If clicking on expand icon, don't select folder
            if ($(e.target).hasClass('expand-icon')) {
                return;
            }

            const $folderItem = $(e.currentTarget);
            const type = $folderItem.data('type');

            if (type === 'library') {
                // For libraries, toggle expansion
                this.toggleLibraryExpansion($folderItem);
            } else {
                // For folders, select and load documents
                this.selectFolder($folderItem);
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
        $('#authStatusText').text(`Logged in as: ${user.username || 'User'}`);
        $('#loginBtn').hide();
        $('#logoutBtn').show();
        $('#searchSection').show();
        $('#mainContent').show();
        $('#loadingSection').hide();
        
        $('.status-indicator').removeClass('offline').addClass('online');
    }

    /**
     * Show unauthenticated state
     */
    showUnauthenticatedState() {
        $('#authStatusText').text('Not authenticated - Please log in to access libraries');
        $('#loginBtn').show();
        $('#logoutBtn').hide();
        $('#searchSection').hide();
        $('#mainContent').hide();
        $('#loadingSection').hide();

        $('.status-indicator').removeClass('online').addClass('offline');

        // Clear any existing tree data
        $('#folderTree').empty();
        $('#documentTableBody').empty();
    }

    /**
     * Show login modal
     */
    showLoginModal() {
        const credentials = window.authManager.getStoredCredentials();

        $('#username').val(credentials.username);
        $('#loginModal').show();
        $('#username').focus();
    }

    /**
     * Hide login modal
     */
    hideLoginModal() {
        $('#loginModal').hide();
        $('#username').val('');
        $('#password').val('');
    }

    /**
     * Handle login form submission
     */
    async handleLogin() {
        try {
            const username = $('#username').val().trim();
            const password = $('#password').val();

            if (!username || !password) {
                this.showError('Please enter username and password');
                return;
            }

            // Attempt login (server URL is pre-configured)
            $('#loginSubmitBtn').prop('disabled', true).text('Logging in...');

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
            $('#loginSubmitBtn').prop('disabled', false).text('Login');
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
        const $treeContainer = $('#folderTree');
        $treeContainer.empty();

        if (Array.isArray(treeData)) {
            treeData.forEach(library => this.renderLibrary(library, $treeContainer));
        }
    }

    /**
     * Render a library with collapsible folders
     */
    renderLibrary(library, $container) {
        console.log('Rendering library with new code:', library.name);
        const hasChildren = library.children && library.children.length > 0;
        const expandIcon = hasChildren ? '▶' : '';

        // Use folder icon for libraries as requested
        const $libraryItem = $(`
            <div class="folder-item library-item"
                 data-library-id="${library.id}"
                 data-folder-id=""
                 data-type="library"
                 data-node-id="${library.id}"
                 data-has-children="${hasChildren}"
                 data-expanded="false">
                <span class="expand-icon">${expandIcon}</span>
                <span class="folder-icon">📁</span>
                <span class="folder-name">${library.name}</span>
            </div>
        `);

        $container.append($libraryItem);

        // Create children container (initially hidden)
        if (hasChildren) {
            const $childrenContainer = $(`
                <div class="children-container"
                     data-parent-id="${library.id}"
                     style="display: none;">
                </div>
            `);

            library.children.forEach(folder => {
                this.renderFolder(folder, $childrenContainer, library.id, 1);
            });

            $container.append($childrenContainer);
        }
    }

    /**
     * Render a folder
     */
    renderFolder(folder, $container, libraryId, level) {
        const hasChildren = folder.children && folder.children.length > 0;
        const expandIcon = hasChildren ? '▶' : '';

        const $folderItem = $(`
            <div class="folder-item"
                 data-library-id="${libraryId}"
                 data-folder-id="${folder.id}"
                 data-type="folder"
                 data-node-id="${folder.id}"
                 data-has-children="${hasChildren}"
                 data-expanded="false"
                 style="margin-left: ${level * 16}px">
                <span class="expand-icon">${expandIcon}</span>
                <span class="folder-icon">📁</span>
                <span class="folder-name">${folder.name}</span>
            </div>
        `);

        $container.append($folderItem);

        // Add children if they exist
        if (hasChildren) {
            const $childrenContainer = $(`
                <div class="children-container"
                     data-parent-id="${folder.id}"
                     style="display: none;">
                </div>
            `);

            folder.children.forEach(child => {
                this.renderFolder(child, $childrenContainer, libraryId, level + 1);
            });

            $container.append($childrenContainer);
        }
    }

    /**
     * Toggle library expansion (show/hide folders)
     */
    toggleLibraryExpansion($libraryItem) {
        const nodeId = $libraryItem.data('node-id');
        const hasChildren = $libraryItem.data('has-children');
        const isExpanded = $libraryItem.data('expanded') === 'true';

        if (!hasChildren) {
            return; // No children to expand
        }

        const $childrenContainer = $(`.children-container[data-parent-id="${nodeId}"]`);
        const $expandIcon = $libraryItem.find('.expand-icon');

        if (isExpanded) {
            // Collapse
            $childrenContainer.slideUp(200);
            $expandIcon.text('▶');
            $libraryItem.data('expanded', 'false');
        } else {
            // Expand
            $childrenContainer.slideDown(200);
            $expandIcon.text('▼');
            $libraryItem.data('expanded', 'true');
        }
    }

    /**
     * Toggle folder expansion (show/hide subfolders)
     */
    toggleFolderExpansion($folderItem) {
        const nodeId = $folderItem.data('node-id');
        const hasChildren = $folderItem.data('has-children');
        const isExpanded = $folderItem.data('expanded') === 'true';

        if (!hasChildren) {
            return; // No children to expand
        }

        const $childrenContainer = $(`.children-container[data-parent-id="${nodeId}"]`);
        const $expandIcon = $folderItem.find('.expand-icon');

        if (isExpanded) {
            // Collapse
            $childrenContainer.slideUp(200);
            $expandIcon.text('▶');
            $folderItem.data('expanded', 'false');
        } else {
            // Expand
            $childrenContainer.slideDown(200);
            $expandIcon.text('▼');
            $folderItem.data('expanded', 'true');
        }
    }

    /**
     * Select a folder and load its documents
     */
    async selectFolder($folderItem) {
        try {
            console.log('Folder selected:', $folderItem);

            // Update UI selection
            $('.folder-item').removeClass('selected');
            $folderItem.addClass('selected');

            const libraryId = $folderItem.data('library-id');
            const folderId = $folderItem.data('folder-id');
            const type = $folderItem.data('type');

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
        const $tbody = $('#documentTableBody');
        $tbody.empty();

        if (!documents || documents.length === 0) {
            $tbody.append('<tr><td colspan="4" class="text-center">No documents found</td></tr>');
            return;
        }

        documents.forEach(doc => {
            const $row = $(`
                <tr class="document-row" data-document-id="${doc.id}">
                    <td>
                        <span class="file-icon">${this.getFileTypeIcon(doc.fileName)}</span>
                    </td>
                    <td>${doc.fileName || 'Untitled'}</td>
                    <td>${this.formatDate(doc.modifiedOn || doc.dateModified)}</td>
                    <td>${this.formatFileSize(doc.size)}</td>
                </tr>
            `);

            $tbody.append($row);
        });

        // Clear selection and disable open button
        this.selectedDocument = null;
        $('#openBtn').prop('disabled', true);
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
                    const $libraryElement = $(`[data-folder-id="${libraryNode.id}"]`);
                    if ($libraryElement.length) {
                        $libraryElement.addClass('selected');
                        // Expand if it has children
                        if (libraryNode.children && libraryNode.children.length > 0) {
                            $libraryElement.find('.folder-toggle').first().click();
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
        $('#loadingMessage').text(message);
        $('#loadingSection').show();
        $('#errorSection').hide();
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
        $('#errorMessage').text(message);
        $('#errorSection').show();
        $('#loadingSection').hide();
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
        const query = $('#fileSearchInput').val().trim();
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
        const query = $('#fileSearchInput').val().trim();
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
        const query = $('#fullTextSearchInput').val().trim();
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
    selectDocument($row) {
        $('.document-row').removeClass('selected');
        $row.addClass('selected');
        this.selectedDocument = $row.data('document-id');

        // Enable open button when document is selected
        $('#openBtn').prop('disabled', false);
    }

    /**
     * Show the modal dialog
     */
    showDialog() {
        $('#dialogOverlay').show();
        $('#openDocumentDialog').show();
    }

    /**
     * Close the dialog/task pane
     */
    closeDialog() {
        $('#dialogOverlay').hide();
        $('#openDocumentDialog').hide();
        console.log('Dialog closed');
    }

    /**
     * Show context menu for document
     */
    showContextMenu(event, $row) {
        this.selectDocument($row);

        const documentId = $row.data('document-id');
        const document = this.documents.find(doc => doc.id === documentId);

        // Enable/disable menu items based on permissions
        $('#editDocument').toggle(document && document.canEdit);
        $('#deleteDocument').toggle(document && document.canDelete);

        $('#contextMenu').css({
            display: 'block',
            left: event.pageX + 'px',
            top: event.pageY + 'px'
        });
    }

    /**
     * Hide context menu
     */
    hideContextMenu() {
        $('#contextMenu').hide();
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
