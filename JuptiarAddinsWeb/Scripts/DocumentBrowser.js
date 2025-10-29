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

        // Initialize DocumentTracker for document detection
        this.initializeDocumentTracker();

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
        $.on('#refreshBtn', 'click', () => this.handleRefreshClick());

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

        // Instead of hiding main content, show a helpful login prompt
        this.showLoginPrompt();

        $.hide('#loadingSection');

        $.removeClass('.status-indicator', 'online');
        $.addClass('.status-indicator', 'offline');

        // Clear any existing tree data
        $.empty('#folderTree');
        $.empty('#documentTableBody');
    }

    /**
     * Show login prompt in main content area
     */
    showLoginPrompt() {
        const mainContent = $.select('#mainContent');
        if (mainContent) {
            mainContent.innerHTML = `
                <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; text-align: center; padding: 40px;">
                    <div style="font-size: 48px; margin-bottom: 20px; color: #0078d4;">🔐</div>
                    <h2 style="color: #323130; margin-bottom: 16px;">Authentication Required</h2>
                    <p style="color: #605e5c; margin-bottom: 24px; max-width: 400px; line-height: 1.5;">
                        Please log in to access your Jupiter Document Management System libraries and documents.
                    </p>
                    <button id="mainLoginBtn" class="ms-Button ms-Button--primary" style="padding: 12px 24px; font-size: 14px;">
                        <span class="ms-Button-label">🚀 Log In to Jupiter DMS</span>
                    </button>
                    <div style="margin-top: 20px; padding: 16px; background: #fff4ce; border: 1px solid #ffb900; border-radius: 4px; max-width: 400px;">
                        <p style="margin: 0; font-size: 13px; color: #8a6914;">
                            <strong>Server:</strong> ${window.JupiterConfig?.get('server.baseUrl') || 'Not configured'}<br>
                            <strong>Status:</strong> Ready to connect
                        </p>
                    </div>
                </div>
            `;

            // Add click handler for the main login button
            const mainLoginBtn = $.select('#mainLoginBtn');
            if (mainLoginBtn) {
                $.on(mainLoginBtn, 'click', () => this.showLoginModal());
            }

            $.show('#mainContent');
        }
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

            // Validate folder ID
            if (!folderId) {
                throw new Error('No folder ID provided');
            }

            // Check authentication
            if (!window.jupiterService || !window.jupiterService.authToken) {
                throw new Error('Authentication required. Please log in to access documents.');
            }

            this.showLoading('Loading documents...');

            const documents = await window.jupiterService.getDocumentsByFolder(folderId);
            console.log('Documents response:', documents);

            this.documents = documents || [];
            console.log('Processed documents:', this.documents);

            this.renderDocumentList(this.documents);
            this.hideLoading();

        } catch (error) {
            console.error('Error loading documents by folder:', error);
            this.hideLoading();

            let errorMessage = 'Failed to load documents';

            if (error.message.includes('401') || error.message.includes('Unauthorized')) {
                errorMessage = 'Authentication required. Please log in to access documents.';
            } else if (error.message.includes('403') || error.message.includes('Forbidden')) {
                errorMessage = 'Access denied. You do not have permission to view this folder.';
            } else if (error.message.includes('404') || error.message.includes('not found')) {
                errorMessage = 'Folder not found. It may have been deleted or moved.';
            } else if (error.message.includes('timeout')) {
                errorMessage = 'Request timeout. Please check your connection and try again.';
            } else if (error.message) {
                errorMessage += ': ' + error.message;
            }

            this.showError(errorMessage);
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
                    <span class="file-icon">${this.getFileTypeIcon(doc.name || doc.fileName)}</span>
                </td>
                <td>${doc.name || doc.fileName || 'Untitled'}</td>
                <td>${this.formatDate(doc.modifiedOn || doc.dateModified)}</td>
                <td>${this.formatFileSize(doc.fileSizeBytes || doc.size)}</td>
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
     * Show delete confirmation dialog (Office Add-ins don't support window.confirm)
     */
    showDeleteConfirmation(fileName, onConfirm) {
        const confirmHtml = `
            <div class="confirmation-dialog" style="background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 4px; position: relative;">
                <div class="confirmation-header" style="display: flex; justify-content: space-between; align-items: center; padding: 15px 15px 10px 15px; border-bottom: 1px solid #ffeaa7;">
                    <h4 style="margin: 0; color: #856404; display: flex; align-items: center; gap: 8px;">
                        <span style="font-size: 18px;">⚠️</span>
                        Delete Document
                    </h4>
                    <button id="confirmDeleteClose" class="close-dialog-btn" style="background: none; border: none; cursor: pointer; padding: 4px; border-radius: 3px; color: #856404; font-size: 16px; line-height: 1; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center;" title="Close">
                        ×
                    </button>
                </div>
                <div class="confirmation-body" style="padding: 15px;">
                    <p style="margin: 0 0 15px 0; color: #856404;">
                        Are you sure you want to delete "<strong>${fileName}</strong>"?<br>
                        This action cannot be undone.
                    </p>
                    <div style="text-align: right; display: flex; justify-content: flex-end; gap: 10px;">
                        <button id="confirmDeleteYes" class="ms-Button ms-Button--primary" style="background-color: #d13438; border-color: #d13438;">
                            <span class="ms-Button-label">Delete</span>
                        </button>
                        <button id="confirmDeleteNo" class="ms-Button">
                            <span class="ms-Button-label">Cancel</span>
                        </button>
                    </div>
                </div>
            </div>
        `;

        // Show the confirmation in the error section (reusing existing UI)
        $.html('#errorMessage', confirmHtml);
        $.show('#errorSection');
        $.hide('#loadingSection');

        // Store handler references for proper cleanup
        this.deleteConfirmationHandlers = {
            confirmYes: () => {
                this.hideDeleteConfirmation();
                onConfirm();
            },
            confirmNo: () => {
                this.hideDeleteConfirmation();
            },
            confirmClose: () => {
                this.hideDeleteConfirmation();
            },
            mouseEnter: (e) => {
                e.target.style.backgroundColor = '#f0e68c';
            },
            mouseLeave: (e) => {
                e.target.style.backgroundColor = 'transparent';
            }
        };

        // Handle confirmation buttons
        $.on('#confirmDeleteYes', 'click', this.deleteConfirmationHandlers.confirmYes);
        $.on('#confirmDeleteNo', 'click', this.deleteConfirmationHandlers.confirmNo);
        $.on('#confirmDeleteClose', 'click', this.deleteConfirmationHandlers.confirmClose);

        // Add hover effect for close button
        $.on('#confirmDeleteClose', 'mouseenter', this.deleteConfirmationHandlers.mouseEnter);
        $.on('#confirmDeleteClose', 'mouseleave', this.deleteConfirmationHandlers.mouseLeave);
    }

    /**
     * Hide delete confirmation dialog and cleanup event handlers
     */
    hideDeleteConfirmation() {
        $.hide('#errorSection');

        // Clean up event handlers if they exist
        if (this.deleteConfirmationHandlers) {
            $.off('#confirmDeleteYes', 'click', this.deleteConfirmationHandlers.confirmYes);
            $.off('#confirmDeleteNo', 'click', this.deleteConfirmationHandlers.confirmNo);
            $.off('#confirmDeleteClose', 'click', this.deleteConfirmationHandlers.confirmClose);
            $.off('#confirmDeleteClose', 'mouseenter', this.deleteConfirmationHandlers.mouseEnter);
            $.off('#confirmDeleteClose', 'mouseleave', this.deleteConfirmationHandlers.mouseLeave);

            this.deleteConfirmationHandlers = null;
        }
    }

    /**
     * Show success message
     */
    showSuccess(message) {
        // You can implement a success message display here
        console.log('Success:', message);
    }

    /**
     * Show temporary success message that auto-hides
     */
    showTemporarySuccess(message) {
        const successHtml = `
            <div class="success-message" style="background: #dff6dd; border: 1px solid #4caf50; padding: 15px; margin: 10px 0; border-radius: 4px; color: #2e7d32;">
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span style="font-size: 16px;">✓</span>
                    <span>${message}</span>
                </div>
            </div>
        `;

        // Show the success message in the error section
        $.html('#errorMessage', successHtml);
        $.show('#errorSection');
        $.hide('#loadingSection');

        // Auto-hide after 3 seconds
        setTimeout(() => {
            $.hide('#errorSection');
        }, 3000);
    }

    /**
     * Handle refresh button click with user feedback
     */
    async handleRefreshClick() {
        try {
            console.log('Refresh button clicked');

            // Show immediate feedback
            this.showLoading('Refreshing...');

            // Perform the refresh
            await this.refreshCurrentView();

            // Show success feedback briefly
            this.showTemporarySuccess('Refreshed successfully');

        } catch (error) {
            console.error('Error during refresh:', error);
            this.showError('Failed to refresh: ' + error.message);
        }
    }

    /**
     * Get current state for debugging
     */
    getCurrentState() {
        return {
            currentLibrary: this.currentLibrary,
            currentFolder: this.currentFolder,
            documentsCount: this.documents ? this.documents.length : 0,
            selectedDocument: this.selectedDocument,
            isAuthenticated: window.authManager ? window.authManager.getAuthStatus().isAuthenticated : false,
            hasAuthToken: window.jupiterService ? !!window.jupiterService.authToken : false
        };
    }

    /**
     * Refresh current view
     */
    async refreshCurrentView() {
        try {
            console.log('=== REFRESH CURRENT VIEW ===');
            console.log('Current state before refresh:', this.getCurrentState());

            // First re-check authentication status
            await this.checkAuthenticationStatus();

            // Check if user is authenticated before proceeding
            if (!window.authManager || !window.authManager.getAuthStatus().isAuthenticated) {
                console.log('User not authenticated, cannot refresh view');
                this.showError('Please log in to access documents');
                return;
            }

            // Then refresh the current view based on what's currently selected
            if (this.currentFolder) {
                // If we have a folder selected, reload documents for that folder
                console.log('Refreshing documents for folder:', this.currentFolder);
                await this.loadDocumentsByFolder(this.currentFolder);
            } else if (this.currentLibrary) {
                // If we have a library selected but no folder, just clear documents
                console.log('Refreshing library view:', this.currentLibrary);
                this.renderDocumentList([]);
            } else {
                // No specific selection, reload the library tree
                console.log('Refreshing library tree');
                await this.loadLibraryTree();

                // If library tree loads successfully but no folder is selected,
                // try to find and select the currently selected folder in the UI
                setTimeout(() => {
                    const selectedFolder = $.select('.folder-item.selected');
                    if (selectedFolder) {
                        const folderId = selectedFolder.getAttribute('data-folder-id');
                        const type = selectedFolder.getAttribute('data-type');

                        if (type === 'folder' && folderId) {
                            console.log('Re-selecting folder from UI:', folderId);
                            this.currentFolder = folderId;
                            this.loadDocumentsByFolder(folderId);
                        }
                    }
                }, 500);
            }

            console.log('Current state after refresh:', this.getCurrentState());
        } catch (error) {
            console.error('Error refreshing current view:', error);
            this.showError('Failed to refresh: ' + error.message);
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

                    // Mark document as Jupiter-managed after successful opening
                    await this.markDocumentAsJupiterManaged(documentId);

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

            // Show confirmation dialog instead of using window.confirm
            this.showDeleteConfirmation(document.fileName, async () => {
                try {
                    this.showLoading('Deleting document...');

                    const result = await window.jupiterService.deleteDocument(documentId);

                    if (result) {
                        // Refresh the document list
                        await this.refreshCurrentView();

                        // Show success message temporarily
                        this.showTemporarySuccess(`Document "${document.fileName}" deleted successfully`);
                    } else {
                        this.showError('Failed to delete document: Unknown error occurred');
                    }
                } catch (error) {
                    console.error('Error deleting document:', error);
                    let errorMessage = 'Failed to delete document';

                    if (error.message) {
                        errorMessage += ': ' + error.message;
                    } else if (error.status) {
                        switch (error.status) {
                            case 403:
                                errorMessage += ': You do not have permission to delete this document';
                                break;
                            case 404:
                                errorMessage += ': Document not found';
                                break;
                            case 409:
                                errorMessage += ': Document is currently checked out and cannot be deleted';
                                break;
                            default:
                                errorMessage += ': Server error occurred';
                        }
                    }

                    this.showError(errorMessage);
                }
            });

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

    /**
     * Mark document as Jupiter-managed by setting appropriate metadata
     * @param {string} documentId - Jupiter document ID
     */
    async markDocumentAsJupiterManaged(documentId) {
        try {
            console.log('🏷️ Marking document as Jupiter-managed:', documentId);

            // Get document information from the API
            const documentInfo = await window.jupiterService.getDocumentById(documentId);
            if (!documentInfo) {
                console.warn('Could not fetch document info for ID:', documentId);
                return;
            }

            // PRIORITY 1: Set the required Jupiter custom properties for DocumentTracker
            console.log('📝 Setting Jupiter custom properties...');
            await this.setJupiterCustomProperties(documentId, documentInfo.libraryId || this.currentLibrary);

            // PRIORITY 2: Try to initialize DocumentStateManager for additional metadata
            try {
                if (!window.documentStateManager && typeof DocumentStateManager !== 'undefined') {
                    window.documentStateManager = new DocumentStateManager();
                    await window.documentStateManager.initialize();
                }

                if (window.documentStateManager) {
                    // Create Jupiter document state
                    const jupiterState = {
                        documentId: documentId,
                        documentName: documentInfo.name,
                        folderId: documentInfo.folderId,
                        folderPath: documentInfo.folderPath || 'Unknown',
                        version: documentInfo.currentVersion || 1,
                        checkoutStatus: documentInfo.checkoutStatus || 'Available',
                        lastSaved: new Date().toISOString(),
                        openedVia: 'DocumentBrowser'
                    };

                    // Set document state in Office settings
                    await window.documentStateManager.setDocumentState(jupiterState);

                    // Also set custom properties for persistence across sessions
                    await window.documentStateManager.setDocumentCustomProperties(jupiterState);

                    console.log('✅ DocumentStateManager metadata set successfully');
                }
            } catch (stateManagerError) {
                console.warn('⚠️ DocumentStateManager not available, but custom properties were set:', stateManagerError.message);
            }

            // PRIORITY 3: Update ribbon if available
            try {
                if (window.ribbonManager) {
                    await window.ribbonManager.showExistingDocumentRibbon();
                }
            } catch (ribbonError) {
                console.warn('⚠️ Could not update ribbon:', ribbonError.message);
            }

            // PRIORITY 4: Initialize DocumentTracker for comprehensive edit detection
            try {
                console.log('🚀 Initializing DocumentTracker after document opened...');
                if (typeof DocumentTracker !== 'undefined') {
                    if (!window.documentTracker) {
                        window.documentTracker = new DocumentTracker();
                    }
                    // Force re-initialization to detect the newly set properties
                    await window.documentTracker.initializeDocumentTracking();
                    console.log('✅ DocumentTracker initialized successfully');
                } else {
                    console.warn('⚠️ DocumentTracker class not available');
                }
            } catch (trackerError) {
                console.error('❌ Could not initialize DocumentTracker:', trackerError);
            }

            // PRIORITY 5: Initialize DocumentEditMonitor for checkout workflow (legacy support)
            try {
                await this.initializeDocumentEditMonitor();
            } catch (monitorError) {
                console.warn('⚠️ Could not initialize DocumentEditMonitor:', monitorError.message);
            }

            console.log('✅ Document successfully marked as Jupiter-managed');

        } catch (error) {
            console.error('❌ Error marking document as Jupiter-managed:', error);
            // Even if there's an error, try to set the basic custom properties
            try {
                await this.setJupiterCustomProperties(documentId, this.currentLibrary);
                console.log('✅ Fallback: Basic custom properties set');
            } catch (fallbackError) {
                console.error('❌ Even fallback failed:', fallbackError);
            }
        }
    }

    /**
     * Initialize DocumentEditMonitor for checkout workflow
     */
    async initializeDocumentEditMonitor() {
        try {
            console.log('🔧 Initializing DocumentEditMonitor...');

            // Ensure all dependencies are available
            if (!window.documentStateManager) {
                console.warn('DocumentStateManager not available - cannot initialize edit monitor');
                return;
            }

            if (!window.jupiterService) {
                console.warn('JupiterService not available - cannot initialize edit monitor');
                return;
            }

            // Create a simple ribbon manager if not available
            if (!window.ribbonManager) {
                window.ribbonManager = {
                    showCheckoutPrompt: async () => {
                        console.log('Ribbon manager not available - using fallback');
                        return false;
                    }
                };
            }

            // Initialize DocumentEditMonitor
            if (!window.documentEditMonitor) {
                window.documentEditMonitor = new DocumentEditMonitor(
                    window.documentStateManager,
                    window.jupiterService,
                    window.ribbonManager
                );
            }

            // Start monitoring for edit attempts
            await window.documentEditMonitor.startMonitoring();

            console.log('✅ DocumentEditMonitor initialized and started');

        } catch (error) {
            console.error('❌ Error initializing DocumentEditMonitor:', error);
        }
    }

    /**
     * Set Jupiter custom properties required by DocumentTracker
     * @param {string} documentId - Jupiter document ID
     * @param {string} libraryId - Jupiter library ID
     */
    async setJupiterCustomProperties(documentId, libraryId) {
        try {
            console.log('🏷️ Setting Jupiter custom properties:', { documentId, libraryId });

            await Word.run(async (context) => {
                const properties = context.document.properties.customProperties;

                // Load existing properties to check for duplicates
                properties.load('items');
                await context.sync();

                // Remove existing Jupiter properties if they exist
                const existingProps = properties.items.filter(p =>
                    p.key === 'JupiterDocumentId' || p.key === 'LibraryId'
                );

                existingProps.forEach(prop => prop.delete());

                // Add the required custom properties
                properties.add('JupiterDocumentId', documentId);
                if (libraryId) {
                    properties.add('LibraryId', libraryId);
                }

                await context.sync();
                console.log('✅ Jupiter custom properties set successfully');
            });

        } catch (error) {
            console.error('❌ Error setting Jupiter custom properties:', error);
            throw error;
        }
    }

    /**
     * Detect if current document is managed by Jupiter DMS
     * @returns {Promise<string|null>} Document ID if Jupiter document, null otherwise
     */
    async detectJupiterDocument() {
        try {
            console.log('🔍 Detecting Jupiter document...');

            return await Word.run(async (context) => {
                const properties = context.document.properties.customProperties;
                properties.load('items');
                await context.sync();

                // Look for JupiterDocumentId property
                const jupiterDocProp = properties.items.find(p => p.key === 'JupiterDocumentId');

                if (jupiterDocProp) {
                    console.log('✅ Jupiter document detected:', jupiterDocProp.value);
                    return jupiterDocProp.value;
                } else {
                    console.log('ℹ️ No Jupiter document ID found - this is not a Jupiter-managed document');
                    return null;
                }
            });

        } catch (error) {
            console.warn('⚠️ Error detecting Jupiter document:', error);
            return null;
        }
    }

    /**
     * Show info message for non-Jupiter documents
     */
    async showNonJupiterDocumentInfo() {
        try {
            const message = "This document is not managed by Jupiter DMS. To manage it, please open a file from the Jupiter library.";
            const options = [{ text: 'OK', action: 'ok' }];

            await this.showJupiterPopup(message, options);

        } catch (error) {
            console.error('Error showing non-Jupiter document info:', error);
        }
    }

    /**
     * Show Jupiter popup dialog
     * @param {string} message - Message to display
     * @param {Array} options - Button options
     * @returns {Promise<string>} User's choice
     */
    async showJupiterPopup(message, options) {
        return new Promise((resolve) => {
            try {
                const encodedMessage = encodeURIComponent(message);
                const encodedOptions = encodeURIComponent(JSON.stringify(options));
                const popupUrl = `${window.location.origin}/jupiter-popup.html?message=${encodedMessage}&options=${encodedOptions}`;

                console.log('🔔 Showing Jupiter popup:', message);

                Office.context.ui.displayDialogAsync(popupUrl, { height: 30, width: 20 }, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const dialog = result.value;

                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                            console.log('📨 Popup response:', arg.message);
                            dialog.close();
                            resolve(arg.message);
                        });

                        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                            console.log('🔔 Dialog event:', arg.error);
                            dialog.close();
                            resolve('cancel');
                        });
                    } else {
                        console.error('Failed to show popup:', result.error);
                        resolve('cancel');
                    }
                });

            } catch (error) {
                console.error('Error showing Jupiter popup:', error);
                resolve('cancel');
            }
        });
    }

    /**
     * Initialize DocumentTracker for document detection and validation
     */
    async initializeDocumentTracker() {
        try {
            console.log('🔍 Initializing DocumentTracker...');

            // Check if DocumentTracker is available
            if (typeof DocumentTracker === 'undefined') {
                console.warn('DocumentTracker class not available');
                return;
            }

            // Initialize DocumentTracker if not already done
            if (!window.documentTracker) {
                window.documentTracker = new DocumentTracker();
                await window.documentTracker.initializeDocumentTracking();
                console.log('✅ DocumentTracker initialized successfully');
            }

        } catch (error) {
            console.error('❌ Error initializing DocumentTracker:', error);
        }
    }

    /**
     * Manual test function - call from console to test DocumentTracker
     */
    async testDocumentTracker() {
        try {
            console.log('🧪 Testing DocumentTracker manually...');

            if (typeof DocumentTracker === 'undefined') {
                console.error('❌ DocumentTracker class not available');
                return;
            }

            if (!window.documentTracker) {
                console.log('🔧 Creating new DocumentTracker instance...');
                window.documentTracker = new DocumentTracker();
            }

            console.log('🚀 Initializing DocumentTracker...');
            await window.documentTracker.initializeDocumentTracking();

            console.log('✅ DocumentTracker test completed - try typing in the document now!');

        } catch (error) {
            console.error('❌ DocumentTracker test failed:', error);
        }
    }
}

// Initialize when Office is ready (only if Office is available)
if (typeof Office !== 'undefined' && Office.onReady) {
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
} else {
    console.warn('DocumentBrowser.js: Office.js not available, skipping Office.onReady initialization');
}
