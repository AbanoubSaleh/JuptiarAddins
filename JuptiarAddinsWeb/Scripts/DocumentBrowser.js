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
        this.checkAuthenticationStatus();
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

        // Search and refresh
        $('#searchBtn').on('click', () => this.performSearch());
        $('#searchInput').on('keypress', (e) => {
            if (e.which === 13) this.performSearch();
        });
        $('#refreshBtn').on('click', () => this.refreshCurrentView());

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

        // Folder selection
        $(document).on('click', '.folder-item', (e) => {
            this.selectFolder($(e.currentTarget));
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
            const authStatus = window.authManager.getAuthStatus();
            this.handleAuthStateChange(authStatus);
        } catch (error) {
            console.error('Error checking auth status:', error);
            this.showError('Failed to check authentication status');
        }
    }

    /**
     * Handle authentication state changes
     */
    handleAuthStateChange(authStatus) {
        if (authStatus.isAuthenticated) {
            this.showAuthenticatedState(authStatus.user);
            this.loadLibraryTree();
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
        $('#authStatusText').text('Not authenticated');
        $('#loginBtn').show();
        $('#logoutBtn').hide();
        $('#searchSection').hide();
        $('#mainContent').hide();
        $('#loadingSection').hide();
        
        $('.status-indicator').removeClass('online').addClass('online');
    }

    /**
     * Show login modal
     */
    showLoginModal() {
        const credentials = window.authManager.getStoredCredentials();
        const settings = window.authManager.getSettings();
        
        $('#username').val(credentials.username);
        $('#serverUrl').val(settings.serverUrl);
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
            const serverUrl = $('#serverUrl').val().trim();

            if (!username || !password || !serverUrl) {
                this.showError('Please fill in all required fields');
                return;
            }

            // Update settings with server URL
            await window.authManager.saveSettings({ serverUrl: serverUrl });

            // Attempt login
            $('#loginSubmitBtn').prop('disabled', true).text('Logging in...');
            
            await window.authManager.login(username, password, true);
            
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
            this.showLoading('Loading libraries...');
            
            const treeData = await window.jupiterService.getLibraryTree();
            this.folderTree = treeData;
            
            this.renderFolderTree(treeData);
            this.hideLoading();
            
        } catch (error) {
            console.error('Error loading library tree:', error);
            this.showError('Failed to load libraries: ' + error.message);
        }
    }

    /**
     * Render folder tree in the UI
     */
    renderFolderTree(treeData) {
        const $treeContainer = $('#folderTree');
        $treeContainer.empty();

        const renderNode = (node, level = 0) => {
            const $item = $(`
                <div class="folder-item" data-library-id="${node.id}" data-folder-id="${node.folderId || ''}" style="margin-left: ${level * 20}px">
                    <span class="folder-icon">📁</span>
                    <span class="folder-name">${node.name}</span>
                </div>
            `);

            $treeContainer.append($item);

            if (node.children && node.children.length > 0) {
                node.children.forEach(child => renderNode(child, level + 1));
            }
        };

        if (Array.isArray(treeData)) {
            treeData.forEach(library => renderNode(library));
        }
    }

    /**
     * Select a folder and load its documents
     */
    async selectFolder($folderItem) {
        try {
            // Update UI selection
            $('.folder-item').removeClass('selected');
            $folderItem.addClass('selected');

            const libraryId = $folderItem.data('library-id');
            const folderId = $folderItem.data('folder-id');

            this.currentLibrary = libraryId;
            this.currentFolder = folderId;

            await this.loadDocuments(libraryId, folderId);
            
        } catch (error) {
            console.error('Error selecting folder:', error);
            this.showError('Failed to load folder contents');
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
            $tbody.append('<tr><td colspan="5" class="text-center">No documents found</td></tr>');
            return;
        }

        documents.forEach(doc => {
            const $row = $(`
                <tr class="document-row" data-document-id="${doc.id}">
                    <td>
                        <span class="file-icon">📄</span>
                        ${this.getFileTypeIcon(doc.fileName)}
                    </td>
                    <td>${doc.fileName || 'Untitled'}</td>
                    <td>${this.formatDate(doc.dateModified)}</td>
                    <td>${this.formatFileSize(doc.size)}</td>
                    <td>
                        <div class="action-buttons">
                            <button class="action-btn view" onclick="documentBrowser.openDocument('${doc.id}')">View</button>
                            <button class="action-btn edit" onclick="documentBrowser.editDocument('${doc.id}')" ${!doc.canEdit ? 'disabled' : ''}>Edit</button>
                            <button class="action-btn delete" onclick="documentBrowser.deleteDocument('${doc.id}')" ${!doc.canDelete ? 'disabled' : ''}>Delete</button>
                        </div>
                    </td>
                </tr>
            `);
            
            $tbody.append($row);
        });
    }

    /**
     * Get file type icon based on extension
     */
    getFileTypeIcon(fileName) {
        if (!fileName) return '';
        
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
     * Show loading indicator
     */
    showLoading(message = 'Loading...') {
        $('#loadingSection p').text(message);
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
        // You can implement a success message display here
        console.log('Success:', message);
    }

    /**
     * Refresh current view
     */
    async refreshCurrentView() {
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
        const query = $('#searchInput').val().trim();
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
     * Select a document row
     */
    selectDocument($row) {
        $('.document-row').removeClass('selected');
        $row.addClass('selected');
        this.selectedDocument = $row.data('document-id');
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
    window.documentBrowser = new DocumentBrowser();
});
