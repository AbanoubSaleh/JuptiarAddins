/**
 * Jupiter Service - Core API communication layer
 * Handles all HTTP requests to the Jupiter document management system
 */
class JupiterService {
    constructor() {
        this.baseUrl = '';
        this.apiEndpoint = '/api';
        this.authToken = null;
        this.timeout = 30000; // 30 seconds
    }
    /**
     * Initialize the service with configuration
     */
    initialize(config) {
        this.baseUrl = config.serverUrl || '';
        this.apiEndpoint = config.apiEndpoint || '/api';
        this.timeout = config.timeout || 30000;
        this.authToken = config.authToken || null;
    }
    /**
     * Get the full API URL
     */
    getApiUrl(endpoint) {
        const fullUrl = `${this.baseUrl}${this.apiEndpoint}${endpoint}`;
        return fullUrl;
    }
    /**
     * Make HTTP request with authentication
     */
    async makeRequest(method, endpoint, data = null, options = {}) {
        const url = this.getApiUrl(endpoint);
        const headers = {
            'Content-Type': 'application/json',
            ...options.headers
        };
        if (this.authToken) {
            headers['Authorization'] = `Bearer ${this.authToken}`;
        }
        const requestOptions = {
            method: method,
            headers: headers,
            ...options
        };
        if (data && (method === 'POST' || method === 'PUT' || method === 'PATCH')) {
            requestOptions.body = JSON.stringify(data);
        }
        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), this.timeout);
            requestOptions.signal = controller.signal;
            const response = await fetch(url, requestOptions);
            clearTimeout(timeoutId);
            if (!response.ok) {
                // Try to get detailed error message from response body
                let errorMessage = `HTTP ${response.status}: ${response.statusText}`;
                try {
                    const contentType = response.headers.get('content-type');
                    if (contentType && contentType.includes('application/json')) {
                        const errorData = await response.json();
                        // Handle different error response formats
                        if (typeof errorData === 'string') {
                            errorMessage = errorData;
                        } else if (errorData.message) {
                            errorMessage = errorData.message;
                        } else if (errorData.error) {
                            errorMessage = errorData.error;
                        } else if (errorData.title) {
                            errorMessage = errorData.title;
                        }
                    } else {
                        const errorText = await response.text();
                        if (errorText) {
                            errorMessage = errorText;
                        }
                    }
                } catch (parseError) {
                    // If we can't parse the error response, use the default message
                    console.warn('Could not parse error response:', parseError);
                }
                throw new Error(errorMessage);
            }
            const contentType = response.headers.get('content-type');
            if (contentType && contentType.includes('application/json')) {
                return await response.json();
            } else {
                return await response.text();
            }
        } catch (error) {
            if (error.name === 'AbortError') {
                error = new Error('Request timeout');
            }
            // Use centralized error handling if available
            if (window.ErrorHandler) {
                window.ErrorHandler.handle(error, {
                    context: `JupiterService.${method}`,
                    showToUser: false, // Let calling code decide whether to show to user
                    logError: true
                });
            }
            throw error;
        }
    }
    // Authentication Methods
    async login(username, password) {
        try {
            const response = await this.makeRequest('POST', '/Auth/login', {
                username: username,
                password: password
            });
            if (response.token) {
                this.authToken = response.token;
                // Transform backend response to match addin expectations
                return {
                    success: true,
                    token: response.token,
                    user: response.user,
                    expiresAt: response.expiresAt
                };
            } else {
                return {
                    success: false,
                    message: 'Login failed - no token received'
                };
            }
        } catch (error) {
            return {
                success: false,
                message: error.message || 'Login failed'
            };
        }
    }
    async logout() {
        try {
            await this.makeRequest('POST', '/Auth/logout');
        } finally {
            this.authToken = null;
        }
    }
    async validateSession() {
        try {
            const response = await this.makeRequest('GET', '/Auth/validate');
            return response; // Backend returns { valid: true/false, ... }
        } catch (error) {
            return { valid: false, message: error.message };
        }
    }
    // Library and Folder Methods
    async getLibraryTree() {
        return await this.makeRequest('GET', '/libraries/tree');
    }
    async getDocuments(libraryId, folderId = null, page = 1, limit = 50) {
        let endpoint = `/documents/list/${libraryId}`;
        const params = new URLSearchParams({
            page: page.toString(),
            limit: limit.toString()
        });
        if (folderId) {
            params.append('folderId', folderId);
        }
        endpoint += `?${params.toString()}`;
        return await this.makeRequest('GET', endpoint);
    }
    /**
     * Get documents by folder ID using the folder endpoint
     */
    async getDocumentsByFolder(folderId) {
        return await this.makeRequest('GET', `/documents/folder/${folderId}`);
    }
    // Document Methods
    async getDocument(documentId) {
        return await this.makeRequest('GET', `/documents/${documentId}`);
    }

    /**
     * Get document by ID (alias for getDocument for clarity)
     * @param {string} documentId - Document ID
     * @returns {Promise<Object>} Document information
     */
    async getDocumentById(documentId) {
        return await this.getDocument(documentId);
    }

    /**
     * Get document status for checkout validation
     * @param {string} documentId - Document ID
     * @returns {Promise<Object>} Document status with checkout information
     */
    async getDocumentStatus(documentId) {
        try {
            console.log(`🔍 Getting document status for: ${documentId}`);

            const response = await this.makeRequest('GET', `/documents/${documentId}/status`);

            console.log('📄 Document status response:', response);

            // Normalize various backend response shapes to a consistent status object
            const currentUserEmail = (await this.getCurrentUserEmail()) || '';
            const rawStatus = response?.checkoutStatus ?? response?.status ?? response?.state;
            const rawIsCheckedOut = typeof response?.isCheckedOut === 'boolean' ? response.isCheckedOut : undefined;

            // Extract checkedOutBy from common fields
            const checkedOutBy = response?.checkedOutBy || response?.lockedBy || response?.checkedOutByEmail || null;

            // Determine isCheckedOut
            let isCheckedOut = false;
            if (typeof rawIsCheckedOut !== 'undefined') {
                isCheckedOut = rawIsCheckedOut;
            } else if (typeof rawStatus === 'string') {
                const s = rawStatus.toLowerCase();
                isCheckedOut = s === 'checkedout' || s === 'checked_out' || s === 'locked' || s === '1';
            } else if (typeof rawStatus === 'number') {
                // Treat 1 as checked out (common enum), 0 as available
                isCheckedOut = rawStatus === 1;
            } else if (checkedOutBy) {
                // If we know who checked it out, treat as checked out
                isCheckedOut = true;
            }

            // Determine lockedByYou
            let lockedByYou = false;
            if (typeof response?.lockedByYou === 'boolean') {
                lockedByYou = response.lockedByYou;
            } else if (checkedOutBy && currentUserEmail) {
                lockedByYou = String(checkedOutBy).toLowerCase() === String(currentUserEmail).toLowerCase();
            }

            const normalized = {
                isCheckedOut,
                checkedOutBy: checkedOutBy || null,
                lockedByYou,
                checkoutStatus: rawStatus ?? (isCheckedOut ? 'CheckedOut' : 'Available'),
                documentInfo: response?.documentInfo || null,
                _raw: response
            };

            return normalized;

        } catch (error) {
            // If the specific status endpoint doesn't exist, fall back to getDocumentById
            console.warn('Status endpoint not available, falling back to document info');

            try {
                const documentInfo = await this.getDocumentById(documentId);
                if (!documentInfo) {
                    throw new Error('Document not found');
                }

                // Get current user email for comparison
                const currentUserEmail = await this.getCurrentUserEmail();

                // Map to expected status format
                return {
                    isCheckedOut: documentInfo.checkoutStatus === 'CheckedOut',
                    checkedOutBy: documentInfo.checkedOutBy || null,
                    lockedByYou: documentInfo.checkoutStatus === 'CheckedOut' &&
                               documentInfo.checkedOutBy === currentUserEmail,
                    checkoutStatus: documentInfo.checkoutStatus,
                    documentInfo: documentInfo
                };

            } catch (fallbackError) {
                console.error('❌ Error getting document status (fallback):', fallbackError);
                throw fallbackError;
            }
        }
    }

    /**
     * Get current user email (or username if email not available)
     * @returns {Promise<string|null>} Current user email/username
     */
    async getCurrentUserEmail() {
        try {
            if (window.authManager) {
                if (window.authManager.currentUser) {
                    const u = window.authManager.currentUser;
                    return u.email || u.username || u.userName || null;
                }
                if (typeof window.authManager.getAuthStatus === 'function') {
                    const status = window.authManager.getAuthStatus();
                    const u = status?.user;
                    return (u && (u.email || u.username || u.userName)) || null;
                }
            }
            return null;
        } catch (error) {
            console.warn('Could not get current user identity:', error);
            return null;
        }
    }
    async downloadDocument(documentId) {
        const url = this.getApiUrl(`/documents/${documentId}/download`);
        const headers = {};
        if (this.authToken) {
            headers['Authorization'] = `Bearer ${this.authToken}`;
        }
        const response = await fetch(url, { headers });
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        return response.blob();
    }
    /**
     * Get document versions
     * @param {string} documentId
     * @returns {Promise<Array>} Array of version DTOs
     */
    async getDocumentVersions(documentId) {
        return await this.makeRequest('GET', `/documents/${documentId}/versions`);
    }

    /**
     * Download a specific document version
     * @param {string} documentId
     * @param {string|number} version
     * @returns {Promise<Blob>} File content
     */
    async downloadDocumentVersion(documentId, version) {
        const url = this.getApiUrl(`/documents/${documentId}/versions/${encodeURIComponent(version)}/download`);
        const headers = {};
        if (this.authToken) {
            headers['Authorization'] = `Bearer ${this.authToken}`;
        }
        const response = await fetch(url, { headers });
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        return response.blob();
    }

    async uploadDocument(libraryId, folderId, file, metadata = {}) {
        const formData = new FormData();
        formData.append('file', file);
        formData.append('libraryId', libraryId);
        if (folderId) {
            formData.append('folderId', folderId);
        }
        // Add metadata
        Object.keys(metadata).forEach(key => {
            if (metadata[key] !== null && metadata[key] !== undefined) {
                formData.append(key, metadata[key]);
            }
        });
        const headers = {};
        if (this.authToken) {
            headers['Authorization'] = `Bearer ${this.authToken}`;
        }
        const url = this.getApiUrl('/documents/upload');
        const response = await fetch(url, {
            method: 'POST',
            headers: headers,
            body: formData
        });
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        return await response.json();
    }
    async updateDocument(documentId, data) {
        return await this.makeRequest('PUT', `/documents/${documentId}`, data);
    }
    async deleteDocument(documentId) {
        try {
            const response = await this.makeRequest('DELETE', `/documents/${documentId}`);
            return response;
        } catch (error) {
            // Extract status code from error message if available
            const statusMatch = error.message.match(/HTTP (\d+):/);
            if (statusMatch) {
                error.status = parseInt(statusMatch[1]);
            }

            // Re-throw with enhanced error information
            throw error;
        }
    }
    // Metadata Methods
    async getDocumentMetadata(documentId) {
        return await this.makeRequest('GET', `/documents/${documentId}/metadata`);
    }
    async updateDocumentMetadata(documentId, metadata) {
        return await this.makeRequest('PUT', `/documents/${documentId}/metadata`, metadata);
    }
    // Search Methods
    // Note: Backend endpoint is GET /api/documents/search with searchTerm, libraryId, folderId.
    // Pagination (page, limit) may or may not be supported by backend; safe to include and ignore if unsupported.
    async searchDocuments(searchTerm, _type = 'filename', libraryId = null, folderId = null, page = null, limit = null) {
        const params = new URLSearchParams();
        params.append('searchTerm', searchTerm || '');
        if (libraryId) params.append('libraryId', libraryId);
        if (folderId) params.append('folderId', folderId);
        if (page !== null && page !== undefined) params.append('page', page.toString());
        if (limit !== null && limit !== undefined) params.append('limit', limit.toString());
        return await this.makeRequest('GET', `/documents/search?${params.toString()}`);
    }
    async fullTextSearch(query) {
        // Kept for backwards-compatibility; maps to same endpoint using searchTerm only
        return await this.searchDocuments(query, 'fulltext');
    }
    // Permission Methods
    async getUserPermissions(documentId) {
        return await this.makeRequest('GET', `/users/permissions/${documentId}`);
    }
    async getUserRoles() {
        return await this.makeRequest('GET', '/users/roles');
    }
    // Utility Methods
    isAuthenticated() {
        return this.authToken !== null;
    }
    setAuthToken(token) {
        this.authToken = token;
    }
    clearAuthToken() {
        this.authToken = null;
    }
    /**
     * Check if document name exists in folder
     * @param {string} name - Document name
     * @param {string} folderId - Folder ID
     * @returns {Promise<Object>} Duplicate check result
     */
    async checkDuplicateName(name, folderId) {
        try {
            const params = new URLSearchParams({
                name: name,
                folderId: folderId
            });
            const response = await this.makeRequest('GET', `/documents/check-duplicate?${params.toString()}`);
            return response;
        } catch (error) {
            console.error('Error checking duplicate name:', error);
            this.handleError(error);
        }
    }
    /**
     * Get all libraries
     * @returns {Promise<Array>} List of libraries
     */
    async getLibraries() {
        try {
            const response = await this.makeRequest('GET', '/libraries');
            return response;
        } catch (error) {
            console.error('Error getting libraries:', error);
            this.handleError(error);
        }
    }
    /**
     * Get library tree with folders (hierarchical structure)
     * @returns {Promise<Array>} Library tree structure
     */
    async getLibraryTree() {
        try {
            const response = await this.makeRequest('GET', '/libraries/tree');
            return response;
        } catch (error) {
            console.error('Error getting library tree:', error);
            this.handleError(error);
        }
    }
    /**
     * Get folders in a library
     * @param {string} libraryId - Library ID (required)
     * @returns {Promise<Array>} List of folders
     */
    async getFolders(libraryId) {
        try {
            if (!libraryId) {
                // If no library ID provided, get the library tree instead
                return await this.getLibraryTree();
            }
            const response = await this.makeRequest('GET', `/folders/library/${libraryId}`);
            return response;
        } catch (error) {
            console.error('Error getting folders:', error);
            this.handleError(error);
        }
    }
    /**
     * Get folder by ID
     * @param {string} folderId - Folder ID
     * @returns {Promise<Object>} Folder details
     */
    async getFolder(folderId) {
        try {
            const response = await this.makeRequest('GET', `/folders/${folderId}`);
            return response;
        } catch (error) {
            console.error('Error getting folder:', error);
            this.handleError(error);
        }
    }
    /**
     * Create a new library (Admin only)
     * @param {Object} libraryData - Library data {name, description, isActive}
     * @returns {Promise<Object>} Created library
     */
    async createLibrary(libraryData) {
        try {
            const response = await this.makeRequest('POST', '/libraries', libraryData);
            return response;
        } catch (error) {
            console.error('Error creating library:', error);
            this.handleError(error);
        }
    }
    /**
     * Create a new folder (Admin only)
     * @param {Object} folderData - Folder data {name, libraryId, parentFolderId, description}
     * @returns {Promise<Object>} Created folder
     */
    async createFolder(folderData) {
        try {
            const response = await this.makeRequest('POST', '/folders', folderData);
            return response;
        } catch (error) {
            console.error('Error creating folder:', error);
            this.handleError(error);
        }
    }

    /**
     * Update a library (Admin only)
     * @param {Object} libraryData - Library data {id, name, description, isActive}
     * @returns {Promise<Object>} Updated library
     */
    async updateLibrary(libraryData) {
        try {
            const response = await this.makeRequest('PUT', '/libraries', libraryData);
            return response;
        } catch (error) {
            console.error('Error updating library:', error);
            this.handleError(error);
        }
    }

    /**
     * Delete a library (Admin only)
     * @param {string} libraryId - Library ID
     * @returns {Promise<Object>} Delete result
     */
    async deleteLibrary(libraryId) {
        try {
            const response = await this.makeRequest('DELETE', `/libraries/${libraryId}`);
            return response;
        } catch (error) {
            console.error('Error deleting library:', error);
            this.handleError(error);
        }
    }

    /**
     * Update a folder (Admin only)
     * @param {Object} folderData - Folder data {id, name, description}
     * @returns {Promise<Object>} Updated folder
     */
    async updateFolder(folderData) {
        try {
            const response = await this.makeRequest('PUT', '/folders', folderData);
            return response;
        } catch (error) {
            console.error('Error updating folder:', error);
            this.handleError(error);
        }
    }

    /**
     * Delete a folder (Admin only)
     * @param {string} folderId - Folder ID
     * @returns {Promise<Object>} Delete result
     */
    async deleteFolder(folderId) {
        try {
            const response = await this.makeRequest('DELETE', `/folders/${folderId}`);
            return response;
        } catch (error) {
            console.error('Error deleting folder:', error);
            this.handleError(error);
        }
    }

    /**
     * Get folder by ID
     * @param {string} folderId - Folder ID
     * @returns {Promise<Object>} Folder details
     */
    async getFolderById(folderId) {
        try {
            const response = await this.makeRequest('GET', `/folders/${folderId}`);
            return response;
        } catch (error) {
            console.error('Error getting folder:', error);
            this.handleError(error);
        }
    }

    /**
     * Get documents in a folder
     * @param {string} folderId - Folder ID
     * @returns {Promise<Array>} List of documents
     */
    async getDocuments(folderId) {
        try {
            const response = await this.makeRequest('GET', `/documents?folderId=${folderId}`);
            return response;
        } catch (error) {
            console.error('Error getting documents:', error);
            this.handleError(error);
        }
    }
    /**
     * Get authentication token
     * @returns {string|null} Current auth token
     */
    getToken() {
        return this.authToken;
    }
    /**
     * Set authentication token
     * @param {string} token - Auth token
     */
    setToken(token) {
        this.authToken = token;
    }
    /**
     * Check out a document for editing
     * @param {string} documentId - Document ID
     * @returns {Promise<Object>} Check-out result
     */
    async checkOutDocument(documentId) {
        try {
            console.log(`🔒 Attempting to check out document: ${documentId}`);

            const response = await this.makeRequest('POST', `/documents/${documentId}/checkout`);

            console.log('✅ Document checked out successfully:', response);
            return {
                success: true,
                message: 'Document checked out successfully',
                ...response
            };
        } catch (error) {
            console.error('❌ Error checking out document:', error);

            // If the backend says it's already checked out, verify who has it.
            try {
                const statusMatch = /HTTP\s+(\d+)/i.exec(error.message || '');
                const statusCode = statusMatch ? parseInt(statusMatch[1], 10) : null;
                if ((statusCode === 400 || statusCode === 409) && /already\s+checked\s+out/i.test(error.message)) {
                    const status = await this.getDocumentStatus(documentId);
                    if (status?.isCheckedOut && status?.lockedByYou) {
                        // Treat as success: the current user already has it checked out
                        console.warn('Checkout returned already-checked-out, but it is checked out by current user. Treating as success.');
                        return { success: true, message: 'Document already checked out by you', ...status };
                    }
                }
            } catch (verifyErr) {
                console.warn('Could not verify checkout ownership after error:', verifyErr);
            }

            // Parse specific error messages from backend
            let errorMessage = error.message;
            if (/already\s+checked\s+out/i.test(error.message)) {
                errorMessage = 'This document is already checked out by another user.';
            } else if (/not\s+found/i.test(error.message)) {
                errorMessage = 'Document not found. It may have been deleted or moved.';
            } else if (/unauthorized|403/i.test(error.message)) {
                errorMessage = 'You do not have permission to check out this document.';
            }

            return {
                success: false,
                error: errorMessage,
                originalError: error.message
            };
        }
    }
    /**
     * Check in a document with new version
     * @param {string} documentId - Document ID
     * @param {FormData} formData - Form data with file and version comment
     * @returns {Promise<Object>} Check-in result
     */
    async checkInDocument(documentId, fileBlob, versionComment = '', _keepCheckedOut = false, fileName = 'document.docx') {
        try {
            console.log(`📥 Attempting to check in document: ${documentId}`);

            // Create FormData for file upload
            const formData = new FormData();
            if (fileBlob) {
                // Include filename so server can persist a meaningful name
                formData.append('file', fileBlob, fileName || 'document.docx');
            }
            if (versionComment) {
                formData.append('versionComment', versionComment);
            }

            const response = await fetch(`${this.baseUrl}${this.apiEndpoint}/documents/${documentId}/checkin`, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${this.authToken}`,
                    'Accept': 'application/json'
                },
                body: formData
            });

            if (!response.ok) {
                let errorText = '';
                try {
                    const errorData = await response.json();
                    errorText = errorData?.message || '';
                } catch (_) {
                    // ignore JSON parse error
                }
                throw new Error(errorText || `HTTP error! status: ${response.status}`);
            }

            const result = await response.json();
            console.log('✅ Document checked in successfully:', result);

            return {
                success: true,
                document: result,
                message: 'Document checked in successfully'
            };
        } catch (error) {
            console.error('❌ Error checking in document:', error);

            // Parse specific error messages from backend
            let errorMessage = error.message;
            if (error.message?.toLowerCase().includes('not checked out')) {
                errorMessage = 'This document is not currently checked out. Please check out the document first.';
            } else if (error.message?.toLowerCase().includes('not found') || error.message?.includes('404')) {
                errorMessage = 'Document not found. It may have been deleted or moved.';
            } else if (error.message?.toLowerCase().includes('unauthorized') || error.message?.includes('403')) {
                errorMessage = 'You do not have permission to check in this document.';
            }

            return {
                success: false,
                error: errorMessage,
                originalError: error.message
            };
        }
    }

    /**
     * Cancel document check-out
     * @param {string} documentId - Document ID
     * @returns {Promise<Object>} Cancel check-out result
     */
    async cancelCheckOut(documentId) {
        try {
            console.log(`🔓 Attempting to cancel check-out for document: ${documentId}`);

            const response = await this.makeRequest('POST', `/documents/${documentId}/cancel-checkout`);

            console.log('✅ Document check-out cancelled successfully:', response);
            return {
                success: true,
                message: 'Document check-out cancelled successfully',
                ...response
            };
        } catch (error) {
            console.error('❌ Error cancelling document check-out:', error);

            let errorMessage = error.message;
            if (error.message.includes('not checked out')) {
                errorMessage = 'This document is not currently checked out.';
            } else if (error.message.includes('not found')) {
                errorMessage = 'Document not found. It may have been deleted or moved.';
            }

            return {
                success: false,
                error: errorMessage,
                originalError: error.message
            };
        }
    }

    // Error handling helper
    handleError(error) {
        console.error('Jupiter Service Error:', error);
        if (error.message.includes('401')) {
            this.clearAuthToken();
            throw new Error('Authentication required. Please login again.');
        } else if (error.message.includes('403')) {
            throw new Error('Access denied. You do not have permission to perform this action.');
        } else if (error.message.includes('404')) {
            throw new Error('Resource not found.');
        } else if (error.message.includes('timeout')) {
            throw new Error('Request timeout. Please check your connection and try again.');
        } else {
            throw error;
        }
    }
}
// Create global instance
window.jupiterService = new JupiterService();
