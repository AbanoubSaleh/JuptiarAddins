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
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
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
    async searchDocuments(query, type = 'filename') {
        const params = new URLSearchParams({
            query: query,
            type: type
        });
        return await this.makeRequest('GET', `/search/documents?${params.toString()}`);
    }
    async fullTextSearch(query) {
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

            // Parse specific error messages from backend
            let errorMessage = error.message;
            if (error.message.includes('already checked out')) {
                errorMessage = 'This document is already checked out by another user. Please try again later.';
            } else if (error.message.includes('not found')) {
                errorMessage = 'Document not found. It may have been deleted or moved.';
            } else if (error.message.includes('unauthorized') || error.message.includes('403')) {
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
    async checkInDocument(documentId, fileBlob, versionComment = '') {
        try {
            console.log(`📥 Attempting to check in document: ${documentId}`);

            // Create FormData for file upload
            const formData = new FormData();
            if (fileBlob) {
                formData.append('file', fileBlob, 'document.docx');
            }
            if (versionComment) {
                formData.append('versionComment', versionComment);
            }

            const response = await fetch(`${this.baseUrl}${this.apiEndpoint}/documents/${documentId}/checkin`, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${this.authToken}`
                },
                body: formData
            });

            if (!response.ok) {
                const errorData = await response.json().catch(() => ({ message: `HTTP ${response.status}` }));
                throw new Error(errorData.message || `HTTP error! status: ${response.status}`);
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
            if (error.message.includes('not checked out')) {
                errorMessage = 'This document is not currently checked out. Please check out the document first.';
            } else if (error.message.includes('not found')) {
                errorMessage = 'Document not found. It may have been deleted or moved.';
            } else if (error.message.includes('unauthorized') || error.message.includes('403')) {
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
