/**
 * Jupiter Service - Core API communication layer
 * Handles all HTTP requests to the Jupiter document management system
 */

class JupiterService {
    constructor() {
        this.baseUrl = '';
        this.apiEndpoint = '/api/v1';
        this.authToken = null;
        this.timeout = 30000; // 30 seconds
    }

    /**
     * Initialize the service with configuration
     */
    initialize(config) {
        this.baseUrl = config.serverUrl || '';
        this.apiEndpoint = config.apiEndpoint || '/api/v1';
        this.timeout = config.timeout || 30000;
        this.authToken = config.authToken || null;
    }

    /**
     * Get the full API URL
     */
    getApiUrl(endpoint) {
        return `${this.baseUrl}${this.apiEndpoint}${endpoint}`;
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
                throw new Error('Request timeout');
            }
            throw error;
        }
    }

    // Authentication Methods
    async login(username, password) {
        const response = await this.makeRequest('POST', '/auth/login', {
            username: username,
            password: password
        });
        
        if (response.token) {
            this.authToken = response.token;
        }
        
        return response;
    }

    async logout() {
        try {
            await this.makeRequest('POST', '/auth/logout');
        } finally {
            this.authToken = null;
        }
    }

    async validateSession() {
        return await this.makeRequest('GET', '/auth/validate');
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

    // Document Methods
    async getDocument(documentId) {
        return await this.makeRequest('GET', `/documents/${documentId}`);
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
        return await this.makeRequest('DELETE', `/documents/${documentId}`);
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
