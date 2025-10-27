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
        console.log('JupiterService: Constructing URL:', {
            baseUrl: this.baseUrl,
            apiEndpoint: this.apiEndpoint,
            endpoint: endpoint,
            fullUrl: fullUrl
        });
        return fullUrl;
    }

    /**
     * Make HTTP request with authentication
     */
    async makeRequest(method, endpoint, data = null, options = {}) {
        const url = this.getApiUrl(endpoint);
        console.log('JupiterService: Making request to:', url);
        console.log('JupiterService: Method:', method);
        console.log('JupiterService: Data:', data);

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

        console.log('JupiterService: Request options:', requestOptions);

        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), this.timeout);

            requestOptions.signal = controller.signal;

            console.log('JupiterService: Sending fetch request...');
            const response = await fetch(url, requestOptions);
            console.log('JupiterService: Response received:', response.status, response.statusText);
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
        try {
            console.log('JupiterService: Login attempt for user:', username);
            console.log('JupiterService: Current configuration:', {
                baseUrl: this.baseUrl,
                apiEndpoint: this.apiEndpoint
            });
            console.log('JupiterService: API URL will be:', this.getApiUrl('/Auth/login'));

            const response = await this.makeRequest('POST', '/Auth/login', {
                username: username,
                password: password
            });

            console.log('JupiterService: Login response received:', response);

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
