/**
 * DocumentStateManager - Manages document state tracking for Jupiter DMS integration
 * Tracks whether a document is new or existing, and stores document metadata
 */
class DocumentStateManager {
    constructor() {
        this.SETTINGS_KEY = 'jupiterDMS';
        this.STATE_KEY = 'documentState';
        this.METADATA_KEY = 'documentMetadata';
    }
    /**
     * Initialize the document state manager
     */
    async initialize() {
        try {
            await Office.onReady();
            this.isInitialized = true;
        } catch (error) {
            console.error('Failed to initialize DocumentStateManager:', error);
            throw error;
        }
    }
    /**
     * Check if document is new (never saved to Jupiter DMS)
     * @returns {Promise<boolean>} True if document is new
     */
    async isNewDocument() {
        try {
            const state = await this.getDocumentState();
            return !state || !state.documentId;
        } catch (error) {
            console.error('Error checking if document is new:', error);
            return true; // Default to new if we can't determine state
        }
    }
    /**
     * Check if document exists in Jupiter DMS
     * @returns {Promise<boolean>} True if document exists in DMS
     */
    async isExistingDocument() {
        return !(await this.isNewDocument());
    }
    /**
     * Get the current document state
     * @returns {Promise<Object|null>} Document state object or null
     */
    async getDocumentState() {
        try {
            // Office.js settings.get() is synchronous and doesn't take a callback
            const state = Office.context.document.settings.get(this.STATE_KEY);
            return state || null;
        } catch (error) {
            console.error('Error getting document state:', error);
            return null;
        }
    }
    /**
     * Set the document state
     * @param {Object} state - Document state object
     * @param {string} state.documentId - Jupiter DMS document ID
     * @param {string} state.documentName - Document name
     * @param {string} state.folderId - Folder ID where document is stored
     * @param {string} state.folderPath - Full folder path
     * @param {number} state.version - Current version number
     * @param {string} state.checkoutStatus - Checkout status
     * @param {Date} state.lastSaved - Last saved timestamp
     */
    async setDocumentState(state) {
        try {
            const stateWithTimestamp = {
                ...state,
                lastUpdated: new Date().toISOString()
            };
            return new Promise((resolve, reject) => {
                Office.context.document.settings.set(this.STATE_KEY, stateWithTimestamp);
                Office.context.document.settings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(true);
                    } else {
                        console.error('Failed to save document state:', result.error);
                        reject(new Error('Failed to save document state'));
                    }
                });
            });
        } catch (error) {
            console.error('Error setting document state:', error);
            throw error;
        }
    }
    /**
     * Mark document as new (clear existing state)
     */
    async markAsNewDocument() {
        try {
            return new Promise((resolve, reject) => {
                Office.context.document.settings.remove(this.STATE_KEY);
                Office.context.document.settings.remove(this.METADATA_KEY);
                Office.context.document.settings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(true);
                    } else {
                        console.error('Failed to mark document as new:', result.error);
                        reject(new Error('Failed to mark document as new'));
                    }
                });
            });
        } catch (error) {
            console.error('Error marking document as new:', error);
            throw error;
        }
    }
    /**
     * Mark document as existing in Jupiter DMS
     * @param {Object} documentInfo - Document information from Jupiter DMS
     */
    async markAsExistingDocument(documentInfo) {
        const state = {
            documentId: documentInfo.id,
            documentName: documentInfo.name,
            folderId: documentInfo.folderId,
            folderPath: documentInfo.folderPath,
            version: documentInfo.currentVersion,
            checkoutStatus: documentInfo.checkoutStatus,
            lastSaved: new Date().toISOString()
        };
        await this.setDocumentState(state);
    }
    /**
     * Get document metadata
     * @returns {Promise<Object|null>} Document metadata or null
     */
    async getDocumentMetadata() {
        try {
            return new Promise((resolve, reject) => {
                Office.context.document.settings.get(this.METADATA_KEY, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value);
                    } else {
                        resolve(null);
                    }
                });
            });
        } catch (error) {
            console.error('Error getting document metadata:', error);
            return null;
        }
    }
    /**
     * Set document metadata
     * @param {Object} metadata - Document metadata
     * @param {string} metadata.title - Document title
     * @param {string} metadata.description - Document description
     * @param {string} metadata.tags - Document tags
     * @param {string} metadata.author - Document author
     */
    async setDocumentMetadata(metadata) {
        try {
            return new Promise((resolve, reject) => {
                Office.context.document.settings.set(this.METADATA_KEY, metadata);
                Office.context.document.settings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(true);
                    } else {
                        console.error('Failed to save document metadata:', result.error);
                        reject(new Error('Failed to save document metadata'));
                    }
                });
            });
        } catch (error) {
            console.error('Error setting document metadata:', error);
            throw error;
        }
    }
    /**
     * Get the document name from Word
     * @returns {Promise<string>} Document name
     */
    async getWordDocumentName() {
        try {
            return new Promise((resolve, reject) => {
                Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        // Extract name from the file or use a default
                        const name = Office.context.document.url ? 
                            Office.context.document.url.split('/').pop().split('\\').pop() : 
                            'Untitled Document.docx';
                        resolve(name);
                    } else {
                        resolve('Untitled Document.docx');
                    }
                });
            });
        } catch (error) {
            console.error('Error getting Word document name:', error);
            return 'Untitled Document.docx';
        }
    }
    /**
     * Check if document has unsaved changes
     * @returns {Promise<boolean>} True if document has unsaved changes
     */
    async hasUnsavedChanges() {
        try {
            return new Promise((resolve, reject) => {
                Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        // This is a simplified check - in reality, you might want to compare
                        // file hashes or use other methods to detect changes
                        resolve(true); // Assume there are always changes for now
                    } else {
                        resolve(false);
                    }
                });
            });
        } catch (error) {
            console.error('Error checking for unsaved changes:', error);
            return false;
        }
    }
    /**
     * Get document checkout information
     * @returns {Promise<Object|null>} Checkout information or null
     */
    async getCheckoutInfo() {
        const state = await this.getDocumentState();
        if (state && state.checkoutStatus) {
            return {
                status: state.checkoutStatus,
                checkedOutBy: state.checkedOutBy,
                checkedOutOn: state.checkedOutOn,
                checkoutExpiry: state.checkoutExpiry
            };
        }
        return null;
    }
    /**
     * Update checkout status
     * @param {string} status - Checkout status
     * @param {Object} checkoutInfo - Additional checkout information
     */
    async updateCheckoutStatus(status, checkoutInfo = {}) {
        const currentState = await this.getDocumentState();
        if (currentState) {
            const updatedState = {
                ...currentState,
                checkoutStatus: status,
                ...checkoutInfo
            };
            await this.setDocumentState(updatedState);
        }
    }
    /**
     * Clear all document state and metadata
     */
    async clearDocumentState() {
        try {
            return new Promise((resolve, reject) => {
                Office.context.document.settings.remove(this.STATE_KEY);
                Office.context.document.settings.remove(this.METADATA_KEY);
                Office.context.document.settings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(true);
                    } else {
                        console.error('Failed to clear document state:', result.error);
                        reject(new Error('Failed to clear document state'));
                    }
                });
            });
        } catch (error) {
            console.error('Error clearing document state:', error);
            throw error;
        }
    }
}
// Export for use in other modules
window.DocumentStateManager = DocumentStateManager;
