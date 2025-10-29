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
     * Scan document for Jupiter DMS indicators on startup
     * This method detects if a document opened via File > Open is a Jupiter document
     * @returns {Promise<Object|null>} Jupiter document info if found, null otherwise
     */
    async scanDocumentForJupiterMetadata() {
        try {
            console.log('🔍 Scanning document for Jupiter metadata...');

            // Method 1: Check Office.context.document.settings for existing Jupiter data
            const existingState = await this.getDocumentState();
            if (existingState && existingState.documentId) {
                console.log('✅ Found Jupiter metadata in document settings:', existingState);
                return existingState;
            }

            // Method 2: Check for legacy Jupiter document ID setting
            const legacyDocumentId = Office.context.document.settings.get('currentJuptiarDocumentId');
            if (legacyDocumentId) {
                console.log('✅ Found legacy Jupiter document ID:', legacyDocumentId);

                // Try to get document info from API
                try {
                    if (window.jupiterService) {
                        const documentInfo = await window.jupiterService.getDocumentById(legacyDocumentId);
                        if (documentInfo) {
                            // Convert legacy setting to new format
                            const jupiterState = {
                                documentId: legacyDocumentId,
                                documentName: documentInfo.name,
                                folderId: documentInfo.folderId,
                                folderPath: documentInfo.folderPath || 'Unknown',
                                version: documentInfo.currentVersion || 1,
                                checkoutStatus: documentInfo.checkoutStatus || 'Available',
                                lastSaved: new Date().toISOString()
                            };

                            // Save in new format
                            await this.setDocumentState(jupiterState);

                            // Remove legacy setting
                            Office.context.document.settings.remove('currentJuptiarDocumentId');
                            await new Promise((resolve) => {
                                Office.context.document.settings.saveAsync(() => resolve());
                            });

                            console.log('✅ Migrated legacy Jupiter document to new format');
                            return jupiterState;
                        }
                    }
                } catch (error) {
                    console.warn('⚠️ Could not fetch document info for legacy ID:', error.message);
                }
            }

            // Method 3: Check document properties for Jupiter metadata
            // This would require custom document properties to be set when documents are saved
            try {
                const documentProperties = await this.getDocumentCustomProperties();
                if (documentProperties && documentProperties.jupiterDocumentId) {
                    console.log('✅ Found Jupiter metadata in document properties:', documentProperties);

                    const jupiterState = {
                        documentId: documentProperties.jupiterDocumentId,
                        documentName: documentProperties.jupiterDocumentName || 'Unknown',
                        folderId: documentProperties.jupiterFolderId || null,
                        folderPath: documentProperties.jupiterFolderPath || 'Unknown',
                        version: parseInt(documentProperties.jupiterVersion) || 1,
                        checkoutStatus: documentProperties.jupiterCheckoutStatus || 'Available',
                        lastSaved: documentProperties.jupiterLastSaved || new Date().toISOString()
                    };

                    // Save to document settings for faster future access
                    await this.setDocumentState(jupiterState);

                    return jupiterState;
                }
            } catch (error) {
                console.warn('⚠️ Could not read document custom properties:', error.message);
            }

            // Method 4: Check document file name patterns (if saved with Jupiter naming convention)
            try {
                const documentUrl = Office.context.document.url;
                if (documentUrl && this.isJupiterDocumentPath(documentUrl)) {
                    console.log('✅ Document path suggests Jupiter origin:', documentUrl);
                    // Could extract document ID from file name if following a pattern
                    // This is implementation-specific based on how Jupiter saves files
                }
            } catch (error) {
                console.warn('⚠️ Could not check document URL:', error.message);
            }

            console.log('❌ No Jupiter metadata found - treating as new document');
            return null;

        } catch (error) {
            console.error('Error scanning document for Jupiter metadata:', error);
            return null;
        }
    }

    /**
     * Get document custom properties (Word-specific)
     * @returns {Promise<Object|null>} Custom properties object or null
     */
    async getDocumentCustomProperties() {
        try {
            return await Word.run(async (context) => {
                const properties = context.document.properties.customProperties;
                properties.load('items');
                await context.sync();

                const jupiterProperties = {};
                properties.items.forEach(prop => {
                    if (prop.key.startsWith('jupiter')) {
                        jupiterProperties[prop.key] = prop.value;
                    }
                });

                return Object.keys(jupiterProperties).length > 0 ? jupiterProperties : null;
            });
        } catch (error) {
            console.warn('Could not read custom properties (might not be Word):', error.message);
            return null;
        }
    }

    /**
     * Set document custom properties for Jupiter metadata
     * @param {Object} jupiterState - Jupiter document state
     */
    async setDocumentCustomProperties(jupiterState) {
        try {
            await Word.run(async (context) => {
                const properties = context.document.properties.customProperties;

                // Set Jupiter metadata as custom properties
                properties.add('jupiterDocumentId', jupiterState.documentId);
                properties.add('jupiterDocumentName', jupiterState.documentName || '');
                properties.add('jupiterFolderId', jupiterState.folderId || '');
                properties.add('jupiterFolderPath', jupiterState.folderPath || '');
                properties.add('jupiterVersion', jupiterState.version?.toString() || '1');
                properties.add('jupiterCheckoutStatus', jupiterState.checkoutStatus || 'Available');
                properties.add('jupiterLastSaved', jupiterState.lastSaved || new Date().toISOString());

                await context.sync();
                console.log('✅ Jupiter metadata saved to document custom properties');
            });
        } catch (error) {
            console.warn('Could not set custom properties (might not be Word):', error.message);
        }
    }

    /**
     * Check if document path suggests Jupiter origin
     * @param {string} documentUrl - Document URL or path
     * @returns {boolean} True if path suggests Jupiter origin
     */
    isJupiterDocumentPath(documentUrl) {
        if (!documentUrl) return false;

        // Check for Jupiter-specific path patterns
        // This is implementation-specific - adjust based on how Jupiter saves files
        const jupiterPatterns = [
            /jupiter/i,
            /dms/i,
            /document.*management/i,
            // Add more patterns based on your Jupiter file naming conventions
        ];

        return jupiterPatterns.some(pattern => pattern.test(documentUrl));
    }

    /**
     * Perform comprehensive document detection on add-in startup
     * This is the main method called during initialization
     * @returns {Promise<string>} Document type: 'new', 'jupiter', or 'external'
     */
    async detectDocumentType() {
        try {
            console.log('🔍 Starting comprehensive document type detection...');

            // First, check if we already have Jupiter metadata
            const existingState = await this.getDocumentState();
            if (existingState && existingState.documentId) {
                console.log('✅ Document type: JUPITER (from existing state)');
                return 'jupiter';
            }

            // Check if this is a completely new document (no content, no file path)
            const isBlankDocument = await this.isBlankDocument();
            if (isBlankDocument) {
                console.log('✅ Document type: NEW (blank document)');
                return 'new';
            }

            // Scan for Jupiter metadata in various locations
            const jupiterMetadata = await this.scanDocumentForJupiterMetadata();
            if (jupiterMetadata) {
                console.log('✅ Document type: JUPITER (detected from metadata)');
                return 'jupiter';
            }

            // If we reach here, it's likely an external document
            console.log('✅ Document type: EXTERNAL (no Jupiter metadata found)');
            return 'external';

        } catch (error) {
            console.error('Error detecting document type:', error);
            // Default to new document if detection fails
            return 'new';
        }
    }

    /**
     * Check if document is blank/new
     * @returns {Promise<boolean>} True if document appears to be blank
     */
    async isBlankDocument() {
        try {
            return await Word.run(async (context) => {
                const body = context.document.body;
                body.load('text');
                await context.sync();

                // Check if document has minimal content (just whitespace/empty)
                const text = body.text.trim();
                const hasMinimalContent = text.length === 0 || text.length < 10;

                // Also check if document has a file path
                const hasFilePath = Office.context.document.url &&
                                  Office.context.document.url !== 'about:blank' &&
                                  !Office.context.document.url.includes('blank');

                return hasMinimalContent && !hasFilePath;
            });
        } catch (error) {
            console.warn('Could not check if document is blank (might not be Word):', error.message);
            // Fallback: check if we have any document state
            const state = await this.getDocumentState();
            return !state;
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
