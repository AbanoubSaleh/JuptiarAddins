/**
 * DocumentUploader - Handles document upload operations with duplicate checking and metadata
 * Manages the complete save workflow for new and existing documents
 */
class DocumentUploader {
    constructor() {
        this.jupiterService = null;
        this.documentStateManager = null;
        this.isInitialized = false;
    }

    /**
     * Initialize the document uploader
     * @param {JupiterService} jupiterService - Jupiter service instance
     * @param {DocumentStateManager} documentStateManager - Document state manager instance
     */
    async initialize(jupiterService, documentStateManager) {
        try {
            this.jupiterService = jupiterService;
            this.documentStateManager = documentStateManager;
            this.isInitialized = true;
            console.log('DocumentUploader initialized');
        } catch (error) {
            console.error('Failed to initialize DocumentUploader:', error);
            throw error;
        }
    }

    /**
     * Save new document to Jupiter DMS
     * @param {Object} saveOptions - Save options
     * @param {string} saveOptions.name - Document name
     * @param {string} saveOptions.folderId - Target folder ID
     * @param {string} saveOptions.title - Document title (optional)
     * @param {string} saveOptions.description - Document description (optional)
     * @param {string} saveOptions.tags - Document tags (optional)
     * @param {string} saveOptions.duplicateAction - Action for duplicates: "rename", "replace", "version"
     * @returns {Promise<Object>} Upload result
     */
    async saveNewDocument(saveOptions) {
        try {
            if (!this.isInitialized) {
                throw new Error('DocumentUploader not initialized');
            }

            // Get document content from Word
            const documentBlob = await this.getWordDocumentBlob();
            
            // Check for duplicates first
            const duplicateCheck = await this.checkDuplicateName(saveOptions.name, saveOptions.folderId);
            
            if (duplicateCheck.exists && !saveOptions.duplicateAction) {
                // Show duplicate dialog and get user choice
                const userChoice = await this.showDuplicateDialog(duplicateCheck);
                saveOptions.duplicateAction = userChoice.action;
                
                if (userChoice.action === 'cancel') {
                    return { success: false, cancelled: true };
                }
                
                if (userChoice.action === 'rename') {
                    saveOptions.name = userChoice.newName || duplicateCheck.suggestedName;
                }
            }

            // Prepare form data for upload
            const formData = new FormData();
            formData.append('file', documentBlob, saveOptions.name);
            formData.append('Name', saveOptions.name);
            formData.append('FolderId', saveOptions.folderId);
            formData.append('duplicateAction', saveOptions.duplicateAction || 'rename');
            
            if (saveOptions.title) formData.append('Title', saveOptions.title);
            if (saveOptions.description) formData.append('Description', saveOptions.description);
            if (saveOptions.tags) formData.append('Tags', saveOptions.tags);

            // Upload document
            const result = await this.uploadWithOptions(formData);
            
            if (result.success) {
                // Update document state to mark as existing
                await this.documentStateManager.markAsExistingDocument(result.document);
                
                // Update metadata
                await this.documentStateManager.setDocumentMetadata({
                    title: saveOptions.title,
                    description: saveOptions.description,
                    tags: saveOptions.tags
                });
            }

            return result;
        } catch (error) {
            console.error('Error saving new document:', error);
            throw error;
        }
    }

    /**
     * Check if document name exists in folder
     * @param {string} name - Document name
     * @param {string} folderId - Folder ID
     * @returns {Promise<Object>} Duplicate check result
     */
    async checkDuplicateName(name, folderId) {
        try {
            const response = await this.jupiterService.checkDuplicateName(name, folderId);
            return response;
        } catch (error) {
            console.error('Error checking duplicate name:', error);
            return { exists: false, duplicateDocument: null, suggestedName: name };
        }
    }

    /**
     * Show duplicate name resolution dialog
     * @param {Object} duplicateInfo - Duplicate information
     * @returns {Promise<Object>} User choice
     */
    async showDuplicateDialog(duplicateInfo) {
        return new Promise((resolve) => {
            // Create and show duplicate dialog
            const dialog = this.createDuplicateDialog(duplicateInfo, resolve);
            document.body.appendChild(dialog);
            
            // Show dialog with animation
            setTimeout(() => {
                dialog.classList.add('show');
            }, 10);
        });
    }

    /**
     * Create duplicate name resolution dialog
     * @param {Object} duplicateInfo - Duplicate information
     * @param {Function} resolve - Promise resolve function
     * @returns {HTMLElement} Dialog element
     */
    createDuplicateDialog(duplicateInfo, resolve) {
        const dialog = document.createElement('div');
        dialog.className = 'duplicate-dialog-overlay';
        
        dialog.innerHTML = `
            <div class="duplicate-dialog">
                <div class="duplicate-dialog-header">
                    <h3>Document Already Exists</h3>
                    <button class="close-btn" onclick="this.closest('.duplicate-dialog-overlay').remove(); resolve({action: 'cancel'})">&times;</button>
                </div>
                <div class="duplicate-dialog-content">
                    <div class="warning-icon">⚠️</div>
                    <p>A document named <strong>"${duplicateInfo.duplicateDocument?.name}"</strong> already exists in this folder.</p>
                    <p>What would you like to do?</p>
                    
                    <div class="duplicate-options">
                        <div class="option-card" data-action="replace">
                            <div class="option-icon">🔄</div>
                            <div class="option-content">
                                <h4>Replace</h4>
                                <p>Replace the existing document with this one</p>
                            </div>
                        </div>
                        
                        <div class="option-card" data-action="rename">
                            <div class="option-icon">📝</div>
                            <div class="option-content">
                                <h4>Keep Both</h4>
                                <p>Save with a new name: <strong>"${duplicateInfo.suggestedName}"</strong></p>
                            </div>
                        </div>
                        
                        <div class="option-card" data-action="cancel">
                            <div class="option-icon">❌</div>
                            <div class="option-content">
                                <h4>Cancel</h4>
                                <p>Don't save the document</p>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="duplicate-dialog-footer">
                    <button class="btn btn-secondary" onclick="this.closest('.duplicate-dialog-overlay').remove(); resolve({action: 'cancel'})">Cancel</button>
                </div>
            </div>
        `;

        // Add event listeners for option cards
        const optionCards = dialog.querySelectorAll('.option-card');
        optionCards.forEach(card => {
            card.addEventListener('click', () => {
                const action = card.getAttribute('data-action');
                const result = { action };
                
                if (action === 'rename') {
                    result.newName = duplicateInfo.suggestedName;
                }
                
                dialog.remove();
                resolve(result);
            });
        });

        return dialog;
    }

    /**
     * Upload document with options
     * @param {FormData} formData - Form data with file and metadata
     * @returns {Promise<Object>} Upload result
     */
    async uploadWithOptions(formData) {
        try {
            const response = await fetch(`${this.jupiterService.baseUrl}/api/documents/upload-with-options`, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${this.jupiterService.getToken()}`
                },
                body: formData
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Upload failed: ${response.status} - ${errorText}`);
            }

            const document = await response.json();
            return { success: true, document };
        } catch (error) {
            console.error('Error uploading document:', error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Get Word document as blob
     * @returns {Promise<Blob>} Document blob
     */
    async getWordDocumentBlob() {
        return new Promise((resolve, reject) => {
            Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const file = result.value;
                    const sliceCount = file.sliceCount;
                    const slices = [];
                    let slicesReceived = 0;

                    const getSlice = (sliceIndex) => {
                        file.getSliceAsync(sliceIndex, (sliceResult) => {
                            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                                slices[sliceIndex] = sliceResult.value.data;
                                slicesReceived++;

                                if (slicesReceived === sliceCount) {
                                    // Combine all slices into a single blob
                                    const blob = new Blob(slices, { 
                                        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
                                    });
                                    file.closeAsync();
                                    resolve(blob);
                                }
                            } else {
                                file.closeAsync();
                                reject(new Error('Failed to get document slice: ' + sliceResult.error.message));
                            }
                        });
                    };

                    // Get all slices
                    for (let i = 0; i < sliceCount; i++) {
                        getSlice(i);
                    }
                } else {
                    reject(new Error('Failed to get document: ' + result.error.message));
                }
            });
        });
    }

    /**
     * Create new version of existing document
     * @param {string} documentId - Document ID
     * @param {string} versionComment - Version comment
     * @returns {Promise<Object>} Version creation result
     */
    async createNewVersion(documentId, versionComment) {
        try {
            const documentBlob = await this.getWordDocumentBlob();
            
            const formData = new FormData();
            formData.append('file', documentBlob);
            if (versionComment) {
                formData.append('versionComment', versionComment);
            }

            const response = await fetch(`${this.jupiterService.baseUrl}/api/documents/${documentId}/versions`, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${this.jupiterService.getToken()}`
                },
                body: formData
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Version creation failed: ${response.status} - ${errorText}`);
            }

            const document = await response.json();
            
            // Update document state
            await this.documentStateManager.markAsExistingDocument(document);
            
            return { success: true, document };
        } catch (error) {
            console.error('Error creating new version:', error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Check in document with changes
     * @param {string} documentId - Document ID
     * @param {string} versionComment - Version comment
     * @returns {Promise<Object>} Check-in result
     */
    async checkInDocument(documentId, versionComment) {
        try {
            const documentBlob = await this.getWordDocumentBlob();
            
            const formData = new FormData();
            formData.append('file', documentBlob);
            if (versionComment) {
                formData.append('versionComment', versionComment);
            }

            const response = await fetch(`${this.jupiterService.baseUrl}/api/documents/${documentId}/checkin`, {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${this.jupiterService.getToken()}`
                },
                body: formData
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Check-in failed: ${response.status} - ${errorText}`);
            }

            const document = await response.json();
            
            // Update document state
            await this.documentStateManager.updateCheckoutStatus('Available');
            
            return { success: true, document };
        } catch (error) {
            console.error('Error checking in document:', error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Get suggested document name from Word
     * @returns {Promise<string>} Suggested document name
     */
    async getSuggestedDocumentName() {
        try {
            // Try to get the document name from Word
            const name = await this.documentStateManager.getWordDocumentName();
            
            // If it's a generic name, suggest a better one
            if (name === 'Untitled Document.docx' || name.startsWith('Document')) {
                const now = new Date();
                return `Document_${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}.docx`;
            }
            
            return name;
        } catch (error) {
            console.error('Error getting suggested document name:', error);
            return 'New Document.docx';
        }
    }
}

// Export for use in other modules
window.DocumentUploader = DocumentUploader;
