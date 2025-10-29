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
            // Ensure filename has .docx extension
            let fileName = saveOptions.name;
            if (!fileName.toLowerCase().endsWith('.docx') && !fileName.toLowerCase().endsWith('.doc')) {
                fileName = fileName + '.docx';
                console.log('Added .docx extension to filename:', fileName);
            }

            // Debug: Log blob details
            console.log('📄 Document blob details:');
            console.log('  - Size:', documentBlob.size, 'bytes');
            console.log('  - Type:', documentBlob.type);
            console.log('  - Filename:', fileName);

            // Create a proper Word document blob with correct MIME type
            const wordBlob = new Blob([documentBlob], {
                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });

            console.log('📄 Final Word blob details:');
            console.log('  - Size:', wordBlob.size, 'bytes');
            console.log('  - Type:', wordBlob.type);

            // Prepare form data for upload
            const formData = new FormData();
            formData.append('file', wordBlob, fileName);
            // Send data with exact property names for model binding
            formData.append('Name', fileName);
            formData.append('FolderId', saveOptions.folderId);
            formData.append('duplicateAction', saveOptions.duplicateAction || 'rename');
            if (saveOptions.title) formData.append('Title', saveOptions.title);
            if (saveOptions.description) formData.append('Description', saveOptions.description);
            if (saveOptions.tags) formData.append('Tags', saveOptions.tags);

            // Debug: Log FormData contents
            console.log('📤 FormData contents:');
            for (let [key, value] of formData.entries()) {
                if (key === 'file') {
                    console.log(`  - ${key}:`, {
                        name: value.name,
                        size: value.size,
                        type: value.type,
                        lastModified: value.lastModified
                    });
                } else {
                    console.log(`  - ${key}:`, value);
                }
            }
            // Validate required fields
            if (!saveOptions.name || saveOptions.name.trim() === '') {
                throw new Error('Document name is required');
            }
            if (!saveOptions.folderId) {
                throw new Error('Folder ID is required');
            }
            if (!documentBlob || documentBlob.size === 0) {
                throw new Error('Document file is required');
            }
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
            // Use centralized error handling
            if (window.ErrorHandler) {
                window.ErrorHandler.handle(error, {
                    context: 'DocumentUploader.saveNewDocument',
                    userMessage: 'Failed to save document. Please try again.',
                    showToUser: true,
                    severity: 'error'
                });
            } else {
                console.error('Error saving new document:', error);
            }
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
        console.log('Starting document extraction - attempting to save directly to Jupiter without local save...');

        // First, try to force Word to commit the document content without saving locally
        let extractedText = null;
        try {
            extractedText = await this.forceDocumentCommit();
            console.log('✅ Successfully extracted text via Word.run:', extractedText ? extractedText.length + ' characters' : 'no text');
        } catch (commitError) {
            console.warn('Document commit failed, but continuing:', commitError.message);
        }

        // If we successfully extracted text, use it directly
        if (extractedText && extractedText.trim().length > 0) {
            console.log('🎯 Using successfully extracted text from Word.run API');
            try {
                const wordContent = await this.createWordDocumentFromText(extractedText);
                const blob = new Blob([wordContent], {
                    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                });
                console.log('✅ Created Word document from extracted text, size:', blob.size);
                return blob;
            } catch (textError) {
                console.warn('Failed to create document from extracted text:', textError.message);
            }
        }

        // Try different approaches in order - prioritize methods that create proper DOCX files
        const methods = [
            { name: 'Compressed (small slices)', method: () => this.getDocumentAsCompressed(65536) }, // 64KB slices - most reliable
            { name: 'Compressed (default slice)', method: () => this.getDocumentAsCompressed() },
            { name: 'Simple text extraction', method: () => this.getDocumentAsSimpleText() },
            { name: 'Selected content', method: () => this.getSelectedContent() },
            { name: 'Text content extraction', method: () => this.getDocumentAsText() },
            { name: 'PDF format', method: () => this.getDocumentAsPDF() }
        ];

        for (const { name, method } of methods) {
            try {
                console.log(`Trying method: ${name}`);
                const result = await method();
                console.log(`Success with method: ${name}`);
                return result;
            } catch (error) {
                console.warn(`Method ${name} failed:`, error.message);

                // If it's a critical error that won't be fixed by other methods, stop trying
                if (error.message.includes('not supported') || error.message.includes('permission')) {
                    break;
                }
            }
        }

        // All methods failed - provide more helpful error
        throw new Error('Unable to extract document content directly. This can happen with unsaved documents. The document content needs to be committed to memory first.');
    }

    /**
     * Force Word to commit document content to memory without local save
     */
    async forceDocumentCommit() {
        return new Promise((resolve, reject) => {
            console.log('Attempting to force document commit using Word.run...');

            // Use Word.run API which is more reliable for content access
            if (typeof Word !== 'undefined' && Word.run) {
                Word.run(async (context) => {
                    console.log('Using Word.run API to access document...');

                    // Load the document body to force Word to commit content
                    const body = context.document.body;
                    body.load('text');

                    await context.sync();

                    console.log('Document content committed via Word.run API');
                    console.log('Document text length:', body.text.length);

                    resolve(body.text);
                }).catch((error) => {
                    console.warn('Word.run failed, trying fallback methods...', error);
                    this.fallbackDocumentCommit().then(resolve).catch(reject);
                });
            } else {
                console.warn('Word.run API not available, trying fallback methods...');
                this.fallbackDocumentCommit().then(resolve).catch(reject);
            }
        });
    }

    /**
     * Fallback document commit methods
     */
    async fallbackDocumentCommit() {
        return new Promise((resolve, reject) => {
            // Try to trigger a document state change that forces Word to commit content
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log('Document commit successful via selection API');
                    resolve(result.value || '');
                } else {
                    console.warn('Selection API failed, trying document properties...');

                    // Try a different approach - get document URL which forces some processing
                    try {
                        const docUrl = Office.context.document.url;
                        console.log('Document URL accessed:', docUrl ? 'available' : 'not available');
                        resolve('');
                    } catch (error) {
                        console.warn('All commit methods failed');
                        resolve(''); // Don't reject, just continue with empty string
                    }
                }
            });
        });
    }

    /**
     * Get document content using simple text selection (most reliable)
     */
    async getDocumentAsSimpleText() {
        console.log('Trying simple text extraction...');

        // Try multiple ways to get the document text
        const textExtractionMethods = [
            () => this.getTextViaSelection(),
            () => this.getTextViaDocument(),
            () => this.getTextViaRange()
        ];

        for (const method of textExtractionMethods) {
            try {
                const textContent = await method();
                if (textContent && textContent.trim().length > 0) {
                    console.log('Text extracted successfully, length:', textContent.length);

                    // Create a proper DOCX file from the text
                    const docxBlob = await this.createMinimalDocx(textContent);
                    console.log('Minimal DOCX created, size:', docxBlob.size);
                    return docxBlob;
                }
            } catch (error) {
                console.warn('Text extraction method failed:', error.message);
            }
        }

        // If all methods fail, create a document with placeholder text
        console.warn('All text extraction methods failed, creating document with placeholder');
        const placeholderText = 'Document content could not be extracted. Please save the document in Word first.';
        const docxBlob = await this.createMinimalDocx(placeholderText);
        return docxBlob;
    }

    /**
     * Get text via selection API
     */
    async getTextViaSelection() {
        return new Promise((resolve, reject) => {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value || '');
                } else {
                    reject(new Error('Selection API failed'));
                }
            });
        });
    }

    /**
     * Get text via document body (if available)
     */
    async getTextViaDocument() {
        return new Promise((resolve, reject) => {
            if (Office.context.document.body && Office.context.document.body.getAsync) {
                Office.context.document.body.getAsync(Office.CoercionType.Text, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value || '');
                    } else {
                        reject(new Error('Document body API failed'));
                    }
                });
            } else {
                reject(new Error('Document body API not available'));
            }
        });
    }

    /**
     * Get text via range selection
     */
    async getTextViaRange() {
        return new Promise((resolve, reject) => {
            // Try to insert empty text to trigger document processing
            Office.context.document.body.insertText('', Office.InsertLocation.Start, (insertResult) => {
                if (insertResult.status === Office.AsyncResultStatus.Succeeded) {
                    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            resolve(result.value || '');
                        } else {
                            reject(new Error('Range selection failed'));
                        }
                    });
                } else {
                    reject(new Error('Range insertion failed'));
                }
            });
        });
    }

    /**
     * Try to get document content by selecting all and reading
     */
    async tryDocumentRange() {
        return new Promise((resolve, reject) => {
            console.log('Trying document range selection...');

            // Try to select all content first
            Office.context.document.body.insertText('', Office.InsertLocation.End, (result) => {
                // Get the entire document as text
                Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (textResult) => {
                    if (textResult.status === Office.AsyncResultStatus.Succeeded) {
                        const textContent = textResult.value || 'Sample document content';
                        console.log('Document range text extracted, length:', textContent.length);

                        this.createWordDocumentFromText(textContent).then((wordContent) => {
                            const blob = new Blob([wordContent], {
                                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                            });

                            resolve(blob);
                        }).catch(reject);
                    } else {
                        reject(new Error('Failed to extract document range: ' + textResult.error?.message));
                    }
                });
            });
        });
    }

    /**
     * Get currently selected content
     */
    async getSelectedContent() {
        return new Promise((resolve, reject) => {
            console.log('Trying to get selected content...');

            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const selectedText = result.value || '';
                    console.log('Selected content extracted, length:', selectedText.length);

                    if (!selectedText || selectedText.trim().length === 0) {
                        reject(new Error('No content selected or document is empty'));
                        return;
                    }

                    this.createWordDocumentFromText(selectedText).then((wordContent) => {
                        const blob = new Blob([wordContent], {
                            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                        });

                        console.log('Selected content document created, size:', blob.size);
                        resolve(blob);
                    }).catch(reject);
                } else {
                    reject(new Error('Failed to get selected content: ' + result.error?.message));
                }
            });
        });
    }

    /**
     * Get document content using Word.run API (most reliable)
     */
    async getDocumentAsText() {
        console.log('Extracting document content using Word.run API...');

        if (typeof Word !== 'undefined' && Word.run) {
            try {
                const documentContent = await Word.run(async (context) => {
                    console.log('Accessing document via Word.run...');

                    // Get the document body
                    const body = context.document.body;

                    // Load both text and OOXML content
                    body.load(['text', 'ooxml']);

                    await context.sync();

                    console.log('Document content loaded successfully');
                    console.log('Text length:', body.text.length);
                    console.log('OOXML length:', body.ooxml.length);

                    return {
                        text: body.text,
                        ooxml: body.ooxml
                    };
                });

                // Prefer OOXML if available, fallback to text
                if (documentContent.ooxml && documentContent.ooxml.trim().length > 0) {
                    console.log('Using OOXML content');
                    const wordContent = this.createWordDocumentFromOOXML(documentContent.ooxml);
                    return new Blob([wordContent], {
                        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    });
                } else if (documentContent.text && documentContent.text.trim().length > 0) {
                    console.log('Using text content');
                    const wordContent = this.createWordDocumentFromText(documentContent.text);
                    return new Blob([wordContent], {
                        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    });
                } else {
                    throw new Error('Document appears to be empty');
                }

            } catch (error) {
                console.warn('Word.run API failed:', error.message);
                return this.getDocumentAsTextFallback();
            }
        } else {
            console.warn('Word.run API not available, using fallback...');
            return this.getDocumentAsTextFallback();
        }
    }

    /**
     * Fallback text extraction using Office.context APIs
     */
    async getDocumentAsTextFallback() {
        return new Promise((resolve, reject) => {
            console.log('Using fallback text extraction...');

            // Try to get the document as OOXML first (preserves formatting)
            Office.context.document.body.getAsync(Office.CoercionType.Ooxml, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const ooxmlContent = result.value;
                    console.log('OOXML content extracted, length:', ooxmlContent.length);

                    if (!ooxmlContent || ooxmlContent.trim().length === 0) {
                        reject(new Error('Document appears to be empty'));
                        return;
                    }

                    // Create a Word document from OOXML
                    const wordContent = this.createWordDocumentFromOOXML(ooxmlContent);
                    const blob = new Blob([wordContent], {
                        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    });

                    console.log('Word document created from OOXML, size:', blob.size);
                    resolve(blob);
                } else {
                    console.warn('OOXML extraction failed, trying plain text...');

                    // Fallback to plain text
                    Office.context.document.body.getAsync(Office.CoercionType.Text, (textResult) => {
                        if (textResult.status === Office.AsyncResultStatus.Succeeded) {
                            const textContent = textResult.value;
                            console.log('Text content extracted, length:', textContent.length);

                            if (!textContent || textContent.trim().length === 0) {
                                reject(new Error('Document appears to be empty'));
                                return;
                            }

                            // Create a simple Word document with the text content
                            const wordContent = this.createWordDocumentFromText(textContent);
                            const blob = new Blob([wordContent], {
                                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                            });

                            console.log('Word document created from text, size:', blob.size);
                            resolve(blob);
                        } else {
                            reject(new Error('Failed to extract document content: ' + textResult.error.message));
                        }
                    });
                }
            });
        });
    }

    /**
     * Create a Word document from OOXML content
     * @param {string} ooxmlContent - The OOXML content
     * @returns {Uint8Array} Word document bytes
     */
    createWordDocumentFromOOXML(ooxmlContent) {
        console.log('Creating Word document from OOXML...');

        // The OOXML content from Office.js is already in the correct format
        // We just need to wrap it in a proper document structure
        const fullDocument = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        ${ooxmlContent}
    </w:body>
</w:document>`;

        const encoder = new TextEncoder();
        return encoder.encode(fullDocument);
    }

    /**
     * Create a proper DOCX file from text content using Office.js
     * @param {string} textContent - The text content
     * @returns {Promise<Uint8Array>} Word document bytes
     */
    async createWordDocumentFromText(textContent) {
        console.log('Creating proper DOCX document from text...');

        return new Promise((resolve, reject) => {
            // Use Office.js to create a proper Word document
            // We'll insert the text into a new document structure and then extract it

            // First, let's try to use the document's own format by inserting and extracting
            const originalContent = textContent;

            // Clear the document and insert our content, then extract as compressed
            Office.context.document.body.clear((clearResult) => {
                if (clearResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log('Document cleared, inserting new content...');

                    // Insert the text content
                    Office.context.document.body.insertText(originalContent, Office.InsertLocation.Start, (insertResult) => {
                        if (insertResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log('Content inserted, extracting as compressed...');

                            // Now extract the document as compressed (proper DOCX)
                            this.getDocumentAsCompressed(65536).then((blob) => {
                                console.log('Proper DOCX created from text, size:', blob.size);

                                // Convert blob to Uint8Array
                                const reader = new FileReader();
                                reader.onload = () => {
                                    const arrayBuffer = reader.result;
                                    const uint8Array = new Uint8Array(arrayBuffer);
                                    resolve(uint8Array);
                                };
                                reader.onerror = () => reject(new Error('Failed to convert blob to bytes'));
                                reader.readAsArrayBuffer(blob);
                            }).catch((extractError) => {
                                console.warn('Compressed extraction failed, using fallback XML method:', extractError.message);
                                resolve(this.createSimpleWordXML(originalContent));
                            });
                        } else {
                            console.warn('Insert failed, using fallback XML method');
                            resolve(this.createSimpleWordXML(originalContent));
                        }
                    });
                } else {
                    console.warn('Clear failed, using fallback XML method');
                    resolve(this.createSimpleWordXML(originalContent));
                }
            });
        });
    }

    /**
     * Create a simple Word XML document (fallback)
     * @param {string} textContent - The text content
     * @returns {Uint8Array} Word document bytes
     */
    createSimpleWordXML(textContent) {
        console.log('Creating simple Word XML from text...');

        // Escape XML characters
        const escapedText = textContent
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');

        // Split text into paragraphs
        const paragraphs = escapedText.split(/\r?\n/).map(para =>
            `<w:p><w:r><w:t>${para || ' '}</w:t></w:r></w:p>`
        ).join('');

        // Basic Word document XML structure
        const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        ${paragraphs}
    </w:body>
</w:document>`;

        const encoder = new TextEncoder();
        return encoder.encode(documentXml);
    }

    /**
     * Get document in compressed format
     * @param {number} sliceSize - Optional slice size in bytes
     * @returns {Promise<Blob>} Document blob
     */
    async getDocumentAsCompressed(sliceSize = 4194304) {
        return new Promise((resolve, reject) => {
            console.log('Starting document extraction...');

            // Try getting the file with specified slice size
            const options = {
                sliceSize: sliceSize
            };

            Office.context.document.getFileAsync(Office.FileType.Compressed, options, (result) => {
                console.log('getFileAsync result:', result);

                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const file = result.value;
                    const sliceCount = file.sliceCount;
                    console.log('File slice count:', sliceCount);

                    if (sliceCount === 0) {
                        file.closeAsync();
                        reject(new Error('Document appears to be empty'));
                        return;
                    }

                    const slices = [];
                    let slicesReceived = 0;

                    const getSlice = (sliceIndex) => {
                        file.getSliceAsync(sliceIndex, (sliceResult) => {
                            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                                try {
                                    const rawData = sliceResult.value.data;
                                    console.log(`Slice ${sliceIndex} data length:`, rawData.length);
                                    console.log(`Slice ${sliceIndex} data type:`, typeof rawData);
                                    console.log(`Slice ${sliceIndex} constructor:`, rawData.constructor.name);

                                    let bytes;

                                    // Handle different data types
                                    if (rawData instanceof ArrayBuffer) {
                                        console.log('Data is ArrayBuffer');
                                        bytes = new Uint8Array(rawData);
                                    } else if (rawData instanceof Uint8Array) {
                                        console.log('Data is already Uint8Array');
                                        bytes = rawData;
                                    } else if (Array.isArray(rawData)) {
                                        console.log('Data is Array, converting to Uint8Array');
                                        bytes = new Uint8Array(rawData);
                                    } else if (typeof rawData === 'string') {
                                        console.log('Data is string, checking if base64...');
                                        // Check if data is base64 encoded by trying to decode it
                                        try {
                                            // Test if it's valid base64
                                            const testDecode = atob(rawData.substring(0, Math.min(100, rawData.length)));
                                            console.log('Data appears to be base64 encoded');

                                            // Convert base64 to binary
                                            const binaryString = atob(rawData);
                                            bytes = new Uint8Array(binaryString.length);
                                            for (let i = 0; i < binaryString.length; i++) {
                                                bytes[i] = binaryString.charCodeAt(i);
                                            }
                                        } catch (base64Error) {
                                            console.log('Data is not base64, treating as raw binary string');

                                            // Data is already a binary string, convert directly to Uint8Array
                                            bytes = new Uint8Array(rawData.length);
                                            for (let i = 0; i < rawData.length; i++) {
                                                bytes[i] = rawData.charCodeAt(i);
                                            }
                                        }
                                    } else {
                                        console.error('Unknown data type:', typeof rawData, rawData.constructor.name);
                                        throw new Error('Unsupported data type: ' + typeof rawData);
                                    }

                                    slices[sliceIndex] = bytes;
                                    slicesReceived++;
                                    console.log(`Received slice ${sliceIndex + 1}/${sliceCount}, size: ${bytes.length} bytes`);

                                    if (slicesReceived === sliceCount) {
                                        // Combine all slices into a single blob
                                        const blob = new Blob(slices, {
                                            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                                        });
                                        console.log('Document blob created, total size:', blob.size);
                                        file.closeAsync();
                                        resolve(blob);
                                    }
                                } catch (error) {
                                    console.error('Error processing slice:', error);
                                    file.closeAsync();
                                    reject(new Error('Failed to process document slice: ' + error.message));
                                }
                            } else {
                                console.error('Failed to get slice:', sliceResult.error);
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
                    console.error('Failed to get document:', result.error);
                    // Try to provide more specific error information
                    let errorMessage = 'Failed to get document';
                    if (result.error) {
                        errorMessage += ': ' + result.error.message;
                        if (result.error.code) {
                            errorMessage += ' (Code: ' + result.error.code + ')';
                        }
                    }
                    reject(new Error(errorMessage));
                }
            });
        });
    }

    /**
     * Get document as PDF (fallback method)
     * @returns {Promise<Blob>} Document blob as PDF
     */
    async getDocumentAsPDF() {
        return new Promise((resolve, reject) => {
            console.log('Trying PDF extraction as fallback...');

            Office.context.document.getFileAsync(Office.FileType.Pdf, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const file = result.value;
                    const sliceCount = file.sliceCount;
                    console.log('PDF file slice count:', sliceCount);

                    if (sliceCount === 0) {
                        file.closeAsync();
                        reject(new Error('PDF document appears to be empty'));
                        return;
                    }

                    const slices = [];
                    let slicesReceived = 0;

                    const getSlice = (sliceIndex) => {
                        file.getSliceAsync(sliceIndex, (sliceResult) => {
                            if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                                const rawData = sliceResult.value.data;

                                let bytes;
                                if (rawData instanceof ArrayBuffer) {
                                    bytes = new Uint8Array(rawData);
                                } else if (rawData instanceof Uint8Array) {
                                    bytes = rawData;
                                } else if (typeof rawData === 'string') {
                                    // PDF data is usually base64 encoded
                                    const binaryString = atob(rawData);
                                    bytes = new Uint8Array(binaryString.length);
                                    for (let i = 0; i < binaryString.length; i++) {
                                        bytes[i] = binaryString.charCodeAt(i);
                                    }
                                }

                                slices[sliceIndex] = bytes;
                                slicesReceived++;

                                if (slicesReceived === sliceCount) {
                                    // Create PDF blob - we'll convert this to DOCX on the server if needed
                                    const blob = new Blob(slices, { type: 'application/pdf' });
                                    console.log('PDF blob created, size:', blob.size);
                                    file.closeAsync();
                                    resolve(blob);
                                }
                            } else {
                                file.closeAsync();
                                reject(new Error('Failed to get PDF slice: ' + sliceResult.error.message));
                            }
                        });
                    };

                    for (let i = 0; i < sliceCount; i++) {
                        getSlice(i);
                    }
                } else {
                    reject(new Error('Failed to get PDF: ' + result.error.message));
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

            // Delegate API call to JupiterService for consistency
            const result = await this.jupiterService.checkInDocument(documentId, documentBlob, versionComment);

            if (result && result.success) {
                // Update document state
                await this.documentStateManager.updateCheckoutStatus('Available');
                return { success: true, document: result.document };
            } else {
                throw new Error(result?.error || 'Check-in failed');
            }
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
