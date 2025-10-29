/**
 * Jupiter Add-in Function File - Main entry point for ribbon functions
 * Handles initialization and ribbon command functions
 */

// Global managers
let documentStateManager = null;
let ribbonManager = null;
let documentUploader = null;
let documentEditMonitor = null;
let documentTracker = null;

// The initialize function must be run each time a new page is loaded.
if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(async () => {
        try {
            console.log('Jupiter Add-in initializing...');
            console.log('Abanoub --Ribbon intiated');

            // Initialize global services first
            await initializeGlobalServices();

            // Initialize managers
            await initializeManagers();

            console.log('Jupiter Add-in initialized successfully');
        } catch (error) {
            console.error('Failed to initialize Jupiter Add-in:', error);
        }
    });
} else {
    console.warn('FunctionFile.js: Office.js not available, skipping Office.onReady initialization');
}

/**
 * Initialize global services
 */
async function initializeGlobalServices() {
    // Initialize JupiterService
    if (!window.jupiterService && typeof JupiterService !== 'undefined') {
        window.jupiterService = new JupiterService();
        window.jupiterService.initialize({
            serverUrl: 'https://localhost:7001',
            apiEndpoint: '/api',
            timeout: 30000
        });
    }

    // Initialize AuthManager if available
    if (!window.authManager && typeof AuthManager !== 'undefined') {
        window.authManager = new AuthManager();
        await window.authManager.initialize();
    }
}

/**
 * Initialize document managers
 */
async function initializeManagers() {
    try {
        // Initialize DocumentStateManager
        if (typeof DocumentStateManager !== 'undefined') {
            documentStateManager = new DocumentStateManager();
            await documentStateManager.initialize();
        }

        // Initialize RibbonManager
        if (typeof RibbonManager !== 'undefined') {
            ribbonManager = new RibbonManager();
            await ribbonManager.initialize(documentStateManager);
        }

        // Initialize DocumentUploader
        if (typeof DocumentUploader !== 'undefined') {
            documentUploader = new DocumentUploader();
            await documentUploader.initialize(window.jupiterService, documentStateManager);
        }

        // Initialize DocumentEditMonitor
        if (typeof DocumentEditMonitor !== 'undefined' && documentStateManager && window.jupiterService && ribbonManager) {
            documentEditMonitor = new DocumentEditMonitor(documentStateManager, window.jupiterService, ribbonManager);
            console.log('✅ DocumentEditMonitor initialized');
        }

        // Initialize DocumentTracker for comprehensive document detection and validation
        if (typeof DocumentTracker !== 'undefined') {
            documentTracker = new DocumentTracker();
            console.log('✅ DocumentTracker initialized');

            // Start document tracking - this will handle all detection and validation
            await documentTracker.initializeDocumentTracking();
        } else {
            // Fallback to legacy document detection if DocumentTracker is not available
            console.warn('DocumentTracker not available, using legacy detection');

            if (documentStateManager && ribbonManager) {
                console.log('🔍 Starting legacy document type detection...');

                const documentType = await documentStateManager.detectDocumentType();
                console.log('📄 Detected document type:', documentType);

                switch (documentType) {
                    case 'new':
                        await ribbonManager.showNewDocumentRibbon();
                        console.log('✅ Applied NEW document ribbon state');
                        break;

                    case 'jupiter':
                        await ribbonManager.showExistingDocumentRibbon();
                        console.log('✅ Applied JUPITER document ribbon state');

                        // Start edit monitoring for Jupiter documents
                        if (documentEditMonitor) {
                            await documentEditMonitor.startMonitoring();
                        }
                        break;

                    case 'external':
                        await ribbonManager.showExternalDocumentRibbon();
                        console.log('✅ Applied EXTERNAL document ribbon state');
                        break;

                    default:
                        // Fallback to new document state
                        await ribbonManager.showNewDocumentRibbon();
                        console.log('⚠️ Unknown document type, defaulting to NEW document ribbon state');
                        break;
                }
            }
        }

        console.log('Document managers initialized successfully');
    } catch (error) {
        console.error('Failed to initialize document managers:', error);
    }
}

/**
 * Check out document function (called from ribbon)
 */
async function checkOutDocument(event) {
    try {
        console.log('Check out document function called');

        // Check if button should be visible/enabled (Office 2019 compatibility)
        if (window.ribbonManager) {
            const isVisible = await window.ribbonManager.isButtonVisible('Jupiter.CheckOutButton');
            if (!isVisible) {
                console.log('❌ Check Out button is disabled for this document type');
                Office.context.ui.displayDialogAsync(
                    'Check Out is not available for this document type. This action is only available for Jupiter-managed documents.',
                    { height: 30, width: 50 }
                );
                event.completed();
                return;
            }
        }

        if (!documentStateManager) {
            throw new Error('Document state manager not initialized');
        }

        // Check if this is a new document
        const isNewDocument = await documentStateManager.isNewDocument();
        if (isNewDocument) {
            Office.context.ui.displayDialogAsync(
                'Check Out is not available for new documents. Please save the document to Jupiter DMS first.',
                { height: 30, width: 50 }
            );
            event.completed();
            return;
        }

        // Get document state
        const documentState = await documentStateManager.getDocumentState();

        if (!documentState || !documentState.documentId) {
            throw new Error('No document information found. This document may not be managed by Jupiter DMS.');
        }

        // Check out the document
        const response = await window.jupiterService.checkOutDocument(documentState.documentId);

        if (response.success) {
            // Update document state
            await documentStateManager.updateCheckoutStatus('CheckedOut');

            // Update ribbon
            if (ribbonManager) {
                await ribbonManager.updateCheckoutButtons();
            }

            // Start edit monitoring if available
            if (documentEditMonitor) {
                await documentEditMonitor.startMonitoring();
            }

            // Show success message
            Office.context.ui.displayDialogAsync(
                'Document checked out successfully! You can now edit the document.',
                { height: 30, width: 50 }
            );

            console.log('✅ Document checked out successfully');
        } else {
            throw new Error(response.error || 'Failed to check out document');
        }

    } catch (error) {
        console.error('Error checking out document:', error);

        // Show error to user
        Office.context.ui.displayDialogAsync(
            `Error: ${error.message}`,
            { height: 30, width: 50 }
        );
    }

    // Required for ribbon functions
    event.completed();
}

/**
 * Check in document function (called from ribbon)
 */
async function checkInDocument(event) {
    try {
        console.log('Check in document function called');

        if (!documentStateManager) {
            throw new Error('Document state manager not initialized');
        }

        // Check if this is a new document
        const isNewDocument = await documentStateManager.isNewDocument();
        if (isNewDocument) {
            Office.context.ui.displayDialogAsync(
                'Check In is not available for new documents. Please save the document to Jupiter DMS first.',
                { height: 30, width: 50 }
            );
            event.completed();
            return;
        }

        // Get document state
        const documentState = await documentStateManager.getDocumentState();
        if (!documentState || !documentState.documentId) {
            throw new Error('No document information found. This document may not be managed by Jupiter DMS.');
        }

        // Show check-in dialog to get version comment
        const checkInResult = await showCheckInDialog();
        if (!checkInResult || checkInResult.cancelled) {
            console.log('Check-in cancelled by user');
            event.completed();
            return;
        }

        // Get current document content as blob
        const documentBlob = await getCurrentDocumentAsBlob();

        // Check in the document with content and version comment
        const response = await window.jupiterService.checkInDocument(
            documentState.documentId,
            documentBlob,
            checkInResult.versionComment
        );

        if (response.success) {
            // Update document state
            await documentStateManager.updateCheckoutStatus('Available');

            // Update document version info
            if (response.document) {
                await documentStateManager.setDocumentState({
                    ...documentState,
                    version: response.document.currentVersion,
                    checkoutStatus: 'Available',
                    lastSaved: new Date().toISOString()
                });
            }

            // Update ribbon
            if (ribbonManager) {
                await ribbonManager.updateCheckoutButtons();
            }

            // Stop edit monitoring
            if (documentEditMonitor) {
                await documentEditMonitor.stopMonitoring();
            }

            // Show success message
            Office.context.ui.displayDialogAsync(
                `Document checked in successfully! New version: ${response.document?.currentVersion || 'Unknown'}`,
                { height: 30, width: 50 }
            );

            console.log('✅ Document checked in successfully');
        } else {
            throw new Error(response.error || 'Failed to check in document');
        }

    } catch (error) {
        console.error('Error checking in document:', error);

        // Show error to user
        Office.context.ui.displayDialogAsync(
            `Error: ${error.message}`,
            { height: 30, width: 50 }
        );
    }

    // Required for ribbon functions
    event.completed();
}

/**
 * Handle Save to Jupiter DMS button click
 * This function checks if the document is new and opens the save dialog
 */
async function saveToJupiterDMS(event) {
    try {
        console.log('Save to Jupiter DMS button clicked');

        if (!documentStateManager) {
            throw new Error('Document state manager not initialized');
        }

        // Check if this is a new document
        const isNewDocument = await documentStateManager.isNewDocument();

        if (!isNewDocument) {
            // Show message that this is for new documents only
            Office.context.ui.displayDialogAsync(
                'This function is for new documents only. Use Properties to edit existing document metadata.',
                { height: 30, width: 50 }
            );
            event.completed();
            return;
        }

        // For new documents, open the Save Dialog taskpane
        console.log('Opening Save Dialog for new document...');

        // Open the Save Dialog using Office.addin.showAsTaskpane
        try {
            // Try to open the SaveDialog as a taskpane
            await Office.addin.showAsTaskpane('SaveDialog.html');
        } catch (taskpaneError) {
            console.warn('Could not open taskpane directly, trying dialog approach:', taskpaneError);

            // Fallback: Open as dialog
            const dialogUrl = Office.context.requirements.isSetSupported('DialogApi', '1.1')
                ? `${window.location.origin}/SaveDialog.html`
                : null;

            if (dialogUrl) {
                Office.context.ui.displayDialogAsync(
                    dialogUrl,
                    { height: 80, width: 60, displayInIframe: true },
                    (result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            console.log('Save Dialog opened successfully');
                        } else {
                            console.error('Failed to open Save Dialog:', result.error);
                        }
                    }
                );
            } else {
                // Final fallback: Show instruction message
                Office.context.ui.displayDialogAsync(
                    'Please use the ribbon to access the Save to Jupiter functionality.',
                    { height: 30, width: 50 }
                );
            }
        }

    } catch (error) {
        console.error('Error in Save to Jupiter DMS:', error);

        // Show error to user
        Office.context.ui.displayDialogAsync(
            `Error: ${error.message}`,
            { height: 30, width: 50 }
        );
    }

    // Required for ribbon functions
    event.completed();
}

/**
 * Handle Properties button click
 * This function checks if the document exists in Jupiter DMS
 */
async function editProperties(event) {
    try {
        console.log('Properties button clicked');

        if (!documentStateManager) {
            throw new Error('Document state manager not initialized');
        }

        // Check if this is an existing document
        const isNewDocument = await documentStateManager.isNewDocument();

        if (isNewDocument) {
            // Show message that this is for existing documents only
            Office.context.ui.displayDialogAsync(
                'This function is for existing documents only. Use "Save to Jupiter DMS" to save new documents first.',
                { height: 30, width: 50 }
            );
            event.completed();
            return;
        }

        // For existing documents, the Properties button should open the Properties taskpane
        // This is handled by the manifest configuration (ShowTaskpane action)
        console.log('Opening Properties Dialog for existing document...');

    } catch (error) {
        console.error('Error in Properties:', error);

        // Show error to user
        Office.context.ui.displayDialogAsync(
            `Error: ${error.message}`,
            { height: 30, width: 50 }
        );
    }

    // Required for ribbon functions
    event.completed();
}

/**
 * Sample function for testing
 */
async function sampleFunction(event) {
    try {
        await Word.run(async (context) => {
            // Insert a paragraph at the end of the document body.
            const body = context.document.body;
            body.insertParagraph("Jupiter Add-in is working! Document state management and ribbon integration are active.", Word.InsertLocation.end);

            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}

/**
 * Get current document content as blob for upload
 */
async function getCurrentDocumentAsBlob() {
    try {
        return await Word.run(async (context) => {
            // Get the document as a base64 string
            const documentBody = context.document.body;
            documentBody.load('*');
            await context.sync();

            // Use Office.js to get the document as a file
            return new Promise((resolve, reject) => {
                Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const file = result.value;
                        const sliceCount = file.sliceCount;
                        let docData = [];

                        // Read all slices
                        const getSlice = (sliceIndex) => {
                            file.getSliceAsync(sliceIndex, (sliceResult) => {
                                if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                                    docData.push(sliceResult.value.data);

                                    if (sliceIndex < sliceCount - 1) {
                                        getSlice(sliceIndex + 1);
                                    } else {
                                        // All slices read, create blob
                                        file.closeAsync();
                                        const blob = new Blob(docData, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                                        resolve(blob);
                                    }
                                } else {
                                    file.closeAsync();
                                    reject(new Error('Failed to read document slice'));
                                }
                            });
                        };

                        getSlice(0);
                    } else {
                        reject(new Error('Failed to get document file'));
                    }
                });
            });
        });
    } catch (error) {
        console.error('Error getting document as blob:', error);
        throw error;
    }
}

/**
 * Show check-in dialog to get version comment
 */
async function showCheckInDialog() {
    try {
        return new Promise((resolve) => {
            const baseUrl = window.location.origin + window.location.pathname.replace(/\/[^\/]*$/, '');
            const dialogUrl = `${baseUrl}/CheckInDialog.html`;

            Office.context.ui.displayDialogAsync(
                dialogUrl,
                { height: 60, width: 80 },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        const dialog = asyncResult.value;

                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                            dialog.close();
                            try {
                                const result = JSON.parse(arg.message);
                                resolve(result);
                            } catch (error) {
                                console.error('Error parsing dialog result:', error);
                                resolve({ cancelled: true });
                            }
                        });

                        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                            dialog.close();
                            resolve({ cancelled: true });
                        });
                    } else {
                        console.error('Failed to open check-in dialog:', asyncResult.error);
                        resolve({ cancelled: true });
                    }
                }
            );
        });
    } catch (error) {
        console.error('Error showing check-in dialog:', error);
        return { cancelled: true };
    }
}

/**
 * Show notification to user (Office 2019 compatible)
 */
function showNotification(title, message) {
    try {
        if (Office.context.ui && Office.context.ui.displayDialogAsync) {
            Office.context.ui.displayDialogAsync(
                `${title}: ${message}`,
                { height: 30, width: 50 }
            );
        } else {
            console.log(`${title}: ${message}`);
            alert(`${title}: ${message}`);
        }
    } catch (error) {
        console.error('Error showing notification:', error);
        alert(`${title}: ${message}`);
    }
}
