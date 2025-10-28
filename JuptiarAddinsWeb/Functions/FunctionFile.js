/**
 * Jupiter Add-in Function File - Main entry point for ribbon functions
 * Handles initialization and ribbon command functions
 */

// Global managers
let documentStateManager = null;
let ribbonManager = null;
let documentUploader = null;

// The initialize function must be run each time a new page is loaded.
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

        // Check document state and update ribbon accordingly
        if (documentStateManager && ribbonManager) {
            const isNewDocument = await documentStateManager.isNewDocument();
            console.log('Document is new:', isNewDocument);

            if (isNewDocument) {
                await ribbonManager.showNewDocumentRibbon();
                console.log('Applied new document ribbon state');
            } else {
                await ribbonManager.showExistingDocumentRibbon();
                console.log('Applied existing document ribbon state');
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
            await ribbonManager.updateCheckoutButtons();

            // Show success message
            console.log('Document checked out successfully');
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

        // Check in the document
        const response = await window.jupiterService.checkInDocument(documentState.documentId);

        if (response.success) {
            // Update document state
            await documentStateManager.updateCheckoutStatus('Available');

            // Update ribbon
            await ribbonManager.updateCheckoutButtons();

            // Show success message
            console.log('Document checked in successfully');
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
