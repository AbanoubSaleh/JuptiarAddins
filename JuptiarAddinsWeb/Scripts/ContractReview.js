/**
 * Contract Review Workflow - Handles the contract review workflow initiation
 */
class ContractReviewController {
    constructor() {
        this.isInitialized = false;
        this.isStartingWorkflow = false;
    }

    /**
     * Initialize the contract review controller
     */
    async initialize() {
        try {
            console.log('Initializing Contract Review Controller...');

            // Check document type first - workflow is only for new/external documents
            await this.checkDocumentEligibility();

            // Initialize UI components
            this.initializeUI();

            // Load initial data
            await this.loadInitialData();

            this.isInitialized = true;
            console.log('Contract Review Controller initialized successfully');
        } catch (error) {
            console.error('Failed to initialize Contract Review Controller:', error);
            this.showError('Failed to initialize contract review form: ' + error.message);
        }
    }

    /**
     * Check if document is eligible for workflow (must be a Jupiter document)
     */
    async checkDocumentEligibility() {
        try {
            // Use DocumentStateManager to check document type
            const dsm = new DocumentStateManager();
            await dsm.initialize();

            // Check if this is an existing Jupiter document
            const isExistingJupiterDoc = await dsm.isExistingDocument();

            if (isExistingJupiterDoc) {
                // This is a Jupiter document - workflow is available
                this.showWorkflowForm();
                return true;
            }

            // This is a new/external document - show message and hide form
            this.showNewDocumentMessage();
            this.hideWorkflowForm();
            return false;

        } catch (error) {
            console.warn('Could not determine document type:', error);
            // Default to hiding workflow if we can't determine
            this.showNewDocumentMessage();
            this.hideWorkflowForm();
            return false;
        }
    }

    /**
     * Show message for new/external documents (workflow not available)
     */
    showNewDocumentMessage() {
        const messageBar = document.querySelector('#messageBar');
        const messageIcon = document.querySelector('#messageIcon i');
        const messageText = document.querySelector('#messageText');
        const messageSection = document.querySelector('#messageSection');

        if (messageBar && messageText && messageSection) {
            messageBar.className = 'ms-MessageBar ms-MessageBar--warning';
            if (messageIcon) {
                messageIcon.className = 'ms-Icon ms-Icon--Info';
            }
            messageText.textContent = 'Workflow is for existing Jupiter documents only. Please open a document from Jupiter DMS first.';
            messageSection.style.display = 'block';
        }
    }

    /**
     * Hide workflow form (for Jupiter documents)
     */
    hideWorkflowForm() {
        const formSection = document.querySelector('#formSection');
        if (formSection) {
            formSection.style.display = 'none';
        }
    }

    /**
     * Show workflow form (for new/external documents)
     */
    showWorkflowForm() {
        const formSection = document.querySelector('#formSection');
        if (formSection) {
            formSection.style.display = 'block';
        }
    }

    /**
     * Initialize UI event handlers
     */
    initializeUI() {
        // Start Workflow button
        this.addEventHandler('#startWorkflowBtn', 'click', (e) => {
            if (e) { e.preventDefault(); e.stopPropagation(); }
            this.handleStartWorkflow();
        });

        // Cancel button
        this.addEventHandler('#cancelBtn', 'click', (e) => {
            if (e) { e.preventDefault(); e.stopPropagation(); }
            this.handleCancel();
        });

        // Form validation on blur
        this.addEventHandler('#contractRefNo', 'blur', () => this.validateContractRefNo());
        this.addEventHandler('#contractType', 'change', () => this.validateContractType());
        this.addEventHandler('#department', 'change', () => this.validateDepartment());
    }

    /**
     * Add event handler with error handling
     */
    addEventHandler(selector, event, handler) {
        try {
            const element = document.querySelector(selector);
            if (element) {
                element.addEventListener(event, handler);
            }
        } catch (error) {
            console.warn(`Failed to add event handler for ${selector}:`, error);
        }
    }

    /**
     * Load initial data and populate form
     */
    async loadInitialData() {
        try {
            console.log('=== CONTRACT REVIEW: Loading initial data ===');

            // Debug: Check what services are available
            console.log('Available services:');
            console.log('- window.jupiterService:', !!window.jupiterService);
            console.log('- window.authManager:', !!window.authManager);
            console.log('- DocumentStateManager:', typeof DocumentStateManager);

            // Auto-fill contract title from document title
            console.log('=== Populating contract title ===');
            await this.populateContractTitle();

            // Auto-fill contract owner from current user
            console.log('=== Populating contract owner ===');
            await this.populateContractOwner();

            console.log('=== Initial data loading completed ===');

        } catch (error) {
            console.error('Error loading initial data:', error);
            // Don't show error to user for auto-population failures
        }
    }

    /**
     * Populate contract title - EXACT copy of PropertiesEditor approach
     */
    async populateContractTitle() {
        try {
            console.log('Step 1: Loading current Word document...');
            // Step 1: Load current document (same as PropertiesEditor.loadCurrentWordDocument)
            await this.loadCurrentWordDocument();
            console.log('Current document loaded:', this.currentDocument);

            console.log('Step 2: Loading document metadata from Jupiter backend...');
            // Step 2: Load metadata from Jupiter backend (same as PropertiesEditor.loadDocumentMetadataFromWord)
            await this.loadDocumentMetadataFromWord();
            console.log('Document metadata loaded:', this.documentMetadata);

            // Step 3: Set the title field
            const title = this.documentMetadata?.title || 'Untitled Document';
            this.setFieldValue('#contractTitle', title);
            console.log('Contract title populated:', title);

        } catch (error) {
            console.error('Could not get document title:', error);
            this.setFieldValue('#contractTitle', 'Untitled Document');
        }
    }

    /**
     * Load current Word document - EXACT copy from PropertiesEditor
     */
    async loadCurrentWordDocument() {
        // Resolve Jupiter-managed document ID from DocumentStateManager first
        let jupiterId = null;
        try {
            const dsm = new DocumentStateManager();
            await dsm.initialize();
            const state = await dsm.getDocumentState();
            jupiterId = (state && state.documentId) ? state.documentId : (Office.context.document.settings.get('currentJuptiarDocumentId') || null);
        } catch (e) {
            try { jupiterId = Office.context.document.settings.get('currentJuptiarDocumentId') || null; } catch (_) { jupiterId = null; }
        }
        return new Promise((resolve, reject) => {
            Word.run(async (context) => {
                try {
                    const document = context.document;
                    const properties = document.properties;

                    properties.load(['title', 'author', 'subject', 'keywords', 'category', 'comments']);

                    await context.sync();

                    this.currentDocument = {
                        id: jupiterId,
                        title: properties.title,
                        author: properties.author,
                        subject: properties.subject,
                        keywords: properties.keywords,
                        category: properties.category,
                        comments: properties.comments,
                        isWordDocument: true
                    };

                    resolve();
                } catch (error) {
                    reject(error);
                }
            });
        });
    }

    /**
     * Load document metadata from Word properties - EXACT copy from PropertiesEditor
     */
    async loadDocumentMetadataFromWord() {
        // Default empty metadata
        this.documentMetadata = {
            title: '',
            description: '',
            tags: ''
        };

        // If document is from Jupiter, load metadata from backend (by Id)
        if (this.currentDocument && this.currentDocument.id) {
            try {
                const doc = await window.jupiterService.getDocument(this.currentDocument.id);
                this.currentDocument = doc;
                this.documentMetadata = {
                    title: doc.title || '',
                    description: doc.description || '',
                    tags: doc.tags || ''
                };
            } catch (error) {
                console.warn('Could not load Jupiter document by id:', error);
                // Fallback: use Word document title if available
                if (this.currentDocument && this.currentDocument.title) {
                    this.documentMetadata.title = this.currentDocument.title;
                }
            }
        } else {
            // Fallback: use Word document title if no Jupiter ID
            if (this.currentDocument && this.currentDocument.title) {
                this.documentMetadata.title = this.currentDocument.title;
            }
        }
    }



    /**
     * Populate contract owner - EXACT copy of Settings approach
     */
    async populateContractOwner() {
        try {
            let userName = 'Unknown User';

            // Get authentication status (same as Settings.updateConnectionStatus)
            if (window.authManager) {
                console.log('AuthManager available, getting auth status...');
                const authStatus = window.authManager.getAuthStatus(); // Fixed: removed await, correct method name
                console.log('Auth status:', authStatus);

                if (authStatus && authStatus.isAuthenticated && authStatus.user) {
                    userName = authStatus.user?.username || 'Unknown User';
                    console.log('User authenticated, username:', userName);
                } else {
                    console.log('User not authenticated or no user data');
                }
            } else {
                console.log('AuthManager not available');
            }

            this.setFieldValue('#contractOwner', userName);
            console.log('Contract owner populated:', userName);

        } catch (error) {
            console.error('Could not get current user:', error);
            this.setFieldValue('#contractOwner', 'Unknown User');
        }
    }

    /**
     * Set field value safely
     */
    setFieldValue(selector, value) {
        const field = document.querySelector(selector);
        if (field) {
            field.value = value;
        }
    }

    /**
     * Get field value safely
     */
    getFieldValue(selector) {
        const field = document.querySelector(selector);
        return field ? field.value.trim() : '';
    }

    /**
     * Validate contract reference number
     */
    validateContractRefNo() {
        const refNo = this.getFieldValue('#contractRefNo');
        const field = document.querySelector('#contractRefNo');
        
        if (!refNo) {
            this.showFieldError(field, 'Contract reference number is required');
            return false;
        }
        
        this.hideFieldError(field);
        return true;
    }

    /**
     * Validate contract type
     */
    validateContractType() {
        const contractType = this.getFieldValue('#contractType');
        const field = document.querySelector('#contractType');
        
        if (!contractType) {
            this.showFieldError(field, 'Contract type is required');
            return false;
        }
        
        this.hideFieldError(field);
        return true;
    }

    /**
     * Validate department
     */
    validateDepartment() {
        const department = this.getFieldValue('#department');
        const field = document.querySelector('#department');
        
        if (!department) {
            this.showFieldError(field, 'Department is required');
            return false;
        }
        
        this.hideFieldError(field);
        return true;
    }

    /**
     * Validate entire form
     */
    validateForm() {
        const validations = [
            this.validateContractRefNo(),
            this.validateContractType(),
            this.validateDepartment()
        ];

        return validations.every(isValid => isValid);
    }

    /**
     * Handle start workflow button click
     */
    async handleStartWorkflow() {
        if (this.isStartingWorkflow) {
            return; // Prevent double-clicks
        }

        try {
            // Double-check document eligibility before starting workflow
            const isEligible = await this.checkDocumentEligibility();
            if (!isEligible) {
                this.showError('Workflow is only available for existing Jupiter documents.');
                return;
            }

            // Validate form
            if (!this.validateForm()) {
                this.showError('Please fill in all required fields correctly.');
                return;
            }

            this.isStartingWorkflow = true;
            this.showLoading('Starting contract review workflow...');

            // Collect form data
            const workflowData = this.collectFormData();

            // Simulate workflow start (static data for prototype)
            await this.startContractReviewWorkflow(workflowData);

            this.showSuccess('Contract review workflow started successfully!');

            // Optionally close the dialog after a delay
            setTimeout(() => {
                this.handleCancel();
            }, 2000);

        } catch (error) {
            console.error('Error starting workflow:', error);
            this.showError('Failed to start workflow: ' + error.message);
        } finally {
            this.isStartingWorkflow = false;
            this.hideLoading();
        }
    }

    /**
     * Collect form data
     */
    collectFormData() {
        return {
            contractTitle: this.getFieldValue('#contractTitle'),
            contractRefNo: this.getFieldValue('#contractRefNo'),
            contractType: this.getFieldValue('#contractType'),
            vendor: this.getFieldValue('#vendor'),
            department: this.getFieldValue('#department'),
            contractOwner: this.getFieldValue('#contractOwner'),
            timestamp: new Date().toISOString()
        };
    }

    /**
     * Start contract review workflow (static implementation for prototype)
     */
    async startContractReviewWorkflow(workflowData) {
        // Simulate API call delay
        await new Promise(resolve => setTimeout(resolve, 1500));

        // Log the workflow data (for prototype demonstration)
        console.log('Contract Review Workflow Started:', workflowData);

        // In the future, this will make an actual API call to Jupiter
        // const response = await fetch('/api/workflows/contract-review', {
        //     method: 'POST',
        //     headers: { 'Content-Type': 'application/json' },
        //     body: JSON.stringify(workflowData)
        // });

        return {
            success: true,
            workflowId: 'WF-' + Date.now(),
            message: 'Contract review workflow initiated successfully'
        };
    }

    /**
     * Handle cancel button click
     */
    handleCancel() {
        try {
            // Close the task pane
            if (Office.context.ui) {
                Office.context.ui.closeContainer();
            } else {
                // Fallback for older Office versions
                window.close();
            }
        } catch (error) {
            console.warn('Could not close dialog:', error);
            // Try alternative method
            try {
                window.close();
            } catch (e) {
                console.warn('Could not close window:', e);
            }
        }
    }

    /**
     * Show field error
     */
    showFieldError(field, message) {
        if (!field) return;

        // Remove existing error
        this.hideFieldError(field);

        // Add error styling
        field.style.borderColor = '#d13438';

        // Create error message element
        const errorElement = document.createElement('div');
        errorElement.className = 'field-error';
        errorElement.style.color = '#d13438';
        errorElement.style.fontSize = '12px';
        errorElement.style.marginTop = '5px';
        errorElement.textContent = message;

        // Insert after field
        field.parentNode.insertBefore(errorElement, field.nextSibling);
    }

    /**
     * Hide field error
     */
    hideFieldError(field) {
        if (!field) return;

        // Reset field styling
        field.style.borderColor = '#8a8886';

        // Remove error message
        const errorElement = field.parentNode.querySelector('.field-error');
        if (errorElement) {
            errorElement.remove();
        }
    }

    /**
     * Show loading state
     */
    showLoading(message = 'Loading...') {
        const loadingSection = document.querySelector('#loadingSection');
        const loadingText = document.querySelector('#loadingText');
        const startBtn = document.querySelector('#startWorkflowBtn');

        if (loadingText) {
            loadingText.textContent = message;
        }

        if (loadingSection) {
            loadingSection.style.display = 'block';
        }

        if (startBtn) {
            startBtn.disabled = true;
        }

        this.hideMessage();
    }

    /**
     * Hide loading state
     */
    hideLoading() {
        const loadingSection = document.querySelector('#loadingSection');
        const startBtn = document.querySelector('#startWorkflowBtn');

        if (loadingSection) {
            loadingSection.style.display = 'none';
        }

        if (startBtn) {
            startBtn.disabled = false;
        }
    }

    /**
     * Show success message
     */
    showSuccess(message) {
        this.showMessage(message, 'success');
    }

    /**
     * Show error message
     */
    showError(message) {
        this.showMessage(message, 'error');
    }

    /**
     * Show message
     */
    showMessage(message, type = 'info') {
        const messageSection = document.querySelector('#messageSection');
        const messageBar = document.querySelector('#messageBar');
        const messageText = document.querySelector('#messageText');
        const messageIcon = document.querySelector('#messageIcon i');

        if (!messageSection || !messageBar || !messageText) {
            return;
        }

        // Set message text
        messageText.textContent = message;

        // Reset classes
        messageBar.className = 'ms-MessageBar';

        // Add type-specific styling
        if (type === 'success') {
            messageBar.classList.add('ms-MessageBar--success');
            if (messageIcon) {
                messageIcon.className = 'ms-Icon ms-Icon--CheckMark';
            }
        } else if (type === 'error') {
            messageBar.classList.add('ms-MessageBar--error');
            if (messageIcon) {
                messageIcon.className = 'ms-Icon ms-Icon--ErrorBadge';
            }
        } else {
            if (messageIcon) {
                messageIcon.className = 'ms-Icon ms-Icon--Info';
            }
        }

        // Show message
        messageSection.style.display = 'block';

        // Auto-hide success messages after 5 seconds
        if (type === 'success') {
            setTimeout(() => {
                this.hideMessage();
            }, 5000);
        }
    }

    /**
     * Hide message
     */
    hideMessage() {
        const messageSection = document.querySelector('#messageSection');
        if (messageSection) {
            messageSection.style.display = 'none';
        }
    }
}

// Global instance
let contractReviewController;

// Initialize when Office is ready
Office.onReady(() => {
    console.log('Office is ready. Initializing Contract Review...');

    // Initialize JupiterConfig first (this might not have been called)
    if (typeof window.JupiterConfig.init === 'function') {
        window.JupiterConfig.init();
    }

    // Initialize global services if not already done
    if (!window.jupiterService) {
        window.jupiterService = new JupiterService();
        window.jupiterService.initialize({
            serverUrl: window.JupiterConfig.get('server.baseUrl'),
            apiEndpoint: window.JupiterConfig.get('server.apiEndpoint') || '/api',
            timeout: window.JupiterConfig.get('server.timeout') || 30000
        });
    }

    if (!window.authManager) {
        window.authManager = new AuthManager();
        // Initialize AuthManager asynchronously and then initialize contract review controller
        window.authManager.initialize().then(() => {
            console.log('AuthManager initialized, initializing contract review controller...');
            contractReviewController = new ContractReviewController();
            contractReviewController.initialize().catch(error => {
                console.error('Failed to initialize Contract Review Controller:', error);
            });
        }).catch(error => {
            console.error('Failed to initialize AuthManager:', error);
            // Still try to initialize the controller even if AuthManager fails
            contractReviewController = new ContractReviewController();
            contractReviewController.initialize().catch(error => {
                console.error('Failed to initialize Contract Review Controller:', error);
            });
        });
    } else {
        // AuthManager already exists, initialize controller directly
        contractReviewController = new ContractReviewController();
        contractReviewController.initialize().catch(error => {
            console.error('Failed to initialize Contract Review Controller:', error);
        });
    }
});
