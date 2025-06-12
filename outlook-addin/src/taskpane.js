/* Opsie Email Assistant - Outlook Addin */

// Visible error reporting for debugging
let initErrors = [];
let loadingSteps = [];

// Create a local log that persists even if console isn't available
let localLog = [];

// Global variable to store the current email data
let currentEmailData = null;

// Define global variable for current email notes
let currentEmailNotes = [];

// Global variables for notes polling
let notesPollingInterval = null;
let lastNotesCheck = null;
let isNotesPollingActive = false;

// Function to log messages to both console and local log
function log(message, type = 'info', data = null) {
    const entry = {
        time: new Date().toISOString(),
        type: type,
        message: message,
        data: data
    };
    
    // Add to local log
    localLog.push(entry);
    
    // Try to log to console
    try {
        if (type === 'error') {
            console.error(`[${entry.time}] ${message}`, data || '');
        } else if (type === 'warn') {
            console.warn(`[${entry.time}] ${message}`, data || '');
        } else {
            console.log(`[${entry.time}] ${message}`, data || '');
        }
    } catch (e) {
        // Console not available
    }
    
    // If debug elements exist, update them
    try {
        updateDebugDisplay();
    } catch (e) {
        // Debug display not available yet
    }
}

// Function to log loading steps visibly
function logStep(step) {
    log(step, 'step');
    loadingSteps.push(`${new Date().toISOString().substr(11, 8)} - ${step}`);
    updateLoadingSteps();
}

// Update loading steps display
function updateLoadingSteps() {
    if (document.getElementById('loading-steps')) {
        document.getElementById('loading-steps').innerHTML = loadingSteps.join('<br>');
    }
}

// Update the full debug display with local log
function updateDebugDisplay() {
    const logContainer = document.getElementById('full-log');
    if (logContainer) {
        const logHtml = localLog.map(entry => {
            const typeClass = entry.type === 'error' ? 'error-log' : 
                              entry.type === 'warn' ? 'warn-log' : 'info-log';
            return `<div class="${typeClass}">
                [${entry.time.substr(11, 8)}] ${entry.message}
                ${entry.data ? `<pre>${JSON.stringify(entry.data, null, 2)}</pre>` : ''}
            </div>`;
        }).join('');
        logContainer.innerHTML = logHtml;
    }
}

// Function to display errors visibly
function displayError(error, context = 'general') {
    const errorMsg = {
        context: context,
        message: error.message || String(error),
        stack: error.stack,
        time: new Date().toISOString()
    };
    
    log(`Error in ${context}: ${errorMsg.message}`, 'error', errorMsg);
    initErrors.push(errorMsg);
    
    // If the DOM is already loaded, update the error display
    if (document.getElementById('error-display')) {
        const errorDisplay = document.getElementById('error-display');
        errorDisplay.style.display = 'block';
        
        const errorHtml = initErrors.map(err => 
            `<div class="error-item">
                <strong>${err.context}:</strong> ${err.message}
                <pre class="error-stack">${err.stack || 'No stack trace available'}</pre>
            </div>`
        ).join('<hr>');
        
        errorDisplay.innerHTML = errorHtml;
    }
}

// Ensure the debug sections are visible on initialization
function showDebugSections() {
    try {
        if (document.getElementById('debug-info')) {
            document.getElementById('debug-info').style.display = 'block';
        }
        if (document.getElementById('loading-steps')) {
            document.getElementById('loading-steps').style.display = 'block';
        }
        if (document.getElementById('toggle-debug')) {
            document.getElementById('toggle-debug').textContent = 'Hide Debug Info';
        }
    } catch (e) {
        // Elements might not be available yet
    }
}

// Call this early and repeatedly until it succeeds
setInterval(showDebugSections, 500);

// Cleanup function to stop polling when page unloads
window.addEventListener('beforeunload', function() {
    stopNotesPolling();
});

// Wrap Office.onReady in try-catch
Office.onReady((info) => {
    try {
        logStep('Office.onReady called');
        
        // Display Office diagnostic info
        if (info) {
            logStep(`Host: ${info.host}, Platform: ${info.platform}`);
        }
        
        // Add version info if available
        try {
            if (Office.context && Office.context.mailbox && Office.context.mailbox.diagnostics) {
                const diagnostics = Office.context.mailbox.diagnostics;
                logStep(`Outlook Version: ${diagnostics.hostVersion}, Host: ${diagnostics.hostName}`);
            }
        } catch (versionError) {
            displayError(versionError, 'version-check');
        }
        
        // Initialize UI
        $(document).ready(() => {
            try {
                logStep('Document ready event fired');
                
                // Update debug info in the UI
                if (document.getElementById('debug-info')) {
                    const debugInfo = document.getElementById('debug-info');
                    debugInfo.innerHTML = `
                        <strong>Add-in Debug Info:</strong><br>
                        Time: ${new Date().toLocaleString()}<br>
                        Office Host: ${info.host || 'Unknown'}<br>
                        Platform: ${info.platform || 'Unknown'}<br>
                        ${Office.context?.mailbox?.diagnostics ? 
                          `Outlook Version: ${Office.context.mailbox.diagnostics.hostVersion}<br>
                           Host Name: ${Office.context.mailbox.diagnostics.hostName}` : 
                          'Mailbox diagnostics not available'}
                    `;
                }
                
                // Then continue with your regular initialization
                logStep('Starting main initialization');
                
                // Initialize the UI and set up event listeners
                setupEventListeners();
                
                // Check if user is authenticated
                checkAuthStatus().then(isAuthenticated => {
                    if (isAuthenticated) {
                        logStep('User is authenticated');
                        
                        // Only call loadCurrentEmail if Office.context.mailbox.item is available
                        if (Office.context.mailbox && Office.context.mailbox.item) {
                            loadCurrentEmail();
                        } else {
                            log('No email item available yet', 'warning');
                        }
                    } else {
                        logStep('User is not authenticated');
                        
                        // Add null checks for DOM elements
                        const authContainer = document.getElementById('auth-container');
                        if (authContainer) {
                            authContainer.style.display = 'block';
                        } else {
                            log('auth-container element not found', 'warning');
                        }
                        
                        const mainContent = document.getElementById('main-content');
                        if (mainContent) {
                            mainContent.style.display = 'none';
                        } else {
                            log('main-content element not found', 'warning');
                        }
                    }
                }).catch(authError => {
                    displayError(authError, 'auth-check');
                });
                
                // Start heartbeat to check authentication status periodically
                startAuthHeartbeat();
                
                logStep('Initialization complete');
            } catch (docReadyError) {
                displayError(docReadyError, 'document-ready');
            }
        });
    } catch (officeReadyError) {
        displayError(officeReadyError, 'office-ready');
    }

    // Initialize document upload functionality
    initializeDocumentUpload();
});

// Catch unhandled rejections
window.addEventListener('unhandledrejection', event => {
    displayError(event.reason, 'unhandled-promise-rejection');
});

// Catch global errors
window.addEventListener('error', event => {
    displayError({
        message: event.message,
        stack: `at ${event.filename}:${event.lineno}:${event.colno}`
    }, 'global-error');
});

// Check if the user is authenticated
async function checkAuthStatus() {
    try {
        window.OpsieApi.log('Checking authentication status', 'info', { context: 'auth-check' });
        
        // Get DOM elements with null checks
        const authErrorContainer = document.getElementById('auth-error-container');
        const authContainer = document.getElementById('auth-container');
        const mainContent = document.getElementById('main-content');
        
        // Store current auth state before checking
        const wasAuthenticated = authErrorContainer && authErrorContainer.style.display === 'none';
        
        // Check authentication status using the API
        const isAuthenticated = await window.OpsieApi.isAuthenticated();
        window.OpsieApi.log(`Authentication status: ${isAuthenticated ? 'Authenticated' : 'Not authenticated'}`, 'info', { context: 'auth-check' });
        
        if (isAuthenticated) {
            // User is authenticated, hide auth containers but DON'T show main content yet
            // Main content will be shown only after team validation in showMainUISections()
            if (authErrorContainer) {
                authErrorContainer.style.display = 'none';
            } else {
                window.OpsieApi.log('auth-error-container element not found', 'warning', { context: 'auth-check' });
            }
            
            if (authContainer) {
                authContainer.style.display = 'none';
            } else {
                window.OpsieApi.log('auth-container element not found', 'warning', { context: 'auth-check' });
            }
            
            // DON'T show main-content yet - wait for team validation
            if (mainContent) {
                mainContent.style.display = 'none'; // Keep it hidden until team validation
            } else {
                window.OpsieApi.log('main-content element not found', 'warning', { context: 'auth-check' });
            }
            
            // If user went from unauthenticated to authenticated
            if (!wasAuthenticated) {
                window.OpsieApi.log('Auth state changed: User is now authenticated', 'info', { context: 'auth-check' });
                
                // Initialize team ID and user information with a callback that will reload email data
                window.OpsieApi.initTeamAndUserInfo(function(teamInfo) {
                    window.OpsieApi.log('Team info initialized within auth status callback', 'info', {
                        teamId: teamInfo.teamId || 'Not available',
                        userId: teamInfo.userId || 'Not available',
                        source: teamInfo.fromCache ? 'cache' : (teamInfo.fromApi ? 'api' : 'token'),
                        context: 'auth-check-callback'
                    });
                    
                    // Check if team ID is available
                    if (!teamInfo.teamId) {
                        window.OpsieApi.log('Team ID not available after team init callback', 'warning', {
                            context: 'auth-check-callback'
                        });
                        window.OpsieApi.showNotification('Warning: Could not initialize team information. Some features may be limited.', 'warning');
                        return;
                    }
                    
                    // Reload current email data if there's an email available
                    if (Office.context.mailbox && Office.context.mailbox.item) {
                        window.OpsieApi.log('Reloading current email data after authentication with team ID: ' + teamInfo.teamId, 'info', {
                            context: 'auth-check-callback',
                            teamId: teamInfo.teamId
                        });
                        window.OpsieApi.showNotification('You are now logged in. Refreshing email data...', 'info');
                        
                        setTimeout(async () => {
                            try {
                                await loadCurrentEmail();
                                window.OpsieApi.showNotification('Email data refreshed successfully!', 'success');
                            } catch (refreshError) {
                                window.OpsieApi.log('Error refreshing email data after auth change', 'error', {
                                    error: refreshError,
                                    context: 'auth-refresh'
                                });
                            }
                        }, 1000); // Short delay to allow UI updates to settle
                    }
                });
            } else {
                // User was already authenticated, just ensure team info is initialized
                try {
                    const teamInitResult = await window.OpsieApi.initTeamAndUserInfo();
                    
                    // New format returns an object
                    if (teamInitResult && typeof teamInitResult === 'object') {
                        window.OpsieApi.log('Team and user initialization result:', 'info', {
                            teamId: teamInitResult.teamId || 'Not available',
                            userId: teamInitResult.userId || 'Not available',
                            source: teamInitResult.fromCache ? 'cache' : (teamInitResult.fromApi ? 'api' : 'token'),
                            context: 'auth-check'
                        });
                        
                        // Store the userId in localStorage for later use if present
                        if (teamInitResult.userId) {
                            localStorage.setItem('userId', teamInitResult.userId);
                            window.OpsieApi.log('Stored userId in localStorage', 'info', {
                                userId: teamInitResult.userId,
                                context: 'auth-check'
                            });
                        }
                        
                        if (!teamInitResult.teamId) {
                            window.OpsieApi.showNotification('Warning: Could not initialize team information. Some features may be limited.', 'warning');
                        } else {
                            // User has team, ensure UI sections are visible
                            window.OpsieApi.log('User has team during auth check, ensuring UI sections are visible', 'info', {
                                teamId: teamInitResult.teamId,
                                context: 'auth-check'
                            });
                            showMainUISections();
                        }
                    } else {
                        // Old format returned a boolean
                        window.OpsieApi.log('Team ID initialization result (legacy format):', 'info', {
                            success: teamInitResult,
                            teamId: teamInitResult ? localStorage.getItem('currentTeamId') : 'Not initialized',
                            context: 'auth-check'
                        });
                        
                        // Try to get userId from localStorage since the old format might have set it
                        const userId = localStorage.getItem('userId');
                        if (userId) {
                            window.OpsieApi.log('Found userId in localStorage after initialization', 'info', {
                                userId: userId,
                                context: 'auth-check'
                            });
                        }
                        
                        if (!teamInitResult) {
                            window.OpsieApi.showNotification('Warning: Could not initialize team information. Some features may be limited.', 'warning');
                        } else {
                            // User has team, ensure UI sections are visible
                            window.OpsieApi.log('User has team during auth check (legacy), ensuring UI sections are visible', 'info', {
                                context: 'auth-check'
                            });
                            showMainUISections();
                        }
                    }
                } catch (teamInitError) {
                    window.OpsieApi.log('Error initializing team information:', 'error', {
                        error: teamInitError,
                        context: 'auth-check'
                    });
                    window.OpsieApi.showNotification('Warning: Could not initialize team information. Some features may be limited.', 'warning');
                }
            }
        } else {
            // User is not authenticated, show auth container
            if (authErrorContainer) {
                authErrorContainer.style.display = 'flex';
            } else {
                window.OpsieApi.log('auth-error-container element not found', 'warning', { context: 'auth-check' });
            }
            
            if (authContainer) {
                authContainer.style.display = 'flex';
            } else {
                window.OpsieApi.log('auth-container element not found', 'warning', { context: 'auth-check' });
            }
            
            if (mainContent) {
                mainContent.style.display = 'none';
            } else {
                window.OpsieApi.log('main-content element not found', 'warning', { context: 'auth-check' });
            }
        }
        
        // Return authentication status
        return isAuthenticated;
    } catch (error) {
        window.OpsieApi.log('Error in auth-check: ' + error.message, 'error', {
            context: 'auth-check',
            message: error.message,
            stack: error.stack,
            time: new Date().toISOString()
        });
        
        // Show error UI in a safe way
        const authErrorContainer = document.getElementById('auth-error-container');
        if (authErrorContainer) {
            authErrorContainer.style.display = 'flex';
        }
        
        const authContainer = document.getElementById('auth-container');
        if (authContainer) {
            authContainer.style.display = 'flex';
        }
        
        const mainContent = document.getElementById('main-content');
        if (mainContent) {
            mainContent.style.display = 'none';
        }
        
        if (document.getElementById('main-content')) {
            document.getElementById('main-content').style.display = 'none';
        }
        
        window.OpsieApi.showNotification('Authentication error: ' + error.message, 'error');
        
        return false;
    }
}

// Setup event listeners for UI components
function setupEventListeners() {
    try {
        // Main action buttons
        const generateSummaryButton = document.getElementById('generate-summary-button');
        if (generateSummaryButton) {
            generateSummaryButton.addEventListener('click', handleGenerateSummary);
        }
        
        const generateContactButton = document.getElementById('generate-contact-button');
        if (generateContactButton) {
            generateContactButton.addEventListener('click', handleGetContact);
        }
        
        const generateReplyButton = document.getElementById('generate-reply-button');
        if (generateReplyButton) {
            generateReplyButton.addEventListener('click', handleGenerateReply);
        }
        
        const saveEmailButton = document.getElementById('save-email-button');
        if (saveEmailButton) {
            saveEmailButton.addEventListener('click', handleSaveEmail);
        }
        
        // Extract Questions button
        const extractQuestionsButton = document.getElementById('extract-qa-button');
        if (extractQuestionsButton) {
            extractQuestionsButton.addEventListener('click', handleExtractQuestions);
        }
        
        // Add event listener for the Mark as Handled button
        const markHandledButton = document.getElementById('mark-handled-button');
        if (markHandledButton) {
            markHandledButton.addEventListener('click', showHandlingModal);
        }
        
        // Add event listeners for the handling modal
        const handlingModalClose = document.getElementById('handling-modal-close');
        if (handlingModalClose) {
            handlingModalClose.addEventListener('click', hideHandlingModal);
        }
        
        const handlingCancel = document.getElementById('handling-cancel');
        if (handlingCancel) {
            handlingCancel.addEventListener('click', hideHandlingModal);
        }
        
        const handlingConfirm = document.getElementById('handling-confirm');
        if (handlingConfirm) {
            handlingConfirm.addEventListener('click', function() {
                const note = document.getElementById('handling-note').value.trim();
                hideHandlingModal();
                markEmailAsHandled(note);
            });
        }
        
        // Add keyboard event listener to handling note textarea
        const handlingNote = document.getElementById('handling-note');
        if (handlingNote) {
            handlingNote.addEventListener('keydown', function(e) {
                // Submit on Ctrl+Enter or Cmd+Enter
                if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                    const note = this.value.trim();
                    hideHandlingModal();
                    markEmailAsHandled(note);
                    e.preventDefault();
                }
            });
        }
        
        // Button that's actually in the reply section
        const btnGenerateReplyNow = document.getElementById('btn-generate-reply-now');
        if (btnGenerateReplyNow) {
            btnGenerateReplyNow.addEventListener('click', handleGenerateReply);
        }
        
        // Reply action buttons
        const copyReplyButton = document.getElementById('copy-reply-button');
        if (copyReplyButton) {
            copyReplyButton.addEventListener('click', handleCopyReply);
        }
        
        const insertReplyButton = document.getElementById('insert-reply-button');
        if (insertReplyButton) {
            insertReplyButton.addEventListener('click', handleInsertReply);
        }
        
        // Search functionality
        const emailSearchButton = document.getElementById('email-search-button');
        if (emailSearchButton) {
            emailSearchButton.addEventListener('click', handleSearch);
        }
        
        // Set up all settings-related event listeners
        setupSettingsEventListeners();
        
        // Login button (for authentication)
        const loginButton = document.getElementById('login-button');
        if (loginButton) {
            loginButton.addEventListener('click', handleLogin);
        }
        
        // Authentication view switching
        const showSignupLink = document.getElementById('show-signup-link');
        if (showSignupLink) {
            showSignupLink.addEventListener('click', function(e) {
                e.preventDefault();
                showSignupView();
            });
        }
        
        const showLoginLink = document.getElementById('show-login-link');
        if (showLoginLink) {
            showLoginLink.addEventListener('click', function(e) {
                e.preventDefault();
                showLoginView();
            });
        }
        
        const forgotPasswordLink = document.getElementById('forgot-password-link');
        if (forgotPasswordLink) {
            forgotPasswordLink.addEventListener('click', function(e) {
                e.preventDefault();
                showPasswordResetView();
            });
        }
        
        const backToLoginLink = document.getElementById('back-to-login-link');
        if (backToLoginLink) {
            backToLoginLink.addEventListener('click', function(e) {
                e.preventDefault();
                showLoginView();
            });
        }
        
        // Signup button
        const signupButton = document.getElementById('signup-button');
        if (signupButton) {
            signupButton.addEventListener('click', function(e) {
                e.preventDefault();
                handleSignup();
            });
        }
        
        // Password reset button
        const resetPasswordButton = document.getElementById('reset-password-button');
        if (resetPasswordButton) {
            resetPasswordButton.addEventListener('click', function(e) {
                e.preventDefault();
                handlePasswordReset();
            });
        }
        
        // Keyboard event listeners for signup form
        const signupInputs = [
            'signup-first-name',
            'signup-last-name', 
            'signup-email',
            'signup-password',
            'signup-confirm-password'
        ];
        
        signupInputs.forEach((inputId, index) => {
            const input = document.getElementById(inputId);
            if (input) {
                input.addEventListener('keydown', function(e) {
                    if (e.key === 'Enter') {
                        e.preventDefault();
                        if (index < signupInputs.length - 1) {
                            // Move to next input
                            const nextInput = document.getElementById(signupInputs[index + 1]);
                            if (nextInput) {
                                nextInput.focus();
                            }
                        } else {
                            // Last input, submit form
                            handleSignup();
                        }
                    }
                });
            }
        });
        
        // Keyboard event listener for password reset form
        const resetEmailInput = document.getElementById('reset-email');
        if (resetEmailInput) {
            resetEmailInput.addEventListener('keydown', function(e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    handlePasswordReset();
                }
            });
        }
        
        // Add keyboard event listeners for the login form
        const authEmailInput = document.getElementById('auth-email');
        const authPasswordInput = document.getElementById('auth-password');
        
        if (authEmailInput) {
            authEmailInput.addEventListener('keydown', function(e) {
                if (e.key === 'Enter') {
                    if (authPasswordInput) {
                        authPasswordInput.focus();
                    } else {
                        handleLogin();
                    }
                    e.preventDefault();
                }
            });
        }
        
        if (authPasswordInput) {
            authPasswordInput.addEventListener('keydown', function(e) {
                if (e.key === 'Enter') {
                    handleLogin();
                    e.preventDefault();
                }
            });
        }
        
        // Auth error login button (shown when authentication errors occur)
        const authErrorLoginButton = document.getElementById('auth-error-login-button');
        if (authErrorLoginButton) {
            authErrorLoginButton.addEventListener('click', function() {
                // Show the regular login form
                const authErrorContainer = document.getElementById('auth-error-container');
                const authContainer = document.getElementById('auth-container');
                
                if (authErrorContainer) {
                    authErrorContainer.style.display = 'none';
                }
                
                if (authContainer) {
                    authContainer.style.display = 'flex';
                }
                
                // Focus the email input if available
                const emailInput = document.getElementById('auth-email');
                if (emailInput) {
                    emailInput.focus();
                }
            });
        }
        
        // Error handling for notification close buttons
        const notificationCloseButtons = document.querySelectorAll('.notification-close');
        notificationCloseButtons.forEach(button => {
            button.addEventListener('click', function() {
                this.parentElement.style.display = 'none';
            });
        });
        
        // Load saved API key
        loadApiKey();
        
        // Manual Question Submission button
        const submitManualQuestionButton = document.getElementById('submit-manual-question');
        if (submitManualQuestionButton) {
            submitManualQuestionButton.addEventListener('click', handleManualQuestionSubmission);
        }
        
        // Add keyboard event listener for manual question input (Enter to submit)
        const manualQuestionInput = document.getElementById('manual-question-input');
        if (manualQuestionInput) {
            manualQuestionInput.addEventListener('keydown', function(e) {
                // Submit on Ctrl+Enter or Cmd+Enter
                if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                    handleManualQuestionSubmission();
                    e.preventDefault();
                }
            });
        }
        
        log('Event listeners set up successfully');
    } catch (error) {
        log('Error setting up event listeners: ' + error.message, 'error', error);
    }
}

// Toggle settings visibility
function toggleSettings() {
    const settingsContainer = document.getElementById('settings-container');
    if (settingsContainer) {
        if (settingsContainer.style.display === 'none' || !settingsContainer.style.display) {
            settingsContainer.style.display = 'block';
            // Load the API key to show in the input
            loadApiKey();
        } else {
            settingsContainer.style.display = 'none';
        }
    }
}

// Load API key from storage
function loadApiKey() {
    try {
        const apiKeyInput = document.getElementById('openai-api-key');
        if (apiKeyInput) {
            const savedKey = localStorage.getItem('openaiApiKey');
            if (savedKey) {
                // Mask the key when displaying
                apiKeyInput.value = '••••••••••••••••••••••' + savedKey.substring(savedKey.length - 5);
                apiKeyInput.dataset.masked = 'true';
                showApiKeyStatus(true, 'API key saved');
            } else {
                apiKeyInput.value = '';
                apiKeyInput.dataset.masked = 'false';
                showApiKeyStatus(false, 'No API key saved');
            }

            // Clear value when clicked if it's masked
            apiKeyInput.addEventListener('focus', function() {
                if (this.dataset.masked === 'true') {
                    this.value = '';
                    this.dataset.masked = 'false';
                }
            });
        }
    } catch (error) {
        log('Error loading API key: ' + error.message, 'error', error);
    }
}

// Save API key to storage
function saveApiKey() {
    try {
        const apiKeyInput = document.getElementById('openai-api-key');
        if (apiKeyInput && apiKeyInput.value.trim()) {
            const apiKey = apiKeyInput.value.trim();
            
            // Only save if it's a new key (not masked)
            if (apiKeyInput.dataset.masked !== 'true') {
                // Validate API key format (basic check)
                if (apiKey.startsWith('sk-')) {
                    localStorage.setItem('openaiApiKey', apiKey);
                    showNotification('API key saved successfully!', 'success');
                    
                    // Mask the key in the UI
                    apiKeyInput.value = '••••••••••••••••••••••' + apiKey.substring(apiKey.length - 5);
                    apiKeyInput.dataset.masked = 'true';
                    showApiKeyStatus(true, 'API key saved');
                } else {
                    showErrorNotification('Invalid API key format. OpenAI API keys start with "sk-"');
                }
            }
        } else {
            showErrorNotification('Please enter an API key');
        }
    } catch (error) {
        log('Error saving API key: ' + error.message, 'error', error);
        showErrorNotification('Error saving API key: ' + error.message);
    }
}

// Show API key status
function showApiKeyStatus(isValid, message) {
    try {
        // Look for existing status element
        let statusElement = document.querySelector('.api-key-status');
        
        // Create if it doesn't exist
        if (!statusElement) {
            statusElement = document.createElement('div');
            statusElement.className = 'api-key-status';
            
            // Insert after the API key input container
            const container = document.querySelector('.api-key-input-container');
            if (container) {
                container.insertAdjacentElement('afterend', statusElement);
            }
        }
        
        // Set the message and style
        if (isValid) {
            statusElement.textContent = message;
            statusElement.style.color = '#28a745';
        } else {
            statusElement.textContent = message;
            statusElement.style.color = '#dc3545';
        }
    } catch (error) {
        log('Error showing API key status: ' + error.message, 'error', error);
    }
}

// Disable login controls during authentication
function disableLoginControls() {
    const loginButton = document.getElementById('login-button');
    const emailInput = document.getElementById('auth-email');
    const passwordInput = document.getElementById('auth-password');
    
    if (loginButton) {
        loginButton.disabled = true;
    }
    
    if (emailInput) {
        emailInput.disabled = true;
    }
    
    if (passwordInput) {
        passwordInput.disabled = true;
    }
    
    log('Login controls disabled', 'info');
}

// Enable login controls after authentication attempt
function enableLoginControls() {
    const loginButton = document.getElementById('login-button');
    const emailInput = document.getElementById('auth-email');
    const passwordInput = document.getElementById('auth-password');
    
    if (loginButton) {
        loginButton.disabled = false;
    }
    
    if (emailInput) {
        emailInput.disabled = false;
    }
    
    if (passwordInput) {
        passwordInput.disabled = false;
    }
    
    // Show the auth container again in case of error
    const authContainer = document.getElementById('auth-container');
    const loginLoading = document.getElementById('login-loading');
    
    if (authContainer) {
        authContainer.style.display = 'flex';
    }
    
    if (loginLoading) {
        loginLoading.style.display = 'none';
    }
    
    log('Login controls enabled', 'info');
}

// Function to show login error message
function showLoginError(message) {
    // Show error notification since there's no specific login error element
    showErrorNotification('Login error: ' + message);
    enableLoginControls();
}

// Handle login button click
async function handleLogin() {
    try {
        disableLoginControls();
        // Show loading indicator and hide the form inputs
        const loginLoading = document.getElementById('login-loading');
        if (loginLoading) {
            loginLoading.style.display = 'flex';
        }
        
        let email = document.getElementById('auth-email').value;
        let password = document.getElementById('auth-password').value;
        
        // Validate input
        if (!email || !password) {
            showLoginError('Please enter email and password');
            return;
        }
        
        // Attempt login with Supabase
        const loginResult = await loginWithSupabase(email, password);
        
        if (loginResult.error) {
            showLoginError(loginResult.error.message || 'Login failed');
            return;
        }
        
        log('Login successful, initializing session', 'info', loginResult);
        
        // Hide auth container
        const authContainer = document.getElementById('auth-container');
        if (authContainer) {
            authContainer.style.display = 'none';
        }
        
        // Don't show main content immediately - wait for team validation
        // The main content will be shown by showMainUISections() after team validation
        const mainContent = document.getElementById('main-content');
        if (mainContent) {
            mainContent.style.display = 'none'; // Keep hidden until team validation
        }
        
        // Make sure to reload user settings information to update the UI with new user info
        // Force a refresh to ensure we're not using cached data
        try {
            await loadUserSettingsInfo(true);
            log('User settings information refreshed after login', 'info');
        } catch (settingsError) {
            log('Error refreshing user settings after login', 'error', settingsError);
        }
        
        // Clear any cached team data to ensure fresh data after login
        localStorage.removeItem('currentTeamId');
        localStorage.removeItem('currentTeamName');
        
        // Initialize team and user info with a callback that will reload the email data once team info is ready
        window.OpsieApi.initTeamAndUserInfo(function(teamInfo) {
            log('Team info initialized within login callback', 'info', teamInfo);
            
            // Check if user has a team
            if (!teamInfo.teamId) {
                log('User does not have a team, showing team selection view', 'info');
                
                // Hide main content and show team selection
                const mainContent = document.getElementById('main-content');
                if (mainContent) {
                    mainContent.style.display = 'none';
                }
                
                // Show team selection view
                if (typeof window.showTeamSelectView === 'function') {
                    window.showTeamSelectView();
                } else {
                    log('showTeamSelectView function not available', 'error');
                    showErrorNotification('Team selection not available. Please refresh the page.');
                }
                return;
            }
            
            // User has a team, proceed with normal flow
            log('User has team, proceeding with normal login flow', 'info', { teamId: teamInfo.teamId });
            
            // Hide team selection view and ensure main content will be shown when ready
            if (typeof window.hideAuthContainer === 'function') {
                window.hideAuthContainer();
                log('Hidden team selection view - user has team', 'info');
            }
            
            // Show the settings button since user has a team
            const settingsButton = document.getElementById('settings-button');
            if (settingsButton) {
                settingsButton.style.display = 'block';
                log('Shown settings button - user has team', 'info');
            }
            
            // Only try to reload email data if we have an active mailbox context and a team ID
            if (Office.context.mailbox && teamInfo.teamId) {
                log('Reloading email data after login with team ID: ' + teamInfo.teamId, 'info');
                
                // Add a small delay to ensure all UI components have updated
                setTimeout(() => {
                    try {
                        loadCurrentEmail();
                        log('Email reload triggered after successful login', 'info');
                    } catch (emailError) {
                        log('Error reloading email after login', 'error', emailError);
                    }
                }, 500);
            } else {
                if (!Office.context.mailbox) {
                    log('Cannot reload email - Office.context.mailbox is not available', 'warning');
                }
                if (!teamInfo.teamId) {
                    log('Cannot reload email - team ID is not available', 'warning');
                }
            }
        });
        
        // Start periodic check for authentication status
        startAuthCheck();
        
    } catch (error) {
        showLoginError(error.message || 'An error occurred during login');
        log('Login error:', 'error', error);
    } finally {
        enableLoginControls();
    }
}

// Load the current email content
function loadCurrentEmail() {
    const item = Office.context.mailbox.item;
    
    if (!item) {
        showErrorNotification('No email is selected');
        return;
    }
    
    // Stop any existing notes polling when loading a new email
    stopNotesPolling();
    
    // Reset email data
    currentEmailData = null;
    
    // Start loading indicator - add null checks
    const loadingIndicator = document.getElementById('loading-indicator');
    const emailDetails = document.getElementById('email-details');
    
    if (loadingIndicator) {
        loadingIndicator.style.display = 'block';
    } else {
        log('Warning: loading-indicator element not found', 'warning');
    }
    
    if (emailDetails) {
        emailDetails.style.display = 'none';
    } else {
        log('Warning: email-details element not found', 'warning');
    }
    
    // First, get basic email data
    const emailData = getEmailData();
    
    // Log the data we have
    log('Initial email data:', 'info', {
        subject: emailData.subject,
        sender: emailData.sender,
        timestamp: emailData.timestamp
    });
    
    // Set loading status
    logStep('Getting email body');
    
    // Get email body
    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Text,
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Add the message body to the email data
                emailData.body = result.value;
                emailData.hasBody = true;
                emailData.bodyLength = result.value.length;
                
                // Log success
                log('Retrieved email body, length: ' + result.value.length + ' characters', 'info');
                
                // Store the current email data globally
                currentEmailData = emailData;
                
                // Extra check for message ID
                ensureMessageId(emailData).then(() => {
                    // Check if email already exists in database
                    if (emailData.messageId) {
                        log('Checking if email exists in database: ' + emailData.messageId, 'info');
                        
                        checkIfEmailExists(emailData.messageId, emailData)
                            .then(checkResult => {
                                // Hide loading indicator - add null checks
                                if (loadingIndicator) {
                                    loadingIndicator.style.display = 'none';
                                }
                                
                                if (emailDetails) {
                                    emailDetails.style.display = 'block';
                                }
                                
                                logStep('Email loaded');
                                
                                if (checkResult.exists) {
                                    log('Email exists in database', 'info', checkResult);
                                    
                                    // Show the message already exists 
                                    // Use function to update saved status in UI
                                    updateSavedStatus(checkResult);
                                    
                                    // If email has summary and urgency, display it
                                    // Note: For secondary matches, this should already be displayed in checkIfEmailExists
                                    if (checkResult.summary && checkResult.foundBy !== 'secondary') {
                                        log('Existing email has summary data, displaying it', 'info');
                                        displayExistingSummary(checkResult.summary, checkResult.urgency);
                                    }
                                    
                                    // Store the existing message data in currentEmailData
                                    if (!currentEmailData.existingMessage) {
                                        currentEmailData.existingMessage = {};
                                    }
                                    
                                    // Update with existing message information
                                    currentEmailData.existingMessage = {
                                        ...currentEmailData.existingMessage,
                                        exists: true,
                                        message: checkResult.message,
                                        user: checkResult.user,
                                        savedAt: checkResult.savedAt,
                                        summary: checkResult.summary,
                                        urgency: checkResult.urgency
                                    };
                                    
                                    // If there is handling information, update that too
                                    if (checkResult.handling) {
                                        currentEmailData.existingMessage.handling = checkResult.handling;
                                        
                                        // Update the handling status display
                                        updateHandlingStatus({
                                            message: checkResult.message,
                                            handling: checkResult.handling
                                        });
                                    }
                                    
                                    // Show the notes section since the email is saved
                                    const notesSection = document.getElementById('notes-section');
                                    if (notesSection) {
                                        notesSection.style.display = 'block';
                                        
                                        // Load notes for the current email
                                        loadNotesForCurrentEmail();
                                    }
                                } else {
                                    log('Email not found in database', 'info');
                                    // Enable the save button
                                    const saveButton = document.getElementById('save-email-button');
                                    if (saveButton) {
                                        saveButton.disabled = false;
                                        saveButton.textContent = 'Save Email';
                                        saveButton.style.backgroundColor = '#4CAF50';
                                    }
                                }
                                
                                // Make sure all appropriate UI sections are visible
                                showMainUISections();
                                
                                // Update email details display
                                updateEmailDetails();
                                
                                // Update notes UI based on whether the email is saved
                                updateNotesUIState();
                            })
                            .catch(error => {
                                log('Error checking email existence:', 'error', error);
                                // Hide loading indicator - add null checks
                                if (loadingIndicator) {
                                    loadingIndicator.style.display = 'none';
                                }
                                
                                if (emailDetails) {
                                    emailDetails.style.display = 'block';
                                }
                                
                                // Reset the save button in case of error
                                const saveButton = document.getElementById('save-email-button');
                                if (saveButton) {
                                    saveButton.disabled = false;
                                    saveButton.textContent = 'Save Email';
                                    saveButton.style.backgroundColor = '#4CAF50';
                                }
                                
                                // Make sure all appropriate UI sections are visible
                                showMainUISections();
                                
                                // Update email details display even if we had an error
                                updateEmailDetails();
                                
                                // Update notes UI based on whether the email is saved
                                updateNotesUIState();
                            });
                    } else {
                        log('No message ID available to check database', 'warning');
                        // Hide loading indicator - add null checks
                        if (loadingIndicator) {
                            loadingIndicator.style.display = 'none';
                        }
                        
                        if (emailDetails) {
                            emailDetails.style.display = 'block';
                        }
                        
                        // Reset the save button since we can't check the database
                        const saveButton = document.getElementById('save-email-button');
                        if (saveButton) {
                            saveButton.disabled = false;
                            saveButton.textContent = 'Save Email';
                            saveButton.style.backgroundColor = '#4CAF50';
                        }
                        
                        // Make sure all appropriate UI sections are visible
                        showMainUISections();
                        
                        // Update email details display anyway
                        updateEmailDetails();
                        
                        // Update notes UI based on whether the email is saved
                        updateNotesUIState();
                    }
                }).catch(error => {
                    log('Error ensuring message ID:', 'error', error);
                    // Hide loading indicator - add null checks
                    if (loadingIndicator) {
                        loadingIndicator.style.display = 'none';
                    }
                    
                    if (emailDetails) {
                        emailDetails.style.display = 'block';
                    }
                    
                    // Reset the save button since there was an error
                    const saveButton = document.getElementById('save-email-button');
                    if (saveButton) {
                        saveButton.disabled = false;
                        saveButton.textContent = 'Save Email';
                        saveButton.style.backgroundColor = '#4CAF50';
                    }
                    
                    // Make sure all appropriate UI sections are visible
                    showMainUISections();
                    
                    // Update email details display even if we had an error
                    updateEmailDetails();
                    
                    // Update notes UI based on whether the email is saved
                    updateNotesUIState();
                });
            } else {
                log('Error getting email body', 'error', result.error);
                // Hide loading indicator - add null checks
                if (loadingIndicator) {
                    loadingIndicator.style.display = 'none';
                }
                
                if (emailDetails) {
                    emailDetails.style.display = 'block';
                }
                
                // Reset the save button since we couldn't get the email body
                const saveButton = document.getElementById('save-email-button');
                if (saveButton) {
                    saveButton.disabled = false;
                    saveButton.textContent = 'Save Email';
                    saveButton.style.backgroundColor = '#4CAF50';
                }
                
                // Make sure all appropriate UI sections are visible
                showMainUISections();
                
                // Update email details display even if we had an error
                updateEmailDetails();
                
                // Update notes UI based on whether the email is saved
                updateNotesUIState();
                
                // Set loading status - error
                logStep('Error loading email body');
            }
        }
    );
}

// Helper function to ensure we have the best message ID available
async function ensureMessageId(emailData) {
    // Wait a short time for any asynchronous ID extraction to complete
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // If we don't have a message ID yet, try once more
    if (!emailData.messageId && Office.context.mailbox.item) {
        const item = Office.context.mailbox.item;
        
        try {
            // Try the REST ID conversion as it's most reliable
            if (item.itemId && Office.context.mailbox.convertToRestId) {
                const restId = Office.context.mailbox.convertToRestId(
                    item.itemId, 
                    Office.MailboxEnums.RestVersion.v2_0
                );
                
                if (restId) {
                    emailData.messageId = restId;
                    log('Used REST ID conversion in ensureMessageId: ' + restId, 'info');
                    
                    // Look for the AAkAL format
                    const idMatch = restId.match(/AAkAL[A-Za-z0-9+/=%]+/);
                    if (idMatch) {
                        emailData.messageId = idMatch[0];
                        log('Extracted URL format ID (AAkAL) in ensureMessageId', 'info');
                    }
                }
            }
        } catch (error) {
            log('Error in ensureMessageId: ' + error.message, 'warning');
        }
    }
    
    log('Final message ID in ensureMessageId: ' + (emailData.messageId || 'Not available'), 'info');
}

// Check if an email already exists in the database
async function checkIfEmailExists(messageId, emailData) {
    try {
        // First, ensure we have authentication
        const isAuthed = await window.OpsieApi.isAuthenticated();
        if (!isAuthed) {
            log('User not authenticated, cannot check email existence', 'warning');
            return { exists: false };
        }
        
        // Get the team ID
        let teamId = localStorage.getItem('currentTeamId');
        if (!teamId) {
            log('No team ID found, cannot check email existence', 'warning');
            return { exists: false };
        }
        
        // Get email data for secondary check if not provided
        emailData = emailData || currentEmailData || getEmailData();
        
        log(`Checking for existing email with primary ID: ${messageId}`, 'info');
        if (emailData && emailData.sender && emailData.sender.email) {
            log(`Fallback check prepared with sender: ${emailData.sender.email} and timestamp near: ${emailData.timestamp}`, 'info');
        }
        
        // Define API URL base - use the API from window.OpsieApi or fall back to direct Supabase URL
        const apiBase = window.OpsieApi.API_BASE_URL || 'https://yourdomain.supabase.co/rest/v1';
        
        // Primary check with message external ID
        if (messageId) {
            const checkResponse = await fetch(`${apiBase}/messages?message_external_id=eq.${encodeURIComponent(messageId)}&team_id=eq.${teamId}&select=*`, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'apikey': window.OpsieApi.SUPABASE_KEY,
                    'Authorization': `Bearer ${localStorage.getItem(window.OpsieApi.STORAGE_KEY_TOKEN)}`
                }
            });
            
            if (checkResponse.ok) {
                const messages = await checkResponse.json();
                
                if (messages && messages.length > 0) {
                    log('Found existing message by primary ID check', 'info', messages[0]);
                    
                    // Use the closest match
                    const closestMessage = messages[0];
                    log('Using closest timestamp match', 'info', {
                        messageTimestamp: closestMessage.timestamp,
                        searchTimestamp: emailData.timestamp,
                        diffSeconds: Math.abs(new Date(closestMessage.timestamp) - new Date(emailData.timestamp)) / 1000
                    });
                    
                    // Get summary and urgency data if available
                    if (closestMessage.summary) {
                        log('Found existing summary in database', 'info', closestMessage.summary);
                    }
                    
                    if (closestMessage.urgency !== null && closestMessage.urgency !== undefined) {
                        log('Found existing urgency score in database', 'info', closestMessage.urgency);
                    }
                    
                    // Try to get user information for who saved this message
                    let userInfo = null;
                    
                    if (closestMessage.user_id) {
                        try {
                            // Fetch user details using RPC function to bypass RLS
                            const userResponse = await fetch(`${apiBase}/rpc/get_user_details`, {
                                method: 'POST',
                                headers: {
                                    'Content-Type': 'application/json',
                                    'apikey': window.OpsieApi.SUPABASE_KEY,
                                    'Authorization': `Bearer ${localStorage.getItem(window.OpsieApi.STORAGE_KEY_TOKEN)}`
                                },
                                body: JSON.stringify({
                                    user_ids: [closestMessage.user_id]
                                })
                            });
                            
                            if (userResponse.ok) {
                                const users = await userResponse.json();
                                
                                if (users && users.length > 0) {
                                    const userData = users[0];
                                    let userName = 'Unknown User';
                                    
                                    if (userData.first_name || userData.last_name) {
                                        userName = `${userData.first_name || ''} ${userData.last_name || ''}`.trim();
                                    } else if (userData.email) {
                                        userName = userData.email;
                                    }
                                    
                                    userInfo = {
                                        name: userName,
                                        email: userData.email
                                    };
                                    
                                    log('Retrieved user information for message (primary check)', 'info', userInfo);
                                }
                            }
                        } catch (userLookupError) {
                            log('Error retrieving user information', 'error', userLookupError);
                        }
                    }
                    
                    // After email is found, prepare the result object
                    const result = {
                        exists: true,
                        message: closestMessage,
                        foundBy: 'primary',
                        user: userInfo || { name: 'Unknown User' },
                        savedAt: closestMessage.created_at || closestMessage.timestamp
                    };
                    
                    // Add summary and urgency to the result if available
                    if (closestMessage.summary) {
                        result.summary = closestMessage.summary;
                    }
                    
                    if (closestMessage.urgency !== null && closestMessage.urgency !== undefined) {
                        result.urgency = closestMessage.urgency;
                    }
                    
                    // After email is found, display any existing summary and urgency
                    if (closestMessage.summary) {
                        // Display the summary in the UI
                        displayExistingSummary(closestMessage.summary, closestMessage.urgency);
                    }
                    
                    // Update currentEmailData with the existing message reference
                    if (!currentEmailData.existingMessage) {
                        currentEmailData.existingMessage = {};
                    }
                    
                    currentEmailData.existingMessage = {
                        ...currentEmailData.existingMessage,
                        exists: true,
                        message: closestMessage,
                        user: userInfo,
                        savedAt: closestMessage.created_at || closestMessage.timestamp,
                        summary: closestMessage.summary,
                        urgency: closestMessage.urgency
                    };
                    
                    // Check for handling status
                    if (closestMessage.handled_by || closestMessage.handled_at) {
                        result.handling = {
                            isHandled: true,
                            handledAt: closestMessage.handled_at,
                            handlingNote: closestMessage.handling_note
                        };
                        
                        currentEmailData.existingMessage.handling = result.handling;
                    }
                    
                    return result;
                }
            }
        }
        
        // Secondary check by sender, timestamp, and optionally subject (less reliable)
        if (emailData && emailData.sender && emailData.sender.email && emailData.timestamp) {
            try {
                // Standardize the timestamp to ISO format
                let timestamp;
                try {
                    timestamp = new Date(emailData.timestamp).toISOString();
                } catch (e) {
                    log('Error converting timestamp for secondary check', 'error', e);
                    timestamp = new Date().toISOString();
                }
                
                // Construct a query with a 2-minute window (±2 minutes) around the timestamp
                const twoMinutesBeforeTimestamp = new Date(new Date(timestamp).getTime() - 2 * 60 * 1000).toISOString();
                const twoMinutesAfterTimestamp = new Date(new Date(timestamp).getTime() + 2 * 60 * 1000).toISOString();
                
                log('Checking with time window', 'info', {
                    sender: emailData.sender.email,
                    timestampOriginal: timestamp,
                    timeWindow: `${twoMinutesBeforeTimestamp} to ${twoMinutesAfterTimestamp}`
                });
                
                // Fetch messages with the specified sender and within the time window
                const query = 
                    `sender_email=eq.${encodeURIComponent(emailData.sender.email)}` +
                    `&team_id=eq.${teamId}` +
                    `&timestamp=gte.${twoMinutesBeforeTimestamp}` +
                    `&timestamp=lte.${twoMinutesAfterTimestamp}` +
                    `&select=*`;
                
                const timeWindowResponse = await fetch(
                    `${apiBase}/messages?${query}`,
                    {
                        method: 'GET',
                        headers: {
                            'Content-Type': 'application/json',
                            'apikey': window.OpsieApi.SUPABASE_KEY,
                            'Authorization': `Bearer ${localStorage.getItem(window.OpsieApi.STORAGE_KEY_TOKEN)}`
                        }
                    }
                );
                
                if (timeWindowResponse.ok) {
                    const timeWindowMessages = await timeWindowResponse.json();
                    
                    if (timeWindowMessages && timeWindowMessages.length > 0) {
                        log(`Found ${timeWindowMessages.length} messages matching sender and timestamp window`, 'info');
                        
                        // Sort messages by timestamp closest to the search timestamp
                        timeWindowMessages.sort((a, b) => {
                            const aDiff = Math.abs(new Date(a.timestamp) - new Date(timestamp));
                            const bDiff = Math.abs(new Date(b.timestamp) - new Date(timestamp));
                            return aDiff - bDiff;
                        });
                        
                        // Use the closest match
                        const closestMessage = timeWindowMessages[0];
                        log('Using closest timestamp match', 'info', {
                            messageTimestamp: closestMessage.timestamp,
                            searchTimestamp: timestamp,
                            diffSeconds: Math.abs(new Date(closestMessage.timestamp) - new Date(timestamp)) / 1000
                        });
                        
                        // Try to get user information for who saved this message
                        let userInfo = null;
                        
                        if (closestMessage.user_id) {
                            try {
                                // Fetch user details from the users table
                                const userResponse = await fetch(`${apiBase}/users?id=eq.${closestMessage.user_id}&select=first_name,last_name,email`, {
                                    method: 'GET',
                                    headers: {
                                        'Content-Type': 'application/json',
                                        'apikey': window.OpsieApi.SUPABASE_KEY,
                                        'Authorization': `Bearer ${localStorage.getItem(window.OpsieApi.STORAGE_KEY_TOKEN)}`
                                    }
                                });
                                
                                if (userResponse.ok) {
                                    const users = await userResponse.json();
                                    
                                    if (users && users.length > 0) {
                                        const userData = users[0];
                                        let userName = 'Unknown User';
                                        
                                        if (userData.first_name || userData.last_name) {
                                            userName = `${userData.first_name || ''} ${userData.last_name || ''}`.trim();
                                        } else if (userData.email) {
                                            userName = userData.email;
                                        }
                                        
                                        userInfo = {
                                            name: userName,
                                            email: userData.email
                                        };
                                        
                                        log('Retrieved user information for message (secondary check)', 'info', userInfo);
                                    }
                                }
                            } catch (userError) {
                                log('Error fetching user details for secondary check', 'error', userError);
                            }
                        }
                        
                        // Create result object
                        const result = {
                            exists: true,
                            message: closestMessage,
                            foundBy: 'secondary',
                            savedAt: closestMessage.created_at,
                            user: userInfo
                        };
                        
                        // Add summary and urgency to the result if available
                        if (closestMessage.summary) {
                            result.summary = closestMessage.summary;
                            log('Found existing summary in secondary match', 'info', closestMessage.summary);
                        }
                        
                        if (closestMessage.urgency !== null && closestMessage.urgency !== undefined) {
                            result.urgency = closestMessage.urgency;
                            log('Found existing urgency in secondary match', 'info', closestMessage.urgency);
                        }
                        
                        // After email is found, display any existing summary and urgency
                        if (closestMessage.summary) {
                            // Display the summary in the UI
                            displayExistingSummary(closestMessage.summary, closestMessage.urgency);
                        }
                        
                        // Check if the message has handling information
                        if (closestMessage.handled_at || closestMessage.handled_by) {
                            log('Message found by secondary check has handling information', 'info');
                            
                            let handlerInfo = null;
                            
                            // If we have a handled_by user ID, get user details
                            if (closestMessage.handled_by) {
                                try {
                                    // Fetch user details from the users table
                                    const handlerResponse = await fetch(`${apiBase}/users?id=eq.${closestMessage.handled_by}&select=first_name,last_name,email`, {
                                        method: 'GET',
                                        headers: {
                                            'Content-Type': 'application/json',
                                            'apikey': window.OpsieApi.SUPABASE_KEY,
                                            'Authorization': `Bearer ${localStorage.getItem(window.OpsieApi.STORAGE_KEY_TOKEN)}`
                                        }
                                    });
                                    
                                    if (handlerResponse.ok) {
                                        const users = await handlerResponse.json();
                                        
                                        if (users && users.length > 0) {
                                            const userData = users[0];
                                            let userName = 'Unknown User';
                                            
                                            if (userData.first_name || userData.last_name) {
                                                userName = `${userData.first_name || ''} ${userData.last_name || ''}`.trim();
                                            } else if (userData.email) {
                                                userName = userData.email;
                                            }
                                            
                                            handlerInfo = {
                                                name: userName,
                                                email: userData.email
                                            };
                                            
                                            log('Retrieved handler information for secondary check', 'info', handlerInfo);
                                        }
                                    }
                                } catch (handlerError) {
                                    log('Error fetching handler details for secondary check', 'error', handlerError);
                                }
                            }
                            
                            // Add handling information to the result
                            result.handling = {
                                handledAt: closestMessage.handled_at,
                                handledBy: handlerInfo || { name: 'Unknown User' },
                                handlingNote: closestMessage.handling_note
                            };
                            
                            log('Added handling info to result for secondary check', 'info', result.handling);
                        }
                        
                        return result;
                    } else {
                        log('No matches found with timestamp window check', 'info');
                    }
                } else {
                    log('API error in secondary check', 'warning', {
                        status: timeWindowResponse.status,
                        statusText: timeWindowResponse.statusText
                    });
                }
            } catch (secondaryCheckError) {
                log('Error during secondary message check', 'error', secondaryCheckError);
            }
        }
        
        // If we get here, email doesn't exist by either check
        return { exists: false };
    } catch (error) {
        log('Exception checking if email exists', 'error', error);
        return { exists: false };
    }
}

// New function to display an existing summary from the database
function displayExistingSummary(summaryText, urgencyScore) {
    try {
        // First, make sure the summary container is visible
        const summaryContainer = document.getElementById('summary-container');
        if (summaryContainer) {
            summaryContainer.style.display = 'block';
        }
        
        // Parse the summary text into an array of items
        let summaryItems = [];
        
        // Check if it's stored as a pipe-separated string (common format)
        if (typeof summaryText === 'string') {
            if (summaryText.includes('|')) {
                summaryItems = summaryText.split('|').map(item => item.trim());
            } 
            // Check if it might be stored as newlines
            else if (summaryText.includes('\n')) {
                summaryItems = summaryText.split('\n').map(item => item.trim());
            }
            // Default to treating it as a single item
            else {
                summaryItems = [summaryText];
            }
        } 
        // If it's already an array, use it directly
        else if (Array.isArray(summaryText)) {
            summaryItems = summaryText;
        }
        
        // Display the summary items
        const summaryItemsElement = document.getElementById('summary-items');
        if (summaryItemsElement && summaryItems.length > 0) {
            summaryItemsElement.innerHTML = '';
            
            summaryItems.forEach(item => {
                const li = document.createElement('li');
                li.textContent = item;
                summaryItemsElement.appendChild(li);
            });
            
            log('Displayed existing summary from database', 'info', summaryItems);
        }
        
        // Display urgency score if available
        if (urgencyScore !== null && urgencyScore !== undefined) {
            const urgencyFill = document.getElementById('urgency-fill');
            if (urgencyFill) {
                const score = parseFloat(urgencyScore);
                const percentage = Math.min(Math.max((score / 10) * 100, 0), 100);
                
                // Update the fill width
                urgencyFill.style.width = `${percentage}%`;
                
                // Set color based on urgency level
                let color;
                if (score <= 3) {
                    color = '#4CAF50'; // Green for low urgency
                } else if (score <= 7) {
                    color = '#FFA500'; // Orange for medium urgency
                } else {
                    color = '#FF0000'; // Red for high urgency
                }
                
                // Update the fill color
                urgencyFill.style.backgroundColor = color;
                
                log('Displayed existing urgency from database', 'info', urgencyScore);
            }
        }
    } catch (error) {
        log('Error displaying existing summary', 'error', error);
    }
}

// Update the UI to show an email has been saved
async function updateSavedStatus(messageData) {
    try {
        // Log the entire message data structure for debugging
        log('Updating saved status with data', 'info', messageData);

        // Update the saved status display
        const savedStatusElement = document.getElementById('saved-status');
        if (!savedStatusElement) {
            log('Warning: saved-status element not found', 'warning');
            return;
        }

        // Try to get user info - check all possible locations
        let userName = 'Unknown user';
        let userEmail = '';
        
        // Modern format (directly in checkResult object)
        if (messageData.user && messageData.user.name) {
            userName = messageData.user.name;
            userEmail = messageData.user.email || '';
            log('Found user info in root user property', 'info');
        }
        // Format from browser extension in message.user_id
        else if (messageData.message && messageData.message.user) {
            userName = messageData.message.user.name || 'Unknown user';
            userEmail = messageData.message.user.email || '';
            log('Found user info in message.user property', 'info');
        }
        // Legacy format
        else if (messageData.user_details && messageData.user_details.name) {
            userName = messageData.user_details.name;
            userEmail = messageData.user_details.email || '';
            log('Found user info in user_details property', 'info');
        }
        // If we still don't have user info, try to fetch it using user ID
        else if (messageData.message && messageData.message.user_id) {
            const userId = messageData.message.user_id;
            log('Found user ID in message.user_id: ' + userId + ', fetching user details', 'info');
            
            const userDetails = await fetchUserDetails(userId);
            if (userDetails && userDetails.name) {
                userName = userDetails.name;
                userEmail = userDetails.email || '';
                log('Successfully fetched user details for saved status: ' + userName, 'info');
            } else {
                userName = `User ${userId}`;
                log('Could not fetch user details for saved status, using fallback name', 'warning');
            }
        }
        
        // Format the saved date - check all possible locations
        let savedDate = 'unknown date';
        try {
            // Check different possible locations for the timestamp
            if (messageData.savedAt) {
                savedDate = new Date(messageData.savedAt).toLocaleString();
                log('Found timestamp in savedAt property', 'info');
            }
            else if (messageData.message && messageData.message.created_at) {
                savedDate = new Date(messageData.message.created_at).toLocaleString();
                log('Found timestamp in message.created_at property', 'info');
            }
            else if (messageData.created_at) {
                savedDate = new Date(messageData.created_at).toLocaleString();
                log('Found timestamp in created_at property', 'info');
            }
        } catch (e) {
            log('Error formatting saved date', 'error', e);
        }
        
        // How the email was found
        let foundByText = '';
        if (messageData.foundBy === 'secondary' || messageData.foundBySecondaryCheck) {
            foundByText = ' (matched by sender and timestamp)';
        }
        
        savedStatusElement.innerHTML = `<div class="saved-badge">✓ Saved</div>
        <div class="saved-details">This email was saved to the database by ${userName} at ${savedDate}${foundByText}</div>`;
        savedStatusElement.style.display = 'block';
        
        log('Updated saved status with user: ' + userName + ' and date: ' + savedDate);
        
        // Disable the save button to prevent saving again - match sidebar behavior
        const saveButton = document.getElementById('save-email-button');
        if (saveButton) {
            saveButton.disabled = true;
            saveButton.textContent = 'Already Saved';
            saveButton.style.backgroundColor = '#999';
            log('Disabled Save Email button (from updateSavedStatus)', 'info');
        }
        
        // Update the Mark as Handled button to be enabled since the email is saved
        const markHandledButton = document.getElementById('mark-handled-button');
        if (markHandledButton) {
            markHandledButton.disabled = false;
            markHandledButton.textContent = 'Mark as Handled';
            markHandledButton.style.backgroundColor = '#ff9800';
        }
        
        // Show the notes section since the email is saved
        const notesSection = document.getElementById('notes-section');
        if (notesSection) {
            notesSection.style.display = 'block';
            
            // Load notes for the current email
            loadNotesForCurrentEmail();
        }
        
        // Check if this message is already handled
        if (messageData.handling || 
            (messageData.message && messageData.message.handled_at) ||
            (messageData.message && messageData.message.handled_by)) {
            
            log('Message is already handled, updating handling status display', 'info');
            
            // Create a handling object if it doesn't exist
            if (!messageData.handling) {
                messageData.handling = {};
                
                // Get the handled_at time
                if (messageData.message && messageData.message.handled_at) {
                    messageData.handling.handledAt = messageData.message.handled_at;
                }
                
                // If we have a handled_by user ID, try to get the user details
                if (messageData.message && messageData.message.handled_by) {
                    messageData.handling.handledBy = { 
                        name: 'Unknown User'
                    };
                    
                    // Get the handling note if it exists
                    if (messageData.message && messageData.message.handling_note) {
                        messageData.handling.handlingNote = messageData.message.handling_note;
                    }
                }
            }
            
            // Update the handling display
            updateHandlingStatus(messageData);
            
            // Disable the Mark as Handled button
            if (markHandledButton) {
                markHandledButton.disabled = true;
                markHandledButton.textContent = 'Already Handled';
                markHandledButton.style.backgroundColor = '#999';
            }
        }
    } catch (error) {
        log('Error updating saved status', 'error', error);
    }
}

// Update the email details section
function updateEmailDetails() {
    const item = Office.context.mailbox.item;
    
    if (!item) {
        log('No email item available', 'error');
        return;
    }
    
    try {
        // Get email data
        const emailData = getEmailData();
        
        // Set the subject
        const subjectElement = document.getElementById('email-subject');
        if (subjectElement) {
            subjectElement.textContent = emailData.subject || '(No subject)';
        }
        
        // Set the sender
        const fromElement = document.getElementById('email-from');
        if (fromElement) {
            fromElement.textContent = `${emailData.sender.name} <${emailData.sender.email}>`;
        }
        
        // Set recipients (To)
        const toElement = document.getElementById('email-to');
        if (toElement && item.to) {
            const recipients = item.to.map(recipient => 
                recipient.displayName || recipient.emailAddress
            ).join(', ');
            toElement.textContent = recipients || '(No recipients)';
        }
        
        // Set the preview
        const previewElement = document.getElementById('email-body-preview');
        if (previewElement) {
            // Will be updated when body is loaded, for now show placeholder
            previewElement.textContent = emailData.body 
                ? emailData.body.substring(0, 100) + '...' 
                : 'Loading message body...';
        }
        
        log('Email details updated successfully');
    } catch (error) {
        log('Error updating email details: ' + error.message, 'error', error);
        showErrorNotification('Error displaying email information');
    }
}

// Get the email data needed for API calls
function getEmailData() {
    const item = Office.context.mailbox.item;
    
    if (!item) return null;
    
    // Create the email data object if it doesn't exist
    if (!currentEmailData) {
        currentEmailData = {
            subject: item.subject || '',
            sender: {
                name: item.from ? item.from.displayName : '',
                email: item.from ? item.from.emailAddress : ''
            },
            timestamp: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : new Date().toISOString(),
            body: '', // Will be populated by loadCurrentEmail
            hasAttachments: item.attachments && item.attachments.length > 0,
            messageId: null // Will be filled in below
        };
        
        // Try to extract the ID using multiple methods
        extractMessageId(item)
            .then(id => {
                if (id) {
                    currentEmailData.messageId = id;
                    log('Successfully extracted message ID asynchronously: ' + id, 'info');
                }
            })
            .catch(error => {
                log('Error extracting message ID asynchronously: ' + error.message, 'warning');
            });
        
        // Meanwhile, try synchronous methods as fallback
        
        // Try to get the restId which is often the correct format we want
        if (item.itemId && Office.context.mailbox.convertToRestId) {
            try {
                const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
                currentEmailData.messageId = restId;
                log('Using REST ID: ' + currentEmailData.messageId, 'info');
            } catch (e) {
                log('Error converting to REST ID: ' + e.message, 'warning');
            }
        }
        
        // If still not found, try other properties
        if (!currentEmailData.messageId) {
            // Try the conversation ID which sometimes has the "AAkAL..." format
            if (item.conversationId) {
                currentEmailData.messageId = item.conversationId;
                log('Using conversationId: ' + currentEmailData.messageId, 'info');
            }
            // Fall back to the item ID
            else if (item.itemId) {
                currentEmailData.messageId = item.itemId;
                log('Using itemId as fallback: ' + currentEmailData.messageId, 'info');
            }
            // Try internet message ID as a last resort
            else if (item.internetMessageId) {
                currentEmailData.messageId = item.internetMessageId;
                log('Using internetMessageId: ' + currentEmailData.messageId, 'info');
            }
            // Generate a synthetic ID if nothing else is available
            else if (Office.context.mailbox.diagnostics && Office.context.mailbox.diagnostics.hostName) {
                log('No standard ID properties found, generating alternative ID', 'warning');
                
                const alternativeId = [
                    Office.context.mailbox.userProfile.emailAddress,
                    Office.context.mailbox.diagnostics.hostName,
                    item.dateTimeCreated,
                    item.subject,
                    item.from ? item.from.emailAddress : ''
                ].join('::');
                
                currentEmailData.messageId = `generated-${btoa(alternativeId).replace(/=/g, '')}`;
                log('Generated alternative message ID: ' + currentEmailData.messageId, 'info');
            }
        }
        
        // Clean and validate the message ID
        if (currentEmailData.messageId) {
            // Remove any prefix if present (sometimes it has "ID:" or similar prefix)
            if (currentEmailData.messageId.includes(':')) {
                currentEmailData.messageId = currentEmailData.messageId.split(':').pop().trim();
                log('Extracted clean ID from prefixed format', 'info');
            }
            
            // Look for the AAkAL format which is the URL ID format
            const idMatch = currentEmailData.messageId.match(/AAkAL[A-Za-z0-9+/=%]+/);
            if (idMatch) {
                currentEmailData.messageId = idMatch[0];
                log('Extracted URL format ID (AAkAL) from messageId', 'info');
            }
        } else {
            log('No message ID could be extracted', 'warning');
        }
        
        // Log the final message ID for debugging
        log('Final message ID: ' + (currentEmailData.messageId || 'Not available'), 
            currentEmailData.messageId ? 'info' : 'warning');
    }
    
    return currentEmailData;
}

// Helper function to extract message ID asynchronously using various methods
async function extractMessageId(item) {
    // Try method 1: Get the ID via REST API (most reliable for getting the AAkAL format)
    if (Office.context.mailbox.getCallbackTokenAsync) {
        try {
            log('Attempting to extract ID using REST API', 'info');
            
            // Get an access token for the REST API
            const tokenResult = await new Promise((resolve, reject) => {
                Office.context.mailbox.getCallbackTokenAsync({isRest: true}, result => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value);
                    } else {
                        reject(new Error('Failed to get callback token: ' + result.error.message));
                    }
                });
            });
            
            const restHost = Office.context.mailbox.restUrl;
            const apiUrl = restHost + '/v2.0/me/messages/' + item.itemId;
            
            log('Making REST API request to: ' + apiUrl, 'info');
            
            const response = await fetch(apiUrl, {
                method: 'GET',
                headers: {
                    'Authorization': 'Bearer ' + tokenResult,
                    'Content-Type': 'application/json'
                }
            });
            
            if (!response.ok) {
                throw new Error('REST API request failed: ' + response.status);
            }
            
            const data = await response.json();
            log('REST API response received', 'info', data);
            
            // Extract the ID from the response
            if (data.id) {
                log('Found ID in REST response: ' + data.id, 'info');
                return data.id;
            }
        } catch (error) {
            log('Error using REST API to get message ID: ' + error.message, 'warning');
        }
    }
    
    // Try method 2: Convert the EWS ID to a REST ID
    if (item.itemId && Office.context.mailbox.convertToRestId) {
        try {
            const restId = Office.context.mailbox.convertToRestId(
                item.itemId,
                Office.MailboxEnums.RestVersion.v2_0
            );
            log('Converted EWS ID to REST ID: ' + restId, 'info');
            return restId;
        } catch (error) {
            log('Error converting EWS ID to REST ID: ' + error.message, 'warning');
        }
    }
    
    // Try method 3: Look for the ID in the URL (if we're in a browser)
    if (typeof window !== 'undefined' && window.location && window.location.href) {
        const url = window.location.href;
        log('Checking current URL for ID: ' + url, 'info');
        
        // Look for the AAkAL pattern in the URL
        const urlIdMatch = url.match(/AAkAL[A-Za-z0-9+%/=]+/);
        if (urlIdMatch) {
            log('Found ID in URL: ' + urlIdMatch[0], 'info');
            return urlIdMatch[0];
        }
    }
    
    // No ID found through these methods
    return null;
}

// Add this function to better handle errors
// Helper function to check if user has a team
function checkTeamMembership() {
    const teamId = localStorage.getItem('currentTeamId');
    if (!teamId) {
        log('Team membership check failed - no team ID found', 'warning');
        showErrorNotification('You need to be assigned to a team to use this feature. Please contact your administrator.');
        return false;
    }
    return true;
}

function handleApiError(error, context) {
    console.error(`Error during ${context}:`, error);
    
    // Check for different error types
    if (error.name === 'TypeError' && error.message.includes('Failed to fetch')) {
        showErrorNotification('Network connection issue. Please check your internet connection and try again.');
    } else if (error.message && error.message.includes('Authentication')) {
        // Handle authentication errors
        checkAuthStatus(); // Re-check auth status to show login screen if needed
        showErrorNotification('Authentication error. Please log in again.');
    } else {
        // Generic error handling
        showErrorNotification(error.message || `Error occurred while ${context}`);
    }
}

// Handle generating summary
async function handleGenerateSummary() {
    try {
        // Check if user has a team
        if (!checkTeamMembership()) {
            return;
        }
        
        // Check if email data is available
        if (!currentEmailData || !currentEmailData.body) {
            showErrorNotification('No email data available to summarize');
            return;
        }
        
        // Show the summary container if it's hidden
        const summaryContainer = document.getElementById('summary-container');
        if (summaryContainer) {
            summaryContainer.style.display = 'block';
        }
        
        // Clear previous items
        const summaryItems = document.getElementById('summary-items');
        if (summaryItems) {
            summaryItems.innerHTML = '<li>Generating summary...</li>';
        }
        
        // Reset urgency meter
        const urgencyFill = document.getElementById('urgency-fill');
        if (urgencyFill) {
            urgencyFill.style.width = '0%';
            urgencyFill.style.backgroundColor = '#e0e0e0';
        }
        
        log('Generating summary for email: ' + currentEmailData.subject);
        
        // Call the API to generate the summary
        const summaryResult = await window.OpsieApi.generateEmailSummary(currentEmailData);
        
        log('Summary generated successfully', 'info', summaryResult);
        
        // Check if the missing API key message is in the summary
        const missingApiKey = summaryResult.summaryItems && 
                             summaryResult.summaryItems.some(item => 
                                typeof item === 'string' && 
                                item.includes('Please add your OpenAI API key in settings'));
        
        if (missingApiKey) {
            log('Missing API key detected in summary result', 'warning');
            showErrorNotification('OpenAI API key is required. Please add your API key in the settings panel.');
            
            // Automatically open settings panel to make it easier for the user
            toggleSettings();
            
            return;
        }
        
        // Store the summary and urgency in the currentEmailData object for saving
        if (summaryResult.summaryItems && summaryResult.summaryItems.length > 0) {
            // Join the summary items with a pipe to store as a single string
            currentEmailData.summary = summaryResult.summaryItems.join(' | ');
            log('Stored summary in currentEmailData:', 'info', currentEmailData.summary);
        }
        
        if (summaryResult.urgencyScore !== undefined) {
            currentEmailData.urgency = summaryResult.urgencyScore;
            log('Stored urgency score in currentEmailData:', 'info', currentEmailData.urgency);
        }
        
        // Display the summary items
        if (summaryItems) {
            summaryItems.innerHTML = '';
            
            // Check if summary items exist and handle different formats
            if (summaryResult.summaryItems && summaryResult.summaryItems.length > 0) {
                summaryResult.summaryItems.forEach(item => {
                    const li = document.createElement('li');
                    // Handle both string items and object items with content property
                    li.textContent = typeof item === 'string' ? item : (item.content || '');
                    summaryItems.appendChild(li);
                });
            } else {
                summaryItems.innerHTML = '<li>No summary items returned. Please try again.</li>';
            }
        }
        
        // Display urgency score if available
        if (urgencyFill) {
            let score = 0;
            
            // Handle different formats of urgency in response
            if (summaryResult.urgencyScore !== undefined) {
                // Numeric score (0-10)
                score = summaryResult.urgencyScore;
            } else if (summaryResult.urgency) {
                // Text-based urgency (low, medium, high)
                switch(summaryResult.urgency.toLowerCase()) {
                    case 'high':
                        score = 9;
                        break;
                    case 'medium':
                        score = 5;
                        break;
                    case 'low':
                        score = 2;
                        break;
                    default:
                        score = 0;
                }
            }
            
            const percentage = Math.min(Math.max((score / 10) * 100, 0), 100);
            
            // Update the fill width
            urgencyFill.style.width = `${percentage}%`;
            
            // Set color based on urgency level
            let color;
            if (score <= 3) {
                color = '#4CAF50'; // Green for low urgency
            } else if (score <= 7) {
                color = '#FFA500'; // Orange for medium urgency
            } else {
                color = '#FF0000'; // Red for high urgency
            }
            
            // Update the fill color
            urgencyFill.style.backgroundColor = color;
        }
        
        // Check if the email is already saved in the database
        if (currentEmailData.existingMessage && 
            currentEmailData.existingMessage.exists && 
            currentEmailData.existingMessage.message && 
            currentEmailData.existingMessage.message.id) {
            
            // Email is already saved, update the summary in the database
            log('Email is already saved, updating summary in database', 'info', {
                messageId: currentEmailData.existingMessage.message.id,
                summary: currentEmailData.summary,
                urgency: currentEmailData.urgency
            });
            
            try {
                // Call the updateEmailSummary function to save the summary to the database
                const updateResult = await window.OpsieApi.updateEmailSummary(
                    currentEmailData.existingMessage.message.id,
                    currentEmailData.summary,
                    currentEmailData.urgency
                );
                
                if (updateResult.success) {
                    log('Summary updated in database successfully', 'info', updateResult);
                    showNotification('Summary generated and saved to database', 'success');
                    
                    // Update the existingMessage with the latest data
                    if (updateResult.data && updateResult.data.message) {
                        currentEmailData.existingMessage.message = updateResult.data.message;
                        currentEmailData.existingMessage.summary = currentEmailData.summary;
                        currentEmailData.existingMessage.urgency = currentEmailData.urgency;
                    }
                } else {
                    log('Error updating summary in database', 'error', updateResult.error);
                    
                    // Check if this is an API key missing error
                    if (updateResult.error && updateResult.error.includes('API key is required')) {
                        showErrorNotification('OpenAI API key is required. Please add your API key in the settings panel.');
                        // Automatically open settings panel to make it easier for the user
                        toggleSettings();
                    } else {
                        showErrorNotification('Error updating summary: ' + updateResult.error);
                    }
                }
            } catch (updateError) {
                log('Exception updating summary in database', 'error', updateError);
                showErrorNotification('Error updating summary: ' + updateError.message);
            }
        } else {
            // Email is not saved yet, show notification that summary will be included when saved
            log('Email is not saved yet, summary will be included when saved', 'info');
            showNotification('Summary generated. It will be saved with the email when you click "Save Email".', 'success');
        }
    } catch (error) {
        handleApiError(error, 'generating summary');
        
        // Check if this is an API key missing error
        if (error.message && error.message.includes('API key')) {
            showErrorNotification('OpenAI API key is required. Please add your API key in the settings panel.');
            // Automatically open settings panel to make it easier for the user
            toggleSettings();
        }
        
        // Clear loading state and show error in UI
        const summaryItems = document.getElementById('summary-items');
        if (summaryItems) {
            summaryItems.innerHTML = '<li>Error generating summary. Please try again.</li>';
        }
    }
}

// Handle getting contact summary
async function handleGetContact() {
    try {
        // Check if user has a team
        if (!checkTeamMembership()) {
            return;
        }
        
        // Check if email data is available
        if (!currentEmailData || !currentEmailData.sender || !currentEmailData.sender.email) {
            showErrorNotification('No contact information available');
            return;
        }
        
        // Show the contact container if it's hidden
        const contactContainer = document.getElementById('contact-container');
        if (contactContainer) {
            contactContainer.style.display = 'block';
        }
        
        // Clear previous items
        const contactItems = document.getElementById('contact-items');
        if (contactItems) {
            contactItems.innerHTML = '<li>Retrieving contact history...</li>';
        }
        
        log('Getting contact history for: ' + currentEmailData.sender.email);
        
        // Call the API to get contact history
        const contactResult = await window.OpsieApi.generateContactHistory(currentEmailData.sender.email);
        
        log('Contact history retrieved successfully', 'info', contactResult);
        
        // Display the contact items
        if (contactItems) {
            contactItems.innerHTML = '';
            
            // Add message count if available
            if (contactResult.messageCount) {
                const countItem = document.createElement('li');
                countItem.style.fontWeight = 'bold';
                countItem.textContent = `Found ${contactResult.messageCount} previous emails with this contact`;
                contactItems.appendChild(countItem);
            }
            
            // Add each summary item
            if (contactResult.summaryItems && contactResult.summaryItems.length > 0) {
                contactResult.summaryItems.forEach(item => {
                    const li = document.createElement('li');
                    // Handle both string items and object items with content property
                    li.textContent = typeof item === 'string' ? item : (item.content || '');
                    contactItems.appendChild(li);
                });
            } else {
                contactItems.innerHTML = '<li>No previous contact history found for this sender</li>';
            }
        }
        
        // Show success notification
        showNotification('Contact history retrieved successfully', 'success');
    } catch (error) {
        handleApiError(error, 'retrieving contact history');
        
        // Clear loading state and show error in UI
        const contactItems = document.getElementById('contact-items');
        if (contactItems) {
            contactItems.innerHTML = '<li>Error retrieving contact history. Please try again.</li>';
        }
    }
}

/**
 * Handle search button click
 */
async function handleSearch() {
    try {
        // Check if user has a team
        if (!checkTeamMembership()) {
            return;
        }
        
        // Get the search query
        const searchInput = document.getElementById('email-search-input');
        const query = searchInput.value.trim();
        
        if (!query) {
            showErrorNotification('Please enter a search query');
            return;
        }
        
        // Set loading state
        setLoading('search', true);
        
        // Show the search results container
        const searchResultsContainer = document.querySelector('.search-results-container');
        if (searchResultsContainer) {
            searchResultsContainer.style.display = 'block';
        }
        
        // Show a placeholder while searching
        const searchAnswerContainer = document.querySelector('.search-answer');
        if (searchAnswerContainer) {
            searchAnswerContainer.textContent = 'Searching...';
        }
        
        // Clear previous references
        const searchReferencesContainer = document.querySelector('.search-references');
        if (searchReferencesContainer) {
            searchReferencesContainer.innerHTML = '';
        }
        
        // Get current email data if available, but don't require it
        let emailData = {};
        try {
            emailData = await getEmailData();
        } catch (emailError) {
            log('Warning: Could not retrieve current email data', 'warn', emailError);
        }
        
        // Get sender's email (from current email or manual entry)
        let senderEmail;
        
        // Try to get sender email from email data
        if (emailData && emailData.sender && emailData.sender.email) {
            senderEmail = emailData.sender.email;
        }
        
        // If no sender email, check if we have a manually selected contact
        if (!senderEmail) {
            const contactSelect = document.getElementById('contact-select');
            if (contactSelect && contactSelect.value) {
                senderEmail = contactSelect.value;
            }
        }
        
        // If still no sender email, show error
        if (!senderEmail) {
            setLoading('search', false);
            showErrorNotification('No sender email available to search');
            
            if (searchAnswerContainer) {
                searchAnswerContainer.textContent = 'Error: No sender email available for searching';
            }
            return;
        }
        
        log('Searching for query:', 'info', query + ' in emails from: ' + senderEmail);
        
        // Search the emails history
        const result = await window.OpsieApi.searchEmailHistory(senderEmail, query);
        
        if (!result.success) {
            setLoading('search', false);
            showErrorNotification(result.error || 'Failed to search emails');
            
            if (searchAnswerContainer) {
                searchAnswerContainer.textContent = 'Error: ' + (result.error || 'Failed to search emails');
            }
            return;
        }
        
        // Show the results
        const { answer, references } = result.data;
        
        if (searchAnswerContainer) {
            searchAnswerContainer.textContent = answer || 'No answer found';
        }
        
        // Clear and update references
        if (searchReferencesContainer) {
            searchReferencesContainer.innerHTML = '';
            
            if (references && references.length > 0) {
                // Create reference items
                references.forEach((ref, index) => {
                    const referenceItem = document.createElement('div');
                    referenceItem.className = 'reference-item';
                    
                    const quoteDiv = document.createElement('div');
                    quoteDiv.className = 'reference-quote';
                    quoteDiv.textContent = `${index + 1}. "${ref.quote}"`;
                    
                    const metaDiv = document.createElement('div');
                    metaDiv.className = 'reference-meta';
                    metaDiv.textContent = ref.meta;
                    
                    referenceItem.appendChild(quoteDiv);
                    referenceItem.appendChild(metaDiv);
                    searchReferencesContainer.appendChild(referenceItem);
                });
            } else {
                // No references found
                const noReferences = document.createElement('div');
                noReferences.className = 'no-references';
                noReferences.textContent = 'No specific references found for this query';
                searchReferencesContainer.appendChild(noReferences);
            }
        }
    } catch (error) {
        log('Error in handle search:', 'error', error);
        showErrorNotification(error.message || 'Failed to search emails');
        
        const searchAnswerContainer = document.querySelector('.search-answer');
        if (searchAnswerContainer) {
            searchAnswerContainer.textContent = 'Error: ' + (error.message || 'An unexpected error occurred');
        }
    } finally {
        setLoading('search', false);
    }
}

// Handle generating reply
async function handleGenerateReply() {
    try {
        // Check if user has a team
        if (!checkTeamMembership()) {
            return;
        }
        
        // Check if email data is available
        if (!currentEmailData || !currentEmailData.body) {
            showErrorNotification('No email data available to generate reply');
            return;
        }
        
        // Get the reply options
        const options = {
            tone: document.getElementById('reply-tone').value,
            length: document.getElementById('reply-length').value,
            language: document.getElementById('reply-language').value,
            additionalContext: document.getElementById('reply-additional-context').value.trim()
        };
        
        // Check if we have Q&A data to include as context
        let qaContext = '';
        if (window.OpsieApi.extractQuestionsAndAnswers && 
            window.OpsieApi.extractQuestionsAndAnswers.questions && 
            window.OpsieApi.extractQuestionsAndAnswers.questions.length > 0) {
            
            const questions = window.OpsieApi.extractQuestionsAndAnswers.questions;
            log('Found Q&A data to include in reply context', 'info', { questionCount: questions.length });
            
            // Format Q&A data for the AI
            const qaItems = questions
                .filter(q => q.answer) // Only include questions that have answers
                .map((q, index) => {
                    let qaItem = `Q${index + 1}: ${q.text}\nA${index + 1}: ${q.answer}`;
                    
                    // Add source information if available
                    if (q.source === 'database' && q.matchType) {
                        if (q.matchType === 'semantic' && q.similarityScore) {
                            qaItem += ` (${Math.round(q.similarityScore * 100)}% match from knowledge base)`;
                        } else if (q.matchType === 'exact') {
                            qaItem += ` (exact match from knowledge base)`;
                        } else if (q.matchType === 'fuzzy') {
                            qaItem += ` (keyword match from knowledge base)`;
                        }
                    } else if (q.source === 'search') {
                        qaItem += ` (found in team emails)`;
                    }
                    
                    // Add verification status
                    if (q.verified || q.isVerified) {
                        qaItem += ` [Verified]`;
                    }
                    
                    return qaItem;
                })
                .slice(0, 5); // Limit to 5 most relevant Q&As to avoid overwhelming the AI
            
            if (qaItems.length > 0) {
                qaContext = `\n\nRelevant Questions & Answers from this email analysis:\n${qaItems.join('\n\n')}`;
                log('Generated Q&A context for reply', 'info', { 
                    contextLength: qaContext.length, 
                    qaCount: qaItems.length 
                });
            }
        }
        
        // Combine additional context with Q&A context
        if (qaContext) {
            if (options.additionalContext) {
                options.additionalContext += qaContext;
            } else {
                options.additionalContext = qaContext.trim();
            }
            log('Enhanced reply context with Q&A data', 'info', { 
                totalContextLength: options.additionalContext.length 
            });
        }
        
        log(`Generating reply with options: ${JSON.stringify({...options, additionalContext: options.additionalContext ? `${options.additionalContext.substring(0, 100)}...` : ''})}`);
        
        // Call the API to generate the reply
        const replyResult = await window.OpsieApi.generateReplySuggestion(currentEmailData, options);
        
        log('Reply generated successfully', 'info', replyResult);
        
        // Display the reply
        const replyContainer = document.getElementById('reply-container');
        const replyPreview = document.getElementById('reply-preview');
        
        if (replyContainer && replyPreview && replyResult.replyText) {
            // Show the reply container
            replyContainer.style.display = 'block';
            
            // Set the reply text
            replyPreview.textContent = replyResult.replyText;
            
            // Scroll to the reply
            replyContainer.scrollIntoView({ behavior: 'smooth' });
        }
        
        // Show success notification with context info
        let successMessage = 'Reply generated successfully!';
        if (qaContext) {
            const qaCount = window.OpsieApi.extractQuestionsAndAnswers.questions.filter(q => q.answer).length;
            successMessage += ` (Enhanced with ${qaCount} Q&A context${qaCount > 1 ? 's' : ''})`;
        }
        showNotification(successMessage, 'success');
    } catch (error) {
        handleApiError(error, 'generating reply');
    }
}

// Handle saving email
async function handleSaveEmail() {
    try {
        // Check if user has a team
        if (!checkTeamMembership()) {
            return;
        }
        
        // Check if email data is available
        if (!currentEmailData || !currentEmailData.body) {
            showErrorNotification('No email data available to save');
            return;
        }
        
        // Check if the email is already saved
        if (currentEmailData.existingMessage && currentEmailData.existingMessage.exists) {
            log('Email is already saved in the database', 'info', currentEmailData.existingMessage);
            
            // Disable the save button visually to match sidebar behavior
            const saveButton = document.getElementById('save-email-button');
            if (saveButton) {
                saveButton.disabled = true;
                saveButton.textContent = 'Already Saved';
                saveButton.style.backgroundColor = '#999';
                log('Disabled Save Email button (already exists check)', 'info');
            }
            
            // Show informational notification
            showNotification('This email was already saved in the database', 'info');
            return;
        }
        
        setLoading('save', true);
        
        log('Saving email to database: ' + currentEmailData.subject);
        
        // Check if we have a message ID
        if (!currentEmailData.messageId) {
            log('No message ID available for this email', 'warning');
            // We'll continue anyway and let the backend generate an ID
        } else {
            log('Using message ID for save: ' + currentEmailData.messageId, 'info');
        }
        
        // Call the API to save the email
        const saveResult = await window.OpsieApi.saveEmail(currentEmailData);
        
        log('Save email result:', 'info', saveResult);
        
        if (saveResult.success) {
            // For both new saves and "already exists" cases, update the button consistently
            // to match sidebar behavior
            const saveButton = document.getElementById('save-email-button');
            if (saveButton) {
                // Force the button to be enabled first to ensure the disabled state change is applied
                saveButton.disabled = false;
                // Apply the disabled state after a short delay to ensure the DOM updates
                setTimeout(() => {
                    saveButton.disabled = true;
                    saveButton.textContent = 'Already Saved';
                    saveButton.style.backgroundColor = '#999';
                    log('Disabled Save Email button after successful save', 'info');
                }, 0);
            } else {
                log('Warning: save-email-button element not found when trying to disable', 'warning');
            }
            
            // If the message was already in the database
            if (saveResult.message && saveResult.message.includes('already exists')) {
                showNotification('This email was already saved in the database', 'info');
            } else {
                showNotification('Email saved successfully!', 'success');
            }
            
            // Update the message details section to show it's been saved
            const savedStatusElement = document.getElementById('saved-status');
            if (savedStatusElement) {
                const timestamp = new Date().toLocaleString();
                
                // Try to get user info from localStorage
                const userData = JSON.parse(localStorage.getItem('userData') || '{}');
                const firstName = userData.first_name || localStorage.getItem('firstName') || '';
                const lastName = userData.last_name || localStorage.getItem('lastName') || '';
                const userName = firstName || lastName ? `${firstName} ${lastName}`.trim() : 'you';
                
                savedStatusElement.innerHTML = `<div class="saved-badge">✓ Saved</div>
                <div class="saved-details">This email was saved to the database by ${userName} at ${timestamp}</div>`;
                savedStatusElement.style.display = 'block';
            }
            
            // If we have a message object in the response, store it
            if (saveResult.data && saveResult.data.message) {
                // Store the message data for future reference
                if (!currentEmailData.existingMessage) {
                    currentEmailData.existingMessage = {};
                }
                currentEmailData.existingMessage.exists = true;
                currentEmailData.existingMessage.message = saveResult.data.message;
                currentEmailData.existingMessage.savedAt = new Date().toISOString();
                currentEmailData.existingMessage.foundBy = 'primary'; // Save was done with primary ID
                
                log('Stored saved message reference', 'info', currentEmailData.existingMessage);
                
                // Show the notes section now that the email is saved
                document.getElementById('notes-section').style.display = 'block';
                
                // Initialize the notes section
                loadNotesForCurrentEmail();
                
                // Update the notes UI state
                updateNotesUIState();
                
                // Enable the "Mark as Handled" button now that the email is saved
                const markHandledButton = document.getElementById('mark-handled-button');
                if (markHandledButton) {
                    markHandledButton.disabled = false;
                    markHandledButton.textContent = 'Mark as Handled';
                    markHandledButton.style.backgroundColor = '#ff9800';
                    log('Enabled Mark as Handled button after save', 'info');
                } else {
                    log('Warning: mark-handled-button element not found', 'warning');
                }
            }
        } else {
            // Failed to save
            showErrorNotification('Failed to save email: ' + (saveResult.error || 'Unknown error'));
            
            // Reset the save button
            const saveButton = document.getElementById('save-email-button');
            if (saveButton) {
                saveButton.disabled = false;
                saveButton.textContent = 'Save Email';
                saveButton.style.backgroundColor = '#4CAF50';
                log('Reset Save Email button after failed save', 'info');
            }
        }
    } catch (error) {
        handleApiError(error, 'saving email');
        
        // Reset the save button
        const saveButton = document.getElementById('save-email-button');
        if (saveButton) {
            saveButton.disabled = false;
            saveButton.textContent = 'Save Email';
            saveButton.style.backgroundColor = '#4CAF50';
            log('Reset Save Email button after save error', 'info');
        }
    } finally {
        setLoading('save', false);
    }
}

// Handle copying reply to clipboard
function handleCopyReply() {
    try {
        const replyPreview = document.getElementById('reply-preview');
        if (!replyPreview || !replyPreview.textContent) {
            showErrorNotification('No reply text to copy');
            return;
        }
        
        // Copy the text to clipboard
        navigator.clipboard.writeText(replyPreview.textContent)
            .then(() => {
                // Show success notification
                showNotification('Reply copied to clipboard!', 'success');
                
                // Visual feedback on the button
                const copyButton = document.getElementById('copy-reply-button');
                if (copyButton) {
                    const originalText = copyButton.textContent;
                    copyButton.textContent = '✓ Copied!';
                    copyButton.style.backgroundColor = '#4CAF50';
                    setTimeout(() => {
                        copyButton.textContent = originalText;
                        copyButton.style.backgroundColor = '';
                    }, 2000);
                }
            })
            .catch(error => {
                console.error('Could not copy text: ', error);
                showErrorNotification('Failed to copy reply. Please select and copy manually.');
            });
    } catch (error) {
        console.error('Error copying reply:', error);
        showErrorNotification('Error copying to clipboard. Please try again.');
    }
}

// Handle inserting reply into email
function handleInsertReply() {
    try {
        const replyPreview = document.getElementById('reply-preview');
        if (!replyPreview || !replyPreview.textContent) {
            showErrorNotification('No reply text to insert');
            return;
        }
        
        // Get the reply text
        const replyText = replyPreview.textContent;
        
        // Insert into the email
        Office.context.mailbox.item.body.setSelectedDataAsync(
            replyText,
            { coercionType: Office.CoercionType.Text },
            result => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('Reply inserted successfully!', 'success');
                } else {
                    console.error('Error inserting reply:', result.error);
                    showErrorNotification('Could not insert reply: ' + result.error.message);
                }
            }
        );
    } catch (error) {
        console.error('Error inserting reply:', error);
        showErrorNotification('Error inserting reply. Please try copying and pasting instead.');
    }
}

// Get the email body as text
function getEmailBody() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(
            Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    console.error('Error getting email body:', result.error);
                    reject(new Error('Failed to get email body'));
                }
            }
        );
    });
}

// Show error notification
function showErrorNotification(message) {
    showNotification(message, 'error');
}

/**
 * Shows a notification to the user
 * @param {string} message - The notification message
 * @param {string} type - The notification type (success, error, warning, info)
 */
function showNotification(message, type = 'info') {
    try {
        if (window.OpsieApi && window.OpsieApi.showNotification) {
            window.OpsieApi.showNotification(message, type);
        } else {
            log('Unable to show notification - OpsieApi.showNotification not available', 'warning');
            console.log(`NOTIFICATION (${type}): ${message}`);
        }
    } catch (error) {
        log('Error showing notification', 'error', error);
    }
}

// Add a window error handler to catch unhandled exceptions
window.onerror = function(message, source, lineno, colno, error) {
    console.error('Unhandled error:', error || message);
    showErrorNotification('An unexpected error occurred. Please try again later.');
    return true; // Prevents default error handling
};

// Add a function to check if Office.js is properly loaded
function checkOfficeAvailability() {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox) {
        showErrorNotification('Office API is not available. Please reload the add-in.');
        return false;
    }
    return true;
}

// Add a function to retry failed operations
function retryOperation(operation, maxRetries = 3, delay = 1000) {
    return new Promise((resolve, reject) => {
        let attempts = 0;
        
        const attempt = () => {
            attempts++;
            operation()
                .then(resolve)
                .catch(error => {
                    if (attempts < maxRetries) {
                        console.log(`Retry attempt ${attempts} after error:`, error);
                        setTimeout(attempt, delay);
                    } else {
                        reject(error);
                    }
                });
        };
        
        attempt();
    });
}

// Add a heartbeat function to periodically check authentication status
function startAuthHeartbeat(interval = 300000) { // 5 minutes
    setInterval(async () => {
        try {
            await checkAuthStatus();
        } catch (error) {
            console.error('Auth heartbeat error:', error);
        }
    }, interval);
}

// Show the handling modal
function showHandlingModal() {
    const modal = document.getElementById('handling-modal');
    if (modal) {
        modal.style.display = 'flex';
        
        // Focus the textarea
        setTimeout(() => {
            const textarea = document.getElementById('handling-note');
            if (textarea) {
                textarea.focus();
            }
        }, 100);
    }
}

// Hide the handling modal
function hideHandlingModal() {
    const modal = document.getElementById('handling-modal');
    if (modal) {
        modal.style.display = 'none';
        
        // Clear the textarea
        const textarea = document.getElementById('handling-note');
        if (textarea) {
            textarea.value = '';
        }
    }
}

// Mark the current email as handled
async function markEmailAsHandled(note) {
    if (!currentEmailData || !currentEmailData.existingMessage || !currentEmailData.existingMessage.message) {
        showNotification('Please save the email first before marking it as handled', 'error');
        return;
    }
    
    try {
        // Get the internal message ID - either directly or from a secondary search match
        const messageId = currentEmailData.existingMessage.message.id;
        const foundBy = currentEmailData.existingMessage.foundBy || 'primary';
        
        log('Marking email as handled', 'info', { 
            messageId: messageId, 
            foundBy: foundBy,
            hasNote: !!note
        });
        
        // Get the user ID from localStorage
        const userData = JSON.parse(localStorage.getItem('userData') || '{}');
        const userId = userData.id || localStorage.getItem('userId');
        if (!userId) {
            showNotification('User ID not found. Please log in again.', 'error');
            return;
        }
        
        // Show loading indicator
        setLoading('handle', true);
        
        // Get the API base URL
        const apiBase = window.OpsieApi.API_BASE_URL || 'https://vewnmfmnvumupdrcraay.supabase.co/rest/v1';
        
        // Fetch current user details from server to ensure we have accurate information
        let userInfo = null;
        try {
            const userResponse = await fetch(`${apiBase}/rpc/get_user_details`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'apikey': window.OpsieApi.SUPABASE_KEY,
                    'Authorization': `Bearer ${localStorage.getItem(window.OpsieApi.STORAGE_KEY_TOKEN)}`
                },
                body: JSON.stringify({
                    user_ids: [userId]
                })
            });
            
            if (userResponse.ok) {
                const users = await userResponse.json();
                if (users && users.length > 0) {
                    const userData = users[0];
                    let userName = 'Unknown User';
                    
                    if (userData.first_name || userData.last_name) {
                        userName = `${userData.first_name || ''} ${userData.last_name || ''}`.trim();
                    } else if (userData.email) {
                        userName = userData.email;
                    }
                    
                    userInfo = {
                        name: userName,
                        email: userData.email
                    };
                    
                    log('Retrieved user information for handling', 'info', userInfo);
                }
            } else {
                log('Failed to fetch user details', 'warning', {
                    status: userResponse.status,
                    statusText: userResponse.statusText
                });
            }
        } catch (userError) {
            log('Error fetching user details for handling', 'error', userError);
        }
        
        // If we couldn't get user details from server, fall back to localStorage
        if (!userInfo) {
            const firstName = localStorage.getItem('firstName') || userData.first_name;
            const lastName = localStorage.getItem('lastName') || userData.last_name;
            const userEmail = localStorage.getItem('userEmail') || userData.email;
            
            let userName = 'Unknown User';
            if (firstName || lastName) {
                userName = `${firstName || ''} ${lastName || ''}`.trim();
            } else if (userEmail) {
                userName = userEmail;
            }
            
            userInfo = {
                name: userName,
                email: userEmail
            };
            
            log('Using localStorage data for user information', 'info', userInfo);
        }
        
        // Send the request to mark the message as handled
        log('Sending PATCH request to mark message as handled', 'info', {
            endpoint: `${apiBase}/messages?id=eq.${messageId}`,
            requestBody: {
                handled_at: new Date().toISOString(),
                handled_by: userId,
                handling_note: note || null
            }
        });
        
        const handleResponse = await fetch(`${apiBase}/messages?id=eq.${messageId}`, {
            method: 'PATCH',
            headers: {
                'Content-Type': 'application/json',
                'apikey': window.OpsieApi.SUPABASE_KEY,
                'Authorization': `Bearer ${localStorage.getItem(window.OpsieApi.STORAGE_KEY_TOKEN)}`
            },
            body: JSON.stringify({
                handled_at: new Date().toISOString(),
                handled_by: userId,
                handling_note: note || null
                // Removed is_handled: true as this column doesn't exist in the database
            })
        });
        
        // Log the response details
        log('PATCH response received', 'info', {
            status: handleResponse.status,
            statusText: handleResponse.statusText,
            ok: handleResponse.ok,
            headers: Array.from(handleResponse.headers.entries())
        });
        
        if (!handleResponse.ok) {
            const responseText = await handleResponse.text();
            log('Error response body', 'error', responseText);
            throw new Error(`Error marking message as handled: ${handleResponse.status} ${handleResponse.statusText}`);
        }
        
        // Update the UI to show that the message has been handled
        log('Email marked as handled successfully', 'info');
        
        // Update handling status in the data model
        if (!currentEmailData.existingMessage.handling) {
            currentEmailData.existingMessage.handling = {};
        }
        
        currentEmailData.existingMessage.handling = {
            handledAt: new Date().toISOString(),
            handledBy: userInfo,
            handlingNote: note || null
        };
        
        // Update the handling status display
        updateHandlingStatus(currentEmailData.existingMessage);
        
        // Disable the Mark as Handled button
        const markHandledButton = document.getElementById('mark-handled-button');
        if (markHandledButton) {
            markHandledButton.disabled = true;
            markHandledButton.textContent = 'Already Handled';
            markHandledButton.style.backgroundColor = '#999';
        }
        
        // If there's a handling note, add it as a regular note too
        if (note && note.trim()) {
            try {
                log('Adding handling note as a regular note', 'info');
                
                // Use the API service function to add the note
                const noteResult = await window.OpsieApi.addNoteToMessage(
                    messageId,
                    userId,
                    note,
                    'Handled'
                );
                
                if (noteResult.success) {
                    log('Handling note added as regular note successfully', 'info');
                } else {
                    log('Failed to add handling note as regular note', 'warning', noteResult.error);
                }
            } catch (noteError) {
                log('Error adding handling note as regular note', 'error', noteError);
            }
        }
        
        // Ensure notes section is visible
        document.getElementById('notes-section').style.display = 'block';
        
        // Reload notes to include the new handling note
        await loadNotesForCurrentEmail();
        
        // Get the latest notes (now including the handling note)
        showNotification('Email marked as handled successfully', 'success');
    } catch (error) {
        log('Error marking email as handled', 'error', error);
        showNotification(`Error: ${error.message}`, 'error');
    } finally {
        setLoading('handle', false);
        hideHandlingModal();
    }
}

// Helper function to fetch user details by user ID
async function fetchUserDetails(userId) {
    try {
        if (!userId) return null;
        
        log('Fetching user details for user ID: ' + userId, 'info');
        
        const response = await window.OpsieApi.apiRequest(`rpc/get_user_details`, 'POST', {
            user_ids: [userId]
        });
        
        if (response.success && response.data && response.data.length > 0) {
            const user = response.data[0];
            const userName = `${user.first_name || ''} ${user.last_name || ''}`.trim() || 'Unknown User';
            log('Successfully fetched user details: ' + userName, 'info');
            return {
                name: userName,
                email: user.email || ''
            };
        }
        
        log('No user details found for user ID: ' + userId, 'warning');
        return null;
    } catch (error) {
        log('Error fetching user details for user ID: ' + userId, 'error', error);
        return null;
    }
}

// Update the handling status display
async function updateHandlingStatus(messageData) {
    try {
        // Get the handling status element
        const handlingStatusElement = document.getElementById('handling-status');
        if (!handlingStatusElement) {
            log('Warning: handling-status element not found', 'warning');
            return;
        }
        
        // Log the message data for debugging
        log('Updating handling status with data', 'info', messageData);
        
        // Check if handling information exists
        if (!messageData || !messageData.handling) {
            log('No handling information available', 'info');
            handlingStatusElement.style.display = 'none';
            return;
        }
        
        // Get the handled by info element
        const handledByInfoElement = document.getElementById('handled-by-info');
        if (!handledByInfoElement) {
            log('Warning: handled-by-info element not found', 'warning');
            return;
        }
        
        // Get the handling note display element
        const handlingNoteDisplayElement = document.getElementById('handling-note-display');
        if (!handlingNoteDisplayElement) {
            log('Warning: handling-note-display element not found', 'warning');
            return;
        }
        
        // Format the handled date
        let handledDate = 'unknown date';
        try {
            if (messageData.handling.handledAt) {
                handledDate = new Date(messageData.handling.handledAt).toLocaleString();
                log('Handling date formatted successfully: ' + handledDate, 'info');
            } else if (messageData.message && messageData.message.handled_at) {
                handledDate = new Date(messageData.message.handled_at).toLocaleString();
                log('Handling date from message.handled_at: ' + handledDate, 'info');
            }
        } catch (e) {
            log('Error formatting handled date: ' + e.message, 'error', e);
        }
        
        // Get who handled it - check multiple possible locations for this information
        let handledByName = 'Unknown user';
        let userInfo = null;
        
        // First try the modern format in the handling property
        if (messageData.handling.handledBy) {
            userInfo = messageData.handling.handledBy;
            if (userInfo.name) {
                handledByName = userInfo.name;
                log('Found handler name in handling.handledBy.name: ' + handledByName, 'info');
            } else if (userInfo.email) {
                handledByName = userInfo.email;
                log('Using email as handler name: ' + handledByName, 'info');
            }
        }
        // Try legacy format if modern format doesn't have a name
        else if (messageData.handling.handledByName) {
            handledByName = messageData.handling.handledByName;
            log('Found handler name in handling.handledByName: ' + handledByName, 'info');
        }
        // If we still don't have a name and we have a user_id, fetch user details
        else if (messageData.message && messageData.message.handled_by) {
            const userId = messageData.message.handled_by;
            log('Found user ID in message.handled_by: ' + userId, 'info');
            
            // Fetch user details from the server
            const userDetails = await fetchUserDetails(userId);
            if (userDetails && userDetails.name) {
                handledByName = userDetails.name;
                log('Successfully fetched handler name: ' + handledByName, 'info');
            } else {
                handledByName = `User ${userId}`;
                log('Could not fetch user details, using fallback name', 'warning');
            }
        }
        
        log('Final handler name: ' + handledByName, 'info');
        
        // Update the handled by info
        handledByInfoElement.textContent = `Marked as handled by ${handledByName} at ${handledDate}`;
        
        // Update the handling note display if a note exists
        const handlingNote = messageData.handling.handlingNote || 
                           (messageData.message && messageData.message.handling_note);
                           
        if (handlingNote) {
            handlingNoteDisplayElement.textContent = `Note: "${handlingNote}"`;
            handlingNoteDisplayElement.style.display = 'block';
            log('Displaying handling note: ' + handlingNote, 'info');
        } else {
            handlingNoteDisplayElement.style.display = 'none';
            log('No handling note to display', 'info');
        }
        
        // Display the handling status
        handlingStatusElement.style.display = 'block';
        
        log('Updated handling status display successfully', 'info');
    } catch (error) {
        log('Error updating handling status display', 'error', error);
    }
}

// Function to set loading state for different sections
function setLoading(section, isLoading) {
    try {
        const saveButton = document.getElementById('save-email-button');
        const handleButton = document.getElementById('mark-handled-button');
        const addNoteButton = document.getElementById('add-note-button');
        
        // Actions specific to notes, save, and handle buttons
        if (section === 'save') {
            if (saveButton) {
                saveButton.disabled = isLoading;
                saveButton.textContent = isLoading ? 'Saving...' : 'Save Email';
            }
            
            const saveLoading = document.getElementById('save-loading');
            if (saveLoading) {
                saveLoading.style.display = isLoading ? 'flex' : 'none';
            }
        } else if (section === 'handle') {
            // Check if email is already handled before updating button state
            const isAlreadyHandled = currentEmailData && 
                                    currentEmailData.existingMessage && 
                                    currentEmailData.existingMessage.handling;
                                    
            if (handleButton) {
                // Only update the button if the email is not already handled
                // or if we're currently in a loading state
                if (isLoading || !isAlreadyHandled) {
                    handleButton.disabled = isLoading;
                    handleButton.textContent = isLoading ? 'Processing...' : 'Mark as Handled';
                    
                    // Reset the background color only if not already handled
                    if (!isLoading && !isAlreadyHandled) {
                        handleButton.style.backgroundColor = '#ff9800'; // Original color
                    }
                }
                
                // If loading is complete and the email is already handled,
                // ensure the button stays in the "handled" state
                if (!isLoading && isAlreadyHandled) {
                    handleButton.disabled = true;
                    handleButton.textContent = 'Already Handled';
                    handleButton.style.backgroundColor = '#999';
                }
            }
        } else if (section === 'notes') {
            if (addNoteButton) {
                addNoteButton.disabled = isLoading;
                addNoteButton.textContent = isLoading ? 'Adding...' : 'Add Note';
            }
        } else if (section === null) {
            // Global loading
            if (saveButton) {
                saveButton.disabled = isLoading;
                saveButton.textContent = isLoading ? 'Processing...' : 'Save Email';
            }
            
            // For the handle button, check if it's already handled
            if (handleButton) {
                const isAlreadyHandled = currentEmailData && 
                                        currentEmailData.existingMessage && 
                                        currentEmailData.existingMessage.handling;
                
                if (!isAlreadyHandled) {
                    handleButton.disabled = isLoading;
                    handleButton.textContent = isLoading ? 'Processing...' : 'Mark as Handled';
                }
            }
            
            if (addNoteButton) {
                addNoteButton.disabled = isLoading;
                addNoteButton.textContent = isLoading ? 'Processing...' : 'Add Note';
            }
        } else {
            // Other specific UI elements
            switch (section) {
                case 'summary':
                    const summaryLoading = document.getElementById('summary-loading');
                    const summaryButton = document.getElementById('generate-summary-button');
                    
                    if (summaryLoading) {
                        summaryLoading.style.display = isLoading ? 'flex' : 'none';
                    }
                    
                    if (summaryButton) {
                        summaryButton.disabled = isLoading;
                        summaryButton.textContent = isLoading ? 'Generating...' : 'Generate AI Summary';
                    }
                    
                    if (isLoading) {
                        // Show the container if we're loading
                        const summaryContainer = document.getElementById('summary-container');
                        if (summaryContainer) {
                            summaryContainer.style.display = 'block';
                        }
                    }
                    break;
                    
                case 'contact':
                    const contactLoading = document.getElementById('contact-loading');
                    const contactButton = document.getElementById('generate-contact-button');
                    
                    if (contactLoading) {
                        contactLoading.style.display = isLoading ? 'flex' : 'none';
                    }
                    
                    if (contactButton) {
                        contactButton.disabled = isLoading;
                        contactButton.textContent = isLoading ? 'Loading...' : 'Get Contact Summaries';
                    }
                    
                    if (isLoading) {
                        // Show the container if we're loading
                        const contactContainer = document.getElementById('contact-container');
                        if (contactContainer) {
                            contactContainer.style.display = 'block';
                        }
                    }
                    break;
                    
                case 'search':
                    const searchLoading = document.getElementById('search-loading');
                    const searchButton = document.getElementById('email-search-button');
                    
                    if (searchLoading) {
                        searchLoading.style.display = isLoading ? 'flex' : 'none';
                    }
                    
                    if (searchButton) {
                        searchButton.disabled = isLoading;
                        searchButton.textContent = isLoading ? 'Searching...' : 'Search';
                    }
                    
                    // Show results container when search is complete, but only if we have results
                    const searchResultsContainer = document.getElementById('search-results-container');
                    if (searchResultsContainer && !isLoading) {
                        const searchAnswer = searchResultsContainer.querySelector('.search-answer');
                        if (searchAnswer && searchAnswer.textContent !== '' && !searchAnswer.classList.contains('placeholder')) {
                            searchResultsContainer.style.display = 'block';
                        }
                    }
                    break;
                    
                case 'reply':
                    const replyLoading = document.getElementById('reply-loading');
                    const replyButton = document.getElementById('btn-generate-reply-now');
                    
                    if (replyLoading) {
                        replyLoading.style.display = isLoading ? 'flex' : 'none';
                    }
                    
                    if (replyButton) {
                        replyButton.disabled = isLoading;
                        replyButton.textContent = isLoading ? 'Generating...' : 'Generate Reply';
                    }
                    break;
                    
                default:
                    console.warn(`Unknown loading section: ${section}`);
            }
        }
    } catch (error) {
        console.error('Error setting loading state:', error);
    }
}

/**
 * Runs when the taskpane is loaded and ready
 */
async function loadTaskpane() {
    try {
        console.log('Loading taskpane...');
        log('Taskpane loading started', 'info');
        
        // Set up event listeners for UI components
        setupEventListeners();
        setupNotesEventListeners(); // Add notes event listeners
        
        // Mark the app as loaded
        isAppLoaded = true;
        
        // Load saved settings from local storage
        await configureWithLocalStorageSettings();
        
        // Immediately check authentication status
        log('Checking authentication status on startup', 'info');
        const isAuthenticated = await window.OpsieApi.isAuthenticated();
        log('Authentication status:', 'info', { authenticated: isAuthenticated });
        
        if (isAuthenticated) {
            // User is authenticated, hide auth UI
            log('User is authenticated, checking team membership', 'info');
            const authErrorContainer = document.getElementById('auth-error-container');
            const authContainer = document.getElementById('auth-container');
            const mainContent = document.getElementById('main-content');
            
            if (authErrorContainer) {
                authErrorContainer.style.display = 'none';
            }
            
            if (authContainer) {
                authContainer.style.display = 'none';
            }
            
            // DON'T show main content yet - wait for team check
            
            // Initialize team and user information and check team membership FIRST
            log('Initializing team and user information', 'info');
            await window.OpsieApi.initTeamAndUserInfo(function(teamInfo) {
                log('Team info initialized on startup', 'info', teamInfo);
                
                // Check if user has a team
                if (!teamInfo || !teamInfo.teamId) {
                    log('User does not have a team, showing team selection view', 'info');
                    
                    // Ensure main content is hidden
                    if (mainContent) {
                        mainContent.style.display = 'none';
                        log('Hidden main content for team selection', 'info');
                    }
                    
                    // Hide all other containers that might interfere
                    if (authContainer) {
                        authContainer.style.display = 'none';
                        log('Hidden auth container for team selection', 'info');
                    }
                    
                    // Show team selection view
                    if (typeof window.showTeamSelectView === 'function') {
                        log('Calling showTeamSelectView function', 'info');
                        window.showTeamSelectView();
                        
                        // Double-check team selection view is visible
                        setTimeout(() => {
                            const teamSelectView = document.getElementById('team-select-view');
                            if (teamSelectView) {
                                log('Team selection view visibility check', 'info', {
                                    display: teamSelectView.style.display,
                                    visible: teamSelectView.offsetParent !== null
                                });
                            } else {
                                log('Team selection view element not found!', 'error');
                            }
                        }, 100);
                    } else {
                        log('showTeamSelectView function not available', 'error');
                        showErrorNotification('Team selection not available. Please refresh the page.');
                    }
                    return;
                }
                
                // User has a team, NOW show main content
                log('User has team, showing main content and loading email', 'info', { teamId: teamInfo.teamId });
            
            if (mainContent) {
                mainContent.style.display = 'block';
            }
                
                // Show the settings button since user has a team
                const settingsButton = document.getElementById('settings-button');
                if (settingsButton) {
                    settingsButton.style.display = 'block';
                    log('Shown settings button - user has team', 'info');
            }
            
            // Make sure all main UI sections are visible
            showMainUISections();
            
                // Load current email
                    loadCurrentEmail();
            });
        } else {
            // User is not authenticated, show auth UI
            log('User is not authenticated, showing login UI', 'info');
            const authErrorContainer = document.getElementById('auth-error-container');
            const authContainer = document.getElementById('auth-container');
            const mainContent = document.getElementById('main-content');
            
            if (authContainer) {
                authContainer.style.display = 'flex';
            }
            
            if (mainContent) {
                mainContent.style.display = 'none';
            }
            
            // Focus the email input field
            const emailInput = document.getElementById('auth-email');
            if (emailInput) {
                setTimeout(() => {
                    emailInput.focus();
                }, 500);
            }
        }
        
        // Start periodic authentication checks
        startAuthCheck();
        
    } catch (error) {
        console.error('Error loading taskpane:', error);
        log('Error loading taskpane:', 'error', error);
        showNotification(
            'Something went wrong while loading the taskpane. Try refreshing or reinstalling the add-in.',
            'error'
        );
    }
}

/**
 * Set up event listeners for notes functionality
 */
function setupNotesEventListeners() {
    window.OpsieApi.log('Setting up notes event listeners', 'info');
    
    // Log all buttons to help with debugging
    const allButtons = document.querySelectorAll('button');
    window.OpsieApi.log('All buttons in the DOM:', 'info', 
        Array.from(allButtons).map(b => ({id: b.id, text: b.textContent, visible: b.offsetParent !== null}))
    );
    
    // Check if notes section exists
    const notesSection = document.getElementById('notes-section');
    window.OpsieApi.log('Notes section element:', 'info', {
        exists: !!notesSection,
        display: notesSection ? notesSection.style.display : 'N/A',
        innerHTML: notesSection ? notesSection.innerHTML.substring(0, 100) + '...' : 'N/A'
    });

    // Force display of notes section for debugging
    if (notesSection) {
        notesSection.style.display = 'block';
        window.OpsieApi.log('Forced notes section to display', 'info');
        
        // Add an alert to confirm this function ran
        setTimeout(() => {
            window.OpsieApi.log('Notes section is now visible', 'info');
        }, 1000);
    }

    // Check if notes form container exists
    const notesFormContainer = document.getElementById('notes-form-container');
    window.OpsieApi.log('Notes form container:', 'info', {
        exists: !!notesFormContainer,
        display: notesFormContainer ? notesFormContainer.style.display : 'N/A'
    });
    
    // Always show notes form for debugging
    if (notesFormContainer) {
        notesFormContainer.style.display = 'block';
        window.OpsieApi.log('Forced notes form container to display', 'info');
    }
    
    // Add debugging attribute to the whole document to check if click events work at all
    document.body.setAttribute('onclick', 'console.log("Document body clicked"); window.OpsieApi.log("Document body clicked", "info");');
    
    window.OpsieApi.log('Notes event listeners setup complete', 'info');
    
    // Execute a test to confirm email data and user data
    setTimeout(() => {
        try {
            window.OpsieApi.log('Testing current email data availability:', 'info', {
                hasEmailData: !!currentEmailData,
                messageId: currentEmailData ? currentEmailData.messageId : 'N/A'
            });
            
            const userDataStr = localStorage.getItem('userData');
            const userData = userDataStr ? JSON.parse(userDataStr) : null;
            
            window.OpsieApi.log('Testing user data availability:', 'info', {
                hasUserData: !!userData,
                userId: userData ? userData.id : 'N/A'
            });
        } catch (error) {
            window.OpsieApi.log('Error checking data availability', 'error', error);
        }
    }, 2000);
}

/**
 * Toggle the display of the notes form
 * Make this global so it can be called from inline HTML
 */
window.toggleNotesForm = function() {
    window.OpsieApi.log('Global toggleNotesForm called', 'info');
    
    try {
        // First check if the email is saved
        const isEmailSaved = currentEmailData && 
                             currentEmailData.existingMessage && 
                             currentEmailData.existingMessage.exists;
        
        if (!isEmailSaved) {
            window.OpsieApi.log('Cannot toggle notes form: Email not saved', 'warning');
            window.OpsieApi.showNotification('You must save the email before adding notes', 'warning');
            return;
        }
        
        // Show a notification instead of alert
        window.OpsieApi.showNotification('Toggling notes form', 'info');
        
        const notesFormContainer = document.getElementById('notes-form-container');
        if (notesFormContainer) {
            const isVisible = notesFormContainer.style.display !== 'none';
            window.OpsieApi.log('Notes form container found', 'info', { 
                currentDisplay: notesFormContainer.style.display,
                isVisible: isVisible,
                newDisplay: isVisible ? 'none' : 'block'
            });
            
            notesFormContainer.style.display = isVisible ? 'none' : 'block';
            
            // Clear form if hiding
            if (isVisible) {
                window.OpsieApi.log('Clearing note form content', 'info');
                document.getElementById('note-body').value = '';
            }
        } else {
            window.OpsieApi.log('Notes form container element not found', 'error');
            window.OpsieApi.showNotification('Notes form container not found!', 'error');
        }
    } catch (error) {
        window.OpsieApi.log('Error in toggleNotesForm', 'error', error);
        window.OpsieApi.showNotification('Error: ' + error.message, 'error');
    }
};

/**
 * Load notes for the current email - global version for direct access
 */
window.loadNotesForCurrentEmail = async function() {
    window.OpsieApi.log('Loading notes for current email', 'info');
    try {
        if (!currentEmailData) {
            window.OpsieApi.log('No email data available for loading notes', 'error');
            window.OpsieApi.showNotification('No email loaded. Cannot fetch notes.', 'error');
            return;
        }
        
        // Check if we have an existing message found in the database
        if (!currentEmailData.existingMessage || !currentEmailData.existingMessage.exists) {
            window.OpsieApi.log('Email is not saved yet, cannot fetch notes', 'warning');
            window.OpsieApi.showNotification('Save the email first to add or view notes.', 'info');
            return;
        }
        
        // Always use the internal message ID from the database when available
        // This is critical for secondary matches where external IDs differ
        let messageId = null;
        
        if (currentEmailData.existingMessage.message && 
            currentEmailData.existingMessage.message.id) {
            // Use the internal database ID if available
            messageId = currentEmailData.existingMessage.message.id;
            window.OpsieApi.log('Using internal database ID for notes lookup', 'info', { 
                internalId: messageId,
                foundBy: currentEmailData.existingMessage.foundBy || 'unknown'
            });
        } else {
            // Fall back to external ID if needed
            messageId = currentEmailData.messageId;
            window.OpsieApi.log('Using external message ID for notes lookup (fallback)', 'info', { 
                externalId: messageId
            });
        }
        
        if (!messageId) {
            window.OpsieApi.log('No message ID available for loading notes', 'error');
            window.OpsieApi.showNotification('Cannot load notes: Message ID not found', 'error');
            return;
        }
        
        window.OpsieApi.log(`Loading notes for message`, 'info', { messageId });
        
        // Use the API service function to get notes
        window.OpsieApi.log('Calling OpsieApi.getMessageNotes', 'info');
        const result = await window.OpsieApi.getMessageNotes(messageId);
        
        window.OpsieApi.log('Result from getMessageNotes', 'info', result);
        
        if (result.success) {
            const notes = result.notes;
            window.OpsieApi.log(`Loaded ${notes.length} notes for message`, 'info', { messageId, notes });
            
            // Store the notes in the global variable
            currentEmailNotes = notes;
            
            // Display the notes in the UI
            window.OpsieApi.log('Displaying notes in UI', 'info');
            window.displayNotes(notes);
            
            // Show the notes section
            window.OpsieApi.log('Making notes section visible', 'info');
            document.getElementById('notes-section').style.display = 'block';
            
            // Start polling for new notes from team members
            startNotesPolling();
            
            window.OpsieApi.showNotification(`Loaded ${notes.length} notes successfully`, 'success');
        } else {
            window.OpsieApi.log('Failed to load notes', 'error', result.error);
            window.OpsieApi.showNotification('Failed to load notes: ' + (result.error || 'Unknown error'), 'error');
        }
    } catch (error) {
        window.OpsieApi.log('Exception in loadNotesForCurrentEmail', 'error', error);
        console.error('Error loading notes:', error);
        window.OpsieApi.showNotification('Error loading notes: ' + error.message, 'error');
    }
};

/**
 * Display notes in the UI - global version for direct access
 */
window.displayNotes = function(notes) {
    window.OpsieApi.log('Global displayNotes function called', 'info', { notesCount: notes ? notes.length : 0 });
    
    try {
        const notesContainer = document.getElementById('notes-container');
        
        if (!notesContainer) {
            window.OpsieApi.log('Notes container element not found', 'error');
            window.OpsieApi.showNotification('Error: Notes container element not found', 'error');
            return;
        }
        
        if (!notes || notes.length === 0) {
            window.OpsieApi.log('No notes to display, showing empty state message', 'info');
            notesContainer.innerHTML = `
                <div class="no-notes-message">
                    No notes for this email yet. Click "Add Note" to create the first note.
                </div>
            `;
            return;
        }
        
        window.OpsieApi.log('Building HTML for notes', 'info', { notes });
        let notesHtml = '';
        
        for (let i = 0; i < notes.length; i++) {
            const note = notes[i];
            window.OpsieApi.log('Processing note', 'info', { noteIndex: i, note });
            
            const noteDate = new Date(note.created_at);
            const formattedDate = noteDate.toLocaleDateString() + ' ' + noteDate.toLocaleTimeString();
            // Map category to proper CSS class
            const category = note.category || 'Other';
            let categoryClass = 'Other';

            if (category === 'Action Required') {
                categoryClass = 'Action';
            } else if (category === 'Pending') {
                categoryClass = 'Pending';
            } else if (category === 'Information') {
                categoryClass = 'Information';
            } else {
                categoryClass = 'Other';
            }
            
            try {
                const noteHtml = `
                    <div class="note-item new-note" data-category="${category}">
                        <div class="note-category ${categoryClass}">${category}</div>
                        <div class="note-text">${window.escapeHtml(note.note_body || note.body || '')}</div>
                        <div class="note-meta">
                            <span class="note-author">${window.escapeHtml(note.user?.name || 'Unknown User')}</span>
                            <span class="note-date">${formattedDate}</span>
                        </div>
                    </div>
                `;
                notesHtml += noteHtml;
                window.OpsieApi.log('Added HTML for note', 'info', { 
                    noteIndex: i, 
                    category: note.category,
                    body: note.note_body || note.body
                });
            } catch (error) {
                window.OpsieApi.log('Error generating HTML for note', 'error', { 
                    noteIndex: i, 
                    note, 
                    error 
                });
            }
        }
        
        window.OpsieApi.log('Setting HTML for notes container', 'info');
        notesContainer.innerHTML = notesHtml;
        window.OpsieApi.log('Finished displaying notes', 'info');
    } catch (error) {
        window.OpsieApi.log('Error in global displayNotes', 'error', error);
        window.OpsieApi.showNotification('Error displaying notes: ' + error.message, 'error');
    }
};

/**
 * Helper function to escape HTML - global version
 */
window.escapeHtml = function(str) {
    if (!str) return '';
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
};

// Add a global function that can be called directly from HTML
window.addTestNoteGlobal = async function() {
    try {
        // Show a notification instead of an alert
        window.OpsieApi.showNotification("Adding test note...", 'info');
        
        window.OpsieApi.log('Global add test note function called', 'info');
        
        if (!currentEmailData) {
            window.OpsieApi.log('No email data available for test note', 'error');
            window.OpsieApi.showNotification('No email loaded. Cannot add test note.', 'error');
            return;
        }

        // Check if the email is saved in the database
        if (!currentEmailData.existingMessage || !currentEmailData.existingMessage.exists) {
            window.OpsieApi.log('Email is not saved yet, cannot add notes', 'warning');
            window.OpsieApi.showNotification('Save the email first to add notes.', 'info');
            return;
        }

        // Always use the internal message ID from the database when available
        let messageId = null;
        
        if (currentEmailData.existingMessage.message && 
            currentEmailData.existingMessage.message.id) {
            // Use the internal database ID if available
            messageId = currentEmailData.existingMessage.message.id;
            window.OpsieApi.log('Using internal database ID for adding note', 'info', { 
                internalId: messageId,
                foundBy: currentEmailData.existingMessage.foundBy || 'unknown'
            });
            } else {
            // Fall back to external ID if needed
            messageId = currentEmailData.messageId;
            window.OpsieApi.log('Using external message ID for adding note (fallback)', 'info', { 
                externalId: messageId
            });
        }
        
        if (!messageId) {
            window.OpsieApi.log('No message ID available for adding note', 'error');
            window.OpsieApi.showNotification('Cannot add note: Message ID not found', 'error');
            return;
            }
            
            try {
            // Get the user ID using our utility function
            const userId = await window.getUserId();
            
            if (!userId) {
                window.OpsieApi.showNotification('Cannot add test note: User ID not found', 'error');
                return;
            }
            
            // Add a test note with timestamp
            const testNoteBody = `Test note created at ${new Date().toLocaleString()}`;
            window.OpsieApi.log('Adding test note', 'info', {
                messageId: messageId,
                userId: userId,
                noteBody: testNoteBody
            });
            
            const result = await window.OpsieApi.addNoteToMessage(
                messageId,
                userId,
                testNoteBody,
                'Information'
            );
            
            window.OpsieApi.log('Test note result:', 'info', result);
            
            if (result.success) {
                window.OpsieApi.showNotification('Test note added successfully!', 'success');
                await window.loadNotesForCurrentEmail();
            } else {
                window.OpsieApi.showNotification(`Failed to add test note: ${result.error}`, 'error');
            }
        } catch (error) {
            window.OpsieApi.log('Error adding test note', 'error', error);
            window.OpsieApi.showNotification('Error adding test note: ' + error.message, 'error');
        }
    } catch (error) {
        window.OpsieApi.log('Exception in global addTestNoteGlobal', 'error', error);
        window.OpsieApi.showNotification('Error: ' + error.message, 'error');
    }
};
    
/**
 * Global function to add a note directly from form
 */
window.directAddNote = async function(noteBody, category) {
    window.OpsieApi.log('Global directAddNote function called', 'info');
    
    try {
        // If parameters not provided, get values from form
        if (!noteBody) {
            noteBody = document.getElementById('note-body').value.trim();
        }
        
        if (!category) {
            category = document.getElementById('note-category').value;
        }
        
        window.OpsieApi.log('Direct note data', 'info', { 
            body: noteBody, 
            category: category 
        });
        
        // Validate input
        if (!noteBody) {
            window.OpsieApi.showNotification('Please enter a note', 'error');
            return;
        }
        
        // Check if we have current email data
        if (!currentEmailData) {
            window.OpsieApi.log('No email data available for adding note', 'error');
            window.OpsieApi.showNotification('No email loaded. Cannot add note.', 'error');
            return;
        }
        
        // Check if the email is saved in the database
        if (!currentEmailData.existingMessage || !currentEmailData.existingMessage.exists) {
            window.OpsieApi.log('Email is not saved yet, cannot add notes', 'warning');
            window.OpsieApi.showNotification('Save the email first to add notes.', 'info');
            return;
        }
        
        // Always use the internal message ID from the database when available
        let messageId = null;
        
        if (currentEmailData.existingMessage.message && 
            currentEmailData.existingMessage.message.id) {
            // Use the internal database ID if available
            messageId = currentEmailData.existingMessage.message.id;
            window.OpsieApi.log('Using internal database ID for adding note', 'info', { 
                internalId: messageId,
                foundBy: currentEmailData.existingMessage.foundBy || 'unknown'
            });
        } else {
            // Fall back to external ID if needed
            messageId = currentEmailData.messageId;
            window.OpsieApi.log('Using external message ID for adding note (fallback)', 'info', { 
                externalId: messageId
            });
        }
        
        if (!messageId) {
            window.OpsieApi.log('No message ID available for adding note', 'error');
            window.OpsieApi.showNotification('Cannot add note: Message ID not found', 'error');
            return;
        }
        
        try {
            // Get the user ID using our utility function
            const userId = await window.getUserId();
            
            if (!userId) {
                window.OpsieApi.showNotification('Cannot add note: User ID not found', 'error');
                return;
            }
            
            window.OpsieApi.log('Adding note with data', 'info', {
                messageId: messageId,
                userId: userId,
                noteBody: noteBody,
                category: category
            });
            
            window.OpsieApi.showNotification('Adding note...', 'info');
            
            // Call the API to add the note
            const result = await window.OpsieApi.addNoteToMessage(
                messageId,
                userId,
                noteBody,
                category
            );
            
            if (result.success) {
                window.OpsieApi.showNotification('Note added successfully', 'success');
                
                // Clear the note body textarea
                document.getElementById('note-body').value = '';
                
                // Toggle the form back to hidden
                window.toggleNotesForm();
                
                // Reload notes to show the new one (this will also restart polling)
                await window.loadNotesForCurrentEmail();
            } else {
                window.OpsieApi.showNotification('Failed to add note: ' + (result.error || 'Unknown error'), 'error');
            }
        } catch (error) {
            window.OpsieApi.log('Error adding note', 'error', error);
            window.OpsieApi.showNotification('Error adding note: ' + error.message, 'error');
        }
    } catch (error) {
        window.OpsieApi.log('Exception in directAddNote', 'error', error);
        window.OpsieApi.showNotification('Error adding note: ' + error.message, 'error');
    }
};

/**
 * Start polling for new notes from other team members
 */
function startNotesPolling() {
    log('Starting notes polling', 'info');
    
    // Don't start polling if already active
    if (isNotesPollingActive) {
        log('Notes polling already active, skipping', 'info');
        return;
    }
    
    // Don't start polling if no email is loaded or saved
    if (!currentEmailData || !currentEmailData.existingMessage || !currentEmailData.existingMessage.exists) {
        log('Cannot start notes polling - no saved email loaded', 'info');
        return;
    }
    
    isNotesPollingActive = true;
    lastNotesCheck = new Date();
    
    // Poll every 10 seconds for new notes
    notesPollingInterval = setInterval(async () => {
        await pollForNewNotes();
    }, 10000);
    
    log('Notes polling started - checking every 10 seconds', 'info');
}

/**
 * Stop notes polling
 */
function stopNotesPolling() {
    log('Stopping notes polling', 'info');
    
    if (notesPollingInterval) {
        clearInterval(notesPollingInterval);
        notesPollingInterval = null;
    }
    
    isNotesPollingActive = false;
    log('Notes polling stopped', 'info');
}

/**
 * Poll for new notes and update UI if changes detected
 */
async function pollForNewNotes() {
    try {
        // Don't poll if no email is loaded
        if (!currentEmailData || !currentEmailData.existingMessage || !currentEmailData.existingMessage.exists) {
            log('Stopping notes polling - no saved email loaded', 'info');
            stopNotesPolling();
            return;
        }
        
        // Get the message ID
        let messageId = null;
        if (currentEmailData.existingMessage.message && currentEmailData.existingMessage.message.id) {
            messageId = currentEmailData.existingMessage.message.id;
        } else {
            messageId = currentEmailData.messageId;
        }
        
        if (!messageId) {
            log('Cannot poll for notes - no message ID available', 'warning');
            return;
        }
        
        log('Polling for new notes', 'info', { messageId });
        
        // Fetch latest notes from API
        const result = await window.OpsieApi.getMessageNotes(messageId);
        
        if (!result.success) {
            log('Failed to poll for notes', 'error', result.error);
            return;
        }
        
        const latestNotes = result.notes || [];
        
        // Compare with current notes to detect changes
        const hasChanges = detectNotesChanges(currentEmailNotes, latestNotes);
        
        if (hasChanges) {
            log(`Detected notes changes - current: ${currentEmailNotes.length}, latest: ${latestNotes.length}`, 'info');
            
            // Show notification about new notes
            const newNotesCount = latestNotes.length - currentEmailNotes.length;
            if (newNotesCount > 0) {
                window.OpsieApi.showNotification(
                    `${newNotesCount} new note${newNotesCount === 1 ? '' : 's'} from team members`, 
                    'info'
                );
            }
            
            // Update the UI with new notes (highlight new ones) - pass old notes for comparison
            displayNotesWithNewIndicator(latestNotes, currentEmailNotes);
            
            // Update the stored notes after displaying
            currentEmailNotes = latestNotes;
        } else {
            log('No new notes detected during polling', 'info');
        }
        
    } catch (error) {
        log('Error polling for notes', 'error', error);
    }
}

/**
 * Detect if there are changes in notes (new notes added)
 */
function detectNotesChanges(oldNotes, newNotes) {
    if (!oldNotes || !newNotes) {
        return newNotes && newNotes.length > 0;
    }
    
    // Check if count changed
    if (oldNotes.length !== newNotes.length) {
        return true;
    }
    
    // Check if any note IDs are different (new notes added)
    const oldNoteIds = new Set(oldNotes.map(note => note.id));
    const newNoteIds = new Set(newNotes.map(note => note.id));
    
    // Check if there are new note IDs
    for (const id of newNoteIds) {
        if (!oldNoteIds.has(id)) {
            return true;
        }
    }
    
    return false;
}

/**
 * Display notes with indicators for new notes
 */
function displayNotesWithNewIndicator(notes, oldNotes = null) {
    log('Displaying notes with new indicators', 'info', { notesCount: notes ? notes.length : 0 });
    
    try {
        const notesContainer = document.getElementById('notes-container');
        
        if (!notesContainer) {
            log('Notes container element not found', 'error');
            return;
        }
        
        if (!notes || notes.length === 0) {
            log('No notes to display, showing empty state message', 'info');
            notesContainer.innerHTML = `
                <div class="no-notes-message">
                    No notes for this email yet. Click "Add Note" to create the first note.
                </div>
            `;
            return;
        }
        
        log('Building HTML for notes with new indicators', 'info');
        let notesHtml = '';
        
        // Get the set of old note IDs to identify new notes
        const oldNoteIds = new Set((oldNotes || []).map(note => note.id));
        
        for (let i = 0; i < notes.length; i++) {
            const note = notes[i];
            const isNewNote = !oldNoteIds.has(note.id);
            
            const noteDate = new Date(note.created_at);
            const formattedDate = noteDate.toLocaleDateString() + ' ' + noteDate.toLocaleTimeString();
            
            // Map category to proper CSS class
            const category = note.category || 'Other';
            let categoryClass = 'Other';

            if (category === 'Action Required') {
                categoryClass = 'Action';
            } else if (category === 'Pending') {
                categoryClass = 'Pending';
            } else if (category === 'Information') {
                categoryClass = 'Information';
            } else {
                categoryClass = 'Other';
            }
            
            try {
                const noteHtml = `
                    <div class="note-item ${isNewNote ? 'new-note-highlight' : ''}" data-category="${category}" data-note-id="${note.id}">
                        ${isNewNote ? '<div class="new-note-badge">NEW</div>' : ''}
                        <div class="note-category ${categoryClass}">${category}</div>
                        <div class="note-text">${window.escapeHtml(note.note_body || note.body || '')}</div>
                        <div class="note-meta">
                            <span class="note-author">${window.escapeHtml(note.user?.name || 'Unknown User')}</span>
                            <span class="note-date">${formattedDate}</span>
                        </div>
                    </div>
                `;
                notesHtml += noteHtml;
                
                log('Added HTML for note', 'info', { 
                    noteIndex: i, 
                    category: note.category,
                    isNew: isNewNote
                });
            } catch (error) {
                log('Error generating HTML for note', 'error', { 
                    noteIndex: i, 
                    note, 
                    error 
                });
            }
        }
        
        log('Setting HTML for notes container with new indicators', 'info');
        notesContainer.innerHTML = notesHtml;
        
        // Remove new note highlights after 10 seconds
        setTimeout(() => {
            const newNoteElements = notesContainer.querySelectorAll('.new-note-highlight');
            newNoteElements.forEach(element => {
                element.classList.remove('new-note-highlight');
                const badge = element.querySelector('.new-note-badge');
                if (badge) {
                    badge.remove();
                }
            });
        }, 10000);
        
        log('Finished displaying notes with new indicators', 'info');
    } catch (error) {
        log('Error in displayNotesWithNewIndicator', 'error', error);
        window.OpsieApi.showNotification('Error displaying notes: ' + error.message, 'error');
    }
}
    
/**
 * Global utility function to get the current user ID
 * This centralizes the user ID retrieval logic to avoid repetition
 */
window.getUserId = async function() {
    try {
        window.OpsieApi.log('Getting user ID', 'info');
        
        // First try to get from localStorage
        let userId = localStorage.getItem('userId');
        
        if (userId) {
            window.OpsieApi.log('Found userId in localStorage', 'info', userId);
            return userId;
        }
        
        // Try to get from Office identity
        return new Promise((resolve, reject) => {
            Office.context.mailbox.getUserIdentityTokenAsync(async function(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    try {
                        // Try to get the team info which should have userId
                        const teamInfoResult = await window.OpsieApi.initTeamAndUserInfo();
                        
                        if (teamInfoResult && typeof teamInfoResult === 'object') {
                            // New return format: object with userId property
                            if (teamInfoResult.userId) {
                                window.OpsieApi.log('Retrieved userId from team info result', 'info', teamInfoResult.userId);
                                resolve(teamInfoResult.userId);
                                return;
                            }
                        } else if (teamInfoResult === true) {
                            // Old format (boolean true)
                            // Try to get userId from localStorage again since initTeamAndUserInfo might have set it
                            userId = localStorage.getItem('userId');
                            if (userId) {
                                window.OpsieApi.log('Found userId in localStorage after initTeamAndUserInfo', 'info', userId);
                                resolve(userId);
                                return;
                            }
                        }
                        
                        // If we still don't have a user ID, try to get it directly
                        window.OpsieApi.log('User ID not found in team info, trying direct user info fetch', 'warning');
                        
                        // Try the new getUserInfo function as a last resort
                        const userInfo = await window.OpsieApi.getUserInfo();
                        if (userInfo && userInfo.id) {
                            window.OpsieApi.log('Retrieved userId from direct API call', 'info', userInfo.id);
                            // Save for future use
                            localStorage.setItem('userId', userInfo.id);
                            resolve(userInfo.id);
                            return;
                        }
                        
                        window.OpsieApi.log('User ID not found in any source', 'error');
                        reject(new Error('User ID not found in any available source'));
                    } catch (error) {
                        window.OpsieApi.log('Error getting team or user info', 'error', error);
                        reject(error);
                    }
                } else {
                    window.OpsieApi.log('Failed to get user identity token', 'error', result.error);
                    reject(new Error('Failed to get user identity token: ' + result.error.message));
                }
            });
        });
    } catch (error) {
        window.OpsieApi.log('Error in getUserId', 'error', error);
        throw error;
    }
};
    
// Add a helper function to display the current message ID information
window.checkCurrentMessageId = function() {
    try {
        window.OpsieApi.log('Checking current message ID details', 'info');
        
        if (!currentEmailData) {
            window.OpsieApi.showNotification('No email data available', 'warning');
            return;
        }
        
        // Determine what ID would be used for note operations
        let messageId = currentEmailData.messageId;
        let idSource = 'external (Office.js)';
        let foundBy = 'not found in database';
        
        // Check if we have internal ID from a database match
        if (currentEmailData.existingMessage && 
            currentEmailData.existingMessage.exists && 
            currentEmailData.existingMessage.message && 
            currentEmailData.existingMessage.message.id) {
            
            if (currentEmailData.existingMessage.foundBy === 'secondary') {
                const internalId = currentEmailData.existingMessage.message.id;
                messageId = internalId;
                idSource = 'internal (database)';
                foundBy = 'secondary match (sender+timestamp)';
            } else {
                foundBy = 'primary match (external ID)';
            }
        }
        
        // Log the details
        window.OpsieApi.log('Message ID details', 'info', {
            currentId: messageId,
            source: idSource,
            foundBy: foundBy,
            externalId: currentEmailData.messageId,
            databaseRecord: currentEmailData.existingMessage || 'none'
        });
        
        // Show a notification with key information
        window.OpsieApi.showNotification(
            `Current message ID: ${messageId.substring(0, 12)}... (${idSource}) - Found by: ${foundBy}`, 
            'info'
        );
        
        return { messageId, idSource, foundBy };
    } catch (error) {
        window.OpsieApi.log('Error checking message ID', 'error', error);
        window.OpsieApi.showNotification('Error checking message ID: ' + error.message, 'error');
    }
};
    
// Function to authenticate with Supabase
async function loginWithSupabase(email, password) {
    try {
        log('Attempting to login with Supabase', 'info');
        
        // Call the login function from the API service
        const result = await window.OpsieApi.login(email, password);
        
        log('Login result received from API service', 'info', {
            success: result.success,
            error: result.error ? 'Error present' : 'No error'
        });
        
        return result;
    } catch (error) {
        log('Error in loginWithSupabase function', 'error', error);
        return {
            success: false,
            error: {
                message: error.message || 'An unexpected error occurred during login'
            }
        };
    }
}

// Function to start periodic authentication checks
function startAuthCheck() {
    // Check immediately
    checkAuthStatus();
    
    // Then check every 5 minutes (300000 ms)
    setInterval(checkAuthStatus, 300000);
    
    log('Started periodic authentication checks', 'info');
}

// Show settings panel with team management
async function showSettings() {
    const settingsContainer = document.getElementById('settings-container');
    
    // Check if user has a team before allowing access to settings
    const teamId = localStorage.getItem('currentTeamId');
    if (!teamId) {
        log('User attempted to access settings without a team', 'warning');
        showErrorNotification('You must be part of a team to access settings');
        
        // Show team selection view instead
        if (typeof window.showTeamSelectView === 'function') {
            window.showTeamSelectView();
        }
        return;
    }
    
    // If settings is already visible, close it
    if (settingsContainer && settingsContainer.style.display === 'block') {
        log('Closing settings panel via settings button', 'info');
        
        // Hide the settings container
        settingsContainer.style.display = 'none';
        
        // Restore saved display states
        try {
            const savedDisplayStates = JSON.parse(localStorage.getItem('sectionDisplayStates') || '{}');
            for (const [id, displayState] of Object.entries(savedDisplayStates)) {
                if (id === 'settings-container') continue;
                
                const element = document.getElementById(id);
                if (element) {
                    element.style.display = displayState;
                    log(`Restored saved display state for ${id}: ${displayState}`, 'info');
                }
            }
        } catch (error) {
            log('Error restoring saved display states: ' + error.message, 'error');
        }
        
        // Show all main UI sections
        showMainUISections();
        return;
    }
    
    // If not visible, show settings
    log('Opening settings panel', 'info');
    
    // Hide all sections first
    hideAllSections();
    
    // Load current user and team information with force refresh
    loadUserSettingsInfo(true);
    
    // Show settings container
    if (settingsContainer) {
        settingsContainer.style.display = 'block';
    }
}

// Load user and team information for settings panel
async function loadUserSettingsInfo(forceRefresh = false) {
    try {
        log('Loading user settings information', 'info', { forceRefresh });
        
        // Get current user information
        const userData = await window.OpsieApi.getUserInfo(forceRefresh);
        log('User info for settings', 'info', userData);
        
        if (userData) {
            // Update user information display
            const userEmailEl = document.getElementById('settings-user-email');
            const teamNameEl = document.getElementById('settings-team-name');
            const userRoleEl = document.getElementById('settings-user-role');
            
            if (userEmailEl) {
                userEmailEl.textContent = userData.email || '-';
            }
            
            if (teamNameEl) {
                // For team name, use stored value or format from team_id
                const teamName = userData.teamName || 
                                localStorage.getItem('currentTeamName') || 
                                (userData.team_id ? `Team ${userData.team_id.substring(0, 6)}...` : '-');
                teamNameEl.textContent = teamName;
            }
            
            if (userRoleEl) {
                userRoleEl.textContent = userData.role || 'member';
            }
            
            // Load team details if user has a team
            if (userData.team_id) {
                await loadTeamDetails(userData.team_id, forceRefresh);
                
                // Set up team management controls based on user role
                setupTeamControls(userData.role);
            } else {
                log('User does not have a team', 'info');
                hideTeamSections();
            }
        } else {
            log('Failed to load user information', 'error');
            showErrorNotification('Failed to load user information');
        }
    } catch (error) {
        log('Error loading user settings information', 'error', error);
        showErrorNotification('Error loading settings information');
    }
}

// Load team details for the settings panel
async function loadTeamDetails(teamId, forceRefresh = false) {
    try {
        log('Loading team details', 'info', { teamId, forceRefresh });
        
        const teamDetailsResult = await window.OpsieApi.getTeamDetails(teamId, forceRefresh);
        log('Team details result', 'info', teamDetailsResult);
        
        if (teamDetailsResult && teamDetailsResult.success) {
            const teamData = teamDetailsResult.data;
            
            // Update team details in the display view
            const organizationEl = document.getElementById('team-organization');
            const invoiceEmailEl = document.getElementById('team-invoice-email');
            const billingAddressEl = document.getElementById('team-billing-address');
            const accessCodeEl = document.getElementById('team-access-code');
            
            if (organizationEl) {
                organizationEl.textContent = teamData.organization || '-';
            }
            
            if (invoiceEmailEl) {
                invoiceEmailEl.textContent = teamData.invoice_email || '-';
            }
            
            if (billingAddressEl) {
                // Format billing address
                const addressParts = [];
                if (teamData.billing_street) addressParts.push(teamData.billing_street);
                if (teamData.billing_city) addressParts.push(teamData.billing_city);
                if (teamData.billing_region) addressParts.push(teamData.billing_region);
                if (teamData.billing_country) addressParts.push(teamData.billing_country);
                
                billingAddressEl.textContent = addressParts.length > 0 ? addressParts.join(', ') : '-';
            }
            
            if (accessCodeEl) {
                accessCodeEl.textContent = teamData.access_code || '-';
            }
            
            // Also set up the edit form with current values
            document.getElementById('edit-team-organization').value = teamData.organization || '';
            document.getElementById('edit-team-invoice-email').value = teamData.invoice_email || '';
            document.getElementById('edit-team-billing-street').value = teamData.billing_street || '';
            document.getElementById('edit-team-billing-city').value = teamData.billing_city || '';
            document.getElementById('edit-team-billing-region').value = teamData.billing_region || '';
            document.getElementById('edit-team-billing-country').value = teamData.billing_country || '';
            
            // Load team members
            await loadTeamMembers(teamId, forceRefresh);
        } else {
            log('Failed to load team details', 'error');
            showErrorNotification('Failed to load team details');
        }
    } catch (error) {
        log('Error loading team details', 'error', error);
        showErrorNotification('Error loading team details');
    }
}

/**
 * Load team members for the specified team
 * @param {string} teamId - The team ID to load members for
 * @param {boolean} forceRefresh - Whether to force refresh from the API instead of cache
 * @returns {Promise<void>}
 */
async function loadTeamMembers(teamId, forceRefresh = false) {
    try {
        log('Loading team members', 'info', { teamId, forceRefresh });
        
        if (!teamId) {
            log('No team ID provided, cannot load team members', 'warning');
            return;
        }
        
        // Get team members from the API, passing the forceRefresh parameter
        const membersResult = await window.OpsieApi.getTeamMembers(teamId, forceRefresh);
        
        if (membersResult.success && membersResult.data) {
            const teamMembersList = document.getElementById('team-members-list');
            if (!teamMembersList) {
                log('Team members list element not found', 'warning');
                return;
            }
            
            // Clear the current list
            teamMembersList.innerHTML = '';
            
            // Get current user ID
            const userData = await window.OpsieApi.getUserInfo();
            const currentUserId = userData.id;
            
            // Sort members: admins first, then alphabetically by name
            const sortedMembers = [...membersResult.data].sort((a, b) => {
                // Sort by role first (admin before member)
                if (a.role === 'admin' && b.role !== 'admin') return -1;
                if (a.role !== 'admin' && b.role === 'admin') return 1;
                
                // Then sort by name
                const nameA = `${a.first_name || ''} ${a.last_name || ''}`.trim();
                const nameB = `${b.first_name || ''} ${b.last_name || ''}`.trim();
                
                // If names are equal or empty, sort by email
                if (!nameA || !nameB || nameA === nameB) {
                    return a.email.localeCompare(b.email);
                }
                
                return nameA.localeCompare(nameB);
            });
            
            // Add each member to the list
            sortedMembers.forEach(member => {
                const memberItem = document.createElement('div');
                memberItem.className = 'team-member-item';
                
                // If this is the current user, add a special class
                if (member.id === currentUserId) {
                    memberItem.classList.add('current-user');
                }
                
                // Format name (either first+last name or just email if no name)
                const memberName = `${member.first_name || ''} ${member.last_name || ''}`.trim();
                const displayName = memberName || member.email;
                
                // Format role with special styling
                const roleClass = member.role === 'admin' ? 'role-admin' : 'role-member';
                
                memberItem.innerHTML = `
                    <div class="member-info">
                        <span class="member-name">${displayName}</span>
                        <span class="team-member-role ${roleClass}">${member.role}</span>
                        ${member.id === currentUserId ? '<span class="current-user-indicator">(You)</span>' : ''}
                    </div>
                    ${userData.role === 'admin' && member.id !== currentUserId ? 
                        `<div class="team-member-actions">
                            <button class="team-member-remove-btn" data-member-id="${member.id}" data-member-email="${member.email}">
                                Remove
                            </button>
                        </div>` : ''
                    }
                `;
                
                teamMembersList.appendChild(memberItem);
            });
            
            // Add event listeners to remove buttons
            const removeButtons = document.querySelectorAll('.team-member-remove-btn');
            removeButtons.forEach(button => {
                button.addEventListener('click', () => {
                    const memberId = button.dataset.memberId;
                    const memberEmail = button.dataset.memberEmail;
                    handleRemoveTeamMember(memberId, memberEmail);
                });
            });
            
            // Also populate the team members dropdown for admin transfer
            const teamMembersSelect = document.getElementById('team-members-select');
            if (teamMembersSelect && userData.role === 'admin') {
                // Clear existing options
                teamMembersSelect.innerHTML = '<option value="">Select a team member...</option>';
                
                // Add team members to the dropdown (exclude current user and other admins)
                sortedMembers.forEach(member => {
                    // Only add non-admin members (excluding current user)
                    if (member.id !== currentUserId && member.role !== 'admin') {
                        const option = document.createElement('option');
                        option.value = member.id;
                        
                        // Format name (either first+last name or just email if no name)
                        const memberName = `${member.first_name || ''} ${member.last_name || ''}`.trim();
                        const displayName = memberName || member.email;
                        
                        option.textContent = `${displayName} (${member.email})`;
                        teamMembersSelect.appendChild(option);
                    }
                });
                
                log('Populated team members dropdown', 'info', { 
                    totalMembers: sortedMembers.length,
                    eligibleForTransfer: sortedMembers.filter(m => m.id !== currentUserId && m.role !== 'admin').length
                });
            }
            
        } else {
            log('Failed to load team members', 'error', membersResult.error);
        }
    } catch (error) {
        log('Error loading team members', 'error', error);
    }
}

// Load pending join requests for admins
async function loadPendingRequests(teamId) {
    try {
        log('=== LOAD PENDING REQUESTS CALLED ===', 'info', { teamId });
        log('Loading pending join requests', 'info', { teamId });
        
        const requestsResult = await window.OpsieApi.getTeamJoinRequests(teamId);
        log('Join requests result', 'info', requestsResult);
        
        if (requestsResult && requestsResult.success) {
            const requests = requestsResult.data;
            const requestsList = document.getElementById('join-requests-list');
            const noRequestsMsg = document.getElementById('no-requests-message');
            
            if (requestsList) {
                // Clear existing requests (except the no requests message)
                if (noRequestsMsg) {
                    requestsList.innerHTML = '';
                    requestsList.appendChild(noRequestsMsg);
                } else {
                    requestsList.innerHTML = '<p id="no-requests-message" style="font-style: italic; color: #666; margin: 5px;">No pending join requests</p>';
                }
                
                if (requests && requests.length > 0) {
                    // Hide the no requests message
                    if (noRequestsMsg) {
                        noRequestsMsg.style.display = 'none';
                    }
                    
                    // Add each request to the list
                    requests.forEach(request => {
                        const requestItem = document.createElement('div');
                        requestItem.className = 'request-item';
                        
                        // Create user name from the correct field names
                        const userName = `${request.user_first_name || ''} ${request.user_last_name || ''}`.trim() || request.user_email || 'Unknown User';
                        
                        // Create the main request info with better styling
                        const requestInfo = document.createElement('div');
                        requestInfo.className = 'request-info';
                        requestInfo.innerHTML = `
                            <div class="request-user">
                                <strong>${userName}</strong>
                                <span class="request-email">(${request.user_email})</span>
                            </div>
                            <div class="request-message">has requested to join your team</div>
                        `;
                        
                        const requestTime = document.createElement('div');
                        requestTime.className = 'request-time';
                        
                        // Try different date field names and format properly
                        let requestDate = null;
                        const dateValue = request.request_date || request.created_at || request.date_created;
                        
                        if (dateValue) {
                            requestDate = new Date(dateValue);
                            // Check if date is valid
                            if (!isNaN(requestDate.getTime())) {
                                const now = new Date();
                                const diffTime = Math.abs(now - requestDate);
                                const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
                                
                                if (diffDays === 0) {
                                    requestTime.textContent = `Requested today at ${requestDate.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}`;
                                } else if (diffDays === 1) {
                                    requestTime.textContent = `Requested yesterday at ${requestDate.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}`;
                                } else if (diffDays < 7) {
                                    requestTime.textContent = `Requested ${diffDays} days ago`;
                                } else {
                                    requestTime.textContent = `Requested on ${requestDate.toLocaleDateString()}`;
                                }
                            } else {
                                requestTime.textContent = 'Request date unavailable';
                            }
                        } else {
                            requestTime.textContent = 'Request date unavailable';
                        }
                        
                        const requestActions = document.createElement('div');
                        requestActions.className = 'request-actions';
                        
                        const approveBtn = document.createElement('button');
                        approveBtn.className = 'request-approve-btn';
                        approveBtn.textContent = 'Approve';
                        approveBtn.onclick = () => handleRequestResponse(request.id, true);
                        
                        const denyBtn = document.createElement('button');
                        denyBtn.className = 'request-deny-btn';
                        denyBtn.textContent = 'Deny';
                        denyBtn.onclick = () => handleRequestResponse(request.id, false);
                        
                        requestActions.appendChild(approveBtn);
                        requestActions.appendChild(denyBtn);
                        
                        requestItem.appendChild(requestInfo);
                        requestItem.appendChild(requestTime);
                        requestItem.appendChild(requestActions);
                        
                        requestsList.appendChild(requestItem);
                    });
                } else {
                    // Show the no requests message
                    if (noRequestsMsg) {
                        noRequestsMsg.style.display = 'block';
                    }
                }
            }
        } else {
            log('Failed to load join requests', 'error');
            showErrorNotification('Failed to load join requests');
        }
    } catch (error) {
        log('Error loading join requests', 'error', error);
        showErrorNotification('Error loading join requests');
    }
}

// Set up team management controls based on user role
function setupTeamControls(userRole) {
    log('Setting up team controls for role', 'info', { userRole });
    
    const teamDetailsSection = document.getElementById('team-details-section');
    const joinRequestsSection = document.getElementById('join-requests-section');
    const adminControls = document.getElementById('admin-controls');
    const memberControls = document.getElementById('member-controls');
    const editTeamDetailsButton = document.getElementById('edit-team-details-button');
    
    if (userRole === 'admin') {
        // User is admin, show admin controls
        if (teamDetailsSection) teamDetailsSection.style.display = 'block';
        if (joinRequestsSection) joinRequestsSection.style.display = 'block';
        if (adminControls) adminControls.style.display = 'block';
        if (memberControls) memberControls.style.display = 'none';
        if (editTeamDetailsButton) editTeamDetailsButton.style.display = 'inline-block';
        
        // For admins, also load pending join requests
        const userInfo = window.OpsieApi.getUserInfo();
        if (userInfo && userInfo.team_id) {
            loadPendingRequests(userInfo.team_id);
        }
    } else {
        // User is a regular member
        if (teamDetailsSection) teamDetailsSection.style.display = 'block';
        if (joinRequestsSection) joinRequestsSection.style.display = 'none';
        if (adminControls) adminControls.style.display = 'none';
        if (memberControls) memberControls.style.display = 'block';
        if (editTeamDetailsButton) editTeamDetailsButton.style.display = 'none';
    }
}

// Hide team-related sections in settings
function hideTeamSections() {
    const teamDetailsSection = document.getElementById('team-details-section');
    const joinRequestsSection = document.getElementById('join-requests-section');
    const adminControls = document.getElementById('admin-controls');
    const memberControls = document.getElementById('member-controls');
    
    if (teamDetailsSection) teamDetailsSection.style.display = 'none';
    if (joinRequestsSection) joinRequestsSection.style.display = 'none';
    if (adminControls) adminControls.style.display = 'none';
    if (memberControls) memberControls.style.display = 'none';
}

// Hide all content sections
function hideAllSections() {
    // First, save the current display state of sections with IDs
    try {
        const displayStates = {};
        const sections = document.querySelectorAll('.section');
        sections.forEach(section => {
            if (section.id) {
                displayStates[section.id] = section.style.display || 'block';
            }
        });
        localStorage.setItem('sectionDisplayStates', JSON.stringify(displayStates));
        log('Saved display states before hiding sections', 'info', displayStates);
    } catch (error) {
        log('Error saving section display states: ' + error.message, 'error');
    }
    
    // Now hide all sections
    const sections = document.querySelectorAll('.section');
    sections.forEach(section => {
        section.style.display = 'none';
    });
}

// Custom confirmation dialog for Office add-ins (since window.confirm is not supported)
function showCustomConfirm(message, title = 'Confirm Action', okText = 'OK', cancelText = 'Cancel') {
    return new Promise((resolve) => {
        const modal = document.getElementById('custom-modal-backdrop');
        const modalTitle = document.getElementById('custom-modal-title');
        const modalBody = document.getElementById('custom-modal-body');
        const modalInput = document.getElementById('custom-modal-input');
        const okButton = document.getElementById('custom-modal-ok');
        const cancelButton = document.getElementById('custom-modal-cancel');
        const closeButton = document.getElementById('custom-modal-close');
        
        // Set up the modal
        modalTitle.textContent = title;
        modalBody.innerHTML = `<p>${message}</p>`;
        modalInput.style.display = 'none'; // Hide input for simple confirmations
        modal.style.display = 'flex';
        
        // Clean up any existing event listeners
        const newOkButton = okButton.cloneNode(true);
        const newCancelButton = cancelButton.cloneNode(true);
        const newCloseButton = closeButton.cloneNode(true);
        okButton.parentNode.replaceChild(newOkButton, okButton);
        cancelButton.parentNode.replaceChild(newCancelButton, cancelButton);
        closeButton.parentNode.replaceChild(newCloseButton, closeButton);
        
        // Set custom button text
        newOkButton.textContent = okText;
        newCancelButton.textContent = cancelText;
        
        // Set up event listeners
        newOkButton.addEventListener('click', () => {
            modal.style.display = 'none';
            resolve(true);
        });
        
        newCancelButton.addEventListener('click', () => {
            modal.style.display = 'none';
            resolve(false);
        });
        
        newCloseButton.addEventListener('click', () => {
            modal.style.display = 'none';
            resolve(false);
        });
        
        // Close on backdrop click
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.style.display = 'none';
                resolve(false);
            }
        });
    });
}

// Handler for removing a team member
async function handleRemoveTeamMember(memberId, memberEmail) {
    try {
        const confirmed = await showCustomConfirm(
            `Are you sure you want to remove ${memberEmail} from the team?`,
            'Remove Team Member',
            'Remove',
            'Cancel'
        );
        
        if (!confirmed) {
            return;
        }
        
        log('Removing team member', 'info', { memberId, memberEmail });
        
        const result = await window.OpsieApi.removeTeamMember(memberId);
        log('Remove team member result', 'info', result);
        
        if (result && result.success) {
            showNotification(`Successfully removed ${memberEmail} from the team`);
            
            // Refresh team members list
            const userInfo = await window.OpsieApi.getUserInfo();
            if (userInfo && userInfo.team_id) {
                await loadTeamMembers(userInfo.team_id);
            }
        } else {
            log('Failed to remove team member', 'error');
            showErrorNotification(`Failed to remove ${memberEmail} from the team`);
        }
    } catch (error) {
        log('Error removing team member', 'error', error);
        showErrorNotification(`Error removing team member: ${error.message}`);
    }
}

// Handler for responding to join requests
async function handleRequestResponse(requestId, approved) {
    try {
        log('Responding to join request', 'info', { requestId, approved });
        
        const result = await window.OpsieApi.respondToJoinRequest(requestId, approved);
        log('Join request response result', 'info', result);
        
        if (result && result.success) {
            const action = approved ? 'approved' : 'denied';
            showNotification(`Successfully ${action} the join request`);
            
            // Get user info for team ID
            const userInfo = await window.OpsieApi.getUserInfo();
            if (userInfo && userInfo.team_id) {
                // Refresh pending requests
                await loadPendingRequests(userInfo.team_id);
                
                // If request was approved, also refresh team members list and admin dropdown
                if (approved) {
                    log('Refreshing team members list after approval', 'info');
                    await loadTeamMembers(userInfo.team_id);
                }
            }
        } else {
            log('Failed to respond to join request', 'error');
            const action = approved ? 'approve' : 'deny';
            showErrorNotification(`Failed to ${action} the join request`);
        }
    } catch (error) {
        log('Error responding to join request', 'error', error);
        showErrorNotification(`Error responding to join request: ${error.message}`);
    }
}

// Handler for team admin transfer
async function handleTransferAdmin() {
    try {
        const teamMembersSelect = document.getElementById('team-members-select');
        const newAdminId = teamMembersSelect.value;
        
        if (!newAdminId) {
            showErrorNotification('Please select a team member to transfer admin rights to');
            return;
        }
        
        const confirmed = await showCustomConfirm(
            'Are you sure you want to transfer admin rights to this user? You will become a regular member.',
            'Transfer Admin Rights',
            'Transfer',
            'Cancel'
        );
        
        if (!confirmed) {
            return;
        }
        
        log('Transferring admin rights', 'info', { newAdminId });
        
        const result = await window.OpsieApi.transferAdminRole(newAdminId);
        log('Transfer admin result', 'info', result);
        
        if (result && result.success) {
            showNotification('Admin rights transferred successfully');
            
            // Refresh user info and update the UI
            await loadUserSettingsInfo(true); // Force refresh to get updated role
            
            // Update UI controls to reflect the new role (user is now a regular member)
            setupTeamControls('member');
            
            // Reload team members to update the list and dropdown with the new admin
            const userInfo = await window.OpsieApi.getUserInfo();
            if (userInfo && userInfo.team_id) {
                await loadTeamMembers(userInfo.team_id, true); // Force refresh
            }
        } else {
            log('Failed to transfer admin rights', 'error');
            showErrorNotification('Failed to transfer admin rights');
        }
    } catch (error) {
        log('Error transferring admin rights', 'error', error);
        showErrorNotification(`Error transferring admin rights: ${error.message}`);
    }
}

// Handler for leaving a team
async function handleLeaveTeam() {
    try {
        const confirmed = await showCustomConfirm(
            'Are you sure you want to leave this team? You will no longer have access to team data.',
            'Leave Team',
            'Leave',
            'Cancel'
        );
        
        if (!confirmed) {
            return;
        }
        
        log('Leaving team', 'info');
        
        const result = await window.OpsieApi.leaveTeam();
        log('Leave team result', 'info', result);
        
        if (result && result.success) {
            showNotification('You have left the team successfully');
            
            // Update UI to reflect the user is no longer in a team
            hideTeamSections();
            
            // Update user info display
            await loadUserSettingsInfo();
            
            // Show team selection view so user can join or create a new team
            if (typeof window.showTeamSelectView === 'function') {
                log('Showing team selection view after leaving team', 'info');
                window.showTeamSelectView();
            } else {
                log('showTeamSelectView function not available', 'error');
            }
        } else {
            log('Failed to leave team', 'error');
            showErrorNotification('Failed to leave team');
        }
    } catch (error) {
        log('Error leaving team', 'error', error);
        showErrorNotification(`Error leaving team: ${error.message}`);
    }
}

// Handler for deleting a team
async function handleDeleteTeam() {
    try {
        const firstConfirmed = await showCustomConfirm(
            'Are you sure you want to delete this team? This action cannot be undone and will remove all team members.',
            'Delete Team',
            'Delete',
            'Cancel'
        );
        
        if (!firstConfirmed) {
            return;
        }
        
        // Double-check with a more explicit confirmation
        const finalConfirmed = await showCustomConfirm(
            'WARNING: This will permanently delete the team and remove all members. This action cannot be undone!',
            'Final Confirmation - Delete Team',
            'Delete Permanently',
            'Cancel'
        );
        
        if (!finalConfirmed) {
            return;
        }
        
        log('Deleting team', 'info');
        
        const result = await window.OpsieApi.deleteTeam();
        log('Delete team result', 'info', result);
        
        if (result && result.success) {
            showNotification('Team deleted successfully');
            
            // Update UI to reflect the user is no longer in a team
            hideTeamSections();
            
            // Update user info display
            await loadUserSettingsInfo();
            
            // Show team selection view so user can join or create a new team
            if (typeof window.showTeamSelectView === 'function') {
                log('Showing team selection view after deleting team', 'info');
                window.showTeamSelectView();
            } else {
                log('showTeamSelectView function not available', 'error');
            }
        } else {
            log('Failed to delete team', 'error');
            showErrorNotification('Failed to delete team');
        }
    } catch (error) {
        log('Error deleting team', 'error', error);
        showErrorNotification(`Error deleting team: ${error.message}`);
    }
}

// Handler for updating team details
async function handleUpdateTeamDetails() {
    try {
        const organization = document.getElementById('edit-team-organization').value;
        const invoiceEmail = document.getElementById('edit-team-invoice-email').value;
        const billingStreet = document.getElementById('edit-team-billing-street').value;
        const billingCity = document.getElementById('edit-team-billing-city').value;
        const billingRegion = document.getElementById('edit-team-billing-region').value;
        const billingCountry = document.getElementById('edit-team-billing-country').value;
        
        // Validate invoice email if provided
        if (invoiceEmail && !isValidEmail(invoiceEmail)) {
            showErrorNotification('Please enter a valid invoice email address');
            return;
        }
        
        const teamDetails = {
            organization,
            invoice_email: invoiceEmail,
            billing_street: billingStreet,
            billing_city: billingCity,
            billing_region: billingRegion,
            billing_country: billingCountry
        };
        
        log('Updating team details', 'info', teamDetails);
        
        const userInfo = await window.OpsieApi.getUserInfo();
        if (!userInfo || !userInfo.team_id) {
            showErrorNotification('No team ID found');
            return;
        }
        
        const result = await window.OpsieApi.updateTeamDetails(userInfo.team_id, teamDetails);
        log('Update team details result', 'info', result);
        
        if (result && result.success) {
            showNotification('Team details updated successfully');
            
            // Toggle back to display view
            toggleTeamDetailsView();
            
            // Refresh team details
            await loadTeamDetails(userInfo.team_id);
        } else {
            log('Failed to update team details', 'error');
            showErrorNotification('Failed to update team details');
        }
    } catch (error) {
        log('Error updating team details', 'error', error);
        showErrorNotification(`Error updating team details: ${error.message}`);
    }
}

// Toggle between team details display and edit views
function toggleTeamDetailsView() {
    const displayView = document.getElementById('team-details-display');
    const editView = document.getElementById('team-details-edit');
    
    if (displayView && editView) {
        const isEditMode = editView.style.display === 'block';
        
        if (isEditMode) {
            // Switch to display mode
            displayView.style.display = 'block';
            editView.style.display = 'none';
        } else {
            // Switch to edit mode
            displayView.style.display = 'none';
            editView.style.display = 'block';
        }
    }
}

// Email validation helper function
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

// Function to handle logout
async function handleLogout() {
    try {
        log('Showing logout confirmation dialog', 'info');
        
        // Use custom dialog instead of window.confirm (which is not supported)
        const modal = document.getElementById('custom-modal-backdrop');
        const modalTitle = document.getElementById('custom-modal-title');
        const modalBody = document.getElementById('custom-modal-body');
        const modalInput = document.getElementById('custom-modal-input');
        const modalOkButton = document.getElementById('custom-modal-ok');
        const modalCancelButton = document.getElementById('custom-modal-cancel');
        const modalCloseButton = document.getElementById('custom-modal-close');
        
        if (!modal || !modalTitle || !modalBody) {
            // Fallback if modal elements don't exist - just logout without confirmation
            log('Modal elements not found, proceeding with logout', 'warning');
            proceedWithLogout();
            return;
        }
        
        // Set up modal for logout confirmation
        modalTitle.textContent = 'Confirm Logout';
        modalBody.innerHTML = '<p>Are you sure you want to log out?</p>';
        
        // Hide the input field since we don't need it for confirmation
        if (modalInput) {
            modalInput.style.display = 'none';
        }
        
        // Set up event handlers for the buttons
        const okHandler = () => {
            // Remove event listeners
            cleanupEventListeners();
            // Hide the modal
            modal.style.display = 'none';
            // Proceed with logout
            proceedWithLogout();
        };
        
        const cancelHandler = () => {
            // Remove event listeners
            cleanupEventListeners();
            // Hide the modal
            modal.style.display = 'none';
        };
        
        // Function to clean up event listeners
        const cleanupEventListeners = () => {
            modalOkButton.removeEventListener('click', okHandler);
            modalCancelButton.removeEventListener('click', cancelHandler);
            if (modalCloseButton) {
                modalCloseButton.removeEventListener('click', cancelHandler);
            }
        };
        
        // Add event listeners
        modalOkButton.addEventListener('click', okHandler);
        modalCancelButton.addEventListener('click', cancelHandler);
        if (modalCloseButton) {
            modalCloseButton.addEventListener('click', cancelHandler);
        }
        
        // Set proper button labels
        modalOkButton.textContent = 'Logout';
        modalCancelButton.textContent = 'Cancel';
        
        // Show the modal
        modal.style.display = 'flex';
    } catch (error) {
        log('Error showing logout confirmation', 'error', error);
        showErrorNotification(`Error during logout process: ${error.message}`);
    }
}

// Function to actually perform the logout
async function proceedWithLogout() {
    try {
        log('Logging out user', 'info');
        
        // Call the API's logout function
        const result = await window.OpsieApi.logout();
        if (!result || !result.success) {
            throw new Error(result?.error || 'Failed to log out');
        }
        
        log('Logout successful, updating UI', 'info');
        showNotification('Logged out successfully');
        
        // Clear any UI elements that might contain user-specific data
        const userEmailEl = document.getElementById('settings-user-email');
        const teamNameEl = document.getElementById('settings-team-name');
        const userRoleEl = document.getElementById('settings-user-role');
        
        if (userEmailEl) userEmailEl.textContent = '-';
        if (teamNameEl) teamNameEl.textContent = '-';
        if (userRoleEl) userRoleEl.textContent = '-';
        
        // Hide any team-related sections
        hideTeamSections();
        
        
        // Clear email-related data
        currentEmailData = null;
        
        // Show the authentication UI
        const authContainer = document.getElementById('auth-container');
        const mainContent = document.getElementById('main-content');
        const settingsContainer = document.getElementById('settings-container');
        
        if (authContainer) {
            authContainer.style.display = 'flex';
        }
        
        if (mainContent) {
            mainContent.style.display = 'none';
        }
        
        // Make sure settings panel is hidden
        if (settingsContainer) {
            settingsContainer.style.display = 'none';
        }
        
        // Reset any other UI elements
        const teamMembersList = document.getElementById('team-members-list');
        if (teamMembersList) {
            teamMembersList.innerHTML = '';
        }
        
        // Focus the email input field
        const emailInput = document.getElementById('auth-email');
        if (emailInput) {
        setTimeout(() => {
                emailInput.focus();
            }, 500);
        }
    } catch (error) {
        log('Error during logout process', 'error', error);
        showErrorNotification(`Error logging out: ${error.message}`);
    }
}

/**
 * Shows all main UI sections that should be visible by default
 */
function showMainUISections() {
    log('Showing main UI sections', 'info');
    
    // First check if user has a team - if not, show team selection view
    const teamId = localStorage.getItem('currentTeamId');
    if (!teamId) {
        log('No team ID found, showing team selection view instead of main UI sections', 'warning');
        
        // Show team selection view instead of error notification
        if (typeof window.showTeamSelectView === 'function') {
            log('Calling showTeamSelectView function', 'info');
            window.showTeamSelectView();
        } else {
            log('showTeamSelectView function not available, showing error notification', 'error');
            showErrorNotification('You need to be assigned to a team to use this application. Please contact your administrator.');
        }
        return;
    }
    
    // User has a team, show the main content container
    const mainContent = document.getElementById('main-content');
    if (mainContent) {
        mainContent.style.display = 'block';
        log('Showed main-content container after team validation', 'info');
    } else {
        log('main-content element not found', 'warning');
    }
    
    // Get settings container ID to exclude it from general restoration
    const settingsContainerId = 'settings-container';
    
    // FIRST: Restore all elements with the .section class that might have been hidden
    // EXCEPT for the settings container
    const allSections = document.querySelectorAll('.section');
    allSections.forEach(section => {
        // Skip the settings container - it should only be shown when settings are opened
        if (section.id === settingsContainerId) {
            section.style.display = 'none';
            log('Keeping settings container hidden', 'info');
            return;
        }
        
        section.style.display = 'block';
        log('Restored section with class .section: ' + (section.id || 'no-id'), 'info');
    });
    
    // Then continue with specific section handling
    // Show the email info section
    const emailInfoSection = document.getElementById('email-info');
    if (emailInfoSection) {
        emailInfoSection.style.display = 'block';
    }
    
    // Show summary section
    const summarySection = document.getElementById('summary-section');
    if (summarySection) {
        summarySection.style.display = 'block';
    }
    
    // Show contact info section
    const contactSection = document.getElementById('contact-section');
    if (contactSection) {
        contactSection.style.display = 'block';
    }
    
    // Show reply section
    const replySection = document.getElementById('reply-section');
    if (replySection) {
        replySection.style.display = 'block';
    }
    
    // Show action buttons section
    const actionButtonsSection = document.getElementById('action-buttons-section');
    if (actionButtonsSection) {
        actionButtonsSection.style.display = 'block';
    }
    
    // Show notes section if the email is saved
    const notesSection = document.getElementById('notes-section');
    if (notesSection) {
        // Only show notes section if we have a saved email
        if (currentEmailData && currentEmailData.existingMessage && currentEmailData.existingMessage.exists) {
            notesSection.style.display = 'block';
        }
    }
    
    // If there are any saved/handled statuses, show those
    const savedStatus = document.getElementById('already-saved-message');
    if (savedStatus) {
        if (currentEmailData && currentEmailData.existingMessage && currentEmailData.existingMessage.exists) {
            savedStatus.style.display = 'block';
        }
    }
    
    const handlingStatus = document.getElementById('handling-status-message');
    if (handlingStatus) {
        if (currentEmailData && currentEmailData.existingMessage && 
            currentEmailData.existingMessage.exists && 
            currentEmailData.existingMessage.handling) {
            handlingStatus.style.display = 'block';
        }
    }
    
    // Make sure settings container stays hidden
    const settingsContainer = document.getElementById(settingsContainerId);
    if (settingsContainer) {
        settingsContainer.style.display = 'none';
    }
    
    // Call updateNotesUIState to ensure notes UI elements are properly enabled/disabled
    // based on whether the email is saved
    updateNotesUIState();
}

// Function to set up event listeners for the settings panel
function setupSettingsEventListeners() {
    log('Setting up settings event listeners', 'info');
    
    // Settings button
    const settingsButton = document.getElementById('settings-button');
    if (settingsButton) {
        settingsButton.addEventListener('click', showSettings);
    }
    
    // Close settings button
    const closeSettingsButton = document.getElementById('close-settings');
    if (closeSettingsButton) {
        closeSettingsButton.addEventListener('click', function() {
            log('Settings closed, restoring main UI sections', 'info');
            
            // Hide the settings container
            const settingsContainer = document.getElementById('settings-container');
            if (settingsContainer) {
                settingsContainer.style.display = 'none';
            }
            
            // First, restore any existing display states that may have been saved
            try {
                const savedDisplayStates = JSON.parse(localStorage.getItem('sectionDisplayStates') || '{}');
                for (const [id, displayState] of Object.entries(savedDisplayStates)) {
                    // Skip restoring the settings container
                    if (id === 'settings-container') continue;
                    
                    const element = document.getElementById(id);
                    if (element) {
                        element.style.display = displayState;
                        log(`Restored saved display state for ${id}: ${displayState}`, 'info');
                    }
                }
    } catch (error) {
                log('Error restoring saved display states: ' + error.message, 'error');
            }
            
            // Show all main UI sections when closing settings
            showMainUISections();
            
            // Additional check for specific sections that might still be hidden
            const emailActionsSection = document.getElementById('summary-section');
            if (emailActionsSection && emailActionsSection.style.display === 'none') {
                log('Forcibly showing email actions section that was still hidden', 'info');
                emailActionsSection.style.display = 'block';
            }
            
            // Final check to make sure settings container is definitely hidden
            if (settingsContainer) {
                settingsContainer.style.display = 'none';
                log('Final check: Ensuring settings container is hidden', 'info');
            }
        });
    }
    
    // Save API key button
    const saveApiKeyButton = document.getElementById('save-api-key');
    if (saveApiKeyButton) {
        saveApiKeyButton.addEventListener('click', saveApiKey);
    }
    
    // Edit team details button
    const editTeamDetailsButton = document.getElementById('edit-team-details-button');
    if (editTeamDetailsButton) {
        editTeamDetailsButton.addEventListener('click', function() {
            toggleTeamDetailsView();
        });
    }
    
    // Cancel edit button
    const cancelEditButton = document.getElementById('cancel-team-edit-button');
    if (cancelEditButton) {
        cancelEditButton.addEventListener('click', function() {
            toggleTeamDetailsView();
        });
    }
    
    // Save team details button
    const saveTeamDetailsButton = document.getElementById('save-team-details-button');
    if (saveTeamDetailsButton) {
        saveTeamDetailsButton.addEventListener('click', handleUpdateTeamDetails);
    }
    
    // Refresh requests button
    const refreshRequestsButton = document.getElementById('refresh-requests-button');
    if (refreshRequestsButton) {
        refreshRequestsButton.addEventListener('click', async function() {
            log('=== REFRESH REQUESTS BUTTON CLICKED ===', 'info');
            const userInfo = await window.OpsieApi.getUserInfo();
            log('User info for refresh:', 'info', userInfo);
            if (userInfo && userInfo.team_id) {
                log('Calling loadPendingRequests with team ID:', 'info', userInfo.team_id);
                await loadPendingRequests(userInfo.team_id);
            } else {
                log('No team ID found for refresh', 'warning');
            }
        });
    }
    
    // Leave team button
    const leaveTeamButton = document.getElementById('leave-team-button');
    if (leaveTeamButton) {
        leaveTeamButton.addEventListener('click', handleLeaveTeam);
    }
    
    // Transfer admin button
    const transferAdminButton = document.getElementById('transfer-admin-button');
    if (transferAdminButton) {
        transferAdminButton.addEventListener('click', handleTransferAdmin);
    }
    
    // Delete team button
    const deleteTeamButton = document.getElementById('delete-team-button');
    if (deleteTeamButton) {
        deleteTeamButton.addEventListener('click', handleDeleteTeam);
    }
    
    // Logout button
    const logoutButton = document.getElementById('logout-button');
    if (logoutButton) {
        logoutButton.addEventListener('click', handleLogout);
    }
}

/**
 * Updates the notes section UI based on whether an email is saved
 * This ensures users can't add notes until the email is saved
 */
function updateNotesUIState() {
    log('Updating notes UI state based on email save status', 'info');
    
    // Check if email is saved
    const isEmailSaved = currentEmailData && 
                          currentEmailData.existingMessage && 
                          currentEmailData.existingMessage.exists;
    
    // Get references to notes-related UI elements
    const notesSection = document.getElementById('notes-section');
    const addNoteButton = document.getElementById('add-note-button');
    const toggleNotesFormButton = document.getElementById('toggle-notes-form-button');
    const notesForm = document.getElementById('notes-form');
    const noteInput = document.getElementById('note-input');
    const saveNoteButton = document.getElementById('save-note-button');
    const cancelNoteButton = document.getElementById('cancel-note-button');
    
    if (notesSection) {
        if (isEmailSaved) {
            // If email is saved, enable notes functionality
            notesSection.style.display = 'block';
            
            // Remove any warning messages
            const warningMsg = notesSection.querySelector('.notes-warning-message');
            if (warningMsg) {
                warningMsg.remove();
            }
            
            // Enable add note button
            if (addNoteButton) {
                addNoteButton.disabled = false;
                addNoteButton.title = "Add a new note";
            }
            
            // Enable toggle form button
            if (toggleNotesFormButton) {
                toggleNotesFormButton.disabled = false;
            }
            
            // Enable note input and save button
            if (noteInput) {
                noteInput.disabled = false;
            }
            
            if (saveNoteButton) {
                saveNoteButton.disabled = false;
            }
            
            log('Notes UI enabled - email is saved', 'info');
        } else {
            // If email is not saved, disable notes functionality
            
            // Keep the notes section visible but add a warning
            notesSection.style.display = 'block';
            
            // Add warning message if it doesn't exist
            if (!notesSection.querySelector('.notes-warning-message')) {
                const warningMsg = document.createElement('div');
                warningMsg.className = 'notes-warning-message';
                warningMsg.style.color = 'red';
                warningMsg.style.padding = '10px';
                warningMsg.style.marginBottom = '10px';
                warningMsg.style.backgroundColor = '#ffeeee';
                warningMsg.style.border = '1px solid #ffaaaa';
                warningMsg.style.borderRadius = '4px';
                warningMsg.innerHTML = '<strong>Note:</strong> You must save the email before adding notes.';
                
                // Insert at the top of notes section
                notesSection.insertBefore(warningMsg, notesSection.firstChild);
            }
            
            // Hide the notes form if it's open
            if (notesForm) {
                notesForm.style.display = 'none';
            }
            
            // Disable add note button
            if (addNoteButton) {
                addNoteButton.disabled = true;
                addNoteButton.title = "Save email before adding notes";
            }
            
            // Disable toggle form button
            if (toggleNotesFormButton) {
                toggleNotesFormButton.disabled = true;
            }
            
            // Disable note input and save button
            if (noteInput) {
                noteInput.disabled = true;
            }
            
            if (saveNoteButton) {
                saveNoteButton.disabled = true;
            }
            
            log('Notes UI disabled - email is not saved', 'info');
        }
    } else {
        log('Notes section not found', 'warning');
    }
}

/**
 * Handles extracting questions from the current email
 */
async function handleExtractQuestions() {
    try {
        log('Extract questions button clicked', 'info');
        
        // Check if user has a team
        if (!checkTeamMembership()) {
            return;
        }
        
        // Check if we have the current email data
        if (!currentEmailData) {
            log('No email data available for question extraction', 'error');
            window.OpsieApi.showNotification('Please load an email before extracting questions.', 'error');
            return;
        }
        
        // Check if email has content
        if (!currentEmailData.body) {
            log('Email has no content for question extraction', 'error');
            window.OpsieApi.showNotification('Email content is missing or empty.', 'error');
            return;
        }
        
        // Get team ID for saving Q&A
        const teamId = localStorage.getItem('currentTeamId');
        
        // Show the QA container and set its initial state
        const qaContainer = document.getElementById('qa-container');
        if (qaContainer) {
            qaContainer.style.display = 'block';
        }
        
        // Clear previous questions
        const qaList = document.getElementById('qa-list');
        if (qaList) {
            qaList.innerHTML = '';
        }
        
        // Show the loading spinner - add null check
        const loadingSpinner = document.getElementById('qa-loading-spinner');
        if (loadingSpinner) {
            loadingSpinner.style.display = 'block';
        } else {
            log('Warning: qa-loading-spinner element not found', 'warning');
        }
        
        // Hide no questions message - add null check
        const noQuestionsMsg = document.getElementById('no-questions-message');
        if (noQuestionsMsg) {
            noQuestionsMsg.style.display = 'none';
        } else {
            log('Warning: no-questions-message element not found', 'warning');
        }
        
        // DEBUG: Check if there's an existing global questions array before extraction
        if (window.OpsieApi.extractQuestionsAndAnswers && window.OpsieApi.extractQuestionsAndAnswers.questions) {
            log('BEFORE EXTRACTION: Existing global questions array', 'debug', {
                length: window.OpsieApi.extractQuestionsAndAnswers.questions.length,
                questions: window.OpsieApi.extractQuestionsAndAnswers.questions.map(q => ({ 
                    text: q.text, 
                    hasAnswer: !!q.answer,
                    answerId: q.answerId
                }))
            });
        } else {
            log('BEFORE EXTRACTION: No existing global questions array', 'debug');
        }
        
        // Call the API to extract questions and find answers
        const result = await window.OpsieApi.extractQuestionsAndAnswers(currentEmailData);
        
        log('Question extraction result:', 'info', result);
        
        // Check for errors
        if (!result.success) {
            log('Failed to extract questions', 'error', result.error);
            window.OpsieApi.showNotification(`Failed to extract questions: ${result.error}`, 'error');
            
            // Hide loading spinner - add null check
            if (loadingSpinner) {
                loadingSpinner.style.display = 'none';
            }
            
            // Show the no questions message - add null check
            if (noQuestionsMsg) {
                noQuestionsMsg.style.display = 'block';
            }
            return;
        }
        
        // If no questions were found
        if (!result.questions || result.questions.length === 0) {
            log('No questions found in the email', 'info');
            window.OpsieApi.showNotification('No questions were found in this email.', 'info');
            
            // Hide loading spinner - add null check
            if (loadingSpinner) {
                loadingSpinner.style.display = 'none';
            }
            
            // Show the no questions message - add null check
            if (noQuestionsMsg) {
                noQuestionsMsg.style.display = 'block';
            }
            return;
        }
        
        // Important: Make sure the global questions array is updated
        if (!window.OpsieApi.extractQuestionsAndAnswers) {
            window.OpsieApi.extractQuestionsAndAnswers = {};
        }
        window.OpsieApi.extractQuestionsAndAnswers.questions = result.questions;
        
        log('AFTER EXTRACTION: Updated global questions array', 'debug', {
            length: window.OpsieApi.extractQuestionsAndAnswers.questions.length,
            questions: window.OpsieApi.extractQuestionsAndAnswers.questions.map(q => ({ 
                text: q.text, 
                hasAnswer: !!q.answer,
                answerId: q.answerId
            }))
        });
        
        // Display the extracted questions
        displayQuestions(result.questions);
        
        // Save any questions that have answers but no answerId (meaning they weren't saved to DB yet)
        // This is a fallback in case the automatic saving in extractQuestionsAndAnswers failed
        if (teamId) {
            let savedCount = 0;
            for (const question of result.questions) {
                if (question.answer && !question.answerId) {
                    try {
                        log('Saving question with answer to database', 'info', {
                            question: question.text,
                            hasAnswer: !!question.answer
                        });
                        
                        // Collect keywords from both the original question and the search
                        const keywords = [
                            ...(question.keywords || []),
                            ...(question.searchKeywords || [])
                        ];
                        
                        const saveResult = await window.OpsieApi.saveQuestionAnswer(
                            question.text,
                            question.answer,
                            question.references || [],
                            teamId,
                            keywords
                        );
                        
                        if (saveResult.success) {
                            question.answerId = saveResult.id;
                            
                            // Also update the question in the global array
                            const globalQuestions = window.OpsieApi.extractQuestionsAndAnswers.questions;
                            const questionIndex = globalQuestions.findIndex(q => q.text === question.text);
                            if (questionIndex >= 0) {
                                globalQuestions[questionIndex] = question;
                                log('Updated question in global array after save', 'debug', {
                                    questionText: question.text,
                                    answerId: saveResult.id
                                });
                            }
                            
                            savedCount++;
                            log('Saved Q&A to database', 'info', { id: saveResult.id });
                        } else {
                            log('Failed to save Q&A to database', 'warning', saveResult.error);
                        }
                    } catch (saveError) {
                        log('Error saving Q&A to database', 'error', saveError);
                    }
                }
            }
            
            if (savedCount > 0) {
                log(`Saved ${savedCount} Q&A pairs to database`, 'info');
                
                // Final check of global questions array after saving
                log('AFTER SAVING: Final global questions array state', 'debug', {
                    length: window.OpsieApi.extractQuestionsAndAnswers.questions.length,
                    questions: window.OpsieApi.extractQuestionsAndAnswers.questions.map(q => ({ 
                        text: q.text, 
                        hasAnswer: !!q.answer,
                        answerId: q.answerId
                    }))
                });
            }
        }
        
        // Show success notification
        window.OpsieApi.showNotification(`Found ${result.questions.length} question${result.questions.length === 1 ? '' : 's'} in the email.`, 'success');
    } catch (error) {
        log('Exception in handleExtractQuestions', 'error', error);
        window.OpsieApi.showNotification(`Error extracting questions: ${error.message}`, 'error');
        
        // Hide loading spinner - add null check
        const loadingSpinner = document.getElementById('qa-loading-spinner');
        if (loadingSpinner) {
            loadingSpinner.style.display = 'none';
        }
        
        // Show the no questions message - add null check
        const noQuestionsMsg = document.getElementById('no-questions-message');
        if (noQuestionsMsg) {
            noQuestionsMsg.style.display = 'block';
        }
    }
}

/**
 * Displays the extracted questions in the UI
 * @param {Array} questions - Array of question objects
 */
function displayQuestions(questions) {
    try {
        log('Displaying questions', 'info', questions);
        
        // Add more detailed logging to understand empty array situation
        log('DEBUG: displayQuestions called with', 'debug', {
            questionsProvided: !!questions,
            isArray: Array.isArray(questions),
            length: questions ? questions.length : 0,
            callerInfo: new Error().stack.split('\n')[2]  // This will show where displayQuestions was called from
        });

        // Check if questions is an empty array and provide a fallback
        if (!questions || !Array.isArray(questions) || questions.length === 0) {
            log('WARNING: displayQuestions called with empty or invalid questions array', 'warn');
            
            // Try to get questions from the global array as a fallback
            if (window.OpsieApi.extractQuestionsAndAnswers && 
                window.OpsieApi.extractQuestionsAndAnswers.questions && 
                window.OpsieApi.extractQuestionsAndAnswers.questions.length > 0) {
                
                log('FALLBACK: Using global questions array instead of empty array', 'info', {
                    length: window.OpsieApi.extractQuestionsAndAnswers.questions.length,
                    source: 'global array fallback'
                });
                
                questions = window.OpsieApi.extractQuestionsAndAnswers.questions;
            } else {
                // No questions to display
                log('No questions to display - hiding qa-container, showing no-questions message', 'info');
                
                // Hide loading spinner
                const loadingSpinner = document.getElementById('qa-loading-spinner');
                if (loadingSpinner) {
                    loadingSpinner.style.display = 'none';
                }
                
                // Show the no questions message
                const noQuestionsMsg = document.getElementById('no-questions-message');
                if (noQuestionsMsg) {
                    noQuestionsMsg.style.display = 'block';
                }
                
                // Hide the qa-container
                const qaContainer = document.getElementById('qa-container');
                if (qaContainer) {
                    qaContainer.style.display = 'block'; // Keep it visible but show the "no questions" message
                }
                
                return;
            }
        }
        
        // Add detailed logging for each question
        questions.forEach((question, index) => {
            log(`Examining question object at index ${index}`, 'info', {
                text: question.text,
                hasAnswer: !!question.answer,
                answerLength: question.answer ? question.answer.length : 0,
                matchType: question.matchType,
                source: question.source,
                verified: question.verified,
                originalQuestion: question.originalQuestion,
                updatedAt: question.updatedAt,
                allProperties: Object.keys(question)
            });
        });
        
        const qaList = document.getElementById('qa-list');
        if (!qaList) {
            log('Q&A list element not found', 'error');
            return;
        }
        
        // Clear existing questions
        qaList.innerHTML = '';
        
        // Sort questions by confidence (highest first)
        questions.sort((a, b) => (b.confidence || 0) - (a.confidence || 0));
        
        // Add each question to the list
        questions.forEach((question, index) => {
            // Create the question item container
            const questionItem = document.createElement('div');
            questionItem.className = 'question-item';
            questionItem.id = `question-${index}`;
            questionItem.style.marginBottom = '20px';
            questionItem.style.padding = '15px';
            questionItem.style.borderRadius = '8px';
            questionItem.style.boxShadow = '0 2px 5px rgba(0,0,0,0.08)';
            questionItem.style.border = '1px solid #e5e5e5';
            questionItem.style.backgroundColor = '#ffffff';
            
            // Create question header with text and confidence
            const questionHeader = document.createElement('div');
            questionHeader.className = 'question-header';
            
            // Question text - main content
            const questionText = document.createElement('div');
            questionText.className = 'question-text';
            questionText.style.fontSize = '1em';
            questionText.style.fontWeight = '600';
            questionText.style.color = '#333';
            questionText.style.marginBottom = '8px';
            questionText.style.lineHeight = '1.4';
            
            // Add manual question badge inline if applicable
            if (question.isManual) {
                const manualBadge = document.createElement('span');
                manualBadge.textContent = 'Manual';
                manualBadge.style.backgroundColor = '#6f42c1';
                manualBadge.style.color = 'white';
                manualBadge.style.padding = '2px 6px';
                manualBadge.style.borderRadius = '8px';
                manualBadge.style.fontSize = '0.7em';
                manualBadge.style.fontWeight = '500';
                manualBadge.style.marginRight = '8px';
                manualBadge.style.display = 'inline-block';
                questionText.appendChild(manualBadge);
            }
            
            // Add the question text content
            const questionTextContent = document.createElement('span');
            questionTextContent.textContent = question.text;
            questionText.appendChild(questionTextContent);
            
            questionHeader.appendChild(questionText);
            

            
            // Question confidence
            if (question.confidence) {
                const confidenceLevel = question.confidence >= 0.8 ? 'high' : 
                                       (question.confidence >= 0.5 ? 'medium' : 'low');
                                       
                const confidenceIndicator = document.createElement('div');
                confidenceIndicator.className = `question-confidence ${confidenceLevel}`;
                confidenceIndicator.textContent = 
                    confidenceLevel.charAt(0).toUpperCase() + confidenceLevel.slice(1);
                questionHeader.appendChild(confidenceIndicator);
            }
            
            // Add a subtle border after the header
            questionHeader.style.paddingBottom = '12px';
            questionHeader.style.borderBottom = '1px solid #eee';
            questionHeader.style.marginBottom = '12px';
            
            // Add header to question item
            questionItem.appendChild(questionHeader);
            
            // If we have semantic match info, display it
            if (question.matchType === 'semantic' && question.originalQuestion && question.similarityScore) {
                const matchInfo = document.createElement('div');
                matchInfo.className = 'semantic-match-info';
                matchInfo.innerHTML = `<strong>Similar to:</strong> "${question.originalQuestion}" <span class="similarity-score">(${Math.round(question.similarityScore * 100)}% match)</span>`;
                matchInfo.style.backgroundColor = '#f8f9fa';
                matchInfo.style.padding = '8px';
                matchInfo.style.borderRadius = '4px';
                matchInfo.style.marginBottom = '10px';
                matchInfo.style.fontSize = '0.9em';
                matchInfo.style.color = '#666';
                matchInfo.style.fontStyle = 'italic';
                
                // Add a similarity badge
                const similarityBadge = document.createElement('span');
                similarityBadge.className = 'similarity-badge';
                similarityBadge.textContent = `${Math.round(question.similarityScore * 100)}%`;
                similarityBadge.style.backgroundColor = question.similarityScore > 0.9 ? '#28a745' : 
                                                     question.similarityScore > 0.8 ? '#17a2b8' : '#ffc107';
                similarityBadge.style.color = question.similarityScore > 0.8 ? 'white' : 'black';
                similarityBadge.style.padding = '2px 6px';
                similarityBadge.style.borderRadius = '10px';
                similarityBadge.style.fontSize = '0.8em';
                similarityBadge.style.marginLeft = '5px';
                
                matchInfo.appendChild(similarityBadge);
                questionItem.appendChild(matchInfo);
            } else if (question.matchType === 'fuzzy' && question.originalQuestion) {
                // For fuzzy matches, also show the original question
                const matchInfo = document.createElement('div');
                matchInfo.className = 'fuzzy-match-info';
                matchInfo.innerHTML = `<strong>Similar to:</strong> "${question.originalQuestion}"`;
                matchInfo.style.backgroundColor = '#fff8e1';
                matchInfo.style.padding = '8px';
                matchInfo.style.borderRadius = '4px';
                matchInfo.style.marginBottom = '10px';
                matchInfo.style.fontSize = '0.9em';
                matchInfo.style.color = '#856404';
                matchInfo.style.fontStyle = 'italic';
                questionItem.appendChild(matchInfo);
            }
            
            // Create answer container
            const answerContainer = document.createElement('div');
            answerContainer.className = 'answer-container';
            
            // Check if we have an answer
            if (question.answer) {
                // Add answer type indicator if available
                if (question.answerType) {
                    const answerTypeLabel = document.createElement('div');
                    answerTypeLabel.className = 'answer-type-label';
                    answerTypeLabel.style.fontSize = '0.8em';
                    answerTypeLabel.style.fontWeight = '600';
                    answerTypeLabel.style.color = '#666';
                    answerTypeLabel.style.marginBottom = '8px';
                    answerTypeLabel.style.textTransform = 'uppercase';
                    answerTypeLabel.style.letterSpacing = '0.5px';
                    
                    // Map answer types to readable labels with icons
                    const typeLabels = {
                        'factual': '📊 Factual Information',
                        'procedural': '📋 Process/Procedure',
                        'contact_info': '👥 Contact Information',
                        'timeline': '⏰ Timeline/Schedule',
                        'technical': '🔧 Technical Details',
                        'status': '📈 Status Update'
                    };
                    
                    answerTypeLabel.textContent = typeLabels[question.answerType] || `💡 ${question.answerType}`;
                    answerContainer.appendChild(answerTypeLabel);
                }
                
                // Create answer text element
                const answerText = document.createElement('div');
                answerText.className = 'answer-text';
                answerText.textContent = question.answer;
                // Enhanced styling to make answer stand out
                answerText.style.fontSize = '1.1em';
                answerText.style.fontWeight = '500';
                answerText.style.padding = '16px 20px';
                answerText.style.margin = '8px 0 12px 0';
                answerText.style.lineHeight = '1.5';
                
                // Different styling based on answer type
                if (question.answerType === 'contact_info') {
                    answerText.style.backgroundColor = '#e8f5e8';
                    answerText.style.border = '1px solid #28a745';
                    answerText.style.color = '#155724';
                } else if (question.answerType === 'timeline') {
                    answerText.style.backgroundColor = '#fff3cd';
                    answerText.style.border = '1px solid #ffc107';
                    answerText.style.color = '#856404';
                } else if (question.answerType === 'technical') {
                    answerText.style.backgroundColor = '#f8d7da';
                    answerText.style.border = '1px solid #dc3545';
                    answerText.style.color = '#721c24';
                } else {
                    // Default styling
                answerText.style.backgroundColor = '#f0f7ff';
                answerText.style.border = '1px solid #cce5ff';
                answerText.style.color = '#0d47a1';
                }
                
                answerText.style.borderRadius = '8px';
                answerText.style.boxShadow = '0 2px 4px rgba(0,0,0,0.08)';
                answerContainer.appendChild(answerText);
                
                // Add source information
                if (question.source) {
                    const sourceInfo = document.createElement('div');
                    sourceInfo.className = 'answer-source';
                    sourceInfo.style.fontSize = '0.85em';
                    sourceInfo.style.fontStyle = 'italic';
                    sourceInfo.style.color = '#666';
                    sourceInfo.style.marginTop = '8px';
                    sourceInfo.style.padding = '8px 10px';
                    sourceInfo.style.backgroundColor = '#f9f9f9';
                    sourceInfo.style.borderRadius = '4px';
                    sourceInfo.style.borderLeft = '3px solid #ddd';
                    sourceInfo.style.fontSize = '0.8em';
                    sourceInfo.style.lineHeight = '1.4';
                    
                    let sourceText = '';
                    
                    if (question.source === 'database') {
                        // Create a more compact display with badges
                        let matchBadge = '';
                        
                        if (question.matchType === 'exact') {
                            matchBadge = '<span style="background-color:#28a745;color:white;padding:2px 6px;border-radius:10px;margin-right:5px;">Exact match</span>';
                        } else if (question.matchType === 'semantic') {
                            const score = Math.round(question.similarityScore * 100);
                            const colorClass = score > 90 ? '#28a745' : score > 75 ? '#17a2b8' : '#ffc107';
                            matchBadge = `<span style="background-color:${colorClass};color:${score > 75 ? 'white' : 'black'};padding:2px 6px;border-radius:10px;margin-right:5px;">Similar (${score}%)</span>`;
                        } else if (question.matchType === 'fuzzy') {
                            matchBadge = '<span style="background-color:#ffc107;color:black;padding:2px 6px;border-radius:10px;margin-right:5px;">Keyword match</span>';
                        } 
                        
                        // Add verification badge if verified
                        const verifiedBadge = question.verified ? 
                            '<span style="background-color:#28a745;color:white;padding:2px 6px;border-radius:10px;margin-left:5px;">✓ Verified</span>' : '';
                        
                        sourceText += `<div style="display:flex;align-items:center;margin-bottom:4px;">${matchBadge}From knowledge base:${verifiedBadge}</div>`;
                        
                        // Add the original question in a more compact form
                        if (question.originalQuestion && (question.matchType === 'semantic' || question.matchType === 'fuzzy')) {
                            sourceText += `<div style="margin-top:4px;color:#555;"><small>Similar to:</small> "${question.originalQuestion}"</div>`;
                        }
                        
                        // Add the last updated timestamp if available
                        if (question.updatedAt) {
                            try {
                                const lastUpdated = new Date(question.updatedAt);
                                const formattedDate = lastUpdated.toLocaleDateString();
                                const timeAgo = getTimeAgo(lastUpdated);
                                sourceText += `<div style="margin-top:4px;color:#666;"><small>Last updated:</small> ${formattedDate} (${timeAgo})</div>`;
                            } catch (dateError) {
                                log('Error formatting date from updatedAt', 'error', {
                                    error: dateError.message,
                                    updatedAt: question.updatedAt
                                });
                            }
                        }
                        
                        // Add source filename if available
                        if (question.sourceFilename && question.sourceFilename.trim() !== '') {
                            sourceText += `<div style="margin-top:4px;color:#666;"><small>Source:</small> ${question.sourceFilename}</div>`;
                        }
                    } else if (question.source === 'search') {
                        sourceText = '<span style="background-color:#6c757d;color:white;padding:2px 6px;border-radius:10px;margin-right:5px;">Email search</span>';
                        
                        // Add date found if available
                        if (question.foundDate) {
                            const foundDate = new Date(question.foundDate);
                            sourceText += ` <small>on</small> ${foundDate.toLocaleDateString()}`;
                        }
                    }
                    
                    sourceInfo.innerHTML = sourceText;
                    answerContainer.appendChild(sourceInfo);
                    
                    // Add action buttons for database matches or search results
                    if (question.source === 'database' || question.source === 'search') {
                        // Create action buttons container
                        const actionContainer = document.createElement('div');
                        actionContainer.className = 'action-buttons-container';
                        actionContainer.style.display = 'flex';
                        actionContainer.style.marginTop = '10px';
                        actionContainer.style.gap = '8px';
                        
                        if (question.source === 'database') {
                            // Edit Answer button - always include this for database answers
                            const editButton = document.createElement('button');
                            editButton.textContent = question.matchType === 'exact' ? 'Edit Answer' : 'Edit Matched Answer';
                            editButton.className = 'qa-button qa-button-edit';
                            editButton.style.backgroundColor = '#4a89dc';
                            editButton.style.color = 'white';
                            editButton.style.border = 'none';
                            editButton.style.padding = '6px 12px';
                            editButton.style.borderRadius = '4px';
                            editButton.style.cursor = 'pointer';
                            editButton.style.fontSize = '0.85em';
                            editButton.style.fontWeight = '500';
                            editButton.onclick = () => question.matchType === 'exact' 
                                ? editAnswer(index, question) 
                                : editMatchedAnswer(index, question);
                            actionContainer.appendChild(editButton);
                            
                            // Only add "New Submission" button for non-exact matches
                            if (question.matchType !== 'exact') {
                                // New Submission button - creates a new entry with the extracted question
                                const newSubmissionButton = document.createElement('button');
                                newSubmissionButton.textContent = 'New Submission';
                                newSubmissionButton.className = 'qa-button qa-button-new';
                                newSubmissionButton.style.backgroundColor = '#5cb85c';
                                newSubmissionButton.style.color = 'white';
                                newSubmissionButton.style.border = 'none';
                                newSubmissionButton.style.padding = '6px 12px';
                                newSubmissionButton.style.borderRadius = '4px';
                                newSubmissionButton.style.cursor = 'pointer';
                                newSubmissionButton.style.fontSize = '0.85em';
                                newSubmissionButton.style.fontWeight = '500';
                                newSubmissionButton.onclick = () => createNewSubmission(index, question);
                                
                                // Add tooltip
                                newSubmissionButton.title = 'Create a new entry with this exact question';
                                actionContainer.appendChild(newSubmissionButton);
                            }
                            
                            // Add tooltip for edit button
                            editButton.title = question.matchType === 'exact' 
                                ? 'Edit this answer in the database' 
                                : 'Edit the answer for the matched question in the database';
                        } else if (question.source === 'search') {
                            // For search results, add both Edit and Save to Database buttons
                            
                            // Edit button for search results
                            const editButton = document.createElement('button');
                            editButton.textContent = 'Edit Answer';
                            editButton.className = 'qa-button qa-button-edit';
                            editButton.style.backgroundColor = '#4a89dc';
                            editButton.style.color = 'white';
                            editButton.style.border = 'none';
                            editButton.style.padding = '6px 12px';
                            editButton.style.borderRadius = '4px';
                            editButton.style.cursor = 'pointer';
                            editButton.style.fontSize = '0.85em';
                            editButton.style.fontWeight = '500';
                            editButton.onclick = () => editSearchAnswer(index, question);
                            editButton.title = 'Edit this search result answer';
                            actionContainer.appendChild(editButton);
                            
                            // Save to Database button
                            const saveButton = document.createElement('button');
                            saveButton.textContent = 'Save to KB';
                            saveButton.className = 'qa-button qa-button-save-to-db';
                            saveButton.style.backgroundColor = '#5cb85c';
                            saveButton.style.color = 'white';
                            saveButton.style.border = 'none';
                            saveButton.style.padding = '6px 12px';
                            saveButton.style.borderRadius = '4px';
                            saveButton.style.cursor = 'pointer';
                            saveButton.style.fontSize = '0.85em';
                            saveButton.style.fontWeight = '500';
                            saveButton.onclick = () => saveSearchToDatabase(index, question);
                            saveButton.title = 'Save this answer to the knowledge base';
                            actionContainer.appendChild(saveButton);
                        }
                        
                        // Add container after source info
                        answerContainer.appendChild(actionContainer);
                    }
                }
            } else {
                // No answer available - show helpful message with more context
                const noAnswerContainer = document.createElement('div');
                noAnswerContainer.className = 'no-answer-container';
                noAnswerContainer.style.padding = '20px';
                noAnswerContainer.style.backgroundColor = '#fff8e1';
                noAnswerContainer.style.border = '1px solid #ffecb3';
                noAnswerContainer.style.borderRadius = '8px';
                noAnswerContainer.style.textAlign = 'center';
                noAnswerContainer.style.marginBottom = '10px';
                
                const noAnswerIcon = document.createElement('div');
                noAnswerIcon.textContent = '🤔';
                noAnswerIcon.style.fontSize = '2em';
                noAnswerIcon.style.marginBottom = '8px';
                noAnswerContainer.appendChild(noAnswerIcon);
                
                const noAnswerTitle = document.createElement('h4');
                noAnswerTitle.textContent = 'No Answer Found';
                noAnswerTitle.style.margin = '0 0 8px 0';
                noAnswerTitle.style.color = '#856404';
                noAnswerTitle.style.fontSize = '1.1em';
                noAnswerContainer.appendChild(noAnswerTitle);
                
                const noAnswerText = document.createElement('p');
                noAnswerText.textContent = 'This question wasn\'t found in your team\'s emails or knowledge base.';
                noAnswerText.style.margin = '0 0 15px 0';
                noAnswerText.style.color = '#856404';
                noAnswerText.style.lineHeight = '1.4';
                noAnswerContainer.appendChild(noAnswerText);
                
                const helpText = document.createElement('p');
                helpText.innerHTML = '<strong>You can help!</strong> Add an answer to build your team\'s knowledge base.';
                helpText.style.margin = '0 0 15px 0';
                helpText.style.color = '#856404';
                helpText.style.fontSize = '0.9em';
                noAnswerContainer.appendChild(helpText);
                
                const addAnswerButton = document.createElement('button');
                addAnswerButton.textContent = '✍️ Add Answer';
                addAnswerButton.className = 'action-button';
                addAnswerButton.style.backgroundColor = '#28a745';
                addAnswerButton.style.color = 'white';
                addAnswerButton.style.padding = '10px 20px';
                addAnswerButton.style.fontSize = '1em';
                addAnswerButton.style.fontWeight = '500';
                addAnswerButton.onclick = () => addAnswer(index, question);
                noAnswerContainer.appendChild(addAnswerButton);
                
                answerContainer.appendChild(noAnswerContainer);
            }
            
            questionItem.appendChild(answerContainer);
            
            // Add the complete question item to the list
            qaList.appendChild(questionItem);
        });
        
        // Hide loading spinner - add null check
        const loadingSpinner = document.getElementById('qa-loading-spinner');
        if (loadingSpinner) {
            loadingSpinner.style.display = 'none';
        } else {
            log('Warning: qa-loading-spinner element not found', 'warning');
        }
        
        // Show container if there are questions - add null checks
        if (questions.length > 0) {
            const qaContainer = document.getElementById('qa-container');
            if (qaContainer) {
                qaContainer.style.display = 'block';
            }
            
            const noQuestionsMsg = document.getElementById('no-questions-message');
            if (noQuestionsMsg) {
                noQuestionsMsg.style.display = 'none';
            }
        } else {
            const noQuestionsMsg = document.getElementById('no-questions-message');
            if (noQuestionsMsg) {
                noQuestionsMsg.style.display = 'block';
            }
        }
    } catch (error) {
        log('Exception in displayQuestions', 'error', error);
        
        // Hide loading spinner - add null check
        const loadingSpinner = document.getElementById('qa-loading-spinner');
        if (loadingSpinner) {
            loadingSpinner.style.display = 'none';
        }
    }
}

/**
 * Verify an answer as correct
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 */
async function verifyAnswer(index, question) {
    try {
        log('Verifying answer', 'info', { index, question });
        
        // Update the UI
        const verifyButton = document.querySelector(`#question-${index} .verify-answer-button`);
        if (verifyButton) {
            verifyButton.textContent = 'Verified ✓';
            verifyButton.classList.add('verified');
            verifyButton.disabled = true;
        }
        
        // Update question object
        question.verified = true;
        
        // If we have an answer ID, update the verification status in the database
        if (question.answerId) {
            try {
                // Call API to update the verification status
                const updateResult = await apiRequest(
                    `qanda?id=eq.${question.answerId}`,
                    'PATCH',
                    { is_verified: true }
                );
                
                if (updateResult.success) {
                    log('Updated verification status in database', 'info', updateResult.data);
                    window.OpsieApi.showNotification('Answer verified and saved to database', 'success');
                } else {
                    log('Failed to update verification status in database', 'error', updateResult.error);
                    window.OpsieApi.showNotification('Answer marked as verified in UI, but failed to update database', 'warning');
                }
            } catch (dbError) {
                log('Database error updating verification status', 'error', dbError);
                window.OpsieApi.showNotification('Answer marked as verified in UI, but failed to update database', 'warning');
            }
        } else {
            // No answer ID means we need to save this as a new Q&A pair
            log('No answer ID available, saving as new verified Q&A pair', 'warning');
            
            // Get team ID
            const teamId = localStorage.getItem('currentTeamId');
            if (teamId) {
                // Collect keywords
                const keywords = [
                    ...(question.keywords || []),
                    ...(question.searchKeywords || [])
                ];
                
                try {
                    const saveResult = await window.OpsieApi.saveQuestionAnswer(
                        question.text,
                        question.answer,
                        question.references || [],
                        teamId,
                        keywords,
                        true // Mark as verified
                    );
                    
                    if (saveResult.success) {
                        question.answerId = saveResult.id;
                        log('Saved verified Q&A to database', 'info', { id: saveResult.id });
                        window.OpsieApi.showNotification('Answer verified and saved to database', 'success');
                    } else {
                        log('Error saving verified Q&A to database', 'error', saveResult.error);
                        window.OpsieApi.showNotification('Answer marked as verified in UI, but failed to save to database', 'warning');
                    }
                } catch (saveError) {
                    log('Exception saving verified Q&A to database', 'error', saveError);
                    window.OpsieApi.showNotification('Answer marked as verified in UI, but failed to save to database', 'warning');
                }
            } else {
                log('No team ID available for saving verified Q&A', 'error');
                window.OpsieApi.showNotification('Answer marked as verified in UI, but could not save to database - missing team ID', 'warning');
            }
        }
    } catch (error) {
        log('Exception in verifyAnswer', 'error', error);
        window.OpsieApi.showNotification(`Error verifying answer: ${error.message}`, 'error');
    }
}

/**
 * Edit an answer
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 */
function editAnswer(index, question) {
    try {
        log('Editing answer', 'info', { index, question });
        
        const answerContainer = document.querySelector(`#question-${index} .answer-container`);
        if (!answerContainer) {
            log('Answer container not found', 'error');
            return;
        }
        
        // Get the current answer text
        const currentAnswerText = question.answer || '';
        
        // Create edit form with enhanced styling
        const editForm = document.createElement('div');
        editForm.className = 'answer-edit-form';
        editForm.style.padding = '20px';
        editForm.style.backgroundColor = '#f8f9fa';
        editForm.style.border = '1px solid #e9ecef';
        editForm.style.borderRadius = '8px';
        editForm.style.marginBottom = '10px';
        
        // Create form header
        const formHeader = document.createElement('div');
        formHeader.style.marginBottom = '15px';
        formHeader.style.display = 'flex';
        formHeader.style.alignItems = 'center';
        formHeader.style.gap = '8px';
        
        const headerIcon = document.createElement('span');
        headerIcon.textContent = '✏️';
        headerIcon.style.fontSize = '1.2em';
        formHeader.appendChild(headerIcon);
        
        const headerText = document.createElement('h4');
        headerText.textContent = 'Edit Answer';
        headerText.style.margin = '0';
        headerText.style.color = '#495057';
        headerText.style.fontSize = '1em';
        headerText.style.fontWeight = '600';
        formHeader.appendChild(headerText);
        
        editForm.appendChild(formHeader);
        
        // Create textarea with enhanced styling
        const textarea = document.createElement('textarea');
        textarea.className = 'answer-edit-textarea';
        textarea.value = currentAnswerText;
        textarea.style.width = '100%';
        textarea.style.minHeight = '120px';
        textarea.style.padding = '15px';
        textarea.style.border = '2px solid #ced4da';
        textarea.style.borderRadius = '6px';
        textarea.style.fontSize = '14px';
        textarea.style.lineHeight = '1.5';
        textarea.style.resize = 'vertical';
        textarea.style.fontFamily = 'inherit';
        textarea.style.boxSizing = 'border-box';
        textarea.style.transition = 'border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out';
        
        // Add focus styling
        textarea.addEventListener('focus', () => {
            textarea.style.borderColor = '#007bff';
            textarea.style.outline = 'none';
            textarea.style.boxShadow = '0 0 0 0.2rem rgba(0, 123, 255, 0.25)';
        });
        
        textarea.addEventListener('blur', () => {
            textarea.style.borderColor = '#ced4da';
            textarea.style.boxShadow = 'none';
        });
        
        editForm.appendChild(textarea);
        
        // Create buttons container with enhanced styling
        const buttonsContainer = document.createElement('div');
        buttonsContainer.className = 'edit-action-buttons';
        buttonsContainer.style.display = 'flex';
        buttonsContainer.style.gap = '12px';
        buttonsContainer.style.marginTop = '15px';
        buttonsContainer.style.justifyContent = 'flex-end';
        
        // Create save button with compact styling
        const saveButton = document.createElement('button');
        saveButton.className = 'save-answer-button';
        saveButton.textContent = 'Save';
        saveButton.style.backgroundColor = '#007bff';
        saveButton.style.color = 'white';
        saveButton.style.border = 'none';
        saveButton.style.padding = '8px 16px';
        saveButton.style.borderRadius = '6px';
        saveButton.style.cursor = 'pointer';
        saveButton.style.fontSize = '13px';
        saveButton.style.fontWeight = '500';
        saveButton.style.transition = 'all 0.15s ease-in-out';
        saveButton.style.boxShadow = '0 2px 4px rgba(0, 123, 255, 0.2)';
        saveButton.style.flex = '1';
        saveButton.style.maxWidth = '80px';
        
        // Add hover effects
        saveButton.addEventListener('mouseenter', () => {
            saveButton.style.backgroundColor = '#0056b3';
            saveButton.style.transform = 'translateY(-1px)';
            saveButton.style.boxShadow = '0 4px 8px rgba(0, 123, 255, 0.3)';
        });
        
        saveButton.addEventListener('mouseleave', () => {
            saveButton.style.backgroundColor = '#007bff';
            saveButton.style.transform = 'translateY(0)';
            saveButton.style.boxShadow = '0 2px 4px rgba(0, 123, 255, 0.2)';
        });
        
        saveButton.onclick = () => saveEditedAnswer(index, question, textarea.value);
        buttonsContainer.appendChild(saveButton);
        
        // Create cancel button with compact styling
        const cancelButton = document.createElement('button');
        cancelButton.className = 'cancel-edit-button';
        cancelButton.textContent = 'Cancel';
        cancelButton.style.backgroundColor = '#6c757d';
        cancelButton.style.color = 'white';
        cancelButton.style.border = 'none';
        cancelButton.style.padding = '8px 16px';
        cancelButton.style.borderRadius = '6px';
        cancelButton.style.cursor = 'pointer';
        cancelButton.style.fontSize = '13px';
        cancelButton.style.fontWeight = '500';
        cancelButton.style.transition = 'all 0.15s ease-in-out';
        cancelButton.style.boxShadow = '0 2px 4px rgba(108, 117, 125, 0.2)';
        cancelButton.style.flex = '1';
        cancelButton.style.maxWidth = '80px';
        
        // Add hover effects
        cancelButton.addEventListener('mouseenter', () => {
            cancelButton.style.backgroundColor = '#5a6268';
            cancelButton.style.transform = 'translateY(-1px)';
            cancelButton.style.boxShadow = '0 4px 8px rgba(108, 117, 125, 0.3)';
        });
        
        cancelButton.addEventListener('mouseleave', () => {
            cancelButton.style.backgroundColor = '#6c757d';
            cancelButton.style.transform = 'translateY(0)';
            cancelButton.style.boxShadow = '0 2px 4px rgba(108, 117, 125, 0.2)';
        });
        
        cancelButton.onclick = () => cancelEdit(index, question);
        buttonsContainer.appendChild(cancelButton);
        
        editForm.appendChild(buttonsContainer);
        
        // Add helpful tip
        const tipText = document.createElement('div');
        tipText.style.fontSize = '12px';
        tipText.style.color = '#6c757d';
        tipText.style.marginTop = '10px';
        tipText.style.padding = '8px 12px';
        tipText.style.backgroundColor = '#e9ecef';
        tipText.style.borderRadius = '4px';
        tipText.style.fontStyle = 'italic';
        tipText.innerHTML = '💡 <strong>Tip:</strong> Your changes will be saved to the team knowledge base and marked as updated.';
        editForm.appendChild(tipText);
        
        // Store the original content and replace with edit form
        answerContainer.dataset.originalContent = answerContainer.innerHTML;
        answerContainer.innerHTML = '';
        answerContainer.appendChild(editForm);
        
        // Focus the textarea and select all text for easy editing
        textarea.focus();
        textarea.select();
    } catch (error) {
        log('Exception in editAnswer', 'error', error);
        window.OpsieApi.showNotification(`Error editing answer: ${error.message}`, 'error');
    }
}

/**
 * Save an edited answer
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 * @param {string} newAnswer - The new answer text
 */
async function saveEditedAnswer(index, question, newAnswer) {
    try {
        log('Saving edited answer', 'info', { index, question, newAnswer });
        
        // Update the question object
        question.answer = newAnswer;
        
        // If we have an answer ID, update the answer in the database
        if (question.answerId) {
            try {
                // Call API to update the answer
                const updateResult = await apiRequest(
                    `qanda?id=eq.${question.answerId}`,
                    'PATCH',
                    { 
                        answer_text: newAnswer,
                        last_updated_by: await getUserId()
                    }
                );
                
                if (updateResult.success) {
                    log('Updated answer in database', 'info', updateResult.data);
                } else {
                    log('Failed to update answer in database', 'error', updateResult.error);
                }
            } catch (dbError) {
                log('Database error updating answer', 'error', dbError);
                // Continue even if the database update fails - at least the UI is updated
            }
        } else {
            log('No answer ID available for answer update - creating new entry', 'warning');
            
            // Get team ID
            const teamId = localStorage.getItem('currentTeamId');
            if (teamId) {
                // Save as a new answer
                const keywords = [
                    ...(question.keywords || []),
                    ...(question.searchKeywords || [])
                ];
                
                try {
                    const saveResult = await window.OpsieApi.saveQuestionAnswer(
                        question.text,
                        newAnswer,
                        [],  // No references for user-edited answers
                        teamId,
                        keywords
                    );
                    
                    if (saveResult.success) {
                        question.answerId = saveResult.id;
                        log('Saved edited answer as new Q&A entry', 'info', { id: saveResult.id });
                    } else {
                        log('Error saving edited answer as new entry', 'warning', saveResult.error);
                    }
                } catch (saveError) {
                    log('Exception saving edited answer as new entry', 'error', saveError);
                }
            } else {
                log('No team ID available for saving edited answer', 'error');
            }
        }
        
        // Update the global questions array and redisplay all questions
        if (window.OpsieApi.extractQuestionsAndAnswers && 
            window.OpsieApi.extractQuestionsAndAnswers.questions) {
            
            // Find and update the question in the global array
            const globalQuestions = window.OpsieApi.extractQuestionsAndAnswers.questions;
            const questionIndex = globalQuestions.findIndex(q => 
                q.text === question.text || 
                (q.id && q.id === question.id) ||
                (q.answerId && q.answerId === question.answerId)
            );
            
            if (questionIndex >= 0) {
                // Update the question in the global array
                globalQuestions[questionIndex] = question;
                log('Updated question in global array after edit', 'info', {
                    index: questionIndex,
                    questionText: question.text,
                    answerId: question.answerId
                });
            } else {
                // If not found, add it to the global array
                globalQuestions.push(question);
                log('Added edited question to global array', 'info', {
                    questionText: question.text,
                    answerId: question.answerId
                });
            }
            
            log('Global questions array after edit', 'debug', {
                length: globalQuestions.length,
                questions: globalQuestions.map(q => ({
                    text: q.text,
                    hasAnswer: !!q.answer,
                    answerId: q.answerId
                }))
            });
            
            // Redisplay all questions
            displayQuestions(globalQuestions);
        } else {
            log('Warning: No global questions array available for update', 'warning');
            
            // Fallback to updating just this specific question's UI without full redisplay
            const questionElement = document.getElementById(`question-${index}`);
            if (questionElement) {
                // Cancel the edit mode
                cancelEdit(index, question);
                
                // Update the answer text
                const answerText = questionElement.querySelector('.answer-text');
                if (answerText) {
                    answerText.textContent = newAnswer;
                }
                
                // Update the source info
                const sourceInfo = questionElement.querySelector('.answer-source');
                if (sourceInfo) {
                    // Only completely replace the source info for new entries that weren't exact matches before
                    if (!question.matchType || question.matchType !== 'exact') {
                        // This is a new entry or a non-exact match that we've edited, so update to show as an exact match
                        sourceInfo.innerHTML = '<div style="display:flex;align-items:center;margin-bottom:4px;"><span style="background-color:#28a745;color:white;padding:2px 6px;border-radius:10px;margin-right:5px;">Exact match</span>From knowledge base:<span style="background-color:#28a745;color:white;padding:2px 6px;border-radius:10px;margin-left:5px;">✓ Verified</span></div>';
                        
                        // Remove the action buttons since this is now an exact match
                        const actionContainer = sourceInfo.querySelector('.action-buttons-container');
                        if (actionContainer) {
                            sourceInfo.removeChild(actionContainer);
                        }
                    } else {
                        // This was already an exact match, so just update the timestamp
                        const lastUpdatedDiv = sourceInfo.querySelector('div[style*="margin-top:4px;color:#666;"]');
                        if (lastUpdatedDiv) {
                            const now = new Date();
                            const formattedDate = now.toLocaleDateString();
                            lastUpdatedDiv.innerHTML = `<small>Last updated:</small> ${formattedDate} (just now)`;
                        }
                    }
                }
                
                // Add a quick animation to highlight the change
                if (answerText) {
                    answerText.style.transition = 'background-color 0.5s ease';
                    answerText.style.backgroundColor = '#e6f7ff';
                    setTimeout(() => {
                        answerText.style.backgroundColor = '#f0f7ff';
                    }, 1000);
                }
                
                log('Updated answer display in UI without redisplay', 'info');
            }
        }
        
        // Show success notification
        window.OpsieApi.showNotification('Answer updated successfully!', 'success');
    } catch (error) {
        log('Exception in saveEditedAnswer', 'error', error);
        window.OpsieApi.showNotification(`Error saving edited answer: ${error.message}`, 'error');
    }
}

/**
 * Cancel editing an answer
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 */
function cancelEdit(index, question) {
    try {
        log('Canceling answer edit', 'info', { index, question });
        
        const answerContainer = document.querySelector(`#question-${index} .answer-container`);
        if (!answerContainer || !answerContainer.dataset.originalContent) {
            log('Original answer content not found', 'error');
            return;
        }
        
        // Instead of just restoring the HTML content, we'll parse it and recreate the DOM elements
        // to ensure the event listeners are properly attached
        
        // First, restore the original HTML content
        answerContainer.innerHTML = answerContainer.dataset.originalContent;
        delete answerContainer.dataset.originalContent;
        
        // Now reattach event listeners to the buttons
        if (question.source === 'database') {
            // Find the action buttons container
            const actionContainer = answerContainer.querySelector('.action-buttons-container');
            if (actionContainer) {
                // Find the edit button
                const editButton = actionContainer.querySelector('.qa-button-edit');
                if (editButton) {
                    // Reattach the correct event listener
                    editButton.onclick = () => question.matchType === 'exact' 
                        ? editAnswer(index, question) 
                        : editMatchedAnswer(index, question);
                }
                
                // Find the new submission button (if it exists)
                if (question.matchType !== 'exact') {
                    const newSubmissionButton = actionContainer.querySelector('.qa-button-new');
                    if (newSubmissionButton) {
                        newSubmissionButton.onclick = () => createNewSubmission(index, question);
                    }
                }
            }
        } else if (!question.answer) {
            // For questions without answers, reattach the add answer button
            const addAnswerButton = answerContainer.querySelector('.action-button');
            if (addAnswerButton) {
                addAnswerButton.onclick = () => addAnswer(index, question);
                log('Reattached Add Answer button event listener', 'info');
            } else {
                log('Add Answer button not found during cancel', 'warning');
            }
        }
        
        log('Successfully restored original content and reattached event listeners', 'info');
    } catch (error) {
        log('Exception in cancelEdit', 'error', error);
        window.OpsieApi.showNotification(`Error canceling edit: ${error.message}`, 'error');
    }
}

/**
 * Add an answer to a question that doesn't have one
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 */
function addAnswer(index, question) {
    try {
        log('Adding answer', 'info', { index, question });
        
        const answerContainer = document.querySelector(`#question-${index} .answer-container`);
        if (!answerContainer) {
            log('Answer container not found', 'error');
            return;
        }
        
        // Create edit form with enhanced styling
        const editForm = document.createElement('div');
        editForm.className = 'answer-edit-form';
        editForm.style.padding = '20px';
        editForm.style.backgroundColor = '#f8f9fa';
        editForm.style.border = '1px solid #e9ecef';
        editForm.style.borderRadius = '8px';
        editForm.style.marginBottom = '10px';
        
        // Create form header
        const formHeader = document.createElement('div');
        formHeader.style.marginBottom = '15px';
        formHeader.style.display = 'flex';
        formHeader.style.alignItems = 'center';
        formHeader.style.gap = '8px';
        
        const headerIcon = document.createElement('span');
        headerIcon.textContent = '✍️';
        headerIcon.style.fontSize = '1.2em';
        formHeader.appendChild(headerIcon);
        
        const headerText = document.createElement('h4');
        headerText.textContent = 'Add Your Answer';
        headerText.style.margin = '0';
        headerText.style.color = '#495057';
        headerText.style.fontSize = '1em';
        headerText.style.fontWeight = '600';
        formHeader.appendChild(headerText);
        
        editForm.appendChild(formHeader);
        
        // Create textarea with enhanced styling
        const textarea = document.createElement('textarea');
        textarea.className = 'answer-edit-textarea';
        textarea.placeholder = 'Enter your answer here... Be specific and helpful!';
        textarea.style.width = '100%';
        textarea.style.minHeight = '120px';
        textarea.style.padding = '15px';
        textarea.style.border = '2px solid #ced4da';
        textarea.style.borderRadius = '6px';
        textarea.style.fontSize = '14px';
        textarea.style.lineHeight = '1.5';
        textarea.style.resize = 'vertical';
        textarea.style.fontFamily = 'inherit';
        textarea.style.boxSizing = 'border-box';
        textarea.style.transition = 'border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out';
        
        // Add focus styling
        textarea.addEventListener('focus', () => {
            textarea.style.borderColor = '#28a745';
            textarea.style.outline = 'none';
            textarea.style.boxShadow = '0 0 0 0.2rem rgba(40, 167, 69, 0.25)';
        });
        
        textarea.addEventListener('blur', () => {
            textarea.style.borderColor = '#ced4da';
            textarea.style.boxShadow = 'none';
        });
        
        editForm.appendChild(textarea);
        
        // Create buttons container with enhanced styling
        const buttonsContainer = document.createElement('div');
        buttonsContainer.className = 'edit-action-buttons';
        buttonsContainer.style.display = 'flex';
        buttonsContainer.style.gap = '12px';
        buttonsContainer.style.marginTop = '15px';
        buttonsContainer.style.justifyContent = 'flex-end';
        
        // Create save button with compact styling
        const saveButton = document.createElement('button');
        saveButton.className = 'save-answer-button';
        saveButton.textContent = 'Save';
        saveButton.style.backgroundColor = '#28a745';
        saveButton.style.color = 'white';
        saveButton.style.border = 'none';
        saveButton.style.padding = '8px 16px';
        saveButton.style.borderRadius = '6px';
        saveButton.style.cursor = 'pointer';
        saveButton.style.fontSize = '13px';
        saveButton.style.fontWeight = '500';
        saveButton.style.transition = 'all 0.15s ease-in-out';
        saveButton.style.boxShadow = '0 2px 4px rgba(40, 167, 69, 0.2)';
        saveButton.style.flex = '1';
        saveButton.style.maxWidth = '80px';
        
        // Add hover effects
        saveButton.addEventListener('mouseenter', () => {
            saveButton.style.backgroundColor = '#218838';
            saveButton.style.transform = 'translateY(-1px)';
            saveButton.style.boxShadow = '0 4px 8px rgba(40, 167, 69, 0.3)';
        });
        
        saveButton.addEventListener('mouseleave', () => {
            saveButton.style.backgroundColor = '#28a745';
            saveButton.style.transform = 'translateY(0)';
            saveButton.style.boxShadow = '0 2px 4px rgba(40, 167, 69, 0.2)';
        });
        
        saveButton.onclick = () => saveNewAnswer(index, question, textarea.value);
        buttonsContainer.appendChild(saveButton);
        
        // Create cancel button with compact styling
        const cancelButton = document.createElement('button');
        cancelButton.className = 'cancel-edit-button';
        cancelButton.textContent = 'Cancel';
        cancelButton.style.backgroundColor = '#6c757d';
        cancelButton.style.color = 'white';
        cancelButton.style.border = 'none';
        cancelButton.style.padding = '8px 16px';
        cancelButton.style.borderRadius = '6px';
        cancelButton.style.cursor = 'pointer';
        cancelButton.style.fontSize = '13px';
        cancelButton.style.fontWeight = '500';
        cancelButton.style.transition = 'all 0.15s ease-in-out';
        cancelButton.style.boxShadow = '0 2px 4px rgba(108, 117, 125, 0.2)';
        cancelButton.style.flex = '1';
        cancelButton.style.maxWidth = '80px';
        
        // Add hover effects
        cancelButton.addEventListener('mouseenter', () => {
            cancelButton.style.backgroundColor = '#5a6268';
            cancelButton.style.transform = 'translateY(-1px)';
            cancelButton.style.boxShadow = '0 4px 8px rgba(108, 117, 125, 0.3)';
        });
        
        cancelButton.addEventListener('mouseleave', () => {
            cancelButton.style.backgroundColor = '#6c757d';
            cancelButton.style.transform = 'translateY(0)';
            cancelButton.style.boxShadow = '0 2px 4px rgba(108, 117, 125, 0.2)';
        });
        
        cancelButton.onclick = () => cancelEdit(index, question);
        buttonsContainer.appendChild(cancelButton);
        
        editForm.appendChild(buttonsContainer);
        
        // Add helpful tip
        const tipText = document.createElement('div');
        tipText.style.fontSize = '12px';
        tipText.style.color = '#6c757d';
        tipText.style.marginTop = '10px';
        tipText.style.padding = '8px 12px';
        tipText.style.backgroundColor = '#e9ecef';
        tipText.style.borderRadius = '4px';
        tipText.style.fontStyle = 'italic';
        tipText.innerHTML = '💡 <strong>Tip:</strong> Your answer will be saved to the team knowledge base to help future team members with similar questions.';
        editForm.appendChild(tipText);
        
        // Store the original content and replace with edit form
        answerContainer.dataset.originalContent = answerContainer.innerHTML;
        answerContainer.innerHTML = '';
        answerContainer.appendChild(editForm);
        
        // Focus the textarea and position cursor at end
        textarea.focus();
        textarea.setSelectionRange(textarea.value.length, textarea.value.length);
    } catch (error) {
        log('Exception in addAnswer', 'error', error);
        window.OpsieApi.showNotification(`Error adding answer: ${error.message}`, 'error');
    }
}

/**
 * Save a new answer
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 * @param {string} newAnswer - The new answer text
 */
async function saveNewAnswer(index, question, newAnswer) {
    try {
        log('Saving new answer', 'info', { index, question, newAnswer });
        
        // Validate input
        if (!newAnswer || newAnswer.trim().length === 0) {
            window.OpsieApi.showNotification('Please enter an answer before saving.', 'error');
            return;
        }
        
        // Update the question object
        question.answer = newAnswer;
        question.verified = true;  // User-provided answers are automatically verified
        question.references = [];
        
        // Get team ID
        const teamId = localStorage.getItem('currentTeamId');
        if (!teamId) {
            log('No team ID available for saving answer', 'error');
            window.OpsieApi.showNotification('Team information is not available. Could not save answer.', 'error');
            return;
        }
        
        // Gather keywords from the question if available
        const keywords = [
            ...(question.keywords || []),
            ...(question.searchKeywords || [])
        ];
        
        // Save to database
        try {
            const saveResult = await window.OpsieApi.saveQuestionAnswer(
                question.text,
                newAnswer,
                [],  // No references for user-added answers
                teamId,
                keywords,
                true  // Mark as verified since it's user-provided
            );
            
            if (saveResult.success) {
                question.answerId = saveResult.id;
                
                // Important: Set additional properties to ensure UI displays correctly
                question.source = 'database';  // Mark as coming from database so edit button shows
                question.matchType = 'exact';  // This is an exact match
                
                // Set updated timestamp
                if (saveResult.updatedAt) {
                    question.updatedAt = saveResult.updatedAt;
                } else {
                    question.updatedAt = new Date().toISOString();
                }
                
                log('Saved new answer to database', 'info', { 
                    id: saveResult.id,
                    source: question.source, 
                    matchType: question.matchType
                });
                
                // IMPORTANT: Update the question in the global array
                if (window.OpsieApi.extractQuestionsAndAnswers && 
                    window.OpsieApi.extractQuestionsAndAnswers.questions) {
                    
                    // Find and update the question in the global array
                    const globalQuestions = window.OpsieApi.extractQuestionsAndAnswers.questions;
                    const questionIndex = globalQuestions.findIndex(q => 
                        q.text === question.text || 
                        (q.id && q.id === question.id)
                    );
                    
                    if (questionIndex >= 0) {
                        globalQuestions[questionIndex] = question;
                        log('Updated question in global array', 'info', {
                            index: questionIndex,
                            questionText: question.text,
                            answerId: saveResult.id
                        });
                    } else {
                        // If the question wasn't found, add it to the global array
                        globalQuestions.push(question);
                        log('Added question to global array', 'info', {
                            questionText: question.text,
                            answerId: saveResult.id
                        });
                    }
                    
                    log('Global questions array after update', 'debug', {
                        length: globalQuestions.length,
                        questions: globalQuestions.map(q => ({
                            text: q.text,
                            hasAnswer: !!q.answer,
                            answerId: q.answerId,
                            source: q.source,
                            matchType: q.matchType
                        }))
                    });
                    
                    // Use the updated global array for display
                    displayQuestions(globalQuestions);
                } else {
                    // If there's no global array, create one with this question
                    if (!window.OpsieApi.extractQuestionsAndAnswers) {
                        window.OpsieApi.extractQuestionsAndAnswers = {};
                    }
                    window.OpsieApi.extractQuestionsAndAnswers.questions = [question];
                    log('Created new global questions array', 'info', {
                        questionText: question.text,
                        answerId: saveResult.id,
                        source: question.source,
                        matchType: question.matchType
                    });
                    
                    // Display the single question
                    displayQuestions([question]);
                }
                
                // Show success notification
                window.OpsieApi.showNotification('Answer saved successfully!', 'success');
            } else {
                log('Error saving answer to database', 'error', saveResult.error);
                window.OpsieApi.showNotification(`Error saving answer: ${saveResult.error}`, 'error');
            }
        } catch (saveError) {
            log('Exception saving answer to database', 'error', saveError);
            window.OpsieApi.showNotification(`Error saving answer: ${saveError.message}`, 'error');
        }
    } catch (error) {
        log('Exception in saveNewAnswer', 'error', error);
        window.OpsieApi.showNotification(`Error saving answer: ${error.message}`, 'error');
    }
}
    
/**
 * Helper function to format time ago from a date
 * @param {Date} date - The date to calculate time from
 * @returns {string} - Formatted time ago string
 */
function getTimeAgo(date) {
    log('getTimeAgo called with date', 'info', {
        date: date.toString(),
        timestamp: date.getTime(),
        now: new Date().toString()
    });
    
    const now = new Date();
    const seconds = Math.floor((now - date) / 1000);
    
    log('Time difference calculated', 'info', {
        differenceMs: now - date,
        differenceSeconds: seconds,
        isNegative: seconds < 0
    });
    
    // Time intervals in seconds
    const intervals = {
        year: 31536000,
        month: 2592000,
        week: 604800,
        day: 86400,
        hour: 3600,
        minute: 60
    };
    
    if (seconds < 60) {
        log('Returning "just now"', 'info');
        return "just now";
    }
    
    for (const [unit, secondsInUnit] of Object.entries(intervals)) {
        const interval = Math.floor(seconds / secondsInUnit);
        
        if (interval >= 1) {
            const result = interval === 1 ? `1 ${unit} ago` : `${interval} ${unit}s ago`;
            log(`Returning time ago result: ${result}`, 'info', {
                unit: unit,
                interval: interval,
                calculation: `${seconds} / ${secondsInUnit} = ${interval}`
            });
            return result;
        }
    }
    
    log('No matching interval found, returning "just now"', 'info');
    return "just now";
}
    
/**
 * Edit the answer for a matched question
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 */
function editMatchedAnswer(index, question) {
    try {
        log('Editing matched answer', 'info', { index, question });
        
        // This is essentially the same as the regular editAnswer function
        // but we're explicitly editing the answer for the matched question
        const answerContainer = document.querySelector(`#question-${index} .answer-container`);
        if (!answerContainer) {
            log('Answer container not found', 'error');
            return;
        }
        
        // Get the current answer text
        const currentAnswerText = question.answer || '';
        
        // Create edit form
        const editForm = document.createElement('div');
        editForm.className = 'answer-edit-form';
        
        // Add a note to explain what's being edited
        const editNote = document.createElement('div');
        editNote.className = 'edit-note';
        editNote.innerHTML = `<strong>Editing answer for matched question:</strong> "${question.originalQuestion}"`;
        editNote.style.marginBottom = '10px';
        editNote.style.padding = '8px';
        editNote.style.backgroundColor = '#f0f4f8';
        editNote.style.borderRadius = '4px';
        editNote.style.fontSize = '0.9em';
        editForm.appendChild(editNote);
        
        // Create textarea for editing
        const textarea = document.createElement('textarea');
        textarea.className = 'answer-edit-textarea';
        textarea.value = currentAnswerText;
        textarea.style.width = '100%';
        textarea.style.minHeight = '120px';
        textarea.style.padding = '8px';
        textarea.style.marginBottom = '10px';
        textarea.style.borderRadius = '4px';
        textarea.style.border = '1px solid #ddd';
        editForm.appendChild(textarea);
        
        // Create buttons container
        const buttonsContainer = document.createElement('div');
        buttonsContainer.className = 'edit-action-buttons';
        buttonsContainer.style.display = 'flex';
        buttonsContainer.style.gap = '8px';
        
        // Create save button
        const saveButton = document.createElement('button');
        saveButton.className = 'save-answer-button';
        saveButton.textContent = 'Save Changes';
        saveButton.style.backgroundColor = '#4a89dc';
        saveButton.style.color = 'white';
        saveButton.style.border = 'none';
        saveButton.style.padding = '8px 16px';
        saveButton.style.borderRadius = '4px';
        saveButton.style.cursor = 'pointer';
        saveButton.onclick = () => saveEditedAnswer(index, question, textarea.value);
        buttonsContainer.appendChild(saveButton);
        
        // Create cancel button
        const cancelButton = document.createElement('button');
        cancelButton.className = 'cancel-edit-button';
        cancelButton.textContent = 'Cancel';
        cancelButton.style.backgroundColor = '#f8f9fa';
        cancelButton.style.color = '#333';
        cancelButton.style.border = '1px solid #ddd';
        cancelButton.style.padding = '8px 16px';
        cancelButton.style.borderRadius = '4px';
        cancelButton.style.cursor = 'pointer';
        cancelButton.onclick = () => cancelEdit(index, question);
        buttonsContainer.appendChild(cancelButton);
        
        editForm.appendChild(buttonsContainer);
        
        // Store the original content and replace with edit form
        answerContainer.dataset.originalContent = answerContainer.innerHTML;
        answerContainer.innerHTML = '';
        answerContainer.appendChild(editForm);
        
        // Focus the textarea
        textarea.focus();
    } catch (error) {
        log('Exception in editMatchedAnswer', 'error', error);
        window.OpsieApi.showNotification(`Error editing matched answer: ${error.message}`, 'error');
    }
}

/**
 * Create a new submission for the extracted question
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 */
function createNewSubmission(index, question) {
    try {
        log('Creating new submission for extracted question', 'info', { index, question });
        
        const answerContainer = document.querySelector(`#question-${index} .answer-container`);
        if (!answerContainer) {
            log('Answer container not found', 'error');
            return;
        }
        
        // Get the current answer text as a starting point
        const currentAnswerText = question.answer || '';
        
        // Create edit form
        const editForm = document.createElement('div');
        editForm.className = 'answer-edit-form';
        
        // Add a note to explain what's being created
        const editNote = document.createElement('div');
        editNote.className = 'edit-note';
        editNote.innerHTML = `<strong>Creating new entry with question:</strong> "${question.text}"`;
        editNote.style.marginBottom = '10px';
        editNote.style.padding = '8px';
        editNote.style.backgroundColor = '#e8f5e9';
        editNote.style.borderRadius = '4px';
        editNote.style.fontSize = '0.9em';
        editForm.appendChild(editNote);
        
        // Create textarea for editing
        const textarea = document.createElement('textarea');
        textarea.className = 'answer-edit-textarea';
        textarea.value = currentAnswerText;
        textarea.style.width = '100%';
        textarea.style.minHeight = '120px';
        textarea.style.padding = '8px';
        textarea.style.marginBottom = '10px';
        textarea.style.borderRadius = '4px';
        textarea.style.border = '1px solid #ddd';
        editForm.appendChild(textarea);
        
        // Create buttons container
        const buttonsContainer = document.createElement('div');
        buttonsContainer.className = 'edit-action-buttons';
        buttonsContainer.style.display = 'flex';
        buttonsContainer.style.gap = '8px';
        
        // Create save button
        const saveButton = document.createElement('button');
        saveButton.className = 'save-answer-button';
        saveButton.textContent = 'Save New Entry';
        saveButton.style.backgroundColor = '#5cb85c';
        saveButton.style.color = 'white';
        saveButton.style.border = 'none';
        saveButton.style.padding = '8px 16px';
        saveButton.style.borderRadius = '4px';
        saveButton.style.cursor = 'pointer';
        saveButton.onclick = () => saveNewSubmission(index, question, textarea.value);
        buttonsContainer.appendChild(saveButton);
        
        // Create cancel button
        const cancelButton = document.createElement('button');
        cancelButton.className = 'cancel-edit-button';
        cancelButton.textContent = 'Cancel';
        cancelButton.style.backgroundColor = '#f8f9fa';
        cancelButton.style.color = '#333';
        cancelButton.style.border = '1px solid #ddd';
        cancelButton.style.padding = '8px 16px';
        cancelButton.style.borderRadius = '4px';
        cancelButton.style.cursor = 'pointer';
        cancelButton.onclick = () => cancelEdit(index, question);
        buttonsContainer.appendChild(cancelButton);
        
        editForm.appendChild(buttonsContainer);
        
        // Store the original content and replace with edit form
        answerContainer.dataset.originalContent = answerContainer.innerHTML;
        answerContainer.innerHTML = '';
        answerContainer.appendChild(editForm);
        
        // Focus the textarea
        textarea.focus();
    } catch (error) {
        log('Exception in createNewSubmission', 'error', error);
        window.OpsieApi.showNotification(`Error creating new submission: ${error.message}`, 'error');
    }
}

/**
 * Save a new submission as a new Q&A entry
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 * @param {string} newAnswer - The new answer text
 */
async function saveNewSubmission(index, question, newAnswer) {
    try {
        log('Saving new submission', 'info', { index, question, newAnswer });
        
        // DEBUG: Log the state of the global questions array before any changes
        if (window.OpsieApi.extractQuestionsAndAnswers && window.OpsieApi.extractQuestionsAndAnswers.questions) {
            log('BEFORE SAVE: Global questions array state', 'debug', {
                length: window.OpsieApi.extractQuestionsAndAnswers.questions.length,
                questions: window.OpsieApi.extractQuestionsAndAnswers.questions.map(q => ({ text: q.text, hasAnswer: !!q.answer }))
            });
        } else {
            log('BEFORE SAVE: Global questions array not found or empty', 'debug');
        }
        
        // Validate input
        if (!newAnswer || newAnswer.trim().length === 0) {
            window.OpsieApi.showNotification('Please enter an answer before saving.', 'error');
            return;
        }
        
        // Get team ID
        const teamId = localStorage.getItem('currentTeamId');
        if (!teamId) {
            log('No team ID available for saving new submission', 'error');
            window.OpsieApi.showNotification('Error: Team ID not found. Please reload the page.', 'error');
            return;
        }
        
        // Collect keywords from both the original question and the current question
        const keywords = [
            ...(question.keywords || []),
            ...(question.searchKeywords || [])
        ];
        
        // Create a new Q&A entry
        try {
            const saveResult = await window.OpsieApi.saveQuestionAnswer(
                question.text, // Use the extracted question text, not the matched one
                newAnswer,
                question.references || [], // Keep any references if available
                teamId,
                keywords,
                true // Set as verified since it's a manual entry
            );
            
            if (saveResult.success) {
                log('Created new Q&A entry', 'info', { id: saveResult.id });
                
                // Update the question object to reflect the new state
                // We're replacing the matched question with the new entry
                question.answerId = saveResult.id;
                question.answer = newAnswer;
                question.verified = true;
                question.matchType = 'exact'; // This is now an exact match
                question.originalQuestion = null; // No original question anymore
                question.source = 'database'; // Mark as coming from the database
                
                // Get the current timestamp for the update
                if (saveResult.updatedAt) {
                    question.updatedAt = saveResult.updatedAt;
                } else {
                    question.updatedAt = new Date().toISOString();
                }

                // Make sure to update the question in the global questions array
                if (window.OpsieApi.extractQuestionsAndAnswers && 
                    window.OpsieApi.extractQuestionsAndAnswers.questions) {
                    // Find the question in the array and update it
                    const globalQuestions = window.OpsieApi.extractQuestionsAndAnswers.questions;
                    const questionIndex = globalQuestions.findIndex(q => q.text === question.text);
                    
                    if (questionIndex >= 0) {
                        globalQuestions[questionIndex] = question;
                        log('Updated question in global questions array', 'info', { 
                            questionIndex, 
                            totalQuestions: globalQuestions.length,
                            updatedQuestion: { 
                                text: question.text, 
                                answer: question.answer ? question.answer.substring(0, 50) + '...' : 'No answer', 
                                answerId: question.answerId
                            }
                        });
                    } else {
                        // If not found, add it to the array
                        globalQuestions.push(question);
                        log('Added question to global questions array', 'info', { 
                            newLength: globalQuestions.length,
                            addedQuestion: { 
                                text: question.text, 
                                answer: question.answer ? question.answer.substring(0, 50) + '...' : 'No answer',
                                answerId: question.answerId
                            }
                        });
                    }

                    // DEBUG: Log the updated global questions array
                    log('AFTER UPDATE: Global questions array state', 'debug', {
                        length: globalQuestions.length,
                        questions: globalQuestions.map(q => ({ 
                            text: q.text, 
                            hasAnswer: !!q.answer,
                            answerId: q.answerId
                        }))
                    });
                    
                    // IMPORTANT: Use the global questions array for display
                    // This ensures consistency with other functions
                    log('Using global questions array for redisplay', 'info', {
                        questionCount: globalQuestions.length
                    });
                    displayQuestions(globalQuestions);
                } else {
                    // Initialize the global questions array if it doesn't exist
                    if (!window.OpsieApi.extractQuestionsAndAnswers) {
                        window.OpsieApi.extractQuestionsAndAnswers = {};
                    }
                    window.OpsieApi.extractQuestionsAndAnswers.questions = [question];
                    log('Created global questions array with new question', 'info', {
                        question: { 
                            text: question.text, 
                            answer: question.answer ? question.answer.substring(0, 50) + '...' : 'No answer',
                            answerId: question.answerId
                        }
                    });

                    // DEBUG: Log the newly created global questions array
                    log('AFTER CREATE: Global questions array state', 'debug', {
                        length: window.OpsieApi.extractQuestionsAndAnswers.questions.length,
                        questions: window.OpsieApi.extractQuestionsAndAnswers.questions.map(q => ({ 
                            text: q.text, 
                            hasAnswer: !!q.answer,
                            answerId: q.answerId
                        }))
                    });
                    
                    // Display the single question
                    displayQuestions([question]);
                }
                
                window.OpsieApi.showNotification('New Q&A entry created successfully!', 'success');
            } else {
                log('Failed to create new Q&A entry', 'error', saveResult.error);
                window.OpsieApi.showNotification(`Error creating new entry: ${saveResult.error}`, 'error');
                
                // Restore the original content
                cancelEdit(index, question);
            }
        } catch (saveError) {
            log('Exception saving new Q&A entry', 'error', saveError);
            window.OpsieApi.showNotification(`Error creating new entry: ${saveError.message}`, 'error');
            
            // Restore the original content
            cancelEdit(index, question);
        }
    } catch (error) {
        log('Exception in saveNewSubmission', 'error', error);
        window.OpsieApi.showNotification(`Error saving new submission: ${error.message}`, 'error');
    }
}
    
/**
 * Edit an answer found through email search
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 */
function editSearchAnswer(index, question) {
    try {
        log('Editing search answer', 'info', { index, question });
        
        const answerContainer = document.querySelector(`#question-${index} .answer-container`);
        if (!answerContainer) {
            log('Answer container not found', 'error');
            return;
        }
        
        // Get the current answer text
        const currentAnswerText = question.answer || '';
        
        // Create edit form
        const editForm = document.createElement('div');
        editForm.className = 'answer-edit-form';
        
        // Add a note to explain what's being edited
        const editNote = document.createElement('div');
        editNote.className = 'edit-note';
        editNote.innerHTML = '<strong>Editing answer from email search:</strong> When you save, this will be added to your knowledge base.';
        editNote.style.marginBottom = '10px';
        editNote.style.padding = '8px';
        editNote.style.backgroundColor = '#e8f5e9';
        editNote.style.borderRadius = '4px';
        editNote.style.fontSize = '0.9em';
        editForm.appendChild(editNote);
        
        // Create textarea for editing
        const textarea = document.createElement('textarea');
        textarea.className = 'answer-edit-textarea';
        textarea.value = currentAnswerText;
        textarea.style.width = '100%';
        textarea.style.minHeight = '120px';
        textarea.style.padding = '8px';
        textarea.style.marginBottom = '10px';
        textarea.style.borderRadius = '4px';
        textarea.style.border = '1px solid #ddd';
        editForm.appendChild(textarea);
        
        // Create buttons container
        const buttonsContainer = document.createElement('div');
        buttonsContainer.className = 'edit-action-buttons';
        buttonsContainer.style.display = 'flex';
        buttonsContainer.style.gap = '8px';
        
        // Create save button
        const saveButton = document.createElement('button');
        saveButton.className = 'save-answer-button';
        saveButton.textContent = 'Save to Knowledge Base';
        saveButton.style.backgroundColor = '#5cb85c';
        saveButton.style.color = 'white';
        saveButton.style.border = 'none';
        saveButton.style.padding = '8px 16px';
        saveButton.style.borderRadius = '4px';
        saveButton.style.cursor = 'pointer';
        saveButton.onclick = () => saveSearchToDatabase(index, question);
        buttonsContainer.appendChild(saveButton);
        
        // Create cancel button
        const cancelButton = document.createElement('button');
        cancelButton.className = 'cancel-edit-button';
        cancelButton.textContent = 'Cancel';
        cancelButton.style.backgroundColor = '#f8f9fa';
        cancelButton.style.color = '#333';
        cancelButton.style.border = '1px solid #ddd';
        cancelButton.style.padding = '8px 16px';
        cancelButton.style.borderRadius = '4px';
        cancelButton.style.cursor = 'pointer';
        cancelButton.onclick = () => cancelEdit(index, question);
        buttonsContainer.appendChild(cancelButton);
        
        editForm.appendChild(buttonsContainer);
        
        // Store the original content and replace with edit form
        answerContainer.dataset.originalContent = answerContainer.innerHTML;
        answerContainer.innerHTML = '';
        answerContainer.appendChild(editForm);
        
        // Focus the textarea
        textarea.focus();
    } catch (error) {
        log('Exception in editSearchAnswer', 'error', error);
        window.OpsieApi.showNotification(`Error editing search answer: ${error.message}`, 'error');
    }
}

/**
 * Save a search answer to the database
 * @param {number} index - Index of the question
 * @param {Object} question - Question object
 * @param {string} [newAnswer] - Optional new answer text, if not provided uses the existing answer
 */
async function saveSearchToDatabase(index, question, newAnswer) {
    try {
        // Use either the provided newAnswer or the existing answer
        const answerText = newAnswer || question.answer;
        
        log('Saving search answer to database', 'info', { 
            index, 
            question, 
            answerText: answerText ? answerText.substring(0, 50) + '...' : 'No answer'
        });
        
        // Validate input
        if (!answerText || answerText.trim().length === 0) {
            window.OpsieApi.showNotification('Cannot save empty answer to database.', 'error');
            return;
        }
        
        // Get team ID
        const teamId = localStorage.getItem('currentTeamId');
        if (!teamId) {
            log('No team ID available for saving to database', 'error');
            window.OpsieApi.showNotification('Team information is not available. Could not save answer.', 'error');
            return;
        }
        
        // Gather keywords from the question if available
        const keywords = [
            ...(question.keywords || []),
            ...(question.searchKeywords || [])
        ];
        
        // Get references from the search result
        const references = question.references || [];
        
        // Save to database
        try {
            const saveResult = await window.OpsieApi.saveQuestionAnswer(
                question.text,
                answerText,
                references,
                teamId,
                keywords,
                true  // Mark as verified since it's manually saved
            );
            
            if (saveResult.success) {
                // Update question properties
                question.answerId = saveResult.id;
                question.source = 'database';  // Now it's from the database
                question.matchType = 'exact';  // It's an exact match
                question.verified = true;      // It's verified
                
                // Update answer if a new one was provided
                if (newAnswer) {
                    question.answer = newAnswer;
                }
                
                // Set updated timestamp
                if (saveResult.updatedAt) {
                    question.updatedAt = saveResult.updatedAt;
                } else {
                    question.updatedAt = new Date().toISOString();
                }
                
                log('Saved search answer to database', 'info', { 
                    id: saveResult.id,
                    source: question.source, 
                    matchType: question.matchType
                });
                
                // Update the global questions array
                if (window.OpsieApi.extractQuestionsAndAnswers && 
                    window.OpsieApi.extractQuestionsAndAnswers.questions) {
                    
                    const globalQuestions = window.OpsieApi.extractQuestionsAndAnswers.questions;
                    const questionIndex = globalQuestions.findIndex(q => 
                        q.text === question.text || 
                        (q.id && q.id === question.id)
                    );
                    
                    if (questionIndex >= 0) {
                        globalQuestions[questionIndex] = question;
                        log('Updated question in global array', 'info', {
                            index: questionIndex,
                            questionText: question.text,
                            answerId: saveResult.id
                        });
                    } else {
                        globalQuestions.push(question);
                        log('Added question to global array', 'info', {
                            questionText: question.text,
                            answerId: saveResult.id
                        });
                    }
                    
                    // Redisplay all questions
                    displayQuestions(globalQuestions);
                } else {
                    // If no global array exists, create one
                    if (!window.OpsieApi.extractQuestionsAndAnswers) {
                        window.OpsieApi.extractQuestionsAndAnswers = {};
                    }
                    window.OpsieApi.extractQuestionsAndAnswers.questions = [question];
                    
                    // Redisplay the single question
                    displayQuestions([question]);
                }
                
                // If we were in edit mode, make sure to cancel it
                if (newAnswer) {
                    const answerContainer = document.querySelector(`#question-${index} .answer-container`);
                    if (answerContainer && answerContainer.dataset.originalContent) {
                        // We were in edit mode, but displayQuestions has been called, so no need to cancel
                        delete answerContainer.dataset.originalContent;
                    }
                }
                
                window.OpsieApi.showNotification('Answer saved to knowledge base successfully!', 'success');
            } else {
                log('Error saving to database', 'error', saveResult.error);
                window.OpsieApi.showNotification(`Error saving to database: ${saveResult.error}`, 'error');
                
                // If we were in edit mode, cancel it
                if (newAnswer) {
                    cancelEdit(index, question);
                }
            }
        } catch (saveError) {
            log('Exception saving to database', 'error', saveError);
            window.OpsieApi.showNotification(`Error saving to database: ${saveError.message}`, 'error');
            
            // If we were in edit mode, cancel it
            if (newAnswer) {
                cancelEdit(index, question);
            }
        }
    } catch (error) {
        log('Exception in saveSearchToDatabase', 'error', error);
        window.OpsieApi.showNotification(`Error saving to database: ${error.message}`, 'error');
    }
}
    
// Document upload handling
function initializeDocumentUpload() {
    const uploadButton = document.getElementById('upload-document-button');
    const fileInput = document.getElementById('knowledge-doc-input');
    const uploadStatus = document.getElementById('document-upload-status');

    if (uploadButton && fileInput) {
        uploadButton.addEventListener('click', () => {
            fileInput.click();
        });

        fileInput.addEventListener('change', async (event) => {
            const file = event.target.files[0];
            if (!file) return;

            try {
                // Show loading state
                uploadStatus.style.display = 'flex';
                uploadButton.disabled = true;

                // Get current team ID
                const teamId = localStorage.getItem('currentTeamId');
                if (!teamId) {
                    throw new Error('No team ID found. Please ensure you are part of a team.');
                }

                // Process the document
                const result = await window.OpsieApi.processDocumentForQA(file, teamId);

                if (result.success) {
                    window.OpsieApi.showNotification(result.message, 'success');
                } else {
                    throw new Error(result.error);
                }
            } catch (error) {
                window.OpsieApi.showNotification(error.message, 'error');
            } finally {
                // Reset the input and UI state
                fileInput.value = '';
                uploadStatus.style.display = 'none';
                uploadButton.disabled = false;
            }
        });
    }
}

// Add to the existing Office.onReady function
Office.onReady((info) => {
    // ... existing code ...

    // Initialize document upload functionality
    initializeDocumentUpload();
});
    
/**
 * Handle manual question submission
 */
async function handleManualQuestionSubmission() {
    try {
        log('Manual question submission started', 'info');
        
        // Get the question text from the input
        const questionInput = document.getElementById('manual-question-input');
        if (!questionInput) {
            log('Manual question input element not found', 'error');
            window.OpsieApi.showNotification('Question input not found.', 'error');
            return;
        }
        
        const questionText = questionInput.value.trim();
        if (!questionText) {
            log('No question text provided', 'warning');
            window.OpsieApi.showNotification('Please enter a question.', 'warning');
            return;
        }
        
        // Check if we have team ID for Q&A operations
        const teamId = localStorage.getItem('currentTeamId');
        if (!teamId) {
            log('No team ID available for manual Q&A', 'warning');
            window.OpsieApi.showNotification('Team information is not available. Please ensure you are logged in.', 'warning');
            return;
        }
        
        // Show the QA container if it's not visible
        const qaContainer = document.getElementById('qa-container');
        if (qaContainer) {
            qaContainer.style.display = 'block';
        }
        
        // Disable the submit button and show loading state
        const submitButton = document.getElementById('submit-manual-question');
        if (submitButton) {
            submitButton.disabled = true;
            submitButton.textContent = 'Searching...';
        }
        
        // Show loading spinner
        const loadingSpinner = document.getElementById('qa-loading-spinner');
        if (loadingSpinner) {
            loadingSpinner.style.display = 'block';
        }
        
        // Hide no questions message
        const noQuestionsMsg = document.getElementById('no-questions-message');
        if (noQuestionsMsg) {
            noQuestionsMsg.style.display = 'none';
        }
        
        log('Processing manual question', 'info', { question: questionText });
        
        // Create a question object similar to extracted questions
        const manualQuestion = {
            text: questionText,
            confidence: 1.0, // High confidence since it's user-provided
            isManual: true,
            answer: null,
            references: [],
            keywords: []
        };
        
        // Try to find existing answer first
        const existingAnswer = await window.OpsieApi.findExistingAnswer(questionText, teamId);
        
        log('Existing answer search result', 'info', existingAnswer);
        
        if (existingAnswer.found && existingAnswer.answer) {
            log('Found existing answer for manual question', 'info', {
                matchType: existingAnswer.matchType,
                similarity: existingAnswer.similarityScore,
                verified: existingAnswer.verified
            });
            manualQuestion.answer = existingAnswer.answer;
            manualQuestion.references = existingAnswer.references || [];
            manualQuestion.keywords = existingAnswer.keywords || [];
            manualQuestion.answerId = existingAnswer.id;
            manualQuestion.isVerified = existingAnswer.verified;
            manualQuestion.confidenceScore = existingAnswer.similarityScore || existingAnswer.confidenceScore;
            manualQuestion.sourceFilename = existingAnswer.sourceFilename;
            manualQuestion.matchType = existingAnswer.matchType;
            manualQuestion.originalQuestion = existingAnswer.originalQuestion;
            manualQuestion.similarityScore = existingAnswer.similarityScore;
            manualQuestion.updatedAt = existingAnswer.updatedAt;
            manualQuestion.verified = existingAnswer.verified;
            manualQuestion.source = 'database'; // This ensures it gets the same display treatment
        } else {
            // Search for answer in team emails
            log('Searching for answer to manual question', 'info');
            const apiKey = await window.OpsieApi.getOpenAIApiKey();
            
            if (!apiKey) {
                log('No OpenAI API key available for manual question search', 'error');
                window.OpsieApi.showNotification('OpenAI API key is required. Please set it in Settings.', 'error');
                return;
            }
            
            const searchResult = await window.OpsieApi.searchForAnswer(questionText, teamId, apiKey);
            
            if (searchResult.success) {
                log('Found answer for manual question', 'info');
                manualQuestion.answer = searchResult.answer;
                manualQuestion.references = searchResult.references || [];
                manualQuestion.keywords = searchResult.keywords || [];
                manualQuestion.confidenceScore = searchResult.confidence;
                
                // Save the Q&A to database
                try {
                    const saveResult = await window.OpsieApi.saveQuestionAnswer(
                        questionText,
                        searchResult.answer,
                        searchResult.references || [],
                        teamId,
                        searchResult.keywords || [],
                        false, // Not user verified initially
                        searchResult.confidence
                    );
                    
                    if (saveResult.success) {
                        manualQuestion.answerId = saveResult.id;
                        log('Saved manual Q&A to database', 'info', { id: saveResult.id });
                    } else {
                        log('Failed to save manual Q&A to database', 'warning', saveResult.error);
                    }
                } catch (saveError) {
                    log('Error saving manual Q&A to database', 'error', saveError);
                }
            } else {
                log('No answer found for manual question', 'info');
                manualQuestion.answer = null;
                manualQuestion.keywords = searchResult.keywords || [];
            }
        }
        
        // Update or create the global questions array
        if (!window.OpsieApi.extractQuestionsAndAnswers) {
            window.OpsieApi.extractQuestionsAndAnswers = {};
        }
        if (!window.OpsieApi.extractQuestionsAndAnswers.questions) {
            window.OpsieApi.extractQuestionsAndAnswers.questions = [];
        }
        
        // Add the manual question to the beginning of the array
        window.OpsieApi.extractQuestionsAndAnswers.questions.unshift(manualQuestion);
        
        // Display all questions (including the new manual one)
        displayQuestions(window.OpsieApi.extractQuestionsAndAnswers.questions);
        
        // Clear the input field
        questionInput.value = '';
        
        // Show success notification
        if (manualQuestion.answer) {
            window.OpsieApi.showNotification('Found an answer to your question!', 'success');
        } else {
            window.OpsieApi.showNotification('Question added, but no answer was found in the knowledge base.', 'info');
        }
        
        log('Manual question submission completed', 'info');
        
    } catch (error) {
        log('Exception in handleManualQuestionSubmission', 'error', error);
        window.OpsieApi.showNotification(`Error processing question: ${error.message}`, 'error');
    } finally {
        // Reset UI state
        const submitButton = document.getElementById('submit-manual-question');
        if (submitButton) {
            submitButton.disabled = false;
            submitButton.textContent = 'Get Answer';
        }
        
        const loadingSpinner = document.getElementById('qa-loading-spinner');
        if (loadingSpinner) {
            loadingSpinner.style.display = 'none';
        }
    }
}
    