/**
 * Team Management Functions for Outlook Add-in
 * Handles team selection, joining, and creation
 */

/**
 * Show the team selection view
 */
function showTeamSelectView() {
    console.log('showTeamSelectView: Starting to show team selection view');
    console.log('showTeamSelectView: Current team ID in localStorage:', localStorage.getItem('currentTeamId'));
    
    // Hide the entire authentication container
    const authContainer = document.getElementById('auth-container');
    if (authContainer) {
        authContainer.style.display = 'none';
    }
    
    // Hide the main content container if visible
    const mainContent = document.getElementById('main-content');
    if (mainContent) {
        mainContent.style.display = 'none';
    }
    
    // Hide settings container if visible
    const settingsContainer = document.getElementById('settings-container');
    if (settingsContainer) {
        settingsContainer.style.display = 'none';
    }
    
    // Hide the settings button since user has no team
    const settingsButton = document.getElementById('settings-button');
    if (settingsButton) {
        settingsButton.style.display = 'none';
        console.log('Hidden settings button - user has no team');
    }
    
    // Show team selection view as full-screen overlay
    const teamSelectView = document.getElementById('team-select-view');
    if (teamSelectView) {
        teamSelectView.style.display = 'block';
        
        // Clear any previous form data
        const joinCodeInput = document.getElementById('join-team-code');
        const createTeamInput = document.getElementById('create-team-name');
        
        if (joinCodeInput) joinCodeInput.value = '';
        if (createTeamInput) createTeamInput.value = '';
        
        // Clear any previous messages
        hideAuthError();
        hideAuthSuccess();
        
        // Check for and display pending join request
        checkAndDisplayPendingRequest();
        
        // Scroll to top
        teamSelectView.scrollTop = 0;
    }
    
    // Setup event listeners
    setupTeamSelectEventListeners();
}

/**
 * Setup event listeners for team selection
 */
function setupTeamSelectEventListeners() {
    // Join team button
    const joinBtn = document.getElementById('join-team-btn');
    if (joinBtn) {
        joinBtn.addEventListener('click', handleJoinTeam);
    }
    
    // Create team button
    const createBtn = document.getElementById('create-team-btn');
    if (createBtn) {
        createBtn.addEventListener('click', handleCreateTeam);
    }
    
    // Refresh request status button
    const refreshBtn = document.getElementById('refresh-request-status-btn');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', handleRefreshPendingRequestStatus);
    }
    
    // Team sign out button
    const signOutBtn = document.getElementById('team-signout-btn');
    if (signOutBtn) {
        signOutBtn.addEventListener('click', handleTeamSignOut);
    }
}

/**
 * Handle joining a team
 */
async function handleJoinTeam() {
    const joinCodeInput = document.getElementById('join-team-code');
    const accessCode = joinCodeInput?.value?.trim();
    
    if (!accessCode) {
        showAuthError('Please enter a team access code');
        return;
    }
    
    showTeamLoading('Requesting to join team...');
    
    try {
        const result = await window.OpsieApi.requestToJoinTeam(accessCode);
        
        if (result.success) {
            showAuthSuccess('Join request sent! Please wait for team admin approval.');
            
            // Clear the input
            joinCodeInput.value = '';
            
            // Refresh the pending request display
            setTimeout(() => {
                checkAndDisplayPendingRequest();
            }, 1000);
            
            // Start polling for approval
            startJoinRequestPolling();
        } else {
            showAuthError(result.error || 'Failed to send join request');
        }
    } catch (error) {
        console.error('Error joining team:', error);
        showAuthError('An error occurred while joining the team');
    } finally {
        hideTeamLoading();
    }
}

/**
 * Handle creating a team
 */
async function handleCreateTeam() {
    const createTeamInput = document.getElementById('create-team-name');
    const teamName = createTeamInput?.value?.trim();
    
    // Collect all form field values
    const organization = document.getElementById('create-team-organization')?.value?.trim();
    const invoiceEmail = document.getElementById('create-team-invoice-email')?.value?.trim();
    const billingStreet = document.getElementById('create-team-billing-street')?.value?.trim();
    const billingCity = document.getElementById('create-team-billing-city')?.value?.trim();
    const billingRegion = document.getElementById('create-team-billing-region')?.value?.trim();
    const billingCountry = document.getElementById('create-team-billing-country')?.value?.trim();
    
    if (!teamName) {
        showAuthError('Please enter a team name');
        return;
    }
    
    if (teamName.length < 2) {
        showAuthError('Team name must be at least 2 characters long');
        return;
    }
    
    // Validate invoice email if provided
    if (invoiceEmail && !isValidEmail(invoiceEmail)) {
        showAuthError('Please enter a valid invoice email address');
        return;
    }
    
    showTeamLoading('Creating team...');
    
    try {
        const result = await window.OpsieApi.createTeam(teamName, {
            organization,
            invoiceEmail,
            billingStreet,
            billingCity,
            billingRegion,
            billingCountry
        });
        
        if (result.success) {
            showAuthSuccess('Team created successfully!');
            
            // Clear all inputs
            createTeamInput.value = '';
            if (document.getElementById('create-team-organization')) document.getElementById('create-team-organization').value = '';
            if (document.getElementById('create-team-invoice-email')) document.getElementById('create-team-invoice-email').value = '';
            if (document.getElementById('create-team-billing-street')) document.getElementById('create-team-billing-street').value = '';
            if (document.getElementById('create-team-billing-city')) document.getElementById('create-team-billing-city').value = '';
            if (document.getElementById('create-team-billing-region')) document.getElementById('create-team-billing-region').value = '';
            if (document.getElementById('create-team-billing-country')) document.getElementById('create-team-billing-country').value = '';
            
            // Hide team selection and show main content
            setTimeout(() => {
                hideAuthContainer();
                showMainContent();
                
                // Show the settings button since user now has a team
                const settingsButton = document.getElementById('settings-button');
                if (settingsButton) {
                    settingsButton.style.display = 'block';
                    console.log('Shown settings button - user now has a team');
                }
                
                // Refresh the UI with new team info
                if (window.OpsieApi && window.OpsieApi.initTeamAndUserInfo) {
                    window.OpsieApi.initTeamAndUserInfo(() => {
                        console.log('Team info refreshed after team creation');
                    });
                }
            }, 1500);
        } else {
            showAuthError(result.error || 'Failed to create team');
        }
    } catch (error) {
        console.error('Error creating team:', error);
        showAuthError('An error occurred while creating the team');
    } finally {
        hideTeamLoading();
    }
}

/**
 * Validate email format
 */
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

/**
 * Check for and display any pending join request
 */
async function checkAndDisplayPendingRequest() {
    try {
        console.log('checkAndDisplayPendingRequest: Starting check for pending requests');
        
        // First check if user already has a team
        const currentTeamId = localStorage.getItem('currentTeamId');
        console.log('checkAndDisplayPendingRequest: Current team ID in localStorage:', currentTeamId);
        
        if (currentTeamId) {
            // User has a team - they shouldn't be in team setup, redirect to main UI
            console.log('checkAndDisplayPendingRequest: User has team but is in team setup - redirecting to main UI');
            hidePendingRequestStatus();
            hideAuthContainer();
            showMainContent();
            
            // Show the settings button since user has a team
            const settingsButton = document.getElementById('settings-button');
            if (settingsButton) {
                settingsButton.style.display = 'block';
                console.log('checkAndDisplayPendingRequest: Shown settings button - user has a team');
            }
            
            // Refresh the UI with team info
            if (window.OpsieApi && window.OpsieApi.initTeamAndUserInfo) {
                window.OpsieApi.initTeamAndUserInfo();
            }
            return;
        }
        
        // Check for pending requests
        const result = await window.OpsieApi.getUserPendingJoinRequest();
        console.log('checkAndDisplayPendingRequest: Pending request result:', result);
        
        if (result.success && result.data) {
            console.log('checkAndDisplayPendingRequest: Found pending request, displaying');
            displayPendingRequestStatus(result.data);
        } else {
            console.log('checkAndDisplayPendingRequest: No pending request found');
            hidePendingRequestStatus();
            
                         // Check if user actually has a team by getting current user info
             console.log('checkAndDisplayPendingRequest: Checking user current team status via API');
             const userInfo = await window.OpsieApi.getUserInfo(true); // Force refresh
             console.log('checkAndDisplayPendingRequest: Current user info:', userInfo);
             
             if (userInfo && userInfo.team_id) {
                 console.log('checkAndDisplayPendingRequest: User currently has a team, redirecting to main UI');
                 hideAuthContainer();
                 
                 // Show the settings button since user has a team
                 const settingsButton = document.getElementById('settings-button');
                 if (settingsButton) {
                     settingsButton.style.display = 'block';
                     console.log('checkAndDisplayPendingRequest: Shown settings button - user has current team');
                 }
                 
                 // Refresh the UI with team info and then show main UI sections
                 if (window.OpsieApi && window.OpsieApi.initTeamAndUserInfo) {
                     console.log('checkAndDisplayPendingRequest: Initializing team info with callback');
                     window.OpsieApi.initTeamAndUserInfo(() => {
                         console.log('checkAndDisplayPendingRequest: Team info initialized, showing main UI sections');
                         // Now that team info is loaded, show the main UI sections
                         if (typeof window.showMainUISections === 'function') {
                             window.showMainUISections();
                         } else {
                             console.log('checkAndDisplayPendingRequest: showMainUISections not available, showing main content manually');
                             showMainContent();
                         }
                     });
                 } else {
                     console.log('checkAndDisplayPendingRequest: initTeamAndUserInfo not available, showing main content manually');
                     showMainContent();
                 }
             } else {
                 console.log('checkAndDisplayPendingRequest: User does not currently have a team - staying in team setup UI');
             }
        }
    } catch (error) {
        console.error('Error checking for pending join request:', error);
        hidePendingRequestStatus();
    }
}

/**
 * Display the pending join request status
 */
function displayPendingRequestStatus(requestData) {
    const pendingRequestSection = document.getElementById('pending-request-status');
    if (!pendingRequestSection) return;
    
    const teamNameElement = document.getElementById('requested-team-name');
    const requestDateElement = document.getElementById('request-date');
    const statusBadge = document.getElementById('request-status-badge');
    
    if (teamNameElement) {
        teamNameElement.textContent = 'a team'; // Generic text instead of team name
    }
    
    if (requestDateElement) {
        const date = new Date(requestData.requestDate);
        requestDateElement.textContent = date.toLocaleDateString();
    }
    
    if (statusBadge) {
        statusBadge.textContent = requestData.status.charAt(0).toUpperCase() + requestData.status.slice(1);
        statusBadge.className = `status-badge status-${requestData.status}`;
    }
    
    pendingRequestSection.style.display = 'block';
    console.log('Displaying pending join request status:', requestData);
}

/**
 * Hide the pending join request status
 */
function hidePendingRequestStatus() {
    const pendingRequestSection = document.getElementById('pending-request-status');
    if (pendingRequestSection) {
        pendingRequestSection.style.display = 'none';
    }
}

/**
 * Handle refreshing the pending request status specifically for the status card
 */
async function handleRefreshPendingRequestStatus() {
    const refreshBtn = document.getElementById('refresh-request-status-btn');
    if (refreshBtn) {
        refreshBtn.disabled = true;
        refreshBtn.textContent = 'Checking...';
    }
    
    try {
        const result = await window.OpsieApi.getUserPendingJoinRequest();
        
        if (result.success && result.data) {
            displayPendingRequestStatus(result.data);
            
                         // Check if the status has changed from pending
             if (result.data.status === 'approved') {
                 showAuthSuccess('Your join request has been approved! Redirecting...');
                 setTimeout(() => {
                     hideAuthContainer();
                     
                     // Show the settings button since user now has a team
                     const settingsButton = document.getElementById('settings-button');
                     if (settingsButton) {
                         settingsButton.style.display = 'block';
                         console.log('Shown settings button - user now has a team');
                     }
                     
                     // Refresh the UI with new team info and show main UI sections
                     if (window.OpsieApi && window.OpsieApi.initTeamAndUserInfo) {
                         console.log('RefreshStatus: Initializing team info with callback');
                         window.OpsieApi.initTeamAndUserInfo(() => {
                             console.log('RefreshStatus: Team info initialized, showing main UI sections');
                             // Now that team info is loaded, show the main UI sections
                             if (typeof window.showMainUISections === 'function') {
                                 window.showMainUISections();
                             } else {
                                 console.log('RefreshStatus: showMainUISections not available, showing main content manually');
                                 showMainContent();
                             }
                         });
                     } else {
                         console.log('RefreshStatus: initTeamAndUserInfo not available, showing main content manually');
                         showMainContent();
                     }
                 }, 1500);
             } else if (result.data.status === 'rejected') {
                showAuthError('Your join request was rejected. Please try with a different team code.');
                hidePendingRequestStatus();
            }
                         } else {
            // No pending request found - check if user now has a team (request was approved)
            hidePendingRequestStatus();
            
            // Check if user now belongs to a team via API
            const userInfo = await window.OpsieApi.getUserInfo(true);
            console.log('Refresh: Current user info check:', userInfo);
            
            if (userInfo && userInfo.team_id) {
                                 // User has a team - request was approved, redirect to main UI
                showAuthSuccess('Your join request has been approved! Redirecting...');
                setTimeout(() => {
                    hideAuthContainer();
                    
                    // Show the settings button since user now has a team
                    const settingsButton = document.getElementById('settings-button');
                    if (settingsButton) {
                        settingsButton.style.display = 'block';
                        console.log('Shown settings button - user now has a team');
                    }
                    
                    // Refresh the UI with new team info and show main UI sections
                    if (window.OpsieApi && window.OpsieApi.initTeamAndUserInfo) {
                        console.log('Refresh: Initializing team info with callback');
                        window.OpsieApi.initTeamAndUserInfo(() => {
                            console.log('Refresh: Team info initialized, showing main UI sections');
                            // Now that team info is loaded, show the main UI sections
                            if (typeof window.showMainUISections === 'function') {
                                window.showMainUISections();
                            } else {
                                console.log('Refresh: showMainUISections not available, showing main content manually');
                                showMainContent();
                            }
                        });
                    } else {
                        console.log('Refresh: initTeamAndUserInfo not available, showing main content manually');
                        showMainContent();
                    }
                }, 1500);
             } else {
                 // Check team status via API to be sure
                 checkTeamMembershipAndRedirect();
             }
         }
    } catch (error) {
        console.error('Error refreshing pending request status:', error);
        showAuthError('Error checking request status. Please try again.');
    } finally {
        if (refreshBtn) {
            refreshBtn.disabled = false;
            refreshBtn.textContent = 'Refresh Status';
        }
    }
}

/**
 * Check team membership via API and redirect if user has a team
 */
async function checkTeamMembershipAndRedirect() {
    try {
        // Check if user has team membership via the API
        const userInfo = await window.OpsieApi.getUserInfo(true);
        console.log('CheckTeamMembership: Current user info check:', userInfo);
        
        if (userInfo && userInfo.team_id) {
            // User has a team - redirect to main UI
            showAuthSuccess('Your join request has been approved! Redirecting...');
            setTimeout(() => {
                hideAuthContainer();
                
                // Show the settings button since user now has a team
                const settingsButton = document.getElementById('settings-button');
                if (settingsButton) {
                    settingsButton.style.display = 'block';
                    console.log('Shown settings button - user now has a team');
                }
                
                // Refresh the UI with new team info and show main UI sections
                if (window.OpsieApi && window.OpsieApi.initTeamAndUserInfo) {
                    console.log('CheckTeamMembership: Initializing team info with callback');
                    window.OpsieApi.initTeamAndUserInfo(() => {
                        console.log('CheckTeamMembership: Team info initialized, showing main UI sections');
                        // Now that team info is loaded, show the main UI sections
                        if (typeof window.showMainUISections === 'function') {
                            window.showMainUISections();
                        } else {
                            console.log('CheckTeamMembership: showMainUISections not available, showing main content manually');
                            showMainContent();
                        }
                    });
                } else {
                    console.log('CheckTeamMembership: initTeamAndUserInfo not available, showing main content manually');
                    showMainContent();
                }
            }, 1500);
        } else {
            // User still has no team - request might have been rejected
            showAuthSuccess('Request status updated. You may need to check with your team admin.');
        }
    } catch (error) {
        console.error('Error checking team membership:', error);
        showAuthSuccess('Request status updated. You may need to check with your team admin.');
    }
}

/**
 * Handle refreshing join request status
 */
async function handleRefreshRequestStatus() {
    showTeamLoading('Checking request status...');
    
    try {
        const result = await window.OpsieApi.checkJoinRequestStatus();
        
        if (result.success) {
            if (result.approved) {
                showAuthSuccess('Your join request has been approved! Redirecting...');
                
                setTimeout(() => {
                    hideAuthContainer();
                    showMainContent();
                    
                    // Refresh the UI with new team info
                    if (window.OpsieApi && window.OpsieApi.initTeamAndUserInfo) {
                        window.OpsieApi.initTeamAndUserInfo(() => {
                            console.log('Team info refreshed after join approval');
                        });
                    }
                }, 1500);
            } else if (result.rejected) {
                showAuthError('Your join request was rejected. Please try with a different team code.');
            } else {
                showAuthSuccess('Your join request is still pending approval.');
            }
        } else {
            showAuthError(result.error || 'Failed to check request status');
        }
    } catch (error) {
        console.error('Error checking request status:', error);
        showAuthError('An error occurred while checking request status');
    } finally {
        hideTeamLoading();
    }
}

/**
 * Start polling for join request approval
 */
function startJoinRequestPolling() {
    let pollCount = 0;
    const maxPolls = 60; // 10 minutes (60 * 10 seconds)
    
    const pollInterval = setInterval(async () => {
        pollCount++;
        
                try {
            // Check user's current team status first
            const userInfo = await window.OpsieApi.getUserInfo(true);
            console.log('Polling - current user info check:', userInfo);
            
            if (userInfo && userInfo.team_id) {
                console.log('Polling detected user now has a team - redirecting');
                clearInterval(pollInterval);
                showAuthSuccess('Your join request has been approved! Redirecting...');
                
                setTimeout(() => {
                    console.log('Executing redirect after approval');
                    hideAuthContainer();
                    
                    // Show the settings button since user now has a team
                    const settingsButton = document.getElementById('settings-button');
                    if (settingsButton) {
                        settingsButton.style.display = 'block';
                        console.log('Shown settings button - user now has a team');
                    }
                    
                    // Refresh the UI with new team info and show main UI sections
                    if (window.OpsieApi && window.OpsieApi.initTeamAndUserInfo) {
                        console.log('Polling: Initializing team info with callback');
                        window.OpsieApi.initTeamAndUserInfo(() => {
                            console.log('Polling: Team info initialized, showing main UI sections');
                            // Now that team info is loaded, show the main UI sections
                            if (typeof window.showMainUISections === 'function') {
                                window.showMainUISections();
                            } else {
                                console.log('Polling: showMainUISections not available, showing main content manually');
                                showMainContent();
                            }
                        });
                    } else {
                        console.log('Polling: initTeamAndUserInfo not available, showing main content manually');
                        showMainContent();
                    }
                }, 1500);
            } else {
                // User still doesn't have a team, check join request status for rejection
                const result = await window.OpsieApi.getUserPendingJoinRequest();
                console.log('Polling - pending request check result:', result);
                
                if (!result.success || !result.data) {
                    // No pending request - might have been rejected, check latest request status
                    const statusResult = await window.OpsieApi.checkJoinRequestStatus();
                    if (statusResult.success && statusResult.data && statusResult.data.status === 'rejected') {
                        console.log('Polling detected rejected request');
                        clearInterval(pollInterval);
                        showAuthError('Your join request was rejected. Please try with a different team code.');
                    } else {
                        console.log('Polling - request status unclear:', statusResult);
                    }
                } else {
                    console.log('Polling - request still pending');
                }
            }
        } catch (error) {
            console.error('Error polling join request status:', error);
        }
        
        // Stop polling after max attempts
        if (pollCount >= maxPolls) {
            clearInterval(pollInterval);
            console.log('Stopped polling for join request approval after 10 minutes');
        }
    }, 10000); // Poll every 10 seconds
}

/**
 * Show team loading state
 */
function showTeamLoading(message = 'Loading...') {
    const loadingElement = document.getElementById('team-loading');
    if (loadingElement) {
        loadingElement.textContent = message;
        loadingElement.style.display = 'block';
    }
}

/**
 * Hide team loading state
 */
function hideTeamLoading() {
    const loadingElement = document.getElementById('team-loading');
    if (loadingElement) {
        loadingElement.style.display = 'none';
    }
}

/**
 * Show authentication error message
 */
function showAuthError(message) {
    const errorElement = document.getElementById('team-auth-error');
    if (errorElement) {
        errorElement.textContent = message;
        errorElement.style.display = 'block';
    }
    
    // Hide success message if showing
    hideAuthSuccess();
}

/**
 * Hide authentication error message
 */
function hideAuthError() {
    const errorElement = document.getElementById('team-auth-error');
    if (errorElement) {
        errorElement.style.display = 'none';
    }
}

/**
 * Show authentication success message
 */
function showAuthSuccess(message) {
    const successElement = document.getElementById('team-auth-success');
    if (successElement) {
        successElement.textContent = message;
        successElement.style.display = 'block';
    }
    
    // Hide error message if showing
    hideAuthError();
}

/**
 * Hide authentication success message
 */
function hideAuthSuccess() {
    const successElement = document.getElementById('team-auth-success');
    if (successElement) {
        successElement.style.display = 'none';
    }
}

/**
 * Hide authentication container
 */
function hideAuthContainer() {
    console.log('hideAuthContainer: Starting to hide auth containers');
    
    const authContainer = document.getElementById('auth-container');
    if (authContainer) {
        console.log('hideAuthContainer: Hiding auth-container');
        authContainer.style.display = 'none';
    } else {
        console.log('hideAuthContainer: auth-container not found');
    }
    
    const teamSelectView = document.getElementById('team-select-view');
    if (teamSelectView) {
        console.log('hideAuthContainer: Hiding team-select-view');
        teamSelectView.style.display = 'none';
    } else {
        console.log('hideAuthContainer: team-select-view not found');
    }
}

/**
 * Show main content
 */
function showMainContent() {
    console.log('showMainContent: Starting to show main content');
    
    const mainContent = document.getElementById('main-content');
    if (mainContent) {
        console.log('showMainContent: Found main-content, setting display to block');
        console.log('showMainContent: Current display style:', mainContent.style.display);
        mainContent.style.display = 'block';
        console.log('showMainContent: New display style:', mainContent.style.display);
    } else {
        console.log('showMainContent: main-content element not found!');
    }
}

/**
 * Handle team sign out - return to login
 */
async function handleTeamSignOut() {
    try {
        // Clear authentication data
        localStorage.removeItem('authToken');
        localStorage.removeItem('refreshToken');
        localStorage.removeItem('userId');
        localStorage.removeItem('currentTeamId');
        localStorage.removeItem('userInfo');
        localStorage.removeItem('teamInfo');
        
        // Hide team selection view
        const teamSelectView = document.getElementById('team-select-view');
        if (teamSelectView) {
            teamSelectView.style.display = 'none';
        }
        
        // Show authentication container with login view
        const authContainer = document.getElementById('auth-container');
        if (authContainer) {
            authContainer.style.display = 'flex';
        }
        
        // Reset to login view
        if (window.showLoginView) {
            window.showLoginView();
        }
        
        // Clear any error/success messages
        hideAuthError();
        hideAuthSuccess();
        
        // Clear form fields
        const joinCodeInput = document.getElementById('join-team-code');
        const teamNameInput = document.getElementById('create-team-name');
        if (joinCodeInput) joinCodeInput.value = '';
        if (teamNameInput) teamNameInput.value = '';
        
        console.log('User signed out from team selection');
        
    } catch (error) {
        console.error('Error during team sign out:', error);
        showAuthError('Error signing out. Please try again.');
    }
}

// Make functions available globally
window.showTeamSelectView = showTeamSelectView;
window.setupTeamSelectEventListeners = setupTeamSelectEventListeners;
window.handleJoinTeam = handleJoinTeam;
window.handleCreateTeam = handleCreateTeam;
window.handleRefreshRequestStatus = handleRefreshRequestStatus; 
window.hideAuthContainer = hideAuthContainer;
window.showMainContent = showMainContent;

// Store a reference to showMainUISections for easier access
if (typeof showMainUISections !== 'undefined') {
    window.showMainUISections = showMainUISections;
} 