/**
 * Popup Script for Opsie Email Assistant
 * Handles authentication and team management
 */

import { 
  signUp, 
  signIn, 
  signOut, 
  isAuthenticated, 
  getUserDetails, 
  createTeam, 
  joinTeam,
  leaveTeam,
  transferAdminRole,
  deleteTeam,
  getAllTeams,
  requestPasswordReset,
  resetPassword
} from '../utils/auth-service.js';

// Import Supabase config (add to top of file)
import { SUPABASE_URL, SUPABASE_KEY } from '../utils/supabase-config.js';

// DOM elements
const views = {
  login: null,
  signup: null,
  teamSelect: null,
  mainApp: null,
  resetPassword: null,
  resetToken: null
};

// Status message element
const statusMessage = document.getElementById('status-message');

// Add more detailed logging at the top of the file for debugging
console.log('Popup script initializing...');

// Check if we have auth data in storage 
chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId'], function(result) {
  console.log('Popup auth data check:', {
    hasAccessToken: !!result.accessToken,
    tokenFirstChars: result.accessToken ? result.accessToken.substring(0, 10) + '...' : 'none',
    tokenLength: result.accessToken ? result.accessToken.length : 0,
    hasUserId: !!result.userId,
    hasTeamId: !!result.currentTeamId
  });

  if (result.accessToken) {
    // Try to decode the JWT to see if it's expired
    try {
      const parts = result.accessToken.split('.');
      if (parts.length === 3) {
        const payload = JSON.parse(atob(parts[1]));
        console.log('Token payload:', {
          exp: payload.exp ? new Date(payload.exp * 1000).toISOString() : 'not found',
          iat: payload.iat ? new Date(payload.iat * 1000).toISOString() : 'not found',
          currentTime: new Date().toISOString(),
          isExpired: payload.exp ? (payload.exp * 1000 < Date.now()) : 'cannot determine'
        });
      } else {
        console.log('Token does not have the expected JWT format');
      }
    } catch (e) {
      console.error('Error decoding JWT token:', e);
    }
  }
});

// Initialize the popup
document.addEventListener('DOMContentLoaded', async function() {
  console.log('Popup DOM content loaded, initializing...');
  
  // Initialize the views object
  views.login = document.getElementById('login-view');
  views.signup = document.getElementById('signup-view');
  views.teamSelect = document.getElementById('team-select-view');
  views.mainApp = document.getElementById('main-app-view');
  views.resetPassword = document.getElementById('reset-password-view');
  views.resetToken = document.getElementById('reset-token-view');

  try {
    // Check if the access token is expired
    const { accessToken, userId } = await chrome.storage.sync.get(['accessToken', 'userId']);
    
    if (accessToken) {
      console.log('Access token found, checking validity...');
      
      // Check if token is expired
      let isExpired = false;
      try {
        const parts = accessToken.split('.');
        if (parts.length === 3) {
          const payload = JSON.parse(atob(parts[1]));
          isExpired = payload.exp ? (payload.exp * 1000 < Date.now()) : false;
        }
      } catch (e) {
        console.error('Error checking token expiration during init:', e);
        isExpired = true; // Assume expired if we can't parse it
      }
      
      if (isExpired) {
        console.log('Token is expired during initialization');
        handleTokenExpiration();
        return;
      }
      
      console.log('Token appears valid, getting user details');
      // Get the user details
      const userDetails = await getUserDetails(userId);
      
      console.log('User details response:', userDetails);
      
      if (!userDetails.success) {
        console.error('Failed to get user details:', userDetails.error);
        
        // Check if token expired
        if (userDetails.error && 
            (userDetails.error.code === 'PGRST301' || 
             (typeof userDetails.error === 'string' && userDetails.error.includes('expired')))) {
          handleTokenExpiration();
        } else {
          showStatus('Failed to retrieve your account details. Please try again.', 'error');
          showLoginView();
        }
        return;
      }
      
      // Check if user is in a team
      if (userDetails.data.team_id) {
        console.log('User has a team, showing main app view');
        await showMainAppView(userDetails.data);
      } else {
        console.log('User has no team, showing team selection view');
        showTeamSelectView(userDetails.data);
      }
    } else {
      console.log('No token found, showing login view');
      showLoginView();
    }
  } catch (error) {
    console.error('Error initializing popup:', error);
    showStatus('An error occurred while initializing the app. Please try again.', 'error');
    showLoginView();
  }
  
  // Set up event listeners
  setupEventListeners();
});

// Set up all event listeners
function setupEventListeners() {
  // Login form
  document.getElementById('login-button').addEventListener('click', handleLogin);
  
  // Signup form
  document.getElementById('signup-button').addEventListener('click', handleSignup);
  
  // Toggle between login and signup views
  document.getElementById('go-to-signup').addEventListener('click', () => showView('signup'));
  document.getElementById('go-to-login').addEventListener('click', () => showView('login'));
  
  // Password reset
  document.getElementById('go-to-reset-password').addEventListener('click', () => showView('resetPassword'));
  document.getElementById('go-back-to-login').addEventListener('click', () => showView('login'));
  document.getElementById('back-to-login').addEventListener('click', () => showView('login'));
  document.getElementById('already-have-token').addEventListener('click', () => {
    // Pre-fill the email field in the token view with the email from the reset request view
    const resetEmail = document.getElementById('reset-email').value.trim();
    if (resetEmail) {
      document.getElementById('reset-token-email').value = resetEmail;
    }
    showView('resetToken');
  });
  document.getElementById('send-reset-email-button').addEventListener('click', handlePasswordResetRequest);
  document.getElementById('reset-password-button').addEventListener('click', handlePasswordReset);
  
  // Team creation
  document.getElementById('create-team-button').addEventListener('click', handleCreateTeam);
  
  // Logout buttons
  document.getElementById('logout-button').addEventListener('click', handleLogout);
  document.getElementById('main-logout-button').addEventListener('click', handleLogout);
  
  // API Key saving
  document.getElementById('save-api-key-button').addEventListener('click', saveApiKey);
  
  // Team selection view events
  document.getElementById('request-join-team-button').addEventListener('click', handleRequestToJoinTeam);
  document.getElementById('refresh-requests-button').addEventListener('click', loadPendingJoinRequests);
}

// Show login view
function showLoginView() {
  showView('login');
}

// Show team selection view
async function showTeamSelectView(userData) {
  document.getElementById('user-email').textContent = userData.email;
  
  // Reset form fields
  document.getElementById('create-team-name').value = '';
  document.getElementById('create-team-organization').value = '';
  document.getElementById('create-team-invoice-email').value = '';
  document.getElementById('create-team-billing-street').value = '';
  document.getElementById('create-team-billing-city').value = '';
  document.getElementById('create-team-billing-region').value = '';
  document.getElementById('create-team-billing-country').value = '';
  document.getElementById('team-access-code-input').value = '';
  
  // Hide the pending request section by default
  document.getElementById('pending-request-section').style.display = 'none';
  
  // Check if the user has any pending join requests
  try {
    const requestStatus = await checkJoinRequestStatus();
    
    if (requestStatus.success && requestStatus.data.requests && requestStatus.data.requests.length > 0) {
      // Find pending requests
      const pendingRequests = requestStatus.data.requests.filter(req => req.status === 'pending');
      
      if (pendingRequests.length > 0) {
        // Show the pending request section
        const pendingSection = document.getElementById('pending-request-section');
        pendingSection.style.display = 'block';
        
        // Update the team name in the pending request section
        const teamName = pendingRequests[0].team ? pendingRequests[0].team.name : 'Unknown team';
        document.getElementById('requested-team-name').textContent = teamName;
      }
      
      // Check if any request has been approved
      if (requestStatus.data.hasApprovedRequest) {
        console.log('Found approved join request, getting user details');
        
        // Refresh user details to get the updated team ID
        const userDetailsResult = await getUserDetails(userData.id);
        
        if (userDetailsResult.success && userDetailsResult.data.team_id) {
          console.log('User has been added to a team, showing main app view');
          await showMainAppView(userDetailsResult.data);
          return;
        }
      }
    }
  } catch (error) {
    console.error('Error checking join request status:', error);
  }
  
  showView('teamSelect');
}

// Show main app view
async function showMainAppView(userData) {
  console.log('showMainAppView called with userData:', userData);
  
  // Update user information display
  document.getElementById('main-user-email').textContent = userData.email;
  document.getElementById('user-role').textContent = userData.role || 'member';
  
  // Get team name if available
  if (userData.teamDetails) {
    console.log('Team details from userData:', userData.teamDetails);
    console.log('Access code value:', userData.teamDetails.access_code);
    
    document.getElementById('team-name').textContent = userData.teamDetails.name;
    
    // Display team organization if available
    const organizationEl = document.getElementById('team-organization');
    organizationEl.textContent = userData.teamDetails.organization || '-';
    
    // Display invoice email if available
    const invoiceEmailEl = document.getElementById('team-invoice-email');
    invoiceEmailEl.textContent = userData.teamDetails.invoice_email || '-';
    
    // Format and display billing address if available
    const billingAddressEl = document.getElementById('team-billing-address');
    const hasBillingStreet = !!userData.teamDetails.billing_street;
    const hasBillingCity = !!userData.teamDetails.billing_city;
    const hasBillingRegion = !!userData.teamDetails.billing_region;
    const hasBillingCountry = !!userData.teamDetails.billing_country;
    
    if (hasBillingStreet || hasBillingCity || hasBillingRegion || hasBillingCountry) {
      let address = [];
      if (hasBillingStreet) address.push(userData.teamDetails.billing_street);
      if (hasBillingCity) address.push(userData.teamDetails.billing_city);
      if (hasBillingRegion) address.push(userData.teamDetails.billing_region);
      if (hasBillingCountry) address.push(userData.teamDetails.billing_country);
      billingAddressEl.textContent = address.join(', ');
    } else {
      billingAddressEl.textContent = '-';
    }
    
    // Display team access code if available
    const accessCodeEl = document.getElementById('team-access-code');
    accessCodeEl.textContent = userData.teamDetails.access_code || '-';
    console.log('Setting access code element text to:', accessCodeEl.textContent);
    
    // Set up the edit form with current values
    document.getElementById('edit-team-organization').value = userData.teamDetails.organization || '';
    document.getElementById('edit-team-invoice-email').value = userData.teamDetails.invoice_email || '';
    document.getElementById('edit-team-billing-street').value = userData.teamDetails.billing_street || '';
    document.getElementById('edit-team-billing-city').value = userData.teamDetails.billing_city || '';
    document.getElementById('edit-team-billing-region').value = userData.teamDetails.billing_region || '';
    document.getElementById('edit-team-billing-country').value = userData.teamDetails.billing_country || '';
    
    // Load and display team members for all users
    if (userData.team_id) {
      displayTeamMembers(userData.team_id);
    }
  } else {
    document.getElementById('team-name').textContent = `Team ID: ${userData.team_id}`;
    document.getElementById('team-organization').textContent = '-';
    document.getElementById('team-invoice-email').textContent = '-';
    document.getElementById('team-billing-address').textContent = '-';
    document.getElementById('team-access-code').textContent = '-';
    
    // Load and display team members if user has a team_id
    if (userData.team_id) {
      displayTeamMembers(userData.team_id);
    }
  }
  
  // Check for stored API key
  chrome.storage.sync.get(['openaiApiKey'], function(result) {
    if (result.openaiApiKey) {
      document.getElementById('openai-api-key').value = result.openaiApiKey;
    }
  });
  
  // Configure team management section based on user role
  const memberControls = document.getElementById('member-controls');
  const adminControls = document.getElementById('admin-controls');
  const editTeamDetailsButton = document.getElementById('edit-team-details-button');
  const joinRequestsSection = document.getElementById('join-requests-section');
  
  if (userData.role === 'admin') {
    // User is an admin, show admin controls and hide member controls
    adminControls.style.display = 'block';
    memberControls.style.display = 'none';
    
    // Show edit team details button and join requests section for admins only
    editTeamDetailsButton.style.display = 'block';
    joinRequestsSection.style.display = 'block';
    
    // Load pending join requests for admins
    await loadPendingJoinRequests();
    
    // Populate team members dropdown for admin transfer
    await loadTeamMembers(userData.team_id);
  } else {
    // User is a regular member, show member controls and hide admin controls
    memberControls.style.display = 'block';
    adminControls.style.display = 'none';
    
    // Hide edit team details button and join requests section for non-admin users
    editTeamDetailsButton.style.display = 'none';
    joinRequestsSection.style.display = 'none';
  }
  
  // Add event listeners for team management buttons
  setupTeamManagementListeners();
  
  // Add event listeners for team details editing
  setupTeamDetailsEditListeners();
  
  showView('mainApp');
}

// Show the specified view
function showView(viewName) {
  // Hide all views
  Object.values(views).forEach(view => {
    if (view) view.classList.remove('active-view');
  });
  
  // Show the selected view
  if (views[viewName]) {
    views[viewName].classList.add('active-view');
  } else {
    console.error(`View "${viewName}" not found!`);
  }
}

// Handle login form submission
async function handleLogin() {
  const email = document.getElementById('login-email').value.trim();
  const password = document.getElementById('login-password').value;
  
  if (!email || !password) {
    showStatus('Please enter both email and password', 'error');
    return;
  }
  
  showStatus('Logging in...', 'info');
  console.log('Attempting login for:', email);
  
  try {
    const result = await signIn(email, password);
    console.log('Login result:', result);
    
    if (result.success) {
      showStatus('Login successful!', 'success');
      
      // Get the user details
      console.log('User details from login:', result.userDetails);
      
      if (!result.userDetails) {
        console.error('No user details received after login. Fetching directly...');
        // Try to get user details directly
        const { userId } = await chrome.storage.sync.get(['userId']);
        console.log('Retrieved userId from storage:', userId);
        
        if (userId) {
          const userDetailsResult = await getUserDetails(userId);
          console.log('Fetched user details:', userDetailsResult);
          
          if (userDetailsResult.success) {
            const userDetails = userDetailsResult.data;
            
            // Store first_name and last_name in Chrome storage for use in signature
            if (userDetails) {
              const firstName = userDetails.first_name || '';
              const lastName = userDetails.last_name || '';
              
              console.log('Storing user name in Chrome storage:', {
                firstName,
                lastName
              });
              
              chrome.storage.sync.set({
                firstName: firstName,
                lastName: lastName
              });
            }
            
            // Check if user is in a team
            if (userDetails.team_id) {
              console.log('User has a team, showing main app view');
              await showMainAppView(userDetails);
            } else {
              console.log('User has no team, showing team selection view');
              showTeamSelectView(userDetails);
            }
          } else {
            console.error('Failed to fetch user details:', userDetailsResult.error);
            showStatus('Failed to fetch user details. Please try again.', 'error');
          }
        } else {
          console.error('No userId found in storage after login');
          showStatus('Login issue: No user ID found. Please try again.', 'error');
        }
      } else {
        // We have user details from the login result
        // Store first_name and last_name in Chrome storage for use in signature
        if (result.userDetails) {
          const firstName = result.userDetails.first_name || '';
          const lastName = result.userDetails.last_name || '';
          
          console.log('Storing user name in Chrome storage:', {
            firstName,
            lastName
          });
          
          chrome.storage.sync.set({
            firstName: firstName,
            lastName: lastName
          });
        }
        
        // Check if user is in a team
        if (result.userDetails.team_id) {
          console.log('User has a team, showing main app view');
          await showMainAppView(result.userDetails);
        } else {
          console.log('User has no team, showing team selection view');
          showTeamSelectView(result.userDetails);
        }
      }
    } else {
      console.error('Login failed:', result.error);
      showStatus(`Login failed: ${result.error}`, 'error');
    }
  } catch (error) {
    console.error('Exception during login:', error);
    showStatus(`Login error: ${error.message}`, 'error');
  }
}

// Handle signup form submission
async function handleSignup() {
  const email = document.getElementById('signup-email').value.trim();
  const password = document.getElementById('signup-password').value;
  const confirmPassword = document.getElementById('signup-confirm-password').value;
  const firstName = document.getElementById('signup-first-name').value.trim();
  const lastName = document.getElementById('signup-last-name').value.trim();
  
  if (!email || !password || !confirmPassword) {
    showStatus('Please fill out all required fields', 'error');
    return;
  }
  
  if (password !== confirmPassword) {
    showStatus('Passwords do not match', 'error');
    return;
  }
  
  if (password.length < 6) {
    showStatus('Password must be at least 6 characters', 'error');
    return;
  }
  
  showStatus('Creating account...', 'info');
  console.log('Attempting to sign up user with email:', email);
  
  try {
    const result = await signUp(email, password, firstName, lastName);
    console.log('Signup result received from auth-service:', result);
    
    if (result.success) {
      // Display warning if one exists
      if (result.warning) {
        console.warn('Signup warning:', result.warning);
        showStatus(`Account created with warning: ${result.warning}`, 'info');
      } else {
        // Check if the response data includes confirmation_sent_at, which indicates email verification is required
        if (result.data && result.data.confirmation_sent_at) {
          showStatus('Account created! Please check your email to verify your account before logging in.', 'success');
        } else {
          showStatus('Account created successfully! You can now log in.', 'success');
        }
      }
      
      // Log details about the created user from the response
      if (result.data && result.data.user) {
        console.log('User created with ID:', result.data.user.id);
        console.log('User email verification status:', result.data.user.email_confirmed_at ? 'Confirmed' : 'Not confirmed');
        
        // Store the user ID temporarily for debugging purposes
        chrome.storage.sync.set({
          'lastSignupUserId': result.data.user.id,
          'lastSignupTime': new Date().toISOString()
        }, function() {
          console.log('Stored signup debug information for reference');
        });
      } else {
        console.warn('User data missing in successful signup response');
      }
      
      // Clear form and switch to login view
      document.getElementById('signup-email').value = '';
      document.getElementById('signup-password').value = '';
      document.getElementById('signup-confirm-password').value = '';
      document.getElementById('signup-first-name').value = '';
      document.getElementById('signup-last-name').value = '';
      
      // Auto-fill login email
      document.getElementById('login-email').value = email;
      
      showView('login');
    } else {
      console.error('Signup failed with error:', result.error);
      
      // Check for specific error codes
      if (result.code === 'email_in_use') {
        showStatus('This email address is already registered. Please try logging in instead.', 'error');
        
        // Auto-fill login email for convenience
        document.getElementById('login-email').value = email;
        
        // Add a button to switch to login view after a short delay
        setTimeout(() => {
          const statusElem = document.getElementById('status-message');
          if (statusElem && statusElem.style.display !== 'none') {
            statusElem.innerHTML += '<br><br><a id="switch-to-login" style="color:white;text-decoration:underline;cursor:pointer;">Switch to Login</a>';
            document.getElementById('switch-to-login').addEventListener('click', () => showView('login'));
          }
        }, 500);
      } else {
        showStatus(`Signup failed: ${result.error}`, 'error');
      }
    }
  } catch (error) {
    console.error('Exception during signup process:', error);
    showStatus(`Signup error: ${error.message}`, 'error');
  }
}

// Handle team creation
async function handleCreateTeam() {
  const teamName = document.getElementById('create-team-name').value.trim();
  const organization = document.getElementById('create-team-organization').value.trim();
  const invoiceEmail = document.getElementById('create-team-invoice-email').value.trim();
  const billingStreet = document.getElementById('create-team-billing-street').value.trim();
  const billingCity = document.getElementById('create-team-billing-city').value.trim();
  const billingRegion = document.getElementById('create-team-billing-region').value.trim(); 
  const billingCountry = document.getElementById('create-team-billing-country').value.trim();
  
  if (!teamName) {
    showStatus('Please enter a team name', 'error');
    return;
  }
  
  // Validate invoice email if provided
  if (invoiceEmail && !isValidEmail(invoiceEmail)) {
    showStatus('Please enter a valid invoice email address', 'error');
    return;
  }
  
  showStatus('Creating team...', 'info');
  
  try {
    const result = await createTeam(teamName, {
      organization,
      invoiceEmail,
      billingStreet,
      billingCity,
      billingRegion,
      billingCountry
    });
    
    console.log('Team creation result:', result);
    
    if (result.success) {
      showStatus(`Team "${teamName}" created successfully!`, 'success');
      
      // After creating a team, refresh user details and show main app view
      const userDetails = await getUserDetails(result.data.teamId);
      
      if (userDetails.success) {
        await showMainAppView({
          ...userDetails.data,
          teamDetails: result.data
        });
      } else {
        console.error('Failed to get user details after team creation:', userDetails.error);
        showStatus('Team created but failed to load details. Please refresh.', 'error');
      }
    } else {
      console.error('Team creation failed:', result.error);
      showStatus(`Failed to create team: ${result.error}`, 'error');
    }
  } catch (error) {
    console.error('Exception during team creation:', error);
    showStatus(`Error creating team: ${error.message}`, 'error');
  }
}

// Helper function for email validation
function isValidEmail(email) {
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

// Handle logout
async function handleLogout() {
  showStatus('Logging out...', 'info');
  
  const result = await signOut();
  
  if (result.success) {
    showStatus('Logged out successfully', 'success');
    showLoginView();
  } else {
    showStatus(`Logout failed: ${result.error}`, 'error');
  }
}

// Save OpenAI API Key
function saveApiKey() {
  const apiKey = document.getElementById('openai-api-key').value.trim();
  
  if (!apiKey) {
    showStatus('Please enter an API key', 'error');
    return;
  }
  
  // Validate API key format (simple check for sk- prefix)
  if (!apiKey.startsWith('sk-')) {
    showStatus('Invalid API key format. OpenAI keys start with "sk-"', 'error');
    return;
  }
  
  chrome.storage.sync.set({ 'openaiApiKey': apiKey }, function() {
    console.log('API Key saved');
    showStatus('API Key saved successfully!', 'success');
  });
}

// Show status message
function showStatus(message, type = 'info') {
  console.log(`Status message (${type}): ${message}`);
  
  if (statusMessage) {
    statusMessage.textContent = message;
    statusMessage.style.display = 'block';
    
    // Remove existing classes
    statusMessage.classList.remove('status-success', 'status-error', 'status-info');
    
    // Add the appropriate class
    if (type === 'success') {
      statusMessage.classList.add('status-success');
    } else if (type === 'error') {
      statusMessage.classList.add('status-error');
    } else {
      statusMessage.classList.add('status-info');
    }
    
    // Hide the message after 5 seconds for success/info messages
    if (type !== 'error') {
      setTimeout(function() {
        if (statusMessage) {
          statusMessage.style.display = 'none';
        }
      }, 5000);
    }
  }
}

// Handle password reset request
async function handlePasswordResetRequest() {
  const resetEmail = document.getElementById('reset-email').value.trim();
  
  // Basic validation
  if (!resetEmail) {
    showStatus('Please enter your email address', 'error');
    return;
  }
  
  // Email format validation
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(resetEmail)) {
    showStatus('Please enter a valid email address', 'error');
    return;
  }
  
  showStatus('Sending reset token...', 'info');
  
  try {
    const result = await requestPasswordReset(resetEmail);
    
    if (result.success) {
      showStatus('Reset token sent to your email address', 'success');
      
      // Clear and pre-fill form fields
      document.getElementById('reset-token').value = '';
      document.getElementById('reset-token-email').value = resetEmail; // Pre-fill the email
      document.getElementById('new-password').value = '';
      document.getElementById('confirm-new-password').value = '';
      
      // Navigate to token reset view
      showView('resetToken');
    } else {
      showStatus(`Failed to send reset token: ${result.error}`, 'error');
    }
  } catch (error) {
    console.error('Error requesting password reset:', error);
    showStatus('An error occurred while requesting the password reset', 'error');
  }
}

// Handle password reset with token
async function handlePasswordReset() {
  const token = document.getElementById('reset-token').value.trim();
  const email = document.getElementById('reset-token-email').value.trim();
  const newPassword = document.getElementById('new-password').value;
  const confirmPassword = document.getElementById('confirm-new-password').value;
  
  // Basic validation
  if (!token) {
    showStatus('Please enter the reset token from your email', 'error');
    return;
  }
  
  if (!email) {
    showStatus('Please enter your email address', 'error');
    return;
  }
  
  // Email format validation
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    showStatus('Please enter a valid email address', 'error');
    return;
  }
  
  if (!newPassword) {
    showStatus('Please enter a new password', 'error');
    return;
  }
  
  if (newPassword !== confirmPassword) {
    showStatus('Passwords do not match', 'error');
    return;
  }
  
  // Password strength validation
  if (newPassword.length < 8) {
    showStatus('Password must be at least 8 characters long', 'error');
    return;
  }
  
  showStatus('Resetting password...', 'info');
  console.log('Starting password reset with token (first few chars):', token.substring(0, 5) + '...');
  
  try {
    const result = await resetPassword(token, newPassword, email);
    console.log('Password reset result:', result);
    
    if (result.success) {
      showStatus('Password reset successful! You can now log in with your new password', 'success');
      // Clear form fields for security
      document.getElementById('reset-token').value = '';
      document.getElementById('reset-token-email').value = '';
      document.getElementById('new-password').value = '';
      document.getElementById('confirm-new-password').value = '';
      
      // Auto-fill email field on login form for convenience
      document.getElementById('login-email').value = email;
      
      // Navigate to login view
      showView('login');
    } else {
      console.error('Failed to reset password:', result);
      
      // Enhanced error display with more details
      let errorMessage = 'Failed to reset password';
      
      if (result.status === 401) {
        errorMessage = 'Invalid or expired token. Please request a new reset token.';
      } else if (result.status === 400) {
        if (result.data && result.data.msg && result.data.msg.includes('Invalid token')) {
          errorMessage = 'Invalid token. Please ensure you\'ve copied the exact token from your email.';
        } else {
          errorMessage = 'The token or email is invalid. Make sure you are using the correct token and email address.';
        }
      } else if (result.error) {
        errorMessage = `Error: ${result.error}`;
      }
      
      showStatus(errorMessage, 'error');
    }
  } catch (error) {
    console.error('Exception during password reset:', error);
    showStatus('An unexpected error occurred while resetting the password. Please try again later.', 'error');
  }
}

// Set up event listeners for team management buttons
function setupTeamManagementListeners() {
  // Leave team button (for members)
  const leaveTeamButton = document.getElementById('leave-team-button');
  if (leaveTeamButton) {
    leaveTeamButton.addEventListener('click', handleLeaveTeam);
  }
  
  // Transfer admin button (for admins)
  const transferAdminButton = document.getElementById('transfer-admin-button');
  if (transferAdminButton) {
    transferAdminButton.addEventListener('click', handleTransferAdmin);
  }
  
  // Delete team button (for admins)
  const deleteTeamButton = document.getElementById('delete-team-button');
  if (deleteTeamButton) {
    deleteTeamButton.addEventListener('click', handleDeleteTeam);
  }
}

// Load team members for the dropdown
async function loadTeamMembers(teamId) {
  console.log('Loading team members for team:', teamId);
  const teamMembersSelect = document.getElementById('team-members-select');
  
  // Clear existing options (except the first one)
  const defaultOption = teamMembersSelect.options[0];
  teamMembersSelect.innerHTML = '';
  teamMembersSelect.appendChild(defaultOption);
  
  try {
    // Get current user ID to exclude from list
    const { userId } = await chrome.storage.sync.get(['userId']);
    
    // Get team members
    const result = await getTeamMembers(teamId);
    console.log('Team members result:', result);
    
    if (result.success && result.data) {
      // Add options for each member except the current user (admin)
      result.data.forEach(member => {
        if (member.id !== userId && member.role !== 'admin') {
          const option = document.createElement('option');
          option.value = member.id;
          option.textContent = member.email;
          teamMembersSelect.appendChild(option);
        }
      });
      
      if (teamMembersSelect.options.length <= 1) {
        // No other members to transfer to
        const option = document.createElement('option');
        option.value = '';
        option.textContent = 'No other members in team';
        option.disabled = true;
        teamMembersSelect.appendChild(option);
        
        // Disable transfer button
        document.getElementById('transfer-admin-button').disabled = true;
      } else {
        // Enable transfer button
        document.getElementById('transfer-admin-button').disabled = false;
      }
    } else {
      console.error('Failed to get team members:', result.error);
      showStatus('Failed to load team members', 'error');
    }
  } catch (error) {
    console.error('Error loading team members:', error);
    showStatus('Error loading team members: ' + error.message, 'error');
  }
}

// Handle leave team action
async function handleLeaveTeam() {
  console.log('Leave team button clicked');
  
  if (!confirm('Are you sure you want to leave this team? You will no longer have access to team data.')) {
    return;
  }
  
  try {
    const result = await leaveTeam();
    console.log('Leave team result:', result);
    
    if (result.success) {
      showStatus('You have left the team successfully', 'success');
      
      // Get updated user details
      const { userId } = await chrome.storage.sync.get(['userId']);
      const userDetails = await getUserDetails(userId);
      
      if (userDetails.success) {
        console.log('User details after leaving team:', userDetails.data);
        
        // Show the team select view since user no longer has a team
        showTeamSelectView(userDetails.data);
      } else {
        console.error('Failed to get user details after leaving team:', userDetails.error);
        showTeamSelectView({ email: document.getElementById('main-user-email').textContent });
      }
    } else {
      showStatus(`Failed to leave team: ${result.error}`, 'error');
    }
  } catch (error) {
    console.error('Error leaving team:', error);
    showStatus(`Error leaving team: ${error.message}`, 'error');
  }
}

// Handle transfer admin role action
async function handleTransferAdmin() {
  console.log('Transfer admin button clicked');
  
  const teamMembersSelect = document.getElementById('team-members-select');
  const newAdminId = teamMembersSelect.value;
  
  if (!newAdminId) {
    showStatus('Please select a team member to transfer admin rights to', 'error');
    return;
  }
  
  if (!confirm(`Are you sure you want to transfer admin rights to this user? You will become a regular member.`)) {
    return;
  }
  
  try {
    const result = await transferAdminRole(newAdminId);
    console.log('Transfer admin result:', result);
    
    if (result.success) {
      showStatus('Admin rights transferred successfully', 'success');
      
      // Get updated user details
      const { userId } = await chrome.storage.sync.get(['userId']);
      const userDetails = await getUserDetails(userId);
      
      if (userDetails.success) {
        console.log('User details after admin transfer:', userDetails.data);
        
        // Refresh the team members display
        refreshTeamMembers();
        
        // Show the updated main app view
        showMainAppView(userDetails.data);
      } else {
        console.error('Failed to get user details after admin transfer:', userDetails.error);
        
        // Just update the role display for now
        document.getElementById('user-role').textContent = 'member';
        
        // Hide admin controls, show member controls
        document.getElementById('admin-controls').style.display = 'none';
        document.getElementById('member-controls').style.display = 'block';
        
        // Hide edit team details button
        document.getElementById('edit-team-details-button').style.display = 'none';
        
        // Refresh team members display
        refreshTeamMembers();
      }
    } else {
      showStatus(`Failed to transfer admin rights: ${result.error}`, 'error');
    }
  } catch (error) {
    console.error('Error transferring admin rights:', error);
    showStatus(`Error transferring admin rights: ${error.message}`, 'error');
  }
}

// Handle delete team button click
async function handleDeleteTeam() {
  if (!confirm('Are you sure you want to delete this team? This action cannot be undone and will remove all team members.')) {
    return;
  }
  
  // Double-check with a more serious warning
  if (!confirm('WARNING: This will permanently delete the team and remove all members. Type "DELETE" to confirm.')) {
    return;
  }
  
  showStatus('Deleting team...', 'info');
  
  try {
    const result = await deleteTeam();
    console.log('Delete team result:', result);
    
    if (result.success) {
      showStatus('Team deleted successfully', 'success');
      
      // Notify content scripts in all tabs about team deletion
      chrome.tabs.query({}, function(tabs) {
        tabs.forEach(function(tab) {
          try {
            chrome.tabs.sendMessage(tab.id, { 
              action: 'teamDeleted'
            });
          } catch (err) {
            console.warn('Error sending message to tab:', tab.id, err);
          }
        });
      });
      
      // Show team selection view
      showTeamSelectView({email: document.getElementById('main-user-email').textContent});
    } else {
      console.error('Failed to delete team:', result.error);
      showStatus(`Failed to delete team: ${result.error}`, 'error');
    }
  } catch (error) {
    console.error('Exception deleting team:', error);
    showStatus('Error deleting team: ' + error.message, 'error');
  }
}

// Setup event listeners for team details editing
function setupTeamDetailsEditListeners() {
  const editButton = document.getElementById('edit-team-details-button');
  const saveButton = document.getElementById('save-team-details-button');
  const cancelButton = document.getElementById('cancel-team-edit-button');
  const displayView = document.getElementById('team-details-display');
  const editView = document.getElementById('team-details-edit');
  
  // Edit button shows the edit form
  editButton.addEventListener('click', function() {
    displayView.style.display = 'none';
    editView.style.display = 'block';
  });
  
  // Cancel button returns to display view without saving
  cancelButton.addEventListener('click', function() {
    displayView.style.display = 'block';
    editView.style.display = 'none';
  });
  
  // Save button updates team details
  saveButton.addEventListener('click', handleUpdateTeamDetails);
}

// Handle team details update
async function handleUpdateTeamDetails() {
  const organization = document.getElementById('edit-team-organization').value.trim();
  const invoiceEmail = document.getElementById('edit-team-invoice-email').value.trim();
  const billingStreet = document.getElementById('edit-team-billing-street').value.trim();
  const billingCity = document.getElementById('edit-team-billing-city').value.trim();
  const billingRegion = document.getElementById('edit-team-billing-region').value.trim();
  const billingCountry = document.getElementById('edit-team-billing-country').value.trim();
  
  // Validate invoice email if provided
  if (invoiceEmail && !isValidEmail(invoiceEmail)) {
    showStatus('Please enter a valid invoice email address', 'error');
    return;
  }
  
  showStatus('Updating team details...', 'info');
  
  try {
    // Get the current user and team info
    const { userId, currentTeamId } = await chrome.storage.sync.get(['userId', 'currentTeamId']);
    
    if (!userId || !currentTeamId) {
      showStatus('Session error. Please log in again.', 'error');
      return;
    }
    
    // Prepare the team details update object
    const teamDetails = {
      organization,
      invoice_email: invoiceEmail,
      billing_street: billingStreet,
      billing_city: billingCity,
      billing_region: billingRegion,
      billing_country: billingCountry
    };
    
    // Call the update team details function
    const result = await updateTeamDetails(currentTeamId, teamDetails);
    
    if (result.success) {
      showStatus('Team details updated successfully!', 'success');
      
      // Update the display view with new values
      document.getElementById('team-organization').textContent = organization || '-';
      document.getElementById('team-invoice-email').textContent = invoiceEmail || '-';
      
      // Format and display updated billing address
      const billingAddressEl = document.getElementById('team-billing-address');
      if (billingStreet || billingCity || billingRegion || billingCountry) {
        let address = [];
        if (billingStreet) address.push(billingStreet);
        if (billingCity) address.push(billingCity);
        if (billingRegion) address.push(billingRegion);
        if (billingCountry) address.push(billingCountry);
        billingAddressEl.textContent = address.join(', ');
      } else {
        billingAddressEl.textContent = '-';
      }
      
      // Switch back to display view
      document.getElementById('team-details-display').style.display = 'block';
      document.getElementById('team-details-edit').style.display = 'none';
    } else {
      showStatus(`Failed to update team details: ${result.error}`, 'error');
    }
  } catch (error) {
    console.error('Exception updating team details:', error);
    showStatus(`Error updating team details: ${error.message}`, 'error');
  }
}

// Function to update team details
async function updateTeamDetails(teamId, teamDetails) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({
      action: 'updateTeamDetails',
      teamId: teamId,
      teamDetails: teamDetails
    }, function(response) {
      resolve(response || { success: false, error: 'No response from background script' });
    });
  });
}

// Handle request to join team
async function handleRequestToJoinTeam() {
  console.log('Handling request to join team');
  
  // Get the access code from the input
  const accessCode = document.getElementById('team-access-code-input').value.trim().toUpperCase();
  
  if (!accessCode) {
    showStatus('Please enter a team access code', 'error');
    return;
  }
  
  showStatus('Submitting join request...', 'info');
  
  try {
    const result = await requestToJoinTeam(accessCode);
    
    if (result.success) {
      showStatus(result.message || 'Join request submitted successfully!', 'success');
      
      // Show the pending request section
      const pendingSection = document.getElementById('pending-request-section');
      pendingSection.style.display = 'block';
      
      // Update the team name in the pending request section
      const teamName = result.data.teamName || 'the team';
      document.getElementById('requested-team-name').textContent = teamName;
    } else {
      showStatus(`Failed to join team: ${result.error}`, 'error');
    }
  } catch (error) {
    console.error('Exception when requesting to join team:', error);
    showStatus(`Error: ${error.message}`, 'error');
  }
}

// Check join request status
async function checkJoinRequestStatus() {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({
      action: 'checkJoinRequestStatus'
    }, function(response) {
      resolve(response || { success: false, error: 'No response from background script' });
    });
  });
}

// Load pending join requests for admin
async function loadPendingJoinRequests() {
  try {
    const result = await getPendingJoinRequests();
    
    if (result.success) {
      const requestsList = document.getElementById('join-requests-list');
      if (!requestsList) {
        console.error('Join requests list element not found');
        return;
      }
      
      // Get the no requests message element
      let noRequestsMessage = document.getElementById('no-requests-message');
      
      // Clear the current list
      requestsList.innerHTML = '';
      
      if (result.data.length === 0) {
        // If the no requests message doesn't exist, create it
        if (!noRequestsMessage) {
          noRequestsMessage = document.createElement('p');
          noRequestsMessage.id = 'no-requests-message';
          noRequestsMessage.style.fontStyle = 'italic';
          noRequestsMessage.style.color = '#666';
          noRequestsMessage.style.margin = '5px';
          noRequestsMessage.textContent = 'No pending join requests';
        }
        
        // Show the no requests message
        requestsList.appendChild(noRequestsMessage);
      } else {
        // We have requests, so just add them to the list
        // No need to try removing the message since we cleared innerHTML already
        
        result.data.forEach(request => {
          const requestItem = createRequestItem(request);
          requestsList.appendChild(requestItem);
        });
      }
    } else {
      console.error('Failed to load join requests:', result.error);
    }
  } catch (error) {
    console.error('Exception when loading join requests:', error);
  }
}

// Create a join request item element
function createRequestItem(request) {
  const itemDiv = document.createElement('div');
  itemDiv.className = 'request-item';
  itemDiv.style.padding = '10px';
  itemDiv.style.borderBottom = '1px solid #eee';
  itemDiv.style.marginBottom = '5px';
  
  // Format the request date
  const requestDate = new Date(request.request_date);
  const formattedDate = requestDate.toLocaleString();
  
  // Get the user's name and email
  const userName = request.user.first_name && request.user.last_name
    ? `${request.user.first_name} ${request.user.last_name}`
    : 'Unknown User';
  
  const userEmail = request.user.email || 'No email';
  
  // Create the request content
  const contentDiv = document.createElement('div');
  contentDiv.innerHTML = `
    <div style="margin-bottom: 5px;">
      <strong>${userName}</strong> (${userEmail})
    </div>
    <div style="font-size: 12px; color: #666; margin-bottom: 10px;">
      Requested on ${formattedDate}
    </div>
  `;
  
  // Create the action buttons
  const buttonsDiv = document.createElement('div');
  buttonsDiv.style.display = 'flex';
  buttonsDiv.style.gap = '10px';
  
  // Approve button
  const approveButton = document.createElement('button');
  approveButton.textContent = 'Approve';
  approveButton.className = 'button button-success';
  approveButton.style.padding = '5px 10px';
  approveButton.style.fontSize = '12px';
  approveButton.style.flex = '1';
  approveButton.addEventListener('click', () => handleRequestResponse(request.id, true));
  
  // Reject button
  const rejectButton = document.createElement('button');
  rejectButton.textContent = 'Reject';
  rejectButton.className = 'button button-danger';
  rejectButton.style.padding = '5px 10px';
  rejectButton.style.fontSize = '12px';
  rejectButton.style.flex = '1';
  rejectButton.addEventListener('click', () => handleRequestResponse(request.id, false));
  
  // Add buttons to the button container
  buttonsDiv.appendChild(approveButton);
  buttonsDiv.appendChild(rejectButton);
  
  // Add content and buttons to the item
  itemDiv.appendChild(contentDiv);
  itemDiv.appendChild(buttonsDiv);
  
  return itemDiv;
}

// Handle join request response (approve/reject)
async function handleRequestResponse(requestId, approved) {
  console.log(`Handling join request response: ${approved ? 'Approve' : 'Reject'} request ${requestId}`);
  
  try {
    const result = await respondToJoinRequest(requestId, approved);
    console.log('Request response result:', result);
    
    if (result.success) {
      showStatus(`Join request ${approved ? 'approved' : 'rejected'} successfully`, 'success');
      
      // Refresh the pending requests list
      await loadPendingJoinRequests();
      
      // If the request was approved, refresh the team members list
      if (approved) {
        refreshTeamMembers();
      }
    } else {
      showStatus(`Failed to process join request: ${result.error}`, 'error');
    }
  } catch (error) {
    console.error('Error responding to join request:', error);
    showStatus(`Error processing join request: ${error.message}`, 'error');
  }
}

// Get pending join requests
async function getPendingJoinRequests() {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({
      action: 'getPendingJoinRequests'
    }, function(response) {
      resolve(response || { success: false, error: 'No response from background script' });
    });
  });
}

// Request to join a team with access code
async function requestToJoinTeam(accessCode) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({
      action: 'requestToJoinTeam',
      accessCode: accessCode
    }, function(response) {
      resolve(response || { success: false, error: 'No response from background script' });
    });
  });
}

// Respond to a join request (approve or reject)
async function respondToJoinRequest(requestId, approved) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({
      action: 'respondToJoinRequest',
      requestId: requestId,
      approved: approved
    }, function(response) {
      resolve(response || { success: false, error: 'No response from background script' });
    });
  });
}

// Function to handle token expiration and reconnection
async function handleTokenExpiration() {
  console.log('Handling token expiration');
  
  // Check if we have refresh token, userId and email stored
  const { refreshToken, userId, userEmail } = await chrome.storage.sync.get(['refreshToken', 'userId', 'userEmail']);
  
  // If we have a refresh token, try to use it first
  if (refreshToken) {
    console.log('Refresh token found, attempting to refresh session automatically');
    showStatus('Your session has expired. Attempting to reconnect automatically...', 'info');
    
    try {
      // Attempt to refresh the token
      const refreshResult = await refreshAccessToken(refreshToken);
      
      if (refreshResult.success) {
        console.log('Token refreshed successfully in popup');
        showStatus('Session refreshed successfully!', 'success');
        
        // Update the UI based on the current view
        const currentView = document.querySelector('.view[style*="display: block"]');
        if (currentView && currentView.id === 'login-view') {
          // If we're on the login view, transition to the main app view
          const userData = { id: userId, email: userEmail };
          await showMainAppView(userData);
        } else {
          // Otherwise, just stay on the current view with refreshed token
        }
        
        // Notify other components about successful token refresh
        chrome.runtime.sendMessage({ 
          action: 'authStateChanged', 
          event: 'tokenRefreshed',
          isLoggedIn: true
        });
        
        return true; // Indicate successful refresh
      } else {
        console.error('Failed to refresh token:', refreshResult.error);
        // Continue with manual login flow
      }
    } catch (error) {
      console.error('Error refreshing token:', error);
      // Continue with manual login flow
    }
  }
  
  // If we reach here, refresh token failed or wasn't available
  if (userId && userEmail) {
    showStatus('Your session has expired. Please log in again.', 'info');
    
    // Clear the expired token but keep userId and userEmail
    chrome.storage.sync.remove(['accessToken', 'refreshToken']);
    
    // Populate the login form with the user's email if available
    const loginEmailField = document.getElementById('login-email');
    if (loginEmailField && userEmail) {
      loginEmailField.value = userEmail;
    }
    
    // Show the login view
    showLoginView();
  } else {
    showStatus('Your session has expired. Please log in again.', 'error');
    showLoginView();
  }
  
  return false; // Indicate manual login is needed
}

// Function to refresh access token
async function refreshAccessToken(refreshToken) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({
      action: 'refreshAccessToken',
      refreshToken: refreshToken
    }, function(response) {
      resolve(response || { success: false, error: 'No response from background script' });
    });
  });
}

// Display team members in the team details section
async function displayTeamMembers(teamId) {
  console.log('Displaying team members for team:', teamId);
  const teamMembersList = document.getElementById('team-members-list');
  const noMembersMessage = document.getElementById('no-members-message');
  
  try {
    // Get current user ID and role to highlight the current user and determine if admin
    const { userId } = await chrome.storage.sync.get(['userId']);
    const userRole = document.getElementById('user-role').textContent.trim();
    const isAdmin = userRole === 'admin';
    
    // Get team members
    const result = await getTeamMembers(teamId);
    console.log('Team members result for display:', result);
    
    if (result.success && result.data) {
      // Clear the loading message
      teamMembersList.innerHTML = '';
      
      if (result.data.length === 0) {
        // No members found
        const noMembersElem = document.createElement('p');
        noMembersElem.textContent = 'No team members found';
        noMembersElem.style.fontStyle = 'italic';
        noMembersElem.style.color = '#666';
        noMembersElem.style.margin = '5px';
        teamMembersList.appendChild(noMembersElem);
      } else {
        // Add each member to the list
        result.data.forEach(member => {
          const memberItem = document.createElement('div');
          memberItem.className = 'team-member-item';
          
          // Highlight current user
          if (member.id === userId) {
            memberItem.classList.add('current-user');
          }
          
          const memberInfo = document.createElement('div');
          memberInfo.style.display = 'flex';
          memberInfo.style.justifyContent = 'space-between';
          memberInfo.style.alignItems = 'center';
          memberInfo.style.width = '100%';
          
          const memberEmail = document.createElement('span');
          memberEmail.textContent = member.email;
          
          const memberRole = document.createElement('span');
          memberRole.textContent = member.role === 'admin' ? 'Admin' : 'Member';
          memberRole.className = 'team-member-role';
          memberRole.classList.add(member.role === 'admin' ? 'role-admin' : 'role-member');
          
          memberInfo.appendChild(memberEmail);
          memberInfo.appendChild(memberRole);
          memberItem.appendChild(memberInfo);
          
          // Add remove button for admins to remove non-admin members (and not themselves)
          if (isAdmin && member.id !== userId && member.role !== 'admin') {
            const actionsDiv = document.createElement('div');
            actionsDiv.className = 'team-member-actions';
            
            const removeButton = document.createElement('button');
            removeButton.textContent = 'Remove';
            removeButton.className = 'team-member-remove-btn';
            
            // Add click event to handle member removal
            removeButton.addEventListener('click', () => handleRemoveTeamMember(member.id, member.email));
            
            actionsDiv.appendChild(removeButton);
            memberItem.appendChild(actionsDiv);
          }
          
          teamMembersList.appendChild(memberItem);
        });
      }
    } else {
      console.error('Failed to get team members for display:', result.error);
      noMembersMessage.textContent = 'Failed to load team members';
    }
  } catch (error) {
    console.error('Error displaying team members:', error);
    noMembersMessage.textContent = 'Error loading team members';
  }
}

// Function to get team members
async function getTeamMembers(teamId) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({
      action: 'getTeamMembers',
      teamId: teamId
    }, function(response) {
      resolve(response || { success: false, error: 'No response from background script' });
    });
  });
}

// Refresh team members list
async function refreshTeamMembers() {
  console.log('Refreshing team members list');
  try {
    const { currentTeamId } = await chrome.storage.sync.get(['currentTeamId']);
    if (currentTeamId) {
      // Update the team members display in the team details section
      displayTeamMembers(currentTeamId);
      
      // Also update the team members dropdown used for transferring admin rights
      loadTeamMembers(currentTeamId);
      
      console.log('Successfully refreshed both team members lists');
    } else {
      console.log('No team ID found, cannot refresh team members');
    }
  } catch (error) {
    console.error('Error refreshing team members:', error);
  }
}

// Handle removing a team member (admin only)
async function handleRemoveTeamMember(memberId, memberEmail) {
  console.log('Remove team member clicked for:', memberEmail, memberId);
  
  // Confirm before removing
  if (!confirm(`Are you sure you want to remove ${memberEmail} from the team? This action cannot be undone.`)) {
    return;
  }
  
  showStatus('Removing team member...', 'info');
  
  try {
    const result = await removeTeamMember(memberId);
    console.log('Remove team member result:', result);
    
    if (result.success) {
      showStatus(`${memberEmail} has been removed from the team`, 'success');
      
      // Refresh the team members list
      refreshTeamMembers();
    } else {
      console.error('Failed to remove team member:', result.error);
      showStatus(`Failed to remove team member: ${result.error}`, 'error');
    }
  } catch (error) {
    console.error('Error removing team member:', error);
    showStatus(`Error removing team member: ${error.message}`, 'error');
  }
}

// Function to remove a team member
async function removeTeamMember(memberId) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({
      action: 'removeTeamMember',
      memberId: memberId
    }, function(response) {
      resolve(response || { success: false, error: 'No response from background script' });
    });
  });
} 