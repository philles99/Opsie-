// Global variables
let currentEmailData = null;
let generateEmailSummary, generateReplySuggestion, generateContactSummary, searchContactEmails;
let isSummaryGenerationInProgress = false;
let isContactSummaryInProgress = false;
let isEmailSearchInProgress = false;
let summaryCache = null; // Reference to the cache from the API service
let contactCache = null; // Reference to the contact cache from the API service

// Load modules dynamically
async function loadModules() {
  try {
    const apiServiceModule = await import(chrome.runtime.getURL('utils/api-service.js'));
    
    generateEmailSummary = apiServiceModule.generateEmailSummary;
    generateReplySuggestion = apiServiceModule.generateReplySuggestion;
    generateContactSummary = apiServiceModule.generateContactSummary;
    searchContactEmails = apiServiceModule.searchContactEmails;
    
    // Get reference to the summary cache if available
    if (apiServiceModule.getSummaryCache) {
      summaryCache = apiServiceModule.getSummaryCache();
    }
    
    // Get reference to the contact cache if available
    if (apiServiceModule.getContactCache) {
      contactCache = apiServiceModule.getContactCache();
    }
    
    console.log('Sidebar modules loaded successfully');
    
    // Initialize after modules are loaded
    init();
  } catch (error) {
    console.error('Error loading sidebar modules:', error);
    showUnauthorizedMessage();
  }
}

// Initialize the sidebar
async function init() {
  console.log('Sidebar initialization started');
  
  // First check if user is authenticated and in a team
  console.log('Checking authentication status...');
  
  // Add storage listener to detect authentication changes
  chrome.storage.onChanged.addListener(function(changes, namespace) {
    if (namespace === 'sync') {
      const authChanged = changes.accessToken || changes.userId || changes.currentTeamId;
      
      if (authChanged) {
        console.log('Auth-related storage changed, rechecking authentication');
        handleAuthChange();
      }
    }
  });
  
  // Also listen for direct messages about auth state changes from background script
  chrome.runtime.onMessage.addListener(function(message, sender, sendResponse) {
    console.log('Sidebar received message:', message);
    
    if (message.action === 'authStateChanged') {
      console.log('Auth state change notification received:', message.event);
      
      // Handle different auth events
      switch (message.event) {
        case 'tokenRefreshed':
          console.log('Token was refreshed successfully');
          // No need to refresh the UI, just log the successful refresh
          showStatusMessage('Session refreshed automatically', 'info');
          break;
          
        case 'tokenRefreshFailed':
        case 'tokenExpired':
          console.log('Token expired or refresh failed, showing auth required message');
          showStatusMessage('Your session has expired. Please log in again.', 'error');
          showUnauthorizedMessage();
          break;
          
        default:
          // For any other auth state change, recheck authentication
          handleAuthChange();
      }
      
      // Acknowledge the message
      if (sendResponse) {
        sendResponse({ received: true });
      }
      return true;
    }
  });
  
  try {
    const isAuthorized = await checkAuthentication();
    console.log('Authentication check result:', isAuthorized);
    
    if (!isAuthorized) {
      console.log('User not authorized, showing unauthorized message');
      showUnauthorizedMessage();
      return;
    }
    
    console.log('User is authenticated and in a team, setting up authorized UI');
    setupAuthorizedUI();
  } catch (error) {
    console.error('Error during sidebar initialization:', error);
    showUnauthorizedMessage();
  }
}

// Setup all UI elements and event listeners for authorized users
function setupAuthorizedUI() {
  console.log('User is authorized, setting up event listeners');

  // Set up message listener for communication with the content script
  window.addEventListener('message', function(event) {
    console.log('Message received in sidebar:', event.data);
    
    if (event.data.type === 'EMAIL_DATA') {
      console.log('EMAIL_DATA message received with data:', {
        sender: event.data.emailData.sender,
        subject: event.data.emailData.subject,
        timestamp: event.data.emailData.timestamp
      });
      
      currentEmailData = event.data.emailData;
      console.log('Email data stored in currentEmailData');
      
      updateEmailInfo(currentEmailData);
    } else if (event.data.type === 'REPLY_INSERT_RESULT') {
      // Handle the result of trying to insert the reply
      if (event.data.success) {
        showStatusMessage('Reply inserted successfully!', 'success');
      } else {
        showStatusMessage('Failed to insert reply: ' + event.data.error + '. You can copy and paste it manually.', 'error');
      }
    }
  });
  
  // Send a message to content script to notify sidebar is loaded
  console.log('Notifying content script that sidebar is loaded');
  window.parent.postMessage({ 
    type: 'SIDEBAR_LOADED'
  }, '*');
  
  // Set up close button
  document.getElementById('close-button').addEventListener('click', function() {
    closeSidebar();
  });
  
  // Set up save button
  document.getElementById('save-email-button').addEventListener('click', function() {
    saveEmailToDatabase();
  });
  
  // Set up mark as handled button
  document.getElementById('mark-handled-button').addEventListener('click', function() {
    markEmailAsHandled();
  });
  
  // Set up generate summary button
  document.getElementById('generate-summary-button').addEventListener('click', function() {
    if (currentEmailData) {
      generateAndDisplaySummary(currentEmailData);
    } else {
      showStatusMessage('No email data available to analyze', 'error');
    }
  });
  
  // Set up generate contact summary button
  document.getElementById('generate-contact-button').addEventListener('click', function() {
    if (currentEmailData && (currentEmailData.sender.name || currentEmailData.sender.email)) {
      generateAndDisplayContactSummary();
    } else {
      showStatusMessage('No contact information available to analyze', 'error');
    }
  });
  
  // Set up reply button
  document.getElementById('generate-reply-button').addEventListener('click', function() {
    generateReply();
  });
  
  // Set up email search button
  const searchButton = document.getElementById('email-search-button');
  const searchInput = document.getElementById('email-search-input');
  
  if (searchButton) {
    searchButton.addEventListener('click', function() {
      const query = searchInput.value.trim();
      if (query) {
        performEmailSearch(query);
        // Clear the search input after search is performed
        searchInput.value = '';
      } else {
        showStatusMessage('Please enter a search query', 'error');
      }
    });
  }
  
  // Add enter key event listener for search input
  if (searchInput) {
    searchInput.addEventListener('keypress', function(e) {
      if (e.key === 'Enter') {
        const query = searchInput.value.trim();
        if (query) {
          performEmailSearch(query);
          // Clear the search input after search is performed
          searchInput.value = '';
        } else {
          showStatusMessage('Please enter a search query', 'error');
        }
      }
    });
  }
  
  // Set up listeners for the reply options to save preferences
  const toneSelect = document.getElementById('reply-tone');
  const lengthSelect = document.getElementById('reply-length');
  const languageSelect = document.getElementById('reply-language');
  
  if (toneSelect) {
    toneSelect.addEventListener('change', function() {
      saveReplyPreferences();
    });
  }
  
  if (lengthSelect) {
    lengthSelect.addEventListener('change', function() {
      saveReplyPreferences();
    });
  }
  
  if (languageSelect) {
    languageSelect.addEventListener('change', function() {
      saveReplyPreferences();
    });
  }
  
  // Set up event listeners for the handling modal
  setupHandlingModalListeners();
  
  // Set up event listeners for the notes section
  setupNotesEventListeners();
  
  // Load saved reply preferences
  loadReplyPreferences();
  
  console.log('Sidebar initialization completed');
}

// Handle authentication state changes
async function handleAuthChange() {
  const wasAuthorized = document.querySelector('.unauthorized-message') === null;
  const isAuthorized = await checkAuthentication();
  
  console.log('Auth change detected - Was authorized:', wasAuthorized, 'Is authorized now:', isAuthorized);
  
  // Notify content script about authentication state change
  window.parent.postMessage({
    type: 'AUTH_STATE_CHANGED',
    isLoggedIn: isAuthorized
  }, '*');
  
  // If authorization status changed, refresh the UI
  if (wasAuthorized !== isAuthorized) {
    console.log('Authorization status changed, refreshing UI');
    
    if (isAuthorized) {
      // User is now authorized, setup proper UI
      const mainContent = document.querySelector('.container');
      if (mainContent && mainContent.querySelector('.unauthorized-message')) {
        console.log('Removing unauthorized message and setting up UI');
        // Clear unauthorized message
        mainContent.innerHTML = `
          <div class="header">
            <h1>Opsie Email Assistant</h1>
            <button id="close-button" class="close-button">×</button>
          </div>
          <div class="content">
            <div id="email-info" class="section">
              <p>Waiting for email data...</p>
            </div>
            <div class="section">
              <h2>AI Summary</h2>
              <button id="generate-summary-button" class="action-button">Generate AI Summary</button>
              <ul id="summary-items" class="summary-list">
                <li>Click "Generate AI Summary" to analyze this email</li>
              </ul>
              <div class="urgency-container">
                <p>Urgency: <span id="urgency-score">-</span>/10</p>
                <div class="urgency-meter">
                  <div id="urgency-fill" class="urgency-fill"></div>
                </div>
              </div>
            </div>
            <div class="section">
              <h2>Contact Info</h2>
              <button id="generate-contact-button" class="action-button">Get Contact Summary</button>
              <ul id="contact-items" class="contact-list">
                <li>Click "Get Contact Summary" to view previous interactions with this sender</li>
              </ul>
            </div>
            <div class="section">
              <button id="generate-reply-button" class="action-button">Generate Reply</button>
              <button id="save-email-button" class="action-button">Save Email</button>
            </div>
          </div>
          <div id="status-message" class="status-message"></div>
        `;
        
        setupAuthorizedUI();
      }
    } else {
      // User is no longer authorized, show unauthorized message
      showUnauthorizedMessage();
    }
  }
}

// Check if user is authenticated and in a team
async function checkAuthentication() {
  console.log('sidebar.js: Starting authentication check');
  
  try {
    // Check authentication status
    console.log('sidebar.js: Sending isAuthenticated request to background script');
    const authResponse = await new Promise(resolve => {
      chrome.runtime.sendMessage({ action: 'isAuthenticated' }, response => {
        console.log('sidebar.js: Received authentication response:', response);
        resolve(response);
      });
    });
    
    if (!authResponse || !authResponse.success) {
      console.error('sidebar.js: Authentication check failed to return valid response');
      return false;
    }
    
    if (!authResponse.authenticated) {
      console.log('sidebar.js: User is not authenticated according to background script');
      return false;
    }
    
    // Check if user has a team
    console.log('sidebar.js: User is authenticated, checking for team ID');
    const storage = await new Promise(resolve => {
      chrome.storage.sync.get(['userId', 'currentTeamId'], result => {
        console.log('sidebar.js: Storage values retrieved:', {
          userId: result.userId ? 'exists' : 'missing',
          currentTeamId: result.currentTeamId ? 'exists' : 'missing',
          rawUserId: result.userId,
          rawTeamId: result.currentTeamId
        });
        resolve(result);
      });
    });
    
    // Strict validation - must have both userId and currentTeamId
    if (!storage.userId) {
      console.error('sidebar.js: User ID missing from storage - authentication failed');
      return false;
    }
    
    if (!storage.currentTeamId) {
      console.error('sidebar.js: User authenticated but has no team - authorization failed');
      return false;
    }
    
    // Verify these are not the default test IDs (for extra security)
    if (storage.userId === '00000000-0000-0000-0000-000000000000') {
      console.error('sidebar.js: Default user ID detected - authentication failed');
      return false;
    }
    
    if (storage.currentTeamId === '11111111-1111-1111-1111-111111111111') {
      console.error('sidebar.js: Default team ID detected - authentication failed');
      return false;
    }
    
    // Get user details to retrieve first_name and last_name
    try {
      console.log('sidebar.js: Getting user details for signature');
      const userDetailsResponse = await new Promise(resolve => {
        chrome.runtime.sendMessage({ 
          action: 'getUserDetails',
          userId: storage.userId 
        }, response => {
          console.log('sidebar.js: Received user details response:', response);
          resolve(response);
        });
      });
      
      if (userDetailsResponse && userDetailsResponse.success && userDetailsResponse.data) {
        // Store first_name and last_name in Chrome storage for use in signature
        const firstName = userDetailsResponse.data.first_name || '';
        const lastName = userDetailsResponse.data.last_name || '';
        
        console.log('sidebar.js: Storing user name in Chrome storage:', {
          firstName,
          lastName
        });
        
        chrome.storage.sync.set({
          firstName: firstName,
          lastName: lastName
        });
      } else {
        console.warn('sidebar.js: Failed to get user details, but continuing authentication');
      }
    } catch (userDetailsError) {
      console.error('sidebar.js: Error getting user details:', userDetailsError);
      // Continue authentication - this is not critical
    }
    
    console.log('sidebar.js: Authentication check passed: User is authenticated and has a team');
    return true;
  } catch (error) {
    console.error('sidebar.js: Error checking authentication:', error);
    return false;
  }
}

// Show message for unauthorized users
function showUnauthorizedMessage() {
  // Replace sidebar content with unauthorized message
  const mainContent = document.querySelector('.container');
  if (mainContent) {
    mainContent.innerHTML = `
      <div class="unauthorized-message" style="text-align: center; padding: 30px; color: #555;">
        <h2>Authentication Required</h2>
        <p>Please log in and join a team to use Opsie Email Assistant.</p>
        <button id="login-button" style="
          padding: 8px 16px;
          background-color: #1a3d5c;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          margin-top: 20px;
        ">Open Login Page</button>
      </div>
    `;
    
    // Add event listener to the login button
    document.getElementById('login-button').addEventListener('click', function() {
      chrome.runtime.sendMessage({ action: 'openPopup' });
    });
  }
}

// Close the sidebar by sending a message to the content script
function closeSidebar() {
  window.parent.postMessage({ type: 'CLOSE_SIDEBAR' }, '*');
}

// Update the email information display
async function updateEmailInfo(emailData) {
  if (!emailData) return;
  
  const emailInfoElement = document.getElementById('email-info');
  
  // Format the date
  const date = new Date(emailData.timestamp);
  const formattedDate = date.toLocaleString();
  
  // Check if we have thread history
  const hasThreadHistory = emailData.threadHistory && emailData.threadHistory.length > 0;
  const threadLabel = hasThreadHistory ? 
    `<span class="thread-label">(Thread with ${emailData.threadHistory.length} messages)</span>` : '';
  
  // Update the email info section
  emailInfoElement.innerHTML = `
    <h3>${emailData.subject || 'No Subject'} ${threadLabel}</h3>
    <p class="email-sender">From: ${emailData.sender.name} (${emailData.sender.email})</p>
    <p class="email-timestamp">Received: ${formattedDate}</p>
    <p>${emailData.message ? emailData.message.substring(0, 150) + '...' : 'No message body'}</p>
  `;
  
  // Add thread history if available
  if (hasThreadHistory) {
    const threadContainer = document.createElement('div');
    threadContainer.className = 'thread-container';
    
    const threadToggle = document.createElement('button');
    threadToggle.className = 'thread-toggle';
    threadToggle.textContent = 'Show Thread History';
    threadToggle.onclick = function() {
      const threadHistory = document.getElementById('thread-history');
      if (threadHistory.style.display === 'none') {
        threadHistory.style.display = 'block';
        this.textContent = 'Hide Thread History';
      } else {
        threadHistory.style.display = 'none';
        this.textContent = 'Show Thread History';
      }
    };
    
    const threadHistory = document.createElement('div');
    threadHistory.id = 'thread-history';
    threadHistory.className = 'thread-history';
    threadHistory.style.display = 'none'; // Hidden by default
    
    // Add each previous message to the thread history
    emailData.threadHistory.forEach((message, index) => {
      const messageDate = new Date(message.timestamp);
      const messageFormattedDate = messageDate.toLocaleString();
      
      const messageElement = document.createElement('div');
      messageElement.className = 'thread-message';
      messageElement.innerHTML = `
        <div class="thread-message-header">
          <strong>${message.sender.name}</strong> (${message.sender.email})
          <span class="thread-message-date">${messageFormattedDate}</span>
        </div>
        <div class="thread-message-content">
          ${message.message ? message.message.substring(0, 150) + '...' : 'No message content'}
        </div>
      `;
      
      threadHistory.appendChild(messageElement);
    });
    
    threadContainer.appendChild(threadToggle);
    threadContainer.appendChild(threadHistory);
    emailInfoElement.appendChild(threadContainer);
  }
  
  // Check if this message is already saved in the database
  const alreadySavedMessage = document.getElementById('already-saved-message');
  const savedByInfo = document.getElementById('saved-by-info');
  
  // Check for handling status
  const handlingStatusMessage = document.getElementById('handling-status-message');
  const handledByInfo = document.getElementById('handled-by-info');
  const handlingNoteDisplay = document.getElementById('handling-note-display');
  
  if (emailData.existingMessage && emailData.existingMessage.exists) {
    // Email already exists in the database
    console.log('Email already saved in database:', emailData.existingMessage);
    
    // Format the saved date
    let savedDate = 'unknown date';
    try {
      if (emailData.existingMessage.savedAt) {
        savedDate = new Date(emailData.existingMessage.savedAt).toLocaleString();
      }
    } catch (e) {
      console.error('Error formatting saved date:', e);
    }
    
    // Get who saved it
    let savedByName = 'Unknown user';
    if (emailData.existingMessage.user && emailData.existingMessage.user.name) {
      savedByName = emailData.existingMessage.user.name;
    }
    
    // Show the message
    if (alreadySavedMessage) {
      alreadySavedMessage.style.display = 'block';
    }
    
    // Update the saved by info
    if (savedByInfo) {
      savedByInfo.textContent = `Message saved by ${savedByName} at ${savedDate}`;
    }
    
    // Disable the save button as it's already saved
    const saveButton = document.getElementById('save-email-button');
    if (saveButton) {
      saveButton.disabled = true;
      saveButton.textContent = 'Already Saved';
      saveButton.style.backgroundColor = '#999';
    }
    
    // Check if there is existing summary or urgency data and display it
    if (emailData.existingMessage.summary) {
      console.log('Found existing summary in saved message:', emailData.existingMessage.summary);
      
      try {
        // Attempt to parse the summary into an array of items
        let summaryItems;
        
        // If the summary is already an array, use it directly
        if (Array.isArray(emailData.existingMessage.summary)) {
          summaryItems = emailData.existingMessage.summary;
        } 
        // If summary is a string, try to split it by common separators
        else if (typeof emailData.existingMessage.summary === 'string') {
          // Check if it's stored as a pipe-separated string (common format)
          if (emailData.existingMessage.summary.includes('|')) {
            summaryItems = emailData.existingMessage.summary.split('|').map(item => item.trim());
          } 
          // Check if it might be stored as newlines
          else if (emailData.existingMessage.summary.includes('\n')) {
            summaryItems = emailData.existingMessage.summary.split('\n').map(item => item.trim());
          }
          // Default to treating it as a single item
          else {
            summaryItems = [emailData.existingMessage.summary];
          }
        }
        
        // If we have valid summary items, display them
        if (summaryItems && summaryItems.length > 0) {
          // Update currentEmailData with the summary for consistency
          currentEmailData.summary = emailData.existingMessage.summary;
          
          // Display the summary
          displayEmailSummary(summaryItems);
          console.log('Displayed existing summary from database');
        }
      } catch (summaryError) {
        console.error('Error parsing existing summary:', summaryError);
      }
    }
    
    // Check if there is urgency data and display it
    if (emailData.existingMessage.urgency !== null && emailData.existingMessage.urgency !== undefined) {
      console.log('Found existing urgency in saved message:', emailData.existingMessage.urgency);
      
      // Update currentEmailData with the urgency for consistency
      currentEmailData.urgency = emailData.existingMessage.urgency;
      
      // Display the urgency score
      displayUrgencyScore(emailData.existingMessage.urgency);
      console.log('Displayed existing urgency from database');
    }
    
    // Check for handling status
    if (emailData.existingMessage.handling) {
      console.log('Email is already handled:', emailData.existingMessage.handling);
      
      // Format the handled date
      let handledDate = 'unknown date';
      try {
        if (emailData.existingMessage.handling.handledAt) {
          handledDate = new Date(emailData.existingMessage.handling.handledAt).toLocaleString();
        }
      } catch (e) {
        console.error('Error formatting handled date:', e);
      }
      
      // Get who handled it
      let handledByName = 'Unknown user';
      if (emailData.existingMessage.handling.handledBy && emailData.existingMessage.handling.handledBy.name) {
        handledByName = emailData.existingMessage.handling.handledBy.name;
      }
      
      // Show the handling message
      if (handlingStatusMessage) {
        handlingStatusMessage.style.display = 'block';
      }
      
      // Update the handled by info
      if (handledByInfo) {
        handledByInfo.textContent = `Marked as handled by ${handledByName} at ${handledDate}`;
      }
      
      // Show handling note if available
      if (handlingNoteDisplay) {
        const note = emailData.existingMessage.handling.handlingNote;
        if (note) {
          handlingNoteDisplay.textContent = `Note: "${note}"`;
          handlingNoteDisplay.style.display = 'block';
        } else {
          handlingNoteDisplay.style.display = 'none';
        }
      }
      
      // Update the button to show it's already handled
      const handleButton = document.getElementById('mark-handled-button');
      if (handleButton) {
        handleButton.disabled = true;
        handleButton.textContent = 'Already Handled';
        handleButton.style.backgroundColor = '#999';
      }
    } else {
      // Email saved but not handled yet
      if (handlingStatusMessage) {
        handlingStatusMessage.style.display = 'none';
      }
      
      // Reset the handle button
      const handleButton = document.getElementById('mark-handled-button');
      if (handleButton) {
        handleButton.disabled = false;
        handleButton.textContent = 'Mark as Handled';
        handleButton.style.backgroundColor = '#ff9800';
      }
    }
    
    // Load notes for this email
    loadNotesForCurrentEmail();
  } else {
    // Email not saved yet
    if (alreadySavedMessage) {
      alreadySavedMessage.style.display = 'none';
    }
    
    // Hide handling status
    if (handlingStatusMessage) {
      handlingStatusMessage.style.display = 'none';
    }
    
    // Reset the save button
    const saveButton = document.getElementById('save-email-button');
    if (saveButton) {
      saveButton.disabled = false;
      saveButton.textContent = 'Save Email to Database';
      saveButton.style.backgroundColor = '#4CAF50';
    }
    
    // Disable handle button since email needs to be saved first
    const handleButton = document.getElementById('mark-handled-button');
    if (handleButton) {
      handleButton.disabled = true;
      handleButton.textContent = 'Save Email First';
      handleButton.style.backgroundColor = '#999';
    }
    
    // Hide notes section since the email is not saved
    document.getElementById('notes-section').style.display = 'none';
  
  // Reset summary and urgency displays for new email
  resetSummaryAndUrgency();
  }
  
  // Reset contact summary for new email
  resetContactSummary();
  
  // Check if we have a cached summary for this email (only if we didn't already display one from the DB)
  if (!(emailData.existingMessage && emailData.existingMessage.exists && 
      (emailData.existingMessage.summary || emailData.existingMessage.urgency !== undefined))) {
  checkAndDisplayCachedSummary(emailData);
  }
  
  // Check if we have a cached contact summary for this contact
  checkAndDisplayCachedContactSummary(emailData.sender);
}

// Reset summary and urgency displays
function resetSummaryAndUrgency() {
  const summaryItemsElement = document.getElementById('summary-items');
  summaryItemsElement.innerHTML = '<li>Click "Generate AI Summary" to analyze this email</li>';
  
  const urgencyScoreElement = document.getElementById('urgency-score');
  urgencyScoreElement.textContent = '-';
  
  const urgencyFillElement = document.getElementById('urgency-fill');
  urgencyFillElement.style.width = '0%';
  urgencyFillElement.style.backgroundColor = '#e0e0e0';
}

// Reset contact summary
function resetContactSummary() {
  const contactItemsElement = document.getElementById('contact-items');
  contactItemsElement.innerHTML = '<li>Click "Get Contact Summary" to view previous interactions with this sender</li>';
}

// Check for cached contact summary and display if available
function checkAndDisplayCachedContactSummary(contact) {
  if (!contactCache) {
    console.log('Contact cache is not available');
    return;
  }
  
  try {
    // Generate a cache key based on the contact information
    const cacheKey = `${contact.name || ''}-${contact.email || ''}-*`;
    
    console.log('Checking contact cache for key:', cacheKey);
    
    // Since we can't use wildcards directly, we need to iterate through the cache
    let cachedResult = null;
    
    // Iterate through the cache keys to find a match
    contactCache.forEach((value, key) => {
      const keyPrefix = `${contact.name || ''}-${contact.email || ''}-`;
      if (key.startsWith(keyPrefix)) {
        cachedResult = value;
        console.log('Found cached contact summary for:', key);
      }
    });
    
    if (cachedResult) {
      displayContactSummary(cachedResult.summaryItems, cachedResult.messageCount);
      showStatusMessage('Loaded cached contact summary', 'info');
    } else {
      console.log('No cached contact summary found');
    }
  } catch (error) {
    console.error('Error checking cached contact summary:', error);
  }
}

// Check for cached summary and display if available
function checkAndDisplayCachedSummary(emailData) {
  // If we don't have access to the cache, we can't check
  if (!summaryCache) {
    console.log('Summary cache is not available');
    return;
  }
  
  try {
    // Generate a cache key similar to what's used in api-service.js
    const threadHistoryLength = emailData.threadHistory ? emailData.threadHistory.length : 0;
    const cacheKey = `${emailData.subject}-${emailData.sender.email}-${threadHistoryLength}`;
    
    console.log('Checking cache for key:', cacheKey);
    
    // Check if we have this cache entry
    if (summaryCache.has(cacheKey)) {
      console.log('Found cached summary for:', cacheKey);
      
      // Get the cached summary
      const cachedResult = summaryCache.get(cacheKey);
      
      // Display the cached summary
      displayEmailSummary(cachedResult.summaryItems);
      displayUrgencyScore(cachedResult.urgencyScore);
      
      // Update the current email data with summary and urgency
      currentEmailData.summary = cachedResult.summaryItems.join(' | ');
      currentEmailData.urgency = cachedResult.urgencyScore;
      
      // Show a status message
      showStatusMessage('Loaded cached summary', 'info');
    } else {
      console.log('No cached summary found for:', cacheKey);
    }
  } catch (error) {
    console.error('Error checking cached summary:', error);
  }
}

// Generate contact summary and display it
async function generateAndDisplayContactSummary() {
  if (!currentEmailData || !currentEmailData.sender) {
    showStatusMessage('No contact information available', 'error');
    return;
  }
  
  // First check authentication
  const isAuthorized = await checkAuthentication();
  if (!isAuthorized) {
    showStatusMessage('Please log in and join a team to view contact history', 'error');
    showUnauthorizedMessage();
    return;
  }
  
  // Show loading state
  const contactItemsElement = document.getElementById('contact-items');
  contactItemsElement.innerHTML = '<li>Searching for contact history...</li>';
  
  // If already in progress, don't start another request
  if (isContactSummaryInProgress) {
    console.log('Contact summary generation already in progress, skipping');
    return;
  }
  
  isContactSummaryInProgress = true;
  
  try {
    // Extract contact information from the email
    const contactData = {
      name: currentEmailData.sender.name,
      email: currentEmailData.sender.email
    };
    
    // Get team ID from chrome storage
    chrome.storage.sync.get(['currentTeamId'], async function(result) {
      const teamId = result.currentTeamId;
      
      // Validate that we have a team ID
      if (!teamId) {
        contactItemsElement.innerHTML = '<li>Authentication required - please log in and join a team</li>';
        showStatusMessage('Authentication required to view contact history', 'error');
        isContactSummaryInProgress = false;
        return;
      }
      
      showStatusMessage('Searching for contact history...', 'info');
      
      try {
        // Step 1: Fetch contact history from database
        const contactHistoryResponse = await new Promise((resolve, reject) => {
          chrome.runtime.sendMessage(
            { 
              action: 'getContactHistory', 
              contactData: contactData,
              teamId: teamId
            },
            function(response) {
              if (chrome.runtime.lastError) {
                reject(chrome.runtime.lastError);
              } else {
                resolve(response);
              }
            }
          );
        });
        
        console.log('Contact history response:', contactHistoryResponse);
        
        if (!contactHistoryResponse.success) {
          throw new Error(contactHistoryResponse.error || 'Failed to retrieve contact history');
        }
        
        const contactHistory = contactHistoryResponse.contactHistory || [];
        console.log('Retrieved contact history count:', contactHistory.length);
        
        if (contactHistory.length === 0) {
          contactItemsElement.innerHTML = '<li>No previous contact history found for this address</li>';
          showStatusMessage('No contact history found', 'info');
          isContactSummaryInProgress = false;
          return;
        }
        
        // Step 2: Generate contact summary using AI
        console.log('Generating AI summary for contact history...');
        const summaryResult = await generateContactSummary(contactData, contactHistory);
        console.log('Contact summary result:', summaryResult);
        
        // Display the summary
        displayContactSummary(summaryResult.summaryItems, summaryResult.messageCount);
        
        // Show success message
        showStatusMessage('Contact summary generated successfully!', 'success');
      } catch (error) {
        console.error('Error generating contact summary:', error);
        contactItemsElement.innerHTML = `
          <li>Error generating contact summary</li>
          <li>${error.message}</li>
        `;
        showStatusMessage('Error generating contact summary: ' + error.message, 'error');
      } finally {
        isContactSummaryInProgress = false;
      }
    });
  } catch (error) {
    console.error('Exception generating contact summary:', error);
    contactItemsElement.innerHTML = `
      <li>Error generating contact summary</li>
      <li>${error.message}</li>
    `;
    showStatusMessage('Error: ' + error.message, 'error');
    isContactSummaryInProgress = false;
  }
}

// Display contact summary
function displayContactSummary(summaryItems, messageCount) {
  const contactItemsElement = document.getElementById('contact-items');
  if (!contactItemsElement) {
    console.error('Contact items container not found');
    return;
  }
  
  // Clear existing content
  contactItemsElement.innerHTML = '';
  
  // Add contact message count
  if (messageCount) {
    const countItem = document.createElement('li');
    countItem.style.fontWeight = 'bold';
    countItem.textContent = `Found ${messageCount} previous emails with this contact`;
    contactItemsElement.appendChild(countItem);
  }
  
  // Add each summary item
  summaryItems.forEach(item => {
    const li = document.createElement('li');
    li.textContent = item;
    contactItemsElement.appendChild(li);
  });
  
  console.log('Contact summary displayed:', summaryItems);
}

// Generate summary using AI and display it
async function generateAndDisplaySummary(emailData) {
  // Show loading state
  const summaryItemsElement = document.getElementById('summary-items');
  summaryItemsElement.innerHTML = '<li>Generating AI summary...</li>';
  
  // Show loading state for urgency
  const urgencyScoreElement = document.getElementById('urgency-score');
  urgencyScoreElement.textContent = '...';
  const urgencyFillElement = document.getElementById('urgency-fill');
  urgencyFillElement.style.width = '0%';
  
  // Direct check of API key in storage
  chrome.storage.sync.get(['openaiApiKey'], function(result) {
    console.log('Direct storage check - API key exists:', result.openaiApiKey ? 'Yes' : 'No');
    console.log('Direct storage check - API key length:', result.openaiApiKey ? result.openaiApiKey.length : 0);
  });
  
  // If already generating, don't start another request
  if (isSummaryGenerationInProgress) {
    console.log('Summary generation already in progress, skipping');
    return;
  }
  
  try {
    isSummaryGenerationInProgress = true;
    
    console.log('Email data for summary:', {
      subject: emailData.subject,
      sender: emailData.sender,
      body: emailData.body || emailData.message,
      bodyLength: (emailData.body || emailData.message || '').length
    });
    
    console.log('Calling generateEmailSummary with:', emailData);
    const summaryResult = await generateEmailSummary(emailData);
    console.log('Summary result:', summaryResult);
    
    if (summaryResult.error) {
      console.error('Error generating summary:', summaryResult.error);
      showStatusMessage('Error generating summary: ' + summaryResult.error, 'error');
    }
    
    // Display the summary items
    displayEmailSummary(summaryResult.summaryItems);
    
    // Display the urgency score
    displayUrgencyScore(summaryResult.urgencyScore);
    
    // Update the email data with summary and urgency for saving
    currentEmailData.summary = summaryResult.summaryItems.join(' | ');
    currentEmailData.urgency = summaryResult.urgencyScore;
    
    // Reset the flag
    isSummaryGenerationInProgress = false;
    
    // Show success message
    showStatusMessage('Summary generated successfully!', 'success');
  } catch (error) {
    console.error('Exception generating summary:', error);
    summaryItemsElement.innerHTML = `
      <li>Error generating AI summary</li>
      <li>${error.message}</li>
    `;
    showStatusMessage('Error generating summary: ' + error.message, 'error');
    
    // Reset urgency display to default
    displayUrgencyScore(5);
    
    // Reset the flag
    isSummaryGenerationInProgress = false;
  }
}

// Display email summary items
function displayEmailSummary(summaryItems) {
  const summaryContainer = document.getElementById('summary-items');
  if (!summaryContainer) {
    console.error('Summary container not found');
    return;
  }
  
  // Clear existing content
  summaryContainer.innerHTML = '';
  
  // Add each summary item
  summaryItems.forEach(item => {
    const li = document.createElement('li');
    li.textContent = item;
    summaryContainer.appendChild(li);
  });
  
  console.log('Summary displayed:', summaryItems);
}

// Display the urgency score with visual indicator
function displayUrgencyScore(score) {
  const urgencyScoreElement = document.getElementById('urgency-score');
  const urgencyFillElement = document.getElementById('urgency-fill');
  
  if (!urgencyScoreElement || !urgencyFillElement) {
    console.error('Urgency elements not found');
    return;
  }
  
  // Update the score text
  urgencyScoreElement.textContent = score;
  
  // Calculate percentage for the meter
  const percentage = Math.min(Math.max((score / 10) * 100, 0), 100);
  
  // Update the fill width
  urgencyFillElement.style.width = `${percentage}%`;
  
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
  urgencyFillElement.style.backgroundColor = color;
}

// Save the current email to the database
async function saveEmailToDatabase() {
  if (!currentEmailData) {
    showStatusMessage('No email data to save', 'error');
    return;
  }
  
  // Check authentication first
  const isAuthorized = await checkAuthentication();
  if (!isAuthorized) {
    showStatusMessage('Please log in and join a team to save emails', 'error');
    showUnauthorizedMessage();
    return;
  }
  
  showStatusMessage('Saving email to database...', 'info');
  
  // Create a copy of the email data to avoid modifying the original
  const emailToSave = { ...currentEmailData };
  
  // Add a thread ID if not present
  if (!emailToSave.threadId) {
    emailToSave.threadId = `thread-${Date.now()}`;
  }
  
  // Ensure timestamp is in ISO format
  if (emailToSave.timestamp) {
    try {
      // If it's already an ISO string, keep it as is
      if (typeof emailToSave.timestamp === 'string' && 
          emailToSave.timestamp.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) {
        // Already in ISO format, no need to change
      }
      // Special handling for Gmail date formats with parenthesis
      else if (typeof emailToSave.timestamp === 'string' && 
              emailToSave.timestamp.includes('(') && 
              emailToSave.timestamp.includes(')')) {
        console.log('Sidebar detected Gmail timestamp format:', emailToSave.timestamp);
        
        // Extract the date part before the parenthesis
        const datePart = emailToSave.timestamp.split('(')[0].trim();
        
        // Check if we have a full date with year (Apr 24, 2025)
        if (/\b\d{4}\b/.test(datePart)) {
          // Date has a year, parse normally
          emailToSave.timestamp = new Date(datePart).toISOString();
        } else {
          // Date doesn't have a year (like "Apr 24"), add current year
          const currentYear = new Date().getFullYear();
          const dateWithYear = `${datePart}, ${currentYear}`;
          console.log('Sidebar: Adding current year to Gmail date:', dateWithYear);
          emailToSave.timestamp = new Date(dateWithYear).toISOString();
        }
      }
      // For other date formats
      else {
      emailToSave.timestamp = new Date(emailToSave.timestamp).toISOString();
      }
    } catch (e) {
      console.error('Error converting timestamp in sidebar:', e);
      // Use current time as fallback
      emailToSave.timestamp = new Date().toISOString();
    }
  } else {
    // If no timestamp, use current time
    emailToSave.timestamp = new Date().toISOString();
  }
  
  // Get user information from chrome storage
  chrome.storage.sync.get(['currentTeamId', 'userId', 'firstName', 'lastName', 'userEmail'], function(result) {
    const { currentTeamId, userId, firstName, lastName, userEmail } = result;
    
    // Validate that we have real user and team IDs
    if (!userId || !currentTeamId) {
      showStatusMessage('Authentication required to save emails', 'error');
      showUnauthorizedMessage();
      return;
    }
    
    // Use stored values (not default values)
    emailToSave.teamId = currentTeamId;
    emailToSave.userId = userId;
    
    console.log('Saving email with formatted timestamp and user data:', emailToSave);
    
    // Send a message to the background script to save the email
    chrome.runtime.sendMessage(
      { action: 'processEmail', emailData: emailToSave },
      function(response) {
        console.log('Received save response:', response);
        
        if (response && response.success) {
          showStatusMessage('Email saved successfully!', 'success');
          
          // Update currentEmailData with existingMessage information
          // This will allow the UI to show it's saved without requiring a reload
          if (!currentEmailData.existingMessage) {
            currentEmailData.existingMessage = {};
          }
          
          // Format the user's name for display
          let userName = 'Unknown User';
          if (firstName || lastName) {
            userName = `${firstName || ''} ${lastName || ''}`.trim();
          } else if (userEmail) {
            userName = userEmail;
          }
          
          // Set the saved information
          currentEmailData.existingMessage.exists = true;
          currentEmailData.existingMessage.savedAt = new Date().toISOString();
          currentEmailData.existingMessage.user = {
            name: userName,
            email: userEmail
          };
          
          // Optional: If response contains the saved message data, use it
          if (response.data && response.data.message) {
            currentEmailData.existingMessage.message = response.data.message;
          }
          
          // Update the UI to show the message is saved
          const alreadySavedMessage = document.getElementById('already-saved-message');
          const savedByInfo = document.getElementById('saved-by-info');
          
          if (alreadySavedMessage) {
            alreadySavedMessage.style.display = 'block';
          }
          
          // Update the saved by info
          if (savedByInfo) {
            const savedDate = new Date().toLocaleString();
            savedByInfo.textContent = `Message saved by ${userName} at ${savedDate}`;
          }
          
          // Disable the save button as it's already saved
          const saveButton = document.getElementById('save-email-button');
          if (saveButton) {
            saveButton.disabled = true;
            saveButton.textContent = 'Already Saved';
            saveButton.style.backgroundColor = '#999';
          }
          
          // Enable the mark as handled button
          const handleButton = document.getElementById('mark-handled-button');
          if (handleButton) {
            handleButton.disabled = false;
            handleButton.textContent = 'Mark as Handled';
            handleButton.style.backgroundColor = '#ff9800';
          }
          
          // Show and initialize the notes section now that the email is saved
          document.getElementById('notes-section').style.display = 'block';
          loadNotesForCurrentEmail();
        } else {
          showStatusMessage('Failed to save email: ' + (response?.error ? JSON.stringify(response.error) : 'Unknown error'), 'error');
        }
      }
    );
  });
}

// Generate a reply for the current email
async function generateReply() {
  if (!currentEmailData) {
    showStatusMessage('No email data to generate reply for', 'error');
    return;
  }
  
  showStatusMessage('Generating AI reply...', 'info');
  
  try {
    // Show loading state
    const generateReplyButton = document.getElementById('generate-reply-button');
    if (generateReplyButton) {
      generateReplyButton.disabled = true;
      generateReplyButton.textContent = 'Generating...';
    }
    
    // Hide any existing reply preview
    const replyPreviewContainer = document.getElementById('reply-preview-container');
    if (replyPreviewContainer) {
      replyPreviewContainer.style.display = 'none';
    }
    
    // Get the selected tone, length, and language options
    const toneSelect = document.getElementById('reply-tone');
    const lengthSelect = document.getElementById('reply-length');
    const languageSelect = document.getElementById('reply-language');
    
    // Get additional context from the textarea
    const additionalContextTextarea = document.getElementById('reply-additional-context');
    const additionalContext = additionalContextTextarea ? additionalContextTextarea.value.trim() : '';
    
    const options = {
      tone: toneSelect ? toneSelect.value : 'professional',
      length: lengthSelect ? lengthSelect.value : 'standard',
      language: languageSelect ? languageSelect.value : 'english',
      additionalContext: additionalContext // Add the additional context to the options
    };
    
    console.log('Generating reply with options:', options);
    
    // Call the actual AI service function from api-service.js
    const result = await generateReplySuggestion(currentEmailData, options);
    
    if (result.error) {
      console.error('Error generating reply:', result.error);
      showStatusMessage('Error generating reply: ' + result.error, 'error');
      return;
    }
    
    // Get the reply text from the response
    const replyText = result.replyText;
    
    // Display the reply in the preview section
    displayReplyPreview(replyText);
    
    showStatusMessage('Reply generated successfully!', 'success');
  } catch (error) {
    console.error('Exception generating reply:', error);
    showStatusMessage('Error generating reply: ' + error.message, 'error');
  } finally {
    // Reset the button state
    const generateReplyButton = document.getElementById('generate-reply-button');
    if (generateReplyButton) {
      generateReplyButton.disabled = false;
      generateReplyButton.textContent = 'Generate Reply';
    }
  }
}

// Display the reply in the preview section
function displayReplyPreview(replyText) {
  const replyPreviewContainer = document.getElementById('reply-preview-container');
  const replyPreview = document.getElementById('reply-preview');
  
  if (replyPreviewContainer && replyPreview) {
    // Set the reply text in the preview
    replyPreview.textContent = replyText;
    
    // Show the preview container
    replyPreviewContainer.style.display = 'block';
    
    // Set up the copy button
    const copyButton = document.getElementById('copy-reply-button');
    if (copyButton) {
      // Remove any existing event listeners
      copyButton.replaceWith(copyButton.cloneNode(true));
      
      // Get a fresh reference to the button
      const newCopyButton = document.getElementById('copy-reply-button');
      
      // Add a new event listener that gets the text from the preview element
      newCopyButton.addEventListener('click', function() {
        const currentReplyText = document.getElementById('reply-preview').textContent;
        copyReplyToClipboard(currentReplyText);
      });
      
      // Also replace the preview click listener
      replyPreview.replaceWith(replyPreview.cloneNode(true));
      
      // Get a fresh reference to the preview element
      const newReplyPreview = document.getElementById('reply-preview');
      newReplyPreview.addEventListener('click', function() {
        const currentReplyText = document.getElementById('reply-preview').textContent;
        copyReplyToClipboard(currentReplyText);
        showStatusMessage('Reply text clicked - copied to clipboard!', 'success');
      });
      newReplyPreview.style.cursor = 'pointer';
      newReplyPreview.title = 'Click to copy to clipboard';
    }
  }
}

// Copy the reply text to clipboard
function copyReplyToClipboard(replyText) {
  // Create a temporary textarea element to hold our text
  const textarea = document.createElement('textarea');
  textarea.value = replyText;
  textarea.style.position = 'fixed';  // Prevent scrolling to the element
  textarea.style.opacity = '0';       // Make it invisible
  document.body.appendChild(textarea);
  
  try {
    // Try the modern async clipboard API first
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(replyText)
        .then(() => {
          showStatusMessage('Reply copied to clipboard!', 'success');
          // Visual feedback on the button
          const copyButton = document.getElementById('copy-reply-button');
          if (copyButton) {
            const originalText = copyButton.textContent;
            copyButton.textContent = '✓ Copied!';
            copyButton.style.backgroundColor = '#4CAF50';
            setTimeout(() => {
              copyButton.textContent = originalText;
              copyButton.style.backgroundColor = '#607D8B';
            }, 2000);
          }
        })
        .catch(err => {
          console.error('Clipboard API error:', err);
          // Fall back to selection method if API fails
          fallbackCopyMethod(textarea);
        });
    } else {
      // If Clipboard API is not available, use the fallback
      fallbackCopyMethod(textarea);
    }
  } catch (err) {
    console.error('Copy to clipboard error:', err);
    fallbackCopyMethod(textarea);
  }
  
  // Clean up by removing the textarea
  setTimeout(() => {
    document.body.removeChild(textarea);
  }, 100);
}

// Fallback method for copying to clipboard
function fallbackCopyMethod(textarea) {
  try {
    // Select the text
    textarea.select();
    textarea.setSelectionRange(0, 99999); // For mobile devices
    
    // Execute copy command
    const successful = document.execCommand('copy');
    
    if (successful) {
      showStatusMessage('Reply copied to clipboard!', 'success');
      // Visual feedback on the button
      const copyButton = document.getElementById('copy-reply-button');
      if (copyButton) {
        const originalText = copyButton.textContent;
        copyButton.textContent = '✓ Copied!';
        copyButton.style.backgroundColor = '#4CAF50';
        setTimeout(() => {
          copyButton.textContent = originalText;
          copyButton.style.backgroundColor = '#607D8B';
        }, 2000);
      }
    } else {
      showStatusMessage('Please press Ctrl+C to copy the reply', 'info');
    }
  } catch (err) {
    console.error('Fallback copy method error:', err);
    showStatusMessage('Could not copy automatically. Please select the text and copy manually.', 'error');
  }
}

// Show a status message
function showStatusMessage(message, type) {
  const statusElement = document.getElementById('status-message');
  statusElement.textContent = message;
  statusElement.style.display = 'block';
  
  // Remove any existing classes
  statusElement.classList.remove('status-success', 'status-error', 'status-info');
  
  // Add the appropriate class
  if (type === 'success') {
    statusElement.classList.add('status-success');
  } else if (type === 'error') {
    statusElement.classList.add('status-error');
  } else {
    statusElement.classList.add('status-info');
  }
  
  // Hide the message after 5 seconds
  setTimeout(() => {
    statusElement.style.display = 'none';
  }, 5000);
}

// Save reply preferences to Chrome storage
function saveReplyPreferences() {
  const toneSelect = document.getElementById('reply-tone');
  const lengthSelect = document.getElementById('reply-length');
  const languageSelect = document.getElementById('reply-language');
  
  if (toneSelect && lengthSelect && languageSelect) {
    const preferences = {
      replyTone: toneSelect.value,
      replyLength: lengthSelect.value,
      replyLanguage: languageSelect.value
    };
    
    console.log('Saving reply preferences:', preferences);
    
    chrome.storage.sync.set(preferences, function() {
      console.log('Reply preferences saved to storage');
    });
  }
}

// Load reply preferences from Chrome storage
function loadReplyPreferences() {
  chrome.storage.sync.get(['replyTone', 'replyLength', 'replyLanguage'], function(result) {
    console.log('Loaded reply preferences:', result);
    
    const toneSelect = document.getElementById('reply-tone');
    const lengthSelect = document.getElementById('reply-length');
    const languageSelect = document.getElementById('reply-language');
    
    if (toneSelect && result.replyTone) {
      toneSelect.value = result.replyTone;
    }
    
    if (lengthSelect && result.replyLength) {
      lengthSelect.value = result.replyLength;
    }
    
    if (languageSelect && result.replyLanguage) {
      languageSelect.value = result.replyLanguage;
    }
  });
}

// Perform email search using the AI search function
async function performEmailSearch(query) {
  // Check if we have current email data
  if (!currentEmailData) {
    showStatusMessage('No email data available to search', 'error');
    return;
  }
  
  // Check if already searching
  if (isEmailSearchInProgress) {
    showStatusMessage('Search already in progress', 'info');
    return;
  }
  
  isEmailSearchInProgress = true;
  
  // First check authentication
  const isAuthorized = await checkAuthentication();
  if (!isAuthorized) {
    showStatusMessage('Please log in and join a team to use search', 'error');
    showUnauthorizedMessage();
    isEmailSearchInProgress = false;
    return;
  }
  
  // Show loading state
  const searchResults = document.getElementById('email-search-results');
  if (searchResults) {
    searchResults.innerHTML = '<p>Searching emails for relevant information...</p>';
    searchResults.style.display = 'block'; // Make sure the container is visible
  }
  
  const searchButton = document.getElementById('email-search-button');
  if (searchButton) {
    searchButton.disabled = true;
    searchButton.textContent = 'Searching...';
  }
  
  showStatusMessage('Searching emails...', 'info');
  
  try {
    // Get team ID from chrome storage
    chrome.storage.sync.get(['currentTeamId'], async function(result) {
      const teamId = result.currentTeamId;
      
      // Validate that we have a team ID
      if (!teamId) {
        searchResults.innerHTML = '<p>Authentication required - please log in and join a team</p>';
        showStatusMessage('Authentication required to search emails', 'error');
        isEmailSearchInProgress = false;
        
        // Reset the search button
        if (searchButton) {
          searchButton.disabled = false;
          searchButton.textContent = 'Search';
        }
        return;
      }
      
      try {
        // Step 1: Get contact history from database if contact information is available
        let contactHistory = [];
        
        if (currentEmailData.sender && (currentEmailData.sender.name || currentEmailData.sender.email)) {
          const contactData = {
            name: currentEmailData.sender.name,
            email: currentEmailData.sender.email
          };
          
          const contactHistoryResponse = await new Promise((resolve, reject) => {
            chrome.runtime.sendMessage(
              { 
                action: 'getContactHistory', 
                contactData: contactData,
                teamId: teamId
              },
              function(response) {
                if (chrome.runtime.lastError) {
                  reject(chrome.runtime.lastError);
                } else {
                  resolve(response);
                }
              }
            );
          });
          
          console.log('Contact history response:', contactHistoryResponse);
          
          if (contactHistoryResponse.success) {
            contactHistory = contactHistoryResponse.contactHistory || [];
          }
        }
        
        // Step 2: Call the search function
        const searchResult = await searchContactEmails(query, currentEmailData, contactHistory);
        console.log('Search result:', searchResult);
        
        if (!searchResult.success) {
          throw new Error(searchResult.error || 'Failed to search emails');
        }
        
        // Step 3: Display the search results
        displaySearchResults(searchResult.mainAnswer, searchResult.references);
        showStatusMessage('Search complete!', 'success');
      } catch (error) {
        console.error('Error during email search:', error);
        searchResults.innerHTML = `<p>Error searching emails: ${error.message}</p>`;
        showStatusMessage('Error searching emails: ' + error.message, 'error');
      } finally {
        isEmailSearchInProgress = false;
        
        // Reset the search button
        if (searchButton) {
          searchButton.disabled = false;
          searchButton.textContent = 'Search';
        }
      }
    });
  } catch (error) {
    console.error('Exception during email search:', error);
    isEmailSearchInProgress = false;
    
    // Reset the search button
    if (searchButton) {
      searchButton.disabled = false;
      searchButton.textContent = 'Search';
    }
    
    showStatusMessage('Error: ' + error.message, 'error');
  }
}

// Display search results in the UI
function displaySearchResults(answer, references) {
  const searchResults = document.getElementById('email-search-results');
  if (!searchResults) return;
  
  // Make sure the results container is visible
  searchResults.style.display = 'block';
  
  // Clear existing content
  searchResults.innerHTML = '';
  
  // Create the main answer element
  const answerElement = document.createElement('div');
  answerElement.className = 'search-answer';
  answerElement.textContent = answer;
  searchResults.appendChild(answerElement);
  
  // Add references if available
  if (references && references.length > 0) {
    const referencesContainer = document.createElement('div');
    referencesContainer.className = 'search-references';
    
    // Add a header for references
    const referencesHeader = document.createElement('div');
    referencesHeader.style.fontWeight = 'bold';
    referencesHeader.style.marginBottom = '8px';
    referencesHeader.textContent = 'References:';
    referencesContainer.appendChild(referencesHeader);
    
    // Add each reference
    references.forEach((ref, index) => {
      const referenceItem = document.createElement('div');
      referenceItem.className = 'reference-item';
      
      // Create quote element with proper formatting
      const quoteElement = document.createElement('div');
      quoteElement.className = 'reference-quote';
      
      // Check if the quote already has quote marks to avoid doubling them
      const quoteText = ref.quote || '';
      const hasQuotes = quoteText.startsWith('"') && quoteText.endsWith('"');
      
      // Format the display text differently based on whether it already has quotes
      if (hasQuotes) {
        quoteElement.textContent = `${index + 1}. ${quoteText}`;
      } else {
        quoteElement.textContent = `${index + 1}. "${quoteText}"`;
      }
      
      // Create metadata element
      const metaElement = document.createElement('div');
      metaElement.className = 'reference-meta';
      
      // Ensure metadata has "Saved by" prefix if it doesn't already
      let metaText = ref.meta || '';
      if (metaText && !metaText.toLowerCase().includes('saved by')) {
        metaText = 'Saved by ' + metaText;
      }
      metaElement.textContent = metaText;
      
      referenceItem.appendChild(quoteElement);
      referenceItem.appendChild(metaElement);
      referencesContainer.appendChild(referenceItem);
    });
    
    searchResults.appendChild(referencesContainer);
  } else if (answer.toLowerCase().includes('no relevant information') || 
            answer.toLowerCase().includes('couldn\'t find') || 
            answer.toLowerCase().includes('no information')) {
    // If no references and the answer indicates nothing was found, add a note
    const noInfoElement = document.createElement('div');
    noInfoElement.className = 'search-no-info';
    noInfoElement.textContent = 'No references found in saved emails.';
    noInfoElement.style.fontStyle = 'italic';
    noInfoElement.style.color = '#777';
    noInfoElement.style.marginTop = '10px';
    searchResults.appendChild(noInfoElement);
  }
}

// Mark the current email as handled
async function markEmailAsHandled() {
  if (!currentEmailData || !currentEmailData.existingMessage || !currentEmailData.existingMessage.message) {
    showStatusMessage('Please save the email first before marking it as handled', 'error');
    return;
  }
  
  // Show the custom modal instead of using the native prompt
  showHandlingModal();
}

// Set up event listeners for the handling modal - will be called once during initialization
function setupHandlingModalListeners() {
  const modal = document.getElementById('custom-modal-backdrop');
  const modalInput = document.getElementById('custom-modal-input');
  const okButton = document.getElementById('custom-modal-ok');
  const cancelButton = document.getElementById('custom-modal-cancel');
  const closeButton = document.getElementById('custom-modal-close');
  
  // Set up the OK button
  okButton.addEventListener('click', function() {
    const note = modalInput.value.trim();
    hideHandlingModal();
    processHandlingNote(note);
  });
  
  // Set up the Cancel button and Close button
  cancelButton.addEventListener('click', hideHandlingModal);
  closeButton.addEventListener('click', hideHandlingModal);
  
  // Close modal when clicking on backdrop (outside the modal)
  modal.addEventListener('click', function(event) {
    if (event.target === modal) {
      hideHandlingModal();
    }
  });
  
  // Handle Enter key press in the textarea
  modalInput.addEventListener('keypress', function(event) {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      const note = modalInput.value.trim();
      hideHandlingModal();
      processHandlingNote(note);
    }
  });
}

// Show the custom modal for handling notes
function showHandlingModal() {
  // Get modal elements
  const modal = document.getElementById('custom-modal-backdrop');
  const modalInput = document.getElementById('custom-modal-input');
  
  // Clear any previous input
  modalInput.value = '';
  
  // Display the modal
  modal.style.display = 'flex';
  
  // Focus the input field
  setTimeout(() => {
    modalInput.focus();
  }, 100);
}

// Hide the custom modal
function hideHandlingModal() {
  const modal = document.getElementById('custom-modal-backdrop');
  modal.style.display = 'none';
}

// Process the handling note and mark the message as handled
function processHandlingNote(note) {
  // Get user information from chrome storage
  chrome.storage.sync.get(['userId', 'firstName', 'lastName', 'userEmail'], async function(result) {
    const { userId, firstName, lastName, userEmail } = result;
    
    // Validate that we have a user ID
    if (!userId) {
      showStatusMessage('Authentication required to mark emails as handled', 'error');
      return;
    }
    
    showStatusMessage('Marking email as handled...', 'info');
    
    // Send a message to the background script
    chrome.runtime.sendMessage(
      { 
        action: 'markMessageAsHandled', 
        messageId: currentEmailData.existingMessage.message.id,
        userId: userId,
        note: note
      },
      function(response) {
        if (response && response.success) {
          showStatusMessage('Email marked as handled!', 'success');
          
          // Update UI to show handling status
          if (!currentEmailData.existingMessage.handling) {
            currentEmailData.existingMessage.handling = {};
          }
          
          // Format user name
          let userName = 'Unknown User';
          if (firstName || lastName) {
            userName = `${firstName || ''} ${lastName || ''}`.trim();
          } else if (userEmail) {
            userName = userEmail;
          }
          
          // Update currentEmailData with handling info
          currentEmailData.existingMessage.handling = {
            handledAt: new Date().toISOString(),
            handledBy: {
              name: userName,
              email: userEmail
            },
            handlingNote: note
          };
          
          // Update the UI
          updateEmailInfo(currentEmailData);
        } else {
          showStatusMessage('Failed to mark email as handled: ' + (response?.error ? JSON.stringify(response.error) : 'Unknown error'), 'error');
        }
      }
    );
  });
}

// Set up event listeners for the notes section
function setupNotesEventListeners() {
  // Toggle notes form visibility
  const toggleNotesFormButton = document.getElementById('toggle-notes-form');
  if (toggleNotesFormButton) {
    toggleNotesFormButton.addEventListener('click', toggleNotesForm);
  }
  
  // Add note button
  const addNoteButton = document.getElementById('add-note-button');
  if (addNoteButton) {
    addNoteButton.addEventListener('click', addNoteToCurrentEmail);
  }
  
  // Allow pressing Enter in the note body textarea to submit (with Ctrl/Cmd key)
  const noteBody = document.getElementById('note-body');
  if (noteBody) {
    noteBody.addEventListener('keydown', function(e) {
      // Check if Ctrl+Enter or Cmd+Enter was pressed
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault(); // Prevent default to avoid newline
        addNoteToCurrentEmail();
      }
    });
  }
}

// Toggle the notes form visibility
function toggleNotesForm() {
  const formContainer = document.getElementById('notes-form-container');
  const toggleButton = document.getElementById('toggle-notes-form');
  
  if (formContainer.style.display === 'none') {
    formContainer.style.display = 'block';
    toggleButton.textContent = 'Cancel';
    
    // Focus the note body textarea
    document.getElementById('note-body').focus();
  } else {
    formContainer.style.display = 'none';
    toggleButton.textContent = 'Add Note';
    
    // Clear the note body textarea
    document.getElementById('note-body').value = '';
  }
}

// Add a note to the current email
async function addNoteToCurrentEmail() {
  // Check if we have a current email
  if (!currentEmailData || !currentEmailData.existingMessage || !currentEmailData.existingMessage.message) {
    showStatusMessage('Cannot add notes until the email is saved', 'error');
    return;
  }
  
  // Get the note body and category
  const noteBody = document.getElementById('note-body').value.trim();
  const category = document.getElementById('note-category').value;
  
  // Validate the note body
  if (!noteBody) {
    showStatusMessage('Please enter a note', 'error');
    return;
  }
  
  try {
    // Get user ID from storage
    const { userId } = await chrome.storage.sync.get(['userId']);
    
    if (!userId) {
      showStatusMessage('You must be logged in to add notes', 'error');
      return;
    }
    
    // Show loading status
    showStatusMessage('Adding note...', 'info');
    
    // Call background script to add the note
    const response = await new Promise((resolve, reject) => {
      chrome.runtime.sendMessage({
        action: 'addNoteToMessage',
        messageId: currentEmailData.existingMessage.message.id,
        userId: userId,
        noteBody: noteBody,
        category: category
      }, response => {
        if (chrome.runtime.lastError) {
          reject(chrome.runtime.lastError);
        } else {
          resolve(response);
        }
      });
    });
    
    if (response.success) {
      showStatusMessage('Note added successfully', 'success');
      
      // Clear the note form
      document.getElementById('note-body').value = '';
      
      // Toggle the form back to hidden
      toggleNotesForm();
      
      // Reload notes to show the new one
      await loadNotesForCurrentEmail();
    } else {
      showStatusMessage('Failed to add note: ' + response.error, 'error');
    }
  } catch (error) {
    console.error('Error adding note:', error);
    showStatusMessage('Error adding note: ' + error.message, 'error');
  }
}

// Load notes for the current email
async function loadNotesForCurrentEmail() {
  // Check if we have a current email with a message ID
  if (!currentEmailData || !currentEmailData.existingMessage || !currentEmailData.existingMessage.message) {
    // Hide notes section if the email is not saved
    document.getElementById('notes-section').style.display = 'none';
    return;
  }
  
  // Show notes section
  document.getElementById('notes-section').style.display = 'block';
  
  try {
    // Show loading in notes container
    const notesContainer = document.getElementById('notes-container');
    notesContainer.innerHTML = '<div class="no-notes-message">Loading notes...</div>';
    
    // Call background script to get notes
    const response = await new Promise((resolve, reject) => {
      chrome.runtime.sendMessage({
        action: 'getMessageNotes',
        messageId: currentEmailData.existingMessage.message.id
      }, response => {
        if (chrome.runtime.lastError) {
          reject(chrome.runtime.lastError);
        } else {
          resolve(response);
        }
      });
    });
    
    if (response.success) {
      // Display the notes
      displayNotes(response.notes);
    } else {
      console.error('Failed to load notes:', response.error);
      notesContainer.innerHTML = '<div class="no-notes-message">Failed to load notes</div>';
    }
  } catch (error) {
    console.error('Error loading notes:', error);
    document.getElementById('notes-container').innerHTML = 
      '<div class="no-notes-message">Error loading notes: ' + error.message + '</div>';
  }
}

// Display notes in the UI
function displayNotes(notes) {
  const notesContainer = document.getElementById('notes-container');
  
  // Clear the container
  notesContainer.innerHTML = '';
  
  // If no notes, show a message
  if (!notes || notes.length === 0) {
    notesContainer.innerHTML = 
      '<div class="no-notes-message">No notes for this email yet. Click "Add Note" to create the first note.</div>';
    return;
  }
  
  // Create a document fragment to improve performance
  const fragment = document.createDocumentFragment();
  
  // Add each note to the fragment
  notes.forEach(note => {
    const noteElement = document.createElement('div');
    noteElement.className = 'note-item';
    
    // Format date
    const noteDate = new Date(note.created_at);
    const formattedDate = noteDate.toLocaleString();
    
    // Map category to CSS class
    let categoryClass = 'category-other';
    if (note.category === 'Action Required') categoryClass = 'category-action';
    else if (note.category === 'Pending') categoryClass = 'category-pending';
    else if (note.category === 'Information') categoryClass = 'category-info';
    
    // Create note HTML
    noteElement.innerHTML = `
      <div class="note-header">
        <span class="note-user">${note.user ? note.user.name : 'Unknown User'}</span>
        <span class="note-time">${formattedDate}</span>
      </div>
      <div class="note-body">${escapeHtml(note.note_body)}</div>
      <div class="note-category ${categoryClass}">${note.category || 'Other'}</div>
    `;
    
    fragment.appendChild(noteElement);
  });
  
  // Add all notes to the container at once
  notesContainer.appendChild(fragment);
}

// Helper function to escape HTML to prevent XSS
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// Initialize when the DOM is fully loaded
document.addEventListener('DOMContentLoaded', function() {
  loadModules(); // Load modules first, then init() will be called
}); 