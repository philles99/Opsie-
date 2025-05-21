// Global variables
let currentEmailData = null;
let sidebarInjected = false;
let emailPlatform = null;
let extractGmailEmailData, extractOutlookEmailData;
let sidebarIsOpen = false;
let lastProcessedEmailId = null; // Track the last email we processed
let outlookProcessingTimeout = null;
let isCurrentlyProcessing = false;
let sidebarManuallyClosedTimestamp = 0;
const SIDEBAR_REOPEN_DELAY = 10000; // 10 seconds delay before sidebar can reopen
let isAuthChecking = false; // Flag to track if we're currently checking auth
let checkMessageExists;

// Load modules dynamically
async function loadModules() {
  try {
    const gmailParserModule = await import(chrome.runtime.getURL('utils/gmail-parser.js'));
    const outlookParserModule = await import(chrome.runtime.getURL('utils/outlook-parser.js'));
    const supabaseClientModule = await import(chrome.runtime.getURL('utils/supabase-client.js'));
    
    extractGmailEmailData = gmailParserModule.extractGmailEmailData;
    extractOutlookEmailData = outlookParserModule.extractOutlookEmailData;
    checkMessageExists = supabaseClientModule.checkMessageExists;
    
    console.log('Modules loaded successfully');
    
    // Initialize after modules are loaded
    init();
  } catch (error) {
    console.error('Error loading modules:', error);
  }
}

// Check if user is authenticated and has a team
async function checkUserAuthAndTeam() {
  console.log('Starting authentication check...');
  
  if (isAuthChecking) {
    console.log('Authentication check already in progress, skipping');
    return false; // Prevent multiple simultaneous checks
  }
  
  try {
    isAuthChecking = true;
    console.log('Authentication check flag set');
    
    let isAuthenticated = false;
    
    try {
      // First check if user is authenticated
      console.log('Requesting authentication status from background script');
      const authResponse = await new Promise((resolve, reject) => {
        try {
          chrome.runtime.sendMessage({ action: 'isAuthenticated' }, response => {
            if (chrome.runtime.lastError) {
              // Handle extension context invalidated error
              console.warn('Chrome runtime error:', chrome.runtime.lastError);
              reject(chrome.runtime.lastError);
            } else {
              console.log('Received authentication response:', response);
              resolve(response);
            }
          });
        } catch (err) {
          reject(err);
        }
      });
      
      console.log('Authentication check result:', authResponse);
      
      if (!authResponse || !authResponse.success) {
        console.error('Authentication check failed to return valid response');
        isAuthChecking = false;
        return false;
      }
      
      isAuthenticated = authResponse.authenticated;
    } catch (error) {
      console.warn('Failed to communicate with background script, falling back to storage check:', error);
      
      // Fallback to checking authentication via storage
      const { accessToken, userId } = await new Promise(resolve => {
        chrome.storage.sync.get(['accessToken', 'userId'], result => {
          resolve(result);
        });
      });
      
      isAuthenticated = !!accessToken && !!userId;
      console.log('Authentication status from storage fallback:', isAuthenticated);
    }
    
    if (!isAuthenticated) {
      console.log('User is not authenticated');
      isAuthChecking = false;
      return false;
    }
    
    // If authenticated, check if user has a team
    console.log('User is authenticated, checking for team ID');
    const { userId, currentTeamId } = await new Promise(resolve => {
      chrome.storage.sync.get(['userId', 'currentTeamId'], result => {
        console.log('Storage values retrieved:', {
          userId: result.userId ? 'exists' : 'missing',
          currentTeamId: result.currentTeamId ? 'exists' : 'missing'
        });
        resolve(result);
      });
    });
    
    console.log('User ID:', userId, 'Team ID:', currentTeamId);
    
    if (!userId) {
      console.log('User ID missing from storage');
      isAuthChecking = false;
      return false;
    }
    
    if (!currentTeamId) {
      console.log('User authenticated but has no team');
      isAuthChecking = false;
      return false;
    }
    
    // User is authenticated and has a team
    console.log('Authentication check passed: User is authenticated and has a team');
    isAuthChecking = false;
    return true;
  } catch (error) {
    console.error('Error checking authentication and team status:', error);
    isAuthChecking = false;
    return false;
  }
}

// Function to show login required message
function showLoginRequiredMessage() {
  // Create a notification element if it doesn't exist
  let notification = document.getElementById('opsie-auth-notification');
  
  if (!notification) {
    notification = document.createElement('div');
    notification.id = 'opsie-auth-notification';
    notification.style.position = 'fixed';
    notification.style.bottom = '60px';
    notification.style.right = '20px';
    notification.style.padding = '12px 16px';
    notification.style.backgroundColor = '#1a3d5c';
    notification.style.color = 'white';
    notification.style.borderRadius = '4px';
    notification.style.boxShadow = '0 2px 10px rgba(0, 0, 0, 0.2)';
    notification.style.zIndex = '9997';
    notification.style.maxWidth = '300px';
    notification.style.lineHeight = '1.5';
    notification.style.fontSize = '14px';
    
    document.body.appendChild(notification);
  }
  
  // Show appropriate message based on authentication status
  chrome.storage.sync.get(['userId', 'currentTeamId'], result => {
    const { userId, currentTeamId } = result;
    
    if (!userId) {
      notification.innerHTML = 'Please <a href="#" id="opsie-login-link" style="color: #ffcc00; text-decoration: underline;">log in</a> to use Opsie Email Assistant.';
    } else if (!currentTeamId) {
      notification.innerHTML = 'Please <a href="#" id="opsie-team-link" style="color: #ffcc00; text-decoration: underline;">create or join a team</a> to use Opsie Email Assistant.';
    } else {
      notification.innerHTML = 'Authentication error. Please try reloading the page.';
    }
    
    // Make the notification visible
    notification.style.display = 'block';
    
    // Add click handlers to open popup
    const loginLink = document.getElementById('opsie-login-link');
    const teamLink = document.getElementById('opsie-team-link');
    
    if (loginLink) {
      loginLink.addEventListener('click', function(e) {
        e.preventDefault();
        chrome.runtime.sendMessage({ action: 'openPopup' });
      });
    }
    
    if (teamLink) {
      teamLink.addEventListener('click', function(e) {
        e.preventDefault();
        chrome.runtime.sendMessage({ action: 'openPopup' });
      });
    }
    
    // Hide after 5 seconds
    setTimeout(() => {
      notification.style.display = 'none';
    }, 5000);
  });
}

// Determine which email platform we're on
function detectEmailPlatform() {
  const url = window.location.href;
  
  if (url.includes('mail.google.com')) {
    emailPlatform = 'gmail';
  } else if (url.includes('outlook.office.com') || url.includes('outlook.live.com')) {
    emailPlatform = 'outlook';
  } else {
    emailPlatform = null;
  }
  
  console.log('Detected email platform:', emailPlatform);
  return emailPlatform;
}

// Initialize the extension
function init() {
  const platform = detectEmailPlatform();
  
  if (!platform) {
    console.log('Not on a supported email platform');
    return;
  }
  
  // Set up observers based on the platform
  if (platform === 'gmail') {
    setupGmailObserver();
  } else if (platform === 'outlook') {
    console.log('Outlook detected, dumping structure for debugging');
    setTimeout(dumpOutlookStructure, 2000); // Wait a bit for the page to load
    setupOutlookObserver();
  }
  
  // Inject sidebar container - but don't show it yet
  injectSidebarContainer();
  
  // Listen for messages from the sidebar iframe
  window.addEventListener('message', function(event) {
    // Only accept messages from our sidebar or expected extension origins
    const isSidebarSource = event.source === document.getElementById('opsie-sidebar-iframe')?.contentWindow;
    
    if (!isSidebarSource && event.origin !== 'chrome-extension://' + chrome.runtime.id) {
      console.log('Ignoring message from unknown source:', event.source);
      return;
    }
    
    console.log('Message received in content script:', event.data);
    
    // Handle different message types
    if (event.data.type === 'CLOSE_SIDEBAR') {
      console.log('Close sidebar message received');
      closeSidebar();
    } else if (event.data.type === 'SIDEBAR_LOADED') {
      console.log('Sidebar loaded message received');
      
      // Check if user is authenticated first
      checkUserAuthAndTeam().then(isAuthorized => {
        console.log('Authentication check on sidebar load:', isAuthorized);
        
        if (isAuthorized) {
          // If we have email data, send it to the sidebar
          if (currentEmailData) {
            console.log('Sending existing email data to loaded sidebar');
            updateSidebar(currentEmailData);
          } else {
            console.log('No email data available to send to loaded sidebar');
          }
        } else {
          console.log('User not authorized, not sending email data to sidebar');
        }
      });
    } else if (event.data.type === 'INSERT_REPLY') {
      insertReplyIntoCompose(event.data.replyText);
    } else if (event.data.type === 'AUTH_STATE_CHANGED') {
      // Handle auth state change (login/logout)
      console.log('Authentication state changed, refreshing state');
      
      // If user logged in and has data waiting, update the sidebar
      if (event.data.isLoggedIn && currentEmailData) {
        console.log('User logged in, updating sidebar with current email data');
        updateSidebar(currentEmailData, true);
      }
    }
  });
  
  // Add listener for extension's authentication state changes
  chrome.storage.onChanged.addListener(function(changes, namespace) {
    if (namespace === 'sync') {
      const authChanged = changes.accessToken || changes.userId || changes.currentTeamId;
      
      if (authChanged) {
        console.log('Auth-related storage changed, refreshing state');
        
        // If we have email data, try updating the sidebar (it will check auth internally)
        if (currentEmailData) {
          console.log('Attempting to update sidebar after auth change');
          updateSidebar(currentEmailData, true);
        }
      }
    }
  });
  
  // Add listener for messages from background script and popup
  chrome.runtime.onMessage.addListener(function(message, sender, sendResponse) {
    console.log('Message received from extension:', message);
    
    // Handle authentication state change notifications
    if (message.action === 'authStateChanged') {
      console.log('Auth state change notification received:', message.event);
      
      // Handle different auth events
      switch (message.event) {
        case 'tokenRefreshed':
          console.log('Token was refreshed successfully');
          // If we have email data, try updating the sidebar to confirm its state
          if (currentEmailData) {
            console.log('Refreshing sidebar with current email data after token refresh');
            updateSidebar(currentEmailData, true);
          }
          break;
          
        case 'tokenRefreshFailed':
        case 'tokenExpired':
          console.log('Token expired or refresh failed, showing auth required message');
          // Close the sidebar if it's open
          if (sidebarIsOpen) {
            closeSidebar();
          }
          showLoginRequiredMessage();
          break;
      }
      
      // Acknowledge the message
      if (sendResponse) {
        sendResponse({ received: true });
      }
      return true;
    }
    
    // Handle team creation/joining messages from the popup
    if (message.action === 'teamCreated' || message.action === 'teamJoined') {
      console.log(`Team ${message.action === 'teamCreated' ? 'created' : 'joined'} message received`);
      
      // Update the current team ID in memory to avoid refresh
      chrome.storage.sync.get(['accessToken', 'userId'], function(data) {
        console.log('Updating auth state after team change, user ID:', message.userId);
        
        // If the user is already authed and we just got team ID
        if (data.accessToken && data.userId && message.teamId) {
          console.log('User authenticated and team joined, showing the sidebar');
          
          // If sidebar is open or we have data, update it
          if (sidebarIsOpen || currentEmailData) {
            updateSidebar(currentEmailData || {}, true);
          }
        }
      });
      
      // Send acknowledgment response
      if (sendResponse) {
        sendResponse({ success: true });
      }
    }
    // Handle team leaving/deletion messages
    else if (message.action === 'teamLeft' || message.action === 'teamDeleted') {
      console.log(`Team ${message.action === 'teamLeft' ? 'left' : 'deleted'} message received`);
      
      // Close the sidebar if it's open
      if (sidebarIsOpen) {
        closeSidebar();
      }
      
      // Show a notification
      let notification = document.getElementById('opsie-auth-notification');
      if (!notification) {
        // Create notification if it doesn't exist
        notification = document.createElement('div');
        notification.id = 'opsie-auth-notification';
        notification.style.position = 'fixed';
        notification.style.bottom = '60px';
        notification.style.right = '20px';
        notification.style.padding = '12px 16px';
        notification.style.backgroundColor = '#1a3d5c';
        notification.style.color = 'white';
        notification.style.borderRadius = '4px';
        notification.style.boxShadow = '0 2px 10px rgba(0, 0, 0, 0.2)';
        notification.style.zIndex = '9997';
        notification.style.maxWidth = '300px';
        notification.style.lineHeight = '1.5';
        notification.style.fontSize = '14px';
        
        document.body.appendChild(notification);
      }
      
      notification.innerHTML = message.action === 'teamLeft' ? 
        'You have left the team. Please <a href="#" id="opsie-team-link" style="color: #ffcc00; text-decoration: underline;">join or create a new team</a> to use Opsie Email Assistant.' : 
        'Your team has been deleted. Please <a href="#" id="opsie-team-link" style="color: #ffcc00; text-decoration: underline;">create a new team</a> to use Opsie Email Assistant.';
      
      notification.style.display = 'block';
      
      // Add click handler to open popup
      const teamLink = document.getElementById('opsie-team-link');
      if (teamLink) {
        teamLink.addEventListener('click', function(e) {
          e.preventDefault();
          chrome.runtime.sendMessage({ action: 'openPopup' });
        });
      }
      
      // Hide after 8 seconds (longer for this important message)
      setTimeout(() => {
        notification.style.display = 'none';
      }, 8000);
      
      // Send acknowledgment response
      if (sendResponse) {
        sendResponse({ success: true });
      }
    }
    // Handle admin role changes
    else if (message.action === 'adminRoleChanged') {
      console.log('Admin role changed to:', message.newRole);
      
      // No UI changes needed in content script for role changes
      // Just acknowledge the message
      if (sendResponse) {
        sendResponse({ success: true });
      }
    }
  });
}

// Create and inject the sidebar container
async function injectSidebarContainer() {
  if (sidebarInjected) return;
  
  console.log('Creating sidebar container...');
  
  const sidebarContainer = document.createElement('div');
  sidebarContainer.id = 'opsie-sidebar-container';
  sidebarContainer.classList.add('opsie-sidebar-container');
  
  // Add styles to the sidebar container
  sidebarContainer.style.position = 'fixed';
  sidebarContainer.style.top = '0';
  sidebarContainer.style.right = '0';
  sidebarContainer.style.width = '300px';
  sidebarContainer.style.height = '100%';
  sidebarContainer.style.backgroundColor = '#f5f4f1'; // Slightly different from header for contrast
  sidebarContainer.style.zIndex = '9999';
  sidebarContainer.style.boxShadow = '-2px 0 5px rgba(0, 0, 0, 0.1)';
  sidebarContainer.style.transition = 'transform 0.3s ease';
  sidebarContainer.style.transform = 'translateX(100%)'; // Start hidden
  
  // Create iframe for the sidebar
  const iframe = document.createElement('iframe');
  iframe.id = 'opsie-sidebar-iframe';
  iframe.src = chrome.runtime.getURL('sidebar/sidebar.html');
  iframe.style.width = '100%';
  iframe.style.height = '100%';
  iframe.style.border = 'none';
  
  // Add listener for when the iframe is loaded
  iframe.addEventListener('load', function() {
    console.log('Sidebar iframe loaded');
    
    // Check if we have email data to send to the sidebar
    if (currentEmailData) {
      console.log('Sending current email data to newly loaded sidebar');
      setTimeout(() => {
        // Allow a brief delay for the sidebar to initialize
        iframe.contentWindow.postMessage({
          type: 'EMAIL_DATA',
          emailData: currentEmailData
        }, '*');
      }, 500);
    } else {
      console.log('No current email data to send to sidebar');
    }
  });
  
  sidebarContainer.appendChild(iframe);
  
  // Add sidebar to body
  document.body.appendChild(sidebarContainer);
  console.log('Sidebar container added to DOM');
  
  // Create toggle button
  const toggleButton = document.createElement('button');
  toggleButton.id = 'opsie-toggle-button';
  toggleButton.textContent = 'Opsie';
  toggleButton.style.position = 'fixed';
  toggleButton.style.right = '20px';
  toggleButton.style.bottom = '20px';
  toggleButton.style.zIndex = '9998';
  toggleButton.style.padding = '8px 16px';
  toggleButton.style.backgroundColor = '#1a3d5c';
  toggleButton.style.color = 'white';
  toggleButton.style.border = 'none';
  toggleButton.style.borderRadius = '4px';
  toggleButton.style.cursor = 'pointer';
  toggleButton.style.boxShadow = '0 2px 5px rgba(0, 0, 0, 0.2)';
  
  toggleButton.onclick = async function() {
    try {
      // Check authentication and team status before toggling
      const isAuthorized = await checkUserAuthAndTeam();
      
      if (!isAuthorized) {
        console.log('User not authorized to use sidebar');
        // Try checking directly from storage as a fallback
        const { accessToken, userId, currentTeamId } = await new Promise(resolve => {
          chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId'], result => {
            resolve(result);
          });
        });
        
        // If we have all the required auth data, consider the user authorized
        if (accessToken && userId && currentTeamId) {
          console.log('Fallback auth check passed, user has required credentials');
          // Continue with opening the sidebar
          const sidebar = document.getElementById('opsie-sidebar-container');
          if (sidebar.style.display === 'none' || sidebar.style.transform === 'translateX(100%)') {
            openSidebar();
            sidebar.style.transform = 'translateX(0)';
          } else {
            sidebar.style.transform = 'translateX(100%)';
            closeSidebar();
          }
          return;
        }
        
        showLoginRequiredMessage();
        return;
      }
      
      const sidebar = document.getElementById('opsie-sidebar-container');
      
      if (sidebar.style.display === 'none' || sidebar.style.transform === 'translateX(100%)') {
        // If sidebar is hidden, show it (manual open)
        openSidebar();
        sidebar.style.transform = 'translateX(0)';
      } else {
        // If sidebar is visible, hide it
        sidebar.style.transform = 'translateX(100%)';
        closeSidebar();
      }
    } catch (error) {
      console.error('Error in toggle button click handler:', error);
      
      // Fallback to direct storage check if there was an error
      try {
        const { accessToken, userId, currentTeamId } = await new Promise(resolve => {
          chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId'], result => {
            resolve(result);
          });
        });
        
        if (accessToken && userId && currentTeamId) {
          console.log('Error handler fallback: User has required credentials');
          const sidebar = document.getElementById('opsie-sidebar-container');
          if (sidebar.style.display === 'none' || sidebar.style.transform === 'translateX(100%)') {
            openSidebar();
            sidebar.style.transform = 'translateX(0)';
          } else {
            sidebar.style.transform = 'translateX(100%)';
            closeSidebar();
          }
          return;
        }
        
        showLoginRequiredMessage();
      } catch (storageError) {
        console.error('Error in fallback authentication check:', storageError);
        showLoginRequiredMessage();
      }
    }
  };
  
  document.body.appendChild(toggleButton);
  console.log('Toggle button added to DOM');
  
  sidebarInjected = true;
}

// Update sidebar with email data
async function updateSidebar(emailData, isManualOpen = false) {
  console.log('UpdateSidebar called with email data:', {
    sender: emailData.sender,
    subject: emailData.subject,
    timestamp: emailData.timestamp,
    messagePreview: emailData.message ? emailData.message.substring(0, 50) + '...' : 'No message'
  });

  // First check if the user is authorized
  console.log('Checking authentication status...');
  const isAuthorized = await checkUserAuthAndTeam();
  console.log('Authentication check result:', isAuthorized);
  
  if (!isAuthorized) {
    console.log('User not authorized to use sidebar, showing login prompt');
    showLoginRequiredMessage();
    return;
  }

  // Check if sidebar was manually closed recently and this is not a manual open
  const timeSinceClose = Date.now() - sidebarManuallyClosedTimestamp;
  if (timeSinceClose < SIDEBAR_REOPEN_DELAY && !isManualOpen) {
    console.log('Sidebar was manually closed', Math.round(timeSinceClose/1000), 'seconds ago. Not reopening yet.');
    return;
  }
  
  console.log('Authentication passed, updating sidebar with email data');
  
  // Get the current team ID from Chrome storage
  const { currentTeamId } = await chrome.storage.sync.get(['currentTeamId']);
  
  // Check if this message already exists in the database
  let existingMessageInfo = { exists: false };
  if (emailData.messageId && currentTeamId) {
    try {
      console.log('Checking if message exists in database:', emailData.messageId);
      existingMessageInfo = await checkMessageExists(emailData.messageId, currentTeamId, emailData);
      console.log('Message existence check result:', existingMessageInfo);
    } catch (error) {
      console.error('Error checking if message exists:', error);
    }
  } else {
    console.log('Cannot check if message exists - missing messageId or teamId', {
      messageId: !!emailData.messageId,
      teamId: !!currentTeamId
    });
  }
  
  // Add existing message info to the email data
  emailData.existingMessage = existingMessageInfo;
  
  // Inject the sidebar if it's not already injected
  if (!sidebarInjected) {
    console.log('Sidebar not injected yet, injecting now');
    injectSidebarContainer();
  }
  
  // Show the sidebar
  const sidebar = document.getElementById('opsie-sidebar-container');
  if (sidebar) {
    console.log('Making sidebar visible');
    sidebar.style.display = 'block';
    sidebar.style.transform = 'translateX(0)'; // Make sure it's visible
    sidebarIsOpen = true;
  } else {
    console.error('Sidebar element not found in DOM');
  }
  
  // Send the email data to the sidebar
  const sidebarFrame = document.getElementById('opsie-sidebar-iframe');
  if (sidebarFrame) {
    console.log('Sending email data to sidebar iframe');
    sidebarFrame.contentWindow.postMessage({
      type: 'EMAIL_DATA',
      emailData: emailData
    }, '*');
  } else {
    console.error('Sidebar iframe not found in DOM');
  }
}

// Helper function to generate a unique ID for an email
function generateEmailId(emailData) {
  // Create a simple hash from sender, subject, and timestamp
  return `${emailData.sender.email}-${emailData.subject}-${emailData.timestamp}`;
}

// Insert reply into compose box
function insertReplyIntoCompose(replyText) {
  let result = { success: false, error: 'Unknown platform' };
  
  if (emailPlatform === 'gmail') {
    result = insertReplyIntoGmailCompose(replyText);
  } else if (emailPlatform === 'outlook') {
    result = insertReplyIntoOutlookCompose(replyText);
  }
  
  // Send the result back to the sidebar
  if (sidebarWindow && sidebarWindow.contentWindow) {
    sidebarWindow.contentWindow.postMessage({
      type: 'REPLY_INSERT_RESULT',
      success: result.success,
      error: result.error
    }, '*');
  }
  
  return result;
}

// Insert reply into Gmail compose box
function insertReplyIntoGmailCompose(replyText) {
  // Find the compose box
  const composeBox = document.querySelector('.Am.Al.editable');
  
  if (composeBox) {
    // Focus on the compose box
    composeBox.focus();
    
    // Insert the reply text
    composeBox.innerHTML = replyText.replace(/\n/g, '<br>');
    
    // Trigger input event to ensure Gmail recognizes the change
    const inputEvent = new Event('input', { bubbles: true });
    composeBox.dispatchEvent(inputEvent);
    
    return { success: true };
  } else {
    // If no compose box is found, try to click the reply button
    const replyButton = document.querySelector('[aria-label="Reply"]');
    if (replyButton) {
      replyButton.click();
      
      // Wait for the compose box to appear
      setTimeout(() => {
        insertReplyIntoGmailCompose(replyText);
      }, 500);
      
      return { success: false, error: 'Attempting to open reply box. Please try inserting again in a moment.' };
    } else {
      console.error('Could not find Gmail compose box or reply button');
      return { success: false, error: 'Please click "Reply" in Gmail first to open the compose box' };
    }
  }
}

// Insert reply into Outlook compose box
function insertReplyIntoOutlookCompose(replyText) {
  // Find the compose box in Outlook - try multiple selectors
  const composeSelectors = [
    '[contenteditable="true"][aria-label="Message body"]',
    '[contenteditable="true"][aria-label*="message"]',
    '[contenteditable="true"][aria-label*="body"]',
    '[contenteditable="true"][role="textbox"]',
    '.elementToProof' // Another possible Outlook compose area class
  ];
  
  let composeBox = null;
  for (const selector of composeSelectors) {
    const element = document.querySelector(selector);
    if (element) {
      composeBox = element;
      console.log('Found Outlook compose box with selector:', selector);
      break;
    }
  }
  
  if (composeBox) {
    // Focus on the compose box
    composeBox.focus();
    
    // Insert the reply text
    composeBox.innerHTML = replyText.replace(/\n/g, '<br>');
    
    // Trigger input event to ensure Outlook recognizes the change
    const inputEvent = new Event('input', { bubbles: true });
    composeBox.dispatchEvent(inputEvent);
    
    return { success: true };
  } else {
    // If no compose box is found, try to click the reply button
    // Try multiple reply button selectors
    const replySelectors = [
      '[aria-label="Reply"]',
      '[aria-label*="reply"]',
      '[title="Reply"]',
      '[title*="reply"]',
      'button:has-text("Reply")',
      '.ms-Button--commandBar[name="reply"]'
    ];
    
    let replyButton = null;
    for (const selector of replySelectors) {
      try {
        const element = document.querySelector(selector);
        if (element) {
          replyButton = element;
          console.log('Found Outlook reply button with selector:', selector);
          break;
        }
      } catch (e) {
        // Some selectors might not be supported in older browsers
        console.log('Selector error:', e.message);
      }
    }
    
    if (replyButton) {
      replyButton.click();
      
      // Wait for the compose box to appear
      setTimeout(() => {
        insertReplyIntoOutlookCompose(replyText);
      }, 1000);
      
      return { success: false, error: 'Attempting to open reply box. Please try inserting again in a moment.' };
    } else {
      console.error('Could not find Outlook compose box or reply button');
      return { 
        success: false, 
        error: 'Please click "Reply" in Outlook first to open the compose box' 
      };
    }
  }
}

// Setup observer for Gmail
function setupGmailObserver() {
  console.log('Setting up Gmail observer');
  
  // Create a MutationObserver to watch for changes in the DOM
  const observer = new MutationObserver(function(mutations) {
    // Check if we're viewing an email
    const emailContainer = document.querySelector('.adn.ads');
    
    if (emailContainer && !currentEmailData) {
      // Extract email data
      processGmailEmail(emailContainer);
    } else if (!emailContainer && currentEmailData) {
      // Email view closed
      currentEmailData = null;
    }
  });
  
  // Start observing the document body for changes
  observer.observe(document.body, {
    childList: true,
    subtree: true
  });
}

// Setup observer for Outlook
function setupOutlookObserver() {
  console.log('Setting up Outlook observer');
  
  // Create a mutation observer to detect when an email is opened
  const observer = new MutationObserver((mutations) => {
    // Skip if too many mutations (likely not an email open event)
    if (mutations.length > 50) {
      console.log('Skipping large mutation batch, count:', mutations.length);
      return;
    }
    
    console.log('Outlook DOM mutation detected, mutations count:', mutations.length);
    
    // Check if an email container is present
    let emailContainer = document.querySelector('[role="main"][aria-label="Reading Pane"]');
    
    if (!emailContainer) {
      console.log('Primary email container not found, trying alternative selectors');
      
      // Try alternative selectors for different Outlook versions/languages
      const alternativeSelectors = [
        '[role="main"][aria-label="Message"]',
        '[role="main"][aria-label="Reading pane"]',
        '[role="main"][aria-label="Läsfönster"]', // Swedish
        '[role="main"][aria-label="Lesebereich"]', // German
        '[role="main"][aria-label="Panneau de lecture"]', // French
        '[role="main"][aria-label*="Message"]',
        '[role="main"][aria-label*="Reading"]',
        '[role="main"][aria-label*="pane"]',
        '[role="main"]', // Last resort - any main role
        '.ms-Panel-main' // Office UI Fabric panel
      ];
      
      // Try each selector in order
      for (const selector of alternativeSelectors) {
        const elements = document.querySelectorAll(selector);
        if (elements.length > 0) {
          console.log(`Found ${elements.length} elements matching selector: ${selector}`);
          // Check each element to see if it contains email content
          for (const element of elements) {
            const hasSender = element.querySelector('.OZZZK') !== null;
            const hasSubject = element.querySelector('[role="heading"]') !== null;
            if (hasSender || hasSubject) {
              emailContainer = element;
              console.log('Found email container with sender/subject using selector:', selector);
              break;
            }
          }
          if (emailContainer) break;
        }
      }
    }
    
    console.log('Outlook email container found after mutation:', !!emailContainer);
    
    if (emailContainer) {
      // Check if this container has any content at all
      const containerText = emailContainer.textContent.trim();
      console.log('Container text length:', containerText.length);
      
      // Check if sidebar was manually closed recently
      const timeSinceClose = Date.now() - sidebarManuallyClosedTimestamp;
      if (timeSinceClose < SIDEBAR_REOPEN_DELAY) {
        console.log('Sidebar was manually closed', Math.round(timeSinceClose/1000), 'seconds ago. Not processing email yet.');
        return;
      }
      
      // Clear any existing timeout to prevent race conditions
      if (outlookProcessingTimeout) {
        clearTimeout(outlookProcessingTimeout);
        outlookProcessingTimeout = null;
      }
      
      // Process the email container
      processOutlookEmail(emailContainer);
    } else {
      console.log('No email container found after trying all selectors');
    }
  });
  
  // Start observing the document body for changes
  observer.observe(document.body, { childList: true, subtree: true });
  
  // Also check immediately in case the email is already loaded
  const emailContainer = document.querySelector('[role="main"][aria-label="Reading Pane"]') || 
                         document.querySelector('[role="main"][aria-label="Message"]') ||
                         document.querySelector('[role="main"][aria-label="Reading pane"]');
  
  if (emailContainer) {
    console.log('Email container found on initial load');
    processOutlookEmail(emailContainer);
  } else {
    console.log('No email container found on initial load, will wait for mutations');
    
    // Add a delayed check in case the observer missed the initial load
    setTimeout(() => {
      const delayedContainer = document.querySelector('[role="main"][aria-label="Reading Pane"]') || 
                               document.querySelector('[role="main"][aria-label="Message"]') ||
                               document.querySelector('[role="main"]');
      
      if (delayedContainer) {
        console.log('Email container found in delayed check');
        processOutlookEmail(delayedContainer);
      }
    }, 2000);
  }
}

// Add CSS styles
function addStyles() {
  const style = document.createElement('style');
  style.textContent = `
    .opsie-sidebar-container {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    #opsie-toggle-button {
      transition: background-color 0.2s ease;
    }
    
    #opsie-toggle-button:hover {
      background-color: #0f2a40;
    }
  `;
  
  document.head.appendChild(style);
}

// Debug function to dump Outlook structure
function dumpOutlookStructure() {
  console.log('Dumping Outlook DOM structure for debugging');
  
  // Find all role="main" elements
  const mainElements = document.querySelectorAll('[role="main"]');
  console.log('Found ' + mainElements.length + ' elements with role="main"');
  
  mainElements.forEach((element, index) => {
    console.log(`Main element #${index}:`, {
      ariaLabel: element.getAttribute('aria-label'),
      className: element.className,
      childCount: element.children.length
    });
  });
  
  // Find expanded elements
  const expandedElements = document.querySelectorAll('[aria-expanded="true"]');
  console.log('Found ' + expandedElements.length + ' elements with aria-expanded="true"');
  
  expandedElements.forEach((element, index) => {
    if (index < 5) { // Limit to first 5 to avoid console spam
      console.log(`Expanded element #${index}:`, {
        tagName: element.tagName,
        className: element.className,
        id: element.id,
        ariaLabel: element.getAttribute('aria-label')
      });
      
      // Look for message body elements inside expanded containers
      const messageBodyElements = element.querySelectorAll('[aria-label="Message body"], [id^="UniqueMessageBody"]');
      console.log(`Found ${messageBodyElements.length} message body elements inside this expanded container`);
      
      if (messageBodyElements.length > 0) {
        const firstBodyElement = messageBodyElements[0];
        console.log('First message body element:', {
          id: firstBodyElement.id,
          className: firstBodyElement.className,
          textPreview: firstBodyElement.textContent.trim().substring(0, 50) + '...'
        });
      }
    }
  });
  
  // Find all elements that might contain email content
  const possibleEmailContainers = [
    ...document.querySelectorAll('[role="region"][aria-label*="Message"]'),
    ...document.querySelectorAll('[role="region"][aria-label*="message"]'),
    ...document.querySelectorAll('[role="region"][aria-label*="body"]'),
    ...document.querySelectorAll('[role="region"][aria-label*="Body"]'),
    ...document.querySelectorAll('.OZZZK')
  ];
  
  console.log('Found ' + possibleEmailContainers.length + ' possible email content containers');
  
  possibleEmailContainers.forEach((element, index) => {
    console.log(`Possible email container #${index}:`, {
      tagName: element.tagName,
      ariaLabel: element.getAttribute('aria-label'),
      className: element.className,
      textContent: element.textContent.substring(0, 50) + '...'
    });
  });
}

// Simplified debug function
function debugEmailData(emailData, platform) {
  console.log(`Email data extracted from ${platform}:`, {
    sender: emailData.sender,
    subject: emailData.subject,
    timestamp: emailData.timestamp,
    messagePreview: emailData.message ? emailData.message.substring(0, 50) + '...' : 'No message'
  });
}

// Function to ensure timestamp is in ISO format
function standardizeTimestamp(timestamp) {
  if (!timestamp) return new Date().toISOString();
  
  // If it's already an ISO string, return it
  if (typeof timestamp === 'string' && timestamp.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) {
    return timestamp;
  }
  
  try {
    // Special handling for Gmail date formats that may have a parenthesis with "(X hours/days ago)"
    if (typeof timestamp === 'string' && timestamp.includes('(') && timestamp.includes(')')) {
      console.log('Detected Gmail timestamp format:', timestamp);
      
      // Extract the date part before the parenthesis
      const datePart = timestamp.split('(')[0].trim();
      
      // Check if we have a full date with year (Apr 24, 2025)
      if (/\b\d{4}\b/.test(datePart)) {
        // Date has a year, parse normally
        return new Date(datePart).toISOString();
      } else {
        // Date doesn't have a year (like "Apr 24"), add current year
        const currentYear = new Date().getFullYear();
        const dateWithYear = `${datePart}, ${currentYear}`;
        console.log('Adding current year to Gmail date:', dateWithYear);
        return new Date(dateWithYear).toISOString();
      }
    }
    
    // Try to convert to a Date object and then to ISO string for non-Gmail formats
    return new Date(timestamp).toISOString();
  } catch (e) {
    console.error('Error converting timestamp:', e);
    // Return current time as fallback
    return new Date().toISOString();
  }
}

// Update the function that processes emails to standardize the timestamp
function processGmailEmail(emailContainer) {
  try {
    console.log('Processing Gmail email container');
    const emailData = extractGmailEmailData(emailContainer);
    
    if (emailData) {
      // Standardize the timestamp
      emailData.timestamp = standardizeTimestamp(emailData.timestamp);
      
      // Add debug output
      debugEmailData(emailData, 'gmail');
      
      // Use messageId for duplicate detection if available, otherwise fall back to composite ID
      let uniqueEmailId;
      
      if (emailData.messageId) {
        // Prefer using the unique message ID when available
        uniqueEmailId = emailData.messageId;
        console.log('Using messageId for duplicate detection:', uniqueEmailId);
      } else {
        // Fall back to the old method if messageId is not available
        const threadHistoryLength = emailData.threadHistory ? emailData.threadHistory.length : 0;
        uniqueEmailId = `${emailData.sender.email}-${emailData.subject}-${threadHistoryLength}`;
        console.log('No messageId available, using composite ID for duplicate detection:', uniqueEmailId);
      }
      
      // Only process if this is a new email or if the ID has changed
      if (uniqueEmailId !== lastProcessedEmailId) {
        lastProcessedEmailId = uniqueEmailId;
        
        // Create a processed data object with all necessary properties
        const processedData = {
          sender: emailData.sender,
          subject: emailData.subject || 'No Subject',
          timestamp: emailData.timestamp,
          message: emailData.message || '',
          threadHistory: emailData.threadHistory || [],
          messagePreview: emailData.message ? emailData.message.substring(0, 100) + '...' : 'No message content',
          isThread: emailData.threadHistory && emailData.threadHistory.length > 0,
          messageUrl: emailData.messageUrl, // Make sure to include the URL
          messageId: emailData.messageId // Include the message ID if available
        };
        
        // If we have thread history, add a simplified preview
        if (processedData.threadHistory && processedData.threadHistory.length > 0) {
          console.log(`Email is part of a thread with ${processedData.threadHistory.length} previous messages`);
          
          // Create a simplified preview of the thread
          processedData.threadPreview = processedData.threadHistory.map(msg => ({
            sender: msg.sender.name,
            timestamp: msg.timestamp,
            previewText: msg.message ? msg.message.substring(0, 50) + '...' : 'No content'
          }));
        }
        
        console.log('Email data extracted from Gmail:', processedData);
        console.log('Message URL captured:', processedData.messageUrl);
        if (processedData.messageId) {
          console.log('Message ID captured:', processedData.messageId);
        }
        
        // Store current email data globally
        currentEmailData = processedData;
        
        // Update the sidebar with the email data
        updateSidebar(processedData);
      } else {
        console.log('Skipping duplicate email processing:', uniqueEmailId);
      }
    }
  } catch (error) {
    console.error('Error processing Gmail email:', error);
  }
}

// Update the processOutlookEmail function
function processOutlookEmail(container) {
  // If already processing, don't start another process
  if (outlookProcessingTimeout) {
    console.log('Already processing an email, will not start another process');
    return;
  }
  
  // Set a flag to show we're processing
  outlookIsProcessing = true;
  
  // Set a timeout to make sure we don't double-process
  outlookProcessingTimeout = setTimeout(() => {
    try {
      // Check if this is an expanded message or contains expanded messages
      const hasExpandedMessages = container.querySelector('[aria-expanded="true"]') !== null;
      const isExpandedMessage = container.getAttribute('aria-expanded') === 'true';
      
      console.log('Processing Outlook email container:', container);
      
      if (isExpandedMessage || hasExpandedMessages) {
        console.log('Container has expanded messages:', {
          isDirectlyExpanded: isExpandedMessage,
          containsExpandedMessages: hasExpandedMessages
        });
      }
      
      // Check for minimum email content indicators with additional debugging
      const senderElement = container.querySelector('.OZZZK');
      const hasSender = senderElement !== null;
      console.log('Found sender element:', hasSender, senderElement ? senderElement.textContent : 'None');
      
      const subjectElement = container.querySelector('[role="heading"]');
      const hasSubject = subjectElement !== null;
      console.log('Found subject element:', hasSubject, subjectElement ? subjectElement.textContent : 'None');
      
      // Expanded message body elements
      const uniqueMessageBody = container.querySelector('[id^="UniqueMessageBody"]');
      const allowTextSelection = container.querySelector('.allowTextSelection');
      const messageBodyLabel = container.querySelector('[aria-label="Message body"]');
      const nzwzClass = container.querySelector('._nzWz');
      const gnovoClass = container.querySelector('[class*="GNoVo"]');
      
      console.log('Message body elements found:', {
        uniqueMessageBody: uniqueMessageBody !== null,
        allowTextSelection: allowTextSelection !== null,
        messageBodyLabel: messageBodyLabel !== null,
        nzwzClass: nzwzClass !== null,
        gnovoClass: gnovoClass !== null
      });
      
      // More lenient check for message body
      const hasMessageBody = messageBodyLabel !== null || 
                           nzwzClass !== null ||
                           uniqueMessageBody !== null ||
                           allowTextSelection !== null ||
                           gnovoClass !== null;
      
      console.log('Container has email content:', hasSender && hasSubject);
      console.log('All email content indicators:', {hasSender, hasSubject, hasMessageBody});
      
      // ** IMPORTANT: Make this check more lenient - only require sender and subject **
      // No longer require hasMessageBody as part of this check
      const hasMinimumEmailContent = hasSender && hasSubject;
      
      // Check if sidebar was manually closed recently
      const timeSinceClose = Date.now() - sidebarManuallyClosedTimestamp;
      const sidebarRecentlyClosed = timeSinceClose < SIDEBAR_REOPEN_DELAY;
      
      if (hasMinimumEmailContent && !sidebarRecentlyClosed) {
        console.log('Attempting to extract Outlook email data');
        
        // Extract email data
        try {
          const result = extractOutlookEmailData(container);
          console.log('Outlook extraction result:', result ? (result.success ? 'Success' : 'Failed') : 'No result');
          
          if (result && result.success && result.data) {
            const emailData = result.data;
            console.log('Extracted data:', emailData);
            
            // Use messageId for duplicate detection if available, otherwise fall back to composite ID
            let uniqueEmailId;
            
            if (emailData.messageId) {
              // Prefer using the unique message ID when available
              uniqueEmailId = emailData.messageId;
              console.log('Using messageId for duplicate detection:', uniqueEmailId);
            } else {
              // Fall back to the old method if messageId is not available
              const threadHistoryLength = emailData.threadHistory ? emailData.threadHistory.length : 0;
              uniqueEmailId = `${emailData.sender.email}-${emailData.subject}-${threadHistoryLength}`;
              console.log('No messageId available, using composite ID for duplicate detection:', uniqueEmailId);
            }
            
            // Only process if this is a new email or if the ID has changed
            if (uniqueEmailId !== lastProcessedEmailId) {
              lastProcessedEmailId = uniqueEmailId;
              
              // Process the email data
              const processedData = {
                sender: emailData.sender,
                subject: emailData.subject || 'No Subject',
                timestamp: emailData.timestamp || new Date().toISOString(),
                message: emailData.message || '',
                threadHistory: emailData.threadHistory || [],
                messagePreview: emailData.message ? emailData.message.substring(0, 100) + '...' : 'No message content',
                isThread: emailData.threadHistory && emailData.threadHistory.length > 0,
                messageUrl: emailData.messageUrl, // Make sure to include the URL in the processed data
                messageId: emailData.messageId // Include the message ID if available
              };
              
              // If we have thread history, add a summary of it
              if (processedData.threadHistory && processedData.threadHistory.length > 0) {
                console.log(`Email is part of a thread with ${processedData.threadHistory.length} previous messages`);
                
                // Create a simplified preview of the thread
                processedData.threadPreview = processedData.threadHistory.map(msg => ({
                  sender: msg.sender.name,
                  timestamp: msg.timestamp,
                  previewText: msg.message ? msg.message.substring(0, 50) + '...' : 'No content'
                }));
              }
              
              console.log('Email data extracted from outlook:', processedData);
              console.log('Message URL captured:', processedData.messageUrl);  // Log the URL
              if (processedData.messageId) {
                console.log('Message ID captured:', processedData.messageId);
              }
              
              // Store current email data globally
              currentEmailData = processedData;
              
              // Update the sidebar with the email data
              updateSidebar(processedData);
            } else {
              console.log('Skipping duplicate email processing:', uniqueEmailId);
            }
          } else {
            console.error('Failed to extract valid email data from Outlook container');
          }
        } catch (extractionError) {
          console.error('Error in Outlook email extraction:', extractionError);
        }
      } else if (sidebarRecentlyClosed) {
        console.log('Sidebar was recently closed manually. Skipping email processing.');
      } else {
        console.log('Container does not have all required email elements');
      }
      
      // Add this inside the processOutlookEmail function, right after checking hasMessageBody
      if (!hasMessageBody) {
        console.log('Message body not found, dumping potential message body elements');
        dumpMessageBodyElements(container);
      }
      
      // Reset processing flag after a delay to prevent rapid re-processing
      setTimeout(() => {
        isCurrentlyProcessing = false;
      }, 2000);
      
    } catch (error) {
      console.error('Error processing Outlook email:', error);
    } finally {
      // Always clear the timeout and reset processing flags
      outlookProcessingTimeout = null;
      outlookIsProcessing = false;
      console.log('Email processing completed, reset processing flags');
    }
  }, 1000); // 1 second debounce time
}

// Add this function to help debug message body elements
function dumpMessageBodyElements(container) {
  console.log('Dumping potential message body elements');
  
  // Find all region elements
  const regions = container.querySelectorAll('[role="region"]');
  console.log('Found', regions.length, 'region elements');
  
  regions.forEach((region, index) => {
    console.log(`Region #${index}:`, {
      ariaLabel: region.getAttribute('aria-label'),
      className: region.className,
      textLength: region.textContent.trim().length,
      textPreview: region.textContent.trim().substring(0, 50) + '...'
    });
  });
  
  // Try other potential message body containers
  const otherContainers = [
    ...container.querySelectorAll('.allowTextSelection'),
    ...container.querySelectorAll('[data-automation-id="message-body"]'),
    ...container.querySelectorAll('.ReadingPaneContent'),
    ...container.querySelectorAll('.ms-font-weight-regular')
  ];
  
  console.log('Found', otherContainers.length, 'other potential message containers');
  
  otherContainers.forEach((element, index) => {
    console.log(`Other container #${index}:`, {
      tagName: element.tagName,
      className: element.className,
      textLength: element.textContent.trim().length,
      textPreview: element.textContent.trim().substring(0, 50) + '...'
    });
  });
}

// Update the closeSidebar function
function closeSidebar() {
  if (sidebarInjected) {
    const sidebar = document.getElementById('opsie-sidebar-container');
    if (sidebar) {
      sidebar.style.display = 'none';
      sidebarIsOpen = false;
      
      // Set the timestamp when sidebar was manually closed
      sidebarManuallyClosedTimestamp = Date.now();
      console.log('Sidebar manually closed, will not reopen for', SIDEBAR_REOPEN_DELAY/1000, 'seconds');
    }
  }
}

// Add this function to manually open the sidebar
async function openSidebar() {
  console.log('Manually opening sidebar');
  
  // Check if the user is authorized
  const isAuthorized = await checkUserAuthAndTeam();
  
  if (!isAuthorized) {
    console.log('User not authorized to use sidebar, showing login prompt');
    showLoginRequiredMessage();
    return;
  }
  
  // Reset the manually closed timestamp to allow opening
  sidebarManuallyClosedTimestamp = 0;
  
  // Inject the sidebar if it's not already injected
  if (!sidebarInjected) {
    injectSidebarContainer();
  }
  
  // Show the sidebar
  const sidebar = document.getElementById('opsie-sidebar-container');
  if (sidebar) {
    sidebar.style.display = 'block';
    sidebarIsOpen = true;
  }
  
  // If we have email data, send it to the sidebar
  if (currentEmailData) {
    const sidebarFrame = document.getElementById('opsie-sidebar-iframe');
    if (sidebarFrame) {
      sidebarFrame.contentWindow.postMessage({
        type: 'EMAIL_DATA',
        emailData: currentEmailData
      }, '*');
    }
  }
}

// Initialize when the DOM is fully loaded
window.addEventListener('load', function() {
  console.log('Opsie Email Assistant content script loaded');
  addStyles();
  loadModules(); // Load modules first, then init() will be called
}); 