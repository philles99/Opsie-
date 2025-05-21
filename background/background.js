// Listen for installation
chrome.runtime.onInstalled.addListener(function() {
  console.log('Opsie Email Assistant installed');
  checkAndRefreshToken();
});

// Also check authentication on extension startup
chrome.runtime.onStartup.addListener(function() {
  console.log('Opsie Email Assistant starting up, checking authentication status');
  checkAndRefreshToken();
});

// Track user activity for session management
let lastActivityTime = Date.now();
let sessionCheckIntervalId = null;
const SESSION_CHECK_INTERVAL = 2 * 60 * 1000; // Check every 2 minutes
const USER_INACTIVITY_THRESHOLD = 30 * 60 * 1000; // 30 minutes of inactivity before stopping checks

// Function to check token status and refresh if needed
async function checkAndRefreshToken() {
  console.log('Checking authentication token status');
  
  try {
    const { accessToken, refreshToken, userId } = await chrome.storage.sync.get(['accessToken', 'refreshToken', 'userId']);
    
    if (!accessToken || !userId) {
      console.log('No access token or user ID found, user needs to log in');
      return;
    }
    
    // Check if token is expired by parsing it
    let isExpired = false;
    let timeUntilExpiry = 0;
    
    try {
      const parts = accessToken.split('.');
      if (parts.length === 3) {
        const payload = JSON.parse(atob(parts[1]));
        const expTime = payload.exp ? new Date(payload.exp * 1000) : null;
        const currentTime = new Date();
        isExpired = expTime ? (expTime < currentTime) : false;
        
        // Calculate minutes until expiry for proactive refresh
        if (expTime) {
          timeUntilExpiry = (expTime - currentTime) / 1000 / 60;
        }
        
        console.log('Token expiration:', {
          expTime: expTime ? expTime.toISOString() : 'unknown',
          currentTime: currentTime.toISOString(),
          isExpired: isExpired,
          timeRemaining: timeUntilExpiry.toFixed(2) + ' minutes'
        });
      }
    } catch (e) {
      console.error('Error checking token expiration:', e);
    }
    
    // If token is expired OR will expire in the next 5 minutes and we have a refresh token, try to refresh
    if ((isExpired || timeUntilExpiry < 5) && refreshToken) {
      console.log(`Token is ${isExpired ? 'expired' : 'expiring soon'}, attempting to refresh using refresh token`);
      
      const refreshResult = await refreshAccessToken(refreshToken);
      
      if (refreshResult.success) {
        console.log('Token refreshed successfully!');
        // Send message to all components to notify of token refresh
        try {
          chrome.runtime.sendMessage({ 
            action: 'authStateChanged', 
            event: 'tokenRefreshed',
            isLoggedIn: true
          });
        } catch (e) {
          console.error('Error sending refresh notification:', e);
        }
        
        // Continue with regular checks since we now have a valid token
        await syncUserTeamData(userId, refreshResult.accessToken);
      } else {
        console.error('Failed to refresh token:', refreshResult.error);
        // Notify components about authentication failure
        try {
          chrome.runtime.sendMessage({ 
            action: 'authStateChanged', 
            event: 'tokenRefreshFailed',
            isLoggedIn: false
          });
        } catch (e) {
          console.error('Error sending refresh failure notification:', e);
        }
        
        // If refresh failed, remove tokens
        chrome.storage.sync.remove(['accessToken', 'refreshToken']);
      }
    } else if (isExpired) {
      console.log('Token is expired and no refresh token available, clearing auth data');
      chrome.storage.sync.remove(['accessToken']);
      
      // Notify components about token expiration
      try {
        chrome.runtime.sendMessage({ 
          action: 'authStateChanged', 
          event: 'tokenExpired',
          isLoggedIn: false
        });
      } catch (e) {
        console.error('Error sending token expired notification:', e);
      }
    } else {
      // Token is valid, check if user has team data synchronized
      console.log('Token is valid, synchronizing user team data');
      await syncUserTeamData(userId, accessToken);
    }
  } catch (error) {
    console.error('Error checking token status:', error);
  }
}

// Helper function to sync user team data
async function syncUserTeamData(userId, accessToken) {
  // Get the current team ID from storage
  const { currentTeamId } = await chrome.storage.sync.get(['currentTeamId']);
  
  // If no team ID in storage, try to get it from the user details
  if (!currentTeamId) {
    console.log('No team ID in storage, attempting to retrieve from user details');
    
    try {
      // Get user details which will also set the team ID in storage if available
      const userDetails = await getUserDetails(userId, accessToken);
      
      if (userDetails.success && userDetails.data && userDetails.data.team_id) {
        console.log('Retrieved team ID from user details:', userDetails.data.team_id);
        
        // Team ID will be automatically stored by the getUserDetails function
        console.log('Team data synchronized successfully');
      } else {
        console.log('User has no team assigned in the database');
      }
    } catch (error) {
      console.error('Error synchronizing team data:', error);
    }
  } else {
    console.log('Team ID already in storage:', currentTeamId);
  }
}

// Import Supabase functions
// Note: Imports must be at the top level in a module
import { 
  saveEmailToSupabase, 
  getThreadsFromSupabase, 
  getThreadMessagesFromSupabase,
  getContactHistoryFromSupabase
} from '../utils/api-service.js';

// Import direct supabase functions for background operations
import { markMessageAsHandled, checkMessageExists, addNoteToMessage, getMessageNotes } from '../utils/supabase-client.js';

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
  getTeamMembers,
  getAllTeams,
  updateTeamDetails,
  requestToJoinTeam,
  getPendingJoinRequests,
  respondToJoinRequest,
  checkJoinRequestStatus,
  removeTeamMember,
  refreshAccessToken
} from '../utils/auth-service.js';

// Listen for messages from content script and popup
chrome.runtime.onMessage.addListener(function(request, sender, sendResponse) {
  // Update last activity time for any message received
  lastActivityTime = Date.now();
  
  // Start the session check interval if it's not already running
  startSessionChecks();
  
  console.log('Background script received message:', request);
  
  if (request.action === 'getApiKey') {
    // Get API key from storage
    chrome.storage.sync.get(['openaiApiKey'], function(result) {
      console.log('API key retrieval request received in background script');
      console.log('API key exists in storage (background):', result.openaiApiKey ? 'Yes' : 'No');
      console.log('API key length (background):', result.openaiApiKey ? result.openaiApiKey.length : 0);
      
      if (result.openaiApiKey) {
        console.log('API key first 5 chars (background):', result.openaiApiKey.substring(0, 5));
        console.log('API key format valid (background):', result.openaiApiKey.startsWith('sk-') ? 'Yes' : 'No');
      }
      
      sendResponse({apiKey: result.openaiApiKey || ''});
    });
    return true; // Required for async sendResponse
  } 
  else if (request.action === 'refreshAccessToken') {
    console.log('Token refresh request received from popup');
    
    // Call the refresh function with the provided refresh token
    refreshAccessToken(request.refreshToken)
      .then(result => {
        console.log('Token refresh result:', result.success ? 'Success' : 'Failed');
        sendResponse(result);
      })
      .catch(error => {
        console.error('Exception during token refresh:', error);
        sendResponse({ success: false, error: error.message });
      });
    
    return true; // Required for async sendResponse
  }
  else if (request.action === 'processEmail') {
    console.log('Processing email in background script:', request.emailData);
    
    // First check if user is authenticated
    isAuthenticated().then(async authenticated => {
      if (!authenticated) {
        console.error('User is not authenticated, cannot process email');
        sendResponse({success: false, error: 'Authentication required'});
        return;
      }
      
      // Check if user has a team
      const { userId, currentTeamId } = await chrome.storage.sync.get(['userId', 'currentTeamId']);
      
      if (!userId || !currentTeamId) {
        console.error('User has no team, cannot process email');
        sendResponse({success: false, error: 'Team membership required'});
        return;
      }
      
      // Store the email in Supabase
      saveEmailToSupabase(request.emailData)
        .then(result => {
          console.log('Supabase storage result:', result);
          if (result.success) {
            console.log('Email successfully stored in Supabase');
            sendResponse({success: true, data: result.data});
          } else {
            console.error('Failed to store email in Supabase:', result.error);
            sendResponse({success: false, error: result.error});
          }
        })
        .catch(error => {
          console.error('Exception when storing email:', error);
          sendResponse({success: false, error: error.message});
        });
    });
    
    return true; // Required for async sendResponse
  }
  else if (request.action === 'getThreads') {
    getThreadsFromSupabase()
      .then(result => {
        if (result.success) {
          sendResponse({success: true, threads: result.data});
        } else {
          sendResponse({success: false, error: result.error});
        }
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true; // Required for async sendResponse
  }
  else if (request.action === 'getThreadMessages') {
    getThreadMessagesFromSupabase(request.threadId)
      .then(result => {
        if (result.success) {
          sendResponse({success: true, messages: result.data});
        } else {
          sendResponse({success: false, error: result.error});
        }
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true; // Required for async sendResponse
  }
  else if (request.action === 'getContactHistory') {
    console.log('Fetching contact history for:', request.contactData);
    
    getContactHistoryFromSupabase(request.contactData, request.teamId)
      .then(result => {
        console.log('Contact history result:', result);
        if (result.success) {
          console.log('Successfully retrieved contact history, count:', result.data ? result.data.length : 0);
          sendResponse({success: true, contactHistory: result.data});
        } else {
          console.error('Failed to retrieve contact history:', result.error);
          sendResponse({success: false, error: result.error});
        }
      })
      .catch(error => {
        console.error('Exception when retrieving contact history:', error);
        sendResponse({success: false, error: error.message});
      });
    return true; // Required for async sendResponse
  }
  // Authentication related message handlers
  else if (request.action === 'signUp') {
    signUp(request.email, request.password)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'signIn') {
    signIn(request.email, request.password)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'signOut') {
    signOut()
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'createTeam') {
    createTeam(request.teamName)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'joinTeam') {
    joinTeam(request.teamId)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'leaveTeam') {
    leaveTeam()
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'transferAdminRole') {
    transferAdminRole(request.newAdminId)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'deleteTeam') {
    deleteTeam()
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'getTeamMembers') {
    getTeamMembers(request.teamId)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'getAllTeams') {
    getAllTeams()
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'getUserDetails') {
    getUserDetails(request.userId)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'isAuthenticated') {
    isAuthenticated()
      .then(result => {
        sendResponse({success: true, authenticated: result});
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'updateTeamDetails') {
    updateTeamDetails(request.teamId, request.teamDetails)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'requestToJoinTeam') {
    requestToJoinTeam(request.accessCode)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'removeTeamMember') {
    removeTeamMember(request.memberId)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'getPendingJoinRequests') {
    getPendingJoinRequests()
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'respondToJoinRequest') {
    respondToJoinRequest(request.requestId, request.approved)
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'checkJoinRequestStatus') {
    checkJoinRequestStatus()
      .then(result => {
        sendResponse(result);
      })
      .catch(error => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'openPopup') {
    // Open the extension popup
    chrome.action.openPopup()
      .then(() => {
        console.log('Popup opened successfully');
        sendResponse({success: true});
      })
      .catch(error => {
        console.error('Error opening popup:', error);
        sendResponse({success: false, error: error.message});
      });
    return true;
  }
  else if (request.action === 'markMessageAsHandled') {
    console.log('Marking message as handled:', request.messageId, 'by user:', request.userId);
    
    markMessageAsHandled(request.messageId, request.userId, request.note)
      .then(result => {
        console.log('Mark as handled result:', result);
        if (result.success) {
          sendResponse({success: true});
        } else {
          console.error('Failed to mark message as handled:', result.error);
          sendResponse({success: false, error: result.error});
        }
      })
      .catch(error => {
        console.error('Exception when marking message as handled:', error);
        sendResponse({success: false, error: error.message});
      });
    return true; // Required for async sendResponse
  }
  else if (request.action === 'addNoteToMessage') {
    console.log('Adding note to message:', request.messageId, 'by user:', request.userId, 'category:', request.category);
    
    addNoteToMessage(request.messageId, request.userId, request.noteBody, request.category)
      .then(result => {
        console.log('Add note result:', result);
        if (result.success) {
          sendResponse({success: true, data: result.data});
        } else {
          console.error('Failed to add note to message:', result.error);
          sendResponse({success: false, error: result.error});
        }
      })
      .catch(error => {
        console.error('Exception when adding note to message:', error);
        sendResponse({success: false, error: error.message});
      });
    return true; // Required for async sendResponse
  }
  else if (request.action === 'getMessageNotes') {
    console.log('Getting notes for message:', request.messageId);
    
    getMessageNotes(request.messageId)
      .then(result => {
        console.log('Get notes result:', result);
        if (result.success) {
          sendResponse({success: true, notes: result.data});
        } else {
          console.error('Failed to get notes for message:', result.error);
          sendResponse({success: false, error: result.error});
        }
      })
      .catch(error => {
        console.error('Exception when getting notes for message:', error);
        sendResponse({success: false, error: error.message});
      });
    return true; // Required for async sendResponse
  }
}); 

// Start periodic session checks
function startSessionChecks() {
  if (sessionCheckIntervalId === null) {
    console.log('Starting periodic session checks');
    // Run an initial check
    checkAndRefreshToken();
    
    // Set up interval for future checks
    sessionCheckIntervalId = setInterval(async () => {
      // Check if user has been inactive for too long
      const currentTime = Date.now();
      const inactiveTime = currentTime - lastActivityTime;
      
      if (inactiveTime > USER_INACTIVITY_THRESHOLD) {
        console.log(`User inactive for ${Math.round(inactiveTime/1000/60)} minutes, pausing session checks`);
        stopSessionChecks();
        return;
      }
      
      // User is still active, check token status
      console.log('Performing periodic token check');
      await checkAndRefreshToken();
    }, SESSION_CHECK_INTERVAL);
  }
}

// Stop periodic session checks
function stopSessionChecks() {
  if (sessionCheckIntervalId !== null) {
    console.log('Stopping periodic session checks due to inactivity');
    clearInterval(sessionCheckIntervalId);
    sessionCheckIntervalId = null;
  }
}

// Track user activity through tab events as well
chrome.tabs.onActivated.addListener(() => {
  lastActivityTime = Date.now();
  startSessionChecks();
});

chrome.tabs.onUpdated.addListener(() => {
  lastActivityTime = Date.now();
  startSessionChecks();
});

// Initialize session management on startup
chrome.runtime.onStartup.addListener(() => {
  console.log('Extension started, initializing session management');
  lastActivityTime = Date.now();
  startSessionChecks();
});

// Also start session management when the background script loads
console.log('Background script loaded, initializing session management');
lastActivityTime = Date.now();
startSessionChecks(); 