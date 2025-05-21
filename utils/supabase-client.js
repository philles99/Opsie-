// Supabase client for handling database operations
// Using a more compatible approach for Chrome extensions

import { SUPABASE_URL, SUPABASE_KEY } from './supabase-config.js';

// Remove default IDs as they're no longer needed
// (These were previously used for testing/development only)

/**
 * Get the user's team ID from Chrome storage
 * @returns {Promise<string|null>} - The team ID or null if not found
 */
async function getUserTeamId() {
  return new Promise((resolve) => {
    chrome.storage.sync.get(['currentTeamId'], (result) => {
      resolve(result.currentTeamId || null);
    });
  });
}

/**
 * Get the user's ID from Chrome storage
 * @returns {Promise<string|null>} - The user ID or null if not found
 */
async function getUserId() {
  return new Promise((resolve) => {
    chrome.storage.sync.get(['userId'], (result) => {
      resolve(result.userId || null);
    });
  });
}

/**
 * Store an email in Supabase using the threads and messages tables
 * @param {Object} emailData - The email data to store
 * @returns {Promise} - The result of the insert operation
 */
export async function storeEmail(emailData) {
  try {
    console.log('Storing email in Supabase:', emailData);
    
    // Check if user is authenticated
    const { accessToken, userId: currentUserId, currentTeamId: userTeamId } = await chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId']);
    
    if (!accessToken || !currentUserId || !userTeamId) {
      console.error('User not authenticated or not in a team. Cannot store email.');
      return { success: false, error: 'Authentication required to store emails' };
    }
    
    // Ensure timestamp is in ISO format
    let timestamp = emailData.timestamp;
    if (timestamp) {
      try {
        // If it's already an ISO string, keep it as is
        if (typeof timestamp === 'string' && timestamp.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) {
          // Already ISO format, keep as is
        } 
        // Special handling for Gmail date formats with parenthesis
        else if (typeof timestamp === 'string' && timestamp.includes('(') && timestamp.includes(')')) {
          console.log('Supabase client detected Gmail timestamp format:', timestamp);
          
          // Extract the date part before the parenthesis
          const datePart = timestamp.split('(')[0].trim();
          
          // Check if we have a full date with year (Apr 24, 2025)
          if (/\b\d{4}\b/.test(datePart)) {
            // Date has a year, parse normally
            timestamp = new Date(datePart).toISOString();
          } else {
            // Date doesn't have a year (like "Apr 24"), add current year
            const currentYear = new Date().getFullYear();
            const dateWithYear = `${datePart}, ${currentYear}`;
            console.log('Supabase client: Adding current year to Gmail date:', dateWithYear);
            timestamp = new Date(dateWithYear).toISOString();
          }
        }
        // For other date formats
        else {
          timestamp = new Date(timestamp).toISOString();
        }
      } catch (e) {
        console.error('Error converting timestamp in Supabase client:', e);
        // Use current time as fallback
        timestamp = new Date().toISOString();
      }
    } else {
      // If no timestamp, use current time
      timestamp = new Date().toISOString();
    }
    
    // Extract the message external ID from URL if available
    let messageExternalId = null;
    
    // First, check if we have a direct messageId from the parser
    if (emailData.messageId) {
      messageExternalId = emailData.messageId;
      console.log('Using messageId directly from email data:', messageExternalId);
    } 
    // If no direct messageId, try to extract from URL
    else if (emailData.messageUrl) {
      try {
        const url = emailData.messageUrl;
        console.log('Extracting message ID from URL:', url);
        
        // Extract ID from Gmail URL
        if (url.includes('mail.google.com')) {
          // Gmail format: https://mail.google.com/mail/u/0/?hl=sv#inbox/FMfcgzQbdrTrrfCfrVdgqZwbpChKPdPC
          const gmailMatches = url.match(/[#/](?:inbox|sent|drafts|trash|spam|category\/\w+)\/([^/?#]+)/i);
          if (gmailMatches && gmailMatches[1]) {
            messageExternalId = gmailMatches[1];
            console.log('Extracted Gmail message ID:', messageExternalId);
          } else {
            console.log('Gmail URL format did not match expected pattern:', url);
          }
        } 
        // Extract ID from Outlook Live URL
        else if (url.includes('outlook.live.com')) {
          // Outlook Live format: https://outlook.live.com/mail/0/inbox/id/AAkALgAAAAAAHYQDEapmEc2byACqAC%2FEWg0A8WNj85utG0G9uhi4NR9gtAAIDM874QAA
          const outlookLiveMatches = url.match(/\/id\/([^/?#]+)/i);
          if (outlookLiveMatches && outlookLiveMatches[1]) {
            messageExternalId = decodeURIComponent(outlookLiveMatches[1]);
            console.log('Extracted Outlook Live message ID:', messageExternalId);
          } else {
            // Try alternative patterns for Outlook Live
            const altOutlookMatches = url.match(/\/([A-Za-z0-9%]+)$/i); // Look for ID at the end of URL
            if (altOutlookMatches && altOutlookMatches[1]) {
              messageExternalId = decodeURIComponent(altOutlookMatches[1]);
              console.log('Extracted Outlook Live message ID (alternative pattern):', messageExternalId);
            } else {
              console.log('Outlook Live URL format did not match any expected patterns:', url);
            }
          }
        }
        // Extract ID from Outlook Office URL
        else if (url.includes('outlook.office.com')) {
          // Outlook Office format: https://outlook.office.com/mail/inbox/id/AAQkADZjZmUzMzkyLTg2OTgtNDNmYS05M2E3LTgxOTQxZmM2MmJlNQAQAFNvYVoeTbVIiOJrAy6zoGk%3D
          const outlookOfficeMatches = url.match(/\/id\/([^/?#]+)/i);
          if (outlookOfficeMatches && outlookOfficeMatches[1]) {
            messageExternalId = decodeURIComponent(outlookOfficeMatches[1]);
            console.log('Extracted Outlook Office message ID:', messageExternalId);
          } else {
            // Try alternative patterns for Outlook Office
            const altOfficeMatches = url.match(/\/([A-Za-z0-9%]+)$/i); // Look for ID at the end of URL
            if (altOfficeMatches && altOfficeMatches[1]) {
              messageExternalId = decodeURIComponent(altOfficeMatches[1]);
              console.log('Extracted Outlook Office message ID (alternative pattern):', messageExternalId);
            } else {
              console.log('Outlook Office URL format did not match any expected patterns:', url);
            }
          }
        } else {
          console.log('URL does not match any known email platform patterns:', url);
        }
      } catch (e) {
        console.error('Error extracting message ID from URL:', e);
      }
    } 
    // If we have neither messageId nor messageUrl, log it
    else {
      console.log('No messageId or messageUrl found in email data:', emailData);
    }
    
    // Generate a thread ID if not present
    const threadId = emailData.threadId || `thread-${Date.now()}`;
    
    // Use the provided team and user IDs or get from storage
    // No longer using default IDs - require authentication
    const teamId = emailData.teamId || userTeamId;
    const userId = emailData.userId || currentUserId;
    
    const headers = {
      'Content-Type': 'application/json',
      'apikey': SUPABASE_KEY,
      'Prefer': 'return=representation',
      'Authorization': `Bearer ${accessToken}`
    };
    
    // 1. First, check if the thread already exists
    const threadResponse = await fetch(`${SUPABASE_URL}/rest/v1/threads?thread_id=eq.${encodeURIComponent(threadId)}`, {
      method: 'GET',
      headers: headers
    });
    
    let threadData = await threadResponse.json();
    let threadUuid;
    
    // 2. If thread doesn't exist, create it
    if (!threadResponse.ok || threadData.length === 0) {
      console.log('Thread not found, creating new thread');
      
      const newThreadResponse = await fetch(`${SUPABASE_URL}/rest/v1/threads`, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify({
          thread_id: threadId,
          subject: emailData.subject,
          sender_email: emailData.sender.email,
          team_id: teamId
        })
      });
      
      if (!newThreadResponse.ok) {
        const error = await newThreadResponse.json();
        console.error('Error creating thread in Supabase:', error);
        return { success: false, error };
      }
      
      threadData = await newThreadResponse.json();
      threadUuid = threadData[0].id;
    } else {
      threadUuid = threadData[0].id;
    }
    
    // 3. Check if message with this external ID already exists for this team
    if (messageExternalId) {
      console.log('Checking if message with external ID already exists:', messageExternalId);
      const existingMessageResponse = await fetch(
        `${SUPABASE_URL}/rest/v1/messages?message_external_id=eq.${encodeURIComponent(messageExternalId)}&team_id=eq.${teamId}`,
        {
          method: 'GET',
          headers: headers
        }
      );
      
      if (existingMessageResponse.ok) {
        const existingMessages = await existingMessageResponse.json();
        if (existingMessages && existingMessages.length > 0) {
          console.log('Message with this external ID already exists. Skipping insertion.');
          return { 
            success: true, 
            data: {
              thread: threadData[0],
              message: existingMessages[0],
              skipped: true
            },
            message: 'Message already exists in database'
          };
        }
      }
    }
    
    // 4. Now add the message to the thread
    const messageResponse = await fetch(`${SUPABASE_URL}/rest/v1/messages`, {
      method: 'POST',
      headers: headers,
      body: JSON.stringify({
        thread_id: threadUuid,
        sender_name: emailData.sender.name,
        sender_email: emailData.sender.email,
        message_body: emailData.body || emailData.message, // Handle different property names
        timestamp: timestamp, // Use our sanitized timestamp
        summary: emailData.summary || null, // Store the AI-generated summary
        urgency: emailData.urgency || null, // Store the urgency score
        team_id: teamId, // Add team ID
        user_id: userId,  // Add user ID
        message_external_id: messageExternalId // Store the external message ID
      })
    });
    
    if (!messageResponse.ok) {
      const error = await messageResponse.json();
      console.error('Error storing message in Supabase:', error);
      return { success: false, error };
    }
    
    const messageData = await messageResponse.json();
    return { 
      success: true, 
      data: {
        thread: threadData[0],
        message: messageData[0]
      }
    };
  } catch (err) {
    console.error('Exception when storing email:', err);
    return { success: false, error: err };
  }
}

/**
 * Get all threads from Supabase
 * @returns {Promise} - The threads from the database
 */
export async function getThreads() {
  try {
    const response = await fetch(`${SUPABASE_URL}/rest/v1/threads?select=*&order=created_at.desc`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`
      }
    });
    
    if (!response.ok) {
      const error = await response.json();
      console.error('Error fetching threads from Supabase:', error);
      return { success: false, error };
    }
    
    const data = await response.json();
    return { success: true, data };
  } catch (err) {
    console.error('Exception when fetching threads:', err);
    return { success: false, error: err };
  }
}

/**
 * Get messages for a specific thread
 * @param {string} threadId - The UUID of the thread
 * @returns {Promise} - The messages from the database
 */
export async function getThreadMessages(threadId) {
  try {
    const response = await fetch(`${SUPABASE_URL}/rest/v1/messages?thread_id=eq.${threadId}&order=timestamp.asc`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`
      }
    });
    
    if (!response.ok) {
      const error = await response.json();
      console.error('Error fetching messages from Supabase:', error);
      return { success: false, error };
    }
    
    const data = await response.json();
    return { success: true, data };
  } catch (err) {
    console.error('Exception when fetching messages:', err);
    return { success: false, error: err };
  }
}

/**
 * Search for contact history by name or email within the same team
 * @param {Object} contactData - Object containing contact information
 * @param {string} contactData.name - The contact's name to search for
 * @param {string} contactData.email - The contact's email to search for
 * @param {string} teamId - The team ID to restrict the search to
 * @returns {Promise} - The contact history from the database
 */
export async function getContactHistory(contactData, teamId) {
  try {
    console.log('Searching for contact:', contactData, 'in team:', teamId);
    
    // Get access token for authenticated requests
    const { accessToken } = await chrome.storage.sync.get(['accessToken']);
    
    // Use the provided team ID, or get from storage, or use default
    const searchTeamId = teamId || await getUserTeamId();
    
    // Try different search approaches in sequence
    let messages = [];
    
    // 1. First try with both name and email if available (complex OR query)
    if (contactData.name && contactData.email) {
      try {
        console.log('Attempting complex OR query with both name and email');
        
        // Create proper PostgREST OR query format
        const params = new URLSearchParams();
        
        // Add the team_id parameter
        params.append('team_id', `eq.${searchTeamId}`);
        
        // Add the OR filter using the PostgREST or operator syntax
        // Handle special characters better
        const encodedName = encodeURIComponent(`%${contactData.name}%`);
        const encodedEmail = encodeURIComponent(`%${contactData.email}%`);
        
        params.append('or', `(sender_name.ilike.${encodedName},sender_email.ilike.${encodedEmail})`);
        params.append('order', 'timestamp.desc');
        // Add user_id to the selection fields
        params.append('select', 'id,sender_name,sender_email,message_body,summary,urgency,timestamp,thread_id,user_id,created_at');
        
        const queryUrl = `${SUPABASE_URL}/rest/v1/messages?${params.toString()}`;
        console.log('Complex query URL:', queryUrl);
        
        // Execute the query with the manually constructed parameters
        const response = await fetch(queryUrl, {
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            'apikey': SUPABASE_KEY,
            'Authorization': accessToken ? `Bearer ${accessToken}` : `Bearer ${SUPABASE_KEY}`
          }
        });
        
        if (!response.ok) {
          const error = await response.json();
          console.error('Error searching with complex query:', error);
          throw new Error('Complex query failed, trying next approach');
        }
        
        messages = await response.json();
        console.log(`Complex query found ${messages.length} messages`);
        
        if (messages.length > 0) {
          // Success! Continue with processing threads
        } else {
          throw new Error('No results from complex query, trying next approach');
        }
      } catch (complexQueryError) {
        console.warn('Complex query approach failed:', complexQueryError);
        
        // 2. Try email-only search as first fallback
        try {
          console.log('Trying email-only search as fallback');
          // Search by email only with ilike (case insensitive partial match)
          const emailQueryString = `team_id=eq.${searchTeamId}&sender_email=ilike.${encodeURIComponent(`%${contactData.email}%`)}`;
          // Add user_id to the selection fields
          const emailQueryUrl = `${SUPABASE_URL}/rest/v1/messages?${emailQueryString}&select=id,sender_name,sender_email,message_body,summary,urgency,timestamp,thread_id,user_id,created_at&order=timestamp.desc`;
          
          console.log('Email query URL:', emailQueryUrl);
          
          const emailResponse = await fetch(emailQueryUrl, {
            method: 'GET',
            headers: {
              'Content-Type': 'application/json',
              'apikey': SUPABASE_KEY,
              'Authorization': accessToken ? `Bearer ${accessToken}` : `Bearer ${SUPABASE_KEY}`
            }
          });
          
          if (!emailResponse.ok) {
            throw new Error('Email-only query failed');
          }
          
          messages = await emailResponse.json();
          console.log(`Email-only search found ${messages.length} messages`);
          
          if (messages.length > 0) {
            // Success! We'll continue with thread processing outside the catch blocks
          } else {
            throw new Error('No results from email-only search, trying direct match');
          }
        } catch (emailQueryError) {
          console.warn('Email-only search failed:', emailQueryError);
          
          // 3. Try exact email match as final fallback
          try {
            console.log('Trying exact email match as final fallback');
            // Use eq instead of ilike for exact matching
            const exactEmailQueryString = `team_id=eq.${searchTeamId}&sender_email=eq.${encodeURIComponent(contactData.email)}`;
            // Add user_id to the selection fields
            const exactEmailQueryUrl = `${SUPABASE_URL}/rest/v1/messages?${exactEmailQueryString}&select=id,sender_name,sender_email,message_body,summary,urgency,timestamp,thread_id,user_id,created_at&order=timestamp.desc`;
            
            console.log('Exact email query URL:', exactEmailQueryUrl);
            
            const exactEmailResponse = await fetch(exactEmailQueryUrl, {
              method: 'GET',
              headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_KEY,
                'Authorization': accessToken ? `Bearer ${accessToken}` : `Bearer ${SUPABASE_KEY}`
              }
            });
            
            if (!exactEmailResponse.ok) {
              throw new Error('Exact email match query failed');
            }
            
            messages = await exactEmailResponse.json();
            console.log(`Exact email match found ${messages.length} messages`);
          } catch (exactEmailError) {
            console.warn('Exact email match failed:', exactEmailError);
          }
        }
      }
    } 
    // Simple case: only name is available
    else if (contactData.name) {
      console.log('Searching by name only');
      // For name-only searches, try to sanitize any problematic characters
      const sanitizedName = contactData.name.replace(/[,()]/g, ' ');
      const nameQueryString = `team_id=eq.${searchTeamId}&sender_name=ilike.${encodeURIComponent(`%${sanitizedName}%`)}`;
      // Add user_id to the selection fields
      const nameQueryUrl = `${SUPABASE_URL}/rest/v1/messages?${nameQueryString}&select=id,sender_name,sender_email,message_body,summary,urgency,timestamp,thread_id,user_id,created_at&order=timestamp.desc`;
      
      console.log('Name-only query URL:', nameQueryUrl);
      
      const nameResponse = await fetch(nameQueryUrl, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': accessToken ? `Bearer ${accessToken}` : `Bearer ${SUPABASE_KEY}`
        }
      });
      
      if (!nameResponse.ok) {
        const error = await nameResponse.json();
        console.error('Error searching by name:', error);
        return { success: false, error };
      }
      
      messages = await nameResponse.json();
      console.log(`Name-only search found ${messages.length} messages`);
    } 
    // Simple case: only email is available
    else if (contactData.email) {
      console.log('Searching by email only');
      const emailQueryString = `team_id=eq.${searchTeamId}&sender_email=ilike.${encodeURIComponent(`%${contactData.email}%`)}`;
      // Add user_id to the selection fields
      const emailQueryUrl = `${SUPABASE_URL}/rest/v1/messages?${emailQueryString}&select=id,sender_name,sender_email,message_body,summary,urgency,timestamp,thread_id,user_id,created_at&order=timestamp.desc`;
      
      console.log('Email-only query URL:', emailQueryUrl);
      
      const emailResponse = await fetch(emailQueryUrl, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': accessToken ? `Bearer ${accessToken}` : `Bearer ${SUPABASE_KEY}`
        }
      });
      
      if (!emailResponse.ok) {
        const error = await emailResponse.json();
        console.error('Error searching by email:', error);
        return { success: false, error };
      }
      
      messages = await emailResponse.json();
      console.log(`Email-only search found ${messages.length} messages`);
      
      // If no results, try exact match
      if (messages.length === 0) {
        console.log('No results with partial match, trying exact email match');
        const exactEmailQueryString = `team_id=eq.${searchTeamId}&sender_email=eq.${encodeURIComponent(contactData.email)}`;
        // Add user_id to the selection fields
        const exactEmailQueryUrl = `${SUPABASE_URL}/rest/v1/messages?${exactEmailQueryString}&select=id,sender_name,sender_email,message_body,summary,urgency,timestamp,thread_id,user_id,created_at&order=timestamp.desc`;
        
        console.log('Exact email query URL:', exactEmailQueryUrl);
        
        const exactEmailResponse = await fetch(exactEmailQueryUrl, {
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            'apikey': SUPABASE_KEY,
            'Authorization': accessToken ? `Bearer ${accessToken}` : `Bearer ${SUPABASE_KEY}`
          }
        });
        
        if (exactEmailResponse.ok) {
          messages = await exactEmailResponse.json();
          console.log(`Exact email match found ${messages.length} messages`);
        }
      }
    } else {
      return { 
        success: false, 
        error: 'No search criteria provided - need name or email' 
      };
    }
    
    // If we found messages, get the thread data to add subject information
    if (messages.length > 0) {
      // Get unique thread IDs
      const threadIds = [...new Set(messages.map(msg => msg.thread_id))];
      
      // Format as a comma-separated list for the "in" operator
      const threadIdList = threadIds.map(id => `"${id}"`).join(',');
      
      const threadsResponse = await fetch(`${SUPABASE_URL}/rest/v1/threads?id=in.(${threadIdList})`, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': accessToken ? `Bearer ${accessToken}` : `Bearer ${SUPABASE_KEY}`
        }
      });
      
      if (threadsResponse.ok) {
        const threads = await threadsResponse.json();
        
        // Create a map of thread_id to subject
        const threadMap = {};
        threads.forEach(thread => {
          threadMap[thread.id] = thread.subject;
        });
        
        // Add subject to each message
        messages.forEach(message => {
          message.subject = threadMap[message.thread_id] || 'No Subject';
        });
      }
    }
    
    // Print some details about results to help with debugging
    if (messages.length > 0) {
      console.log('Found messages with sender emails:', messages.map(m => m.sender_email));
    } else {
      console.log('No messages found for any search approach');
    }
    
    return { 
      success: true, 
      data: messages,
      contactData: contactData // Return the original contact data for reference
    };
  } catch (err) {
    console.error('Exception when searching for contact history:', err);
    return { success: false, error: err };
  }
}

/**
 * Mark a message as handled
 * @param {string} messageId - The ID of the message to mark as handled
 * @param {string} userId - The ID of the user marking the message
 * @param {string} note - Optional note about how the message was handled
 * @returns {Promise} - Result of the update operation
 */
export async function markMessageAsHandled(messageId, userId, note = null) {
  try {
    console.log('Marking message as handled:', messageId, 'by user:', userId, 'with note:', note);
    
    const { accessToken } = await chrome.storage.sync.get(['accessToken']);
    
    if (!accessToken || !userId || !messageId) {
      console.error('Missing required data for marking message as handled');
      return { success: false, error: 'Missing required data (accessToken, userId, or messageId)' };
    }
    
    const headers = {
      'Content-Type': 'application/json',
      'apikey': SUPABASE_KEY,
      'Authorization': `Bearer ${accessToken}`
    };
    
    const response = await fetch(`${SUPABASE_URL}/rest/v1/messages?id=eq.${messageId}`, {
      method: 'PATCH',
      headers: headers,
      body: JSON.stringify({
        handled_at: new Date().toISOString(),
        handled_by: userId,
        handling_note: note
      })
    });
    
    if (!response.ok) {
      const error = await response.json();
      console.error('Error marking message as handled:', error);
      return { success: false, error };
    }
    
    console.log('Message successfully marked as handled');
    return { success: true };
  } catch (err) {
    console.error('Exception when marking message as handled:', err);
    return { success: false, error: err };
  }
}

/**
 * Check if a message with the given external ID already exists in the database for this team
 * @param {string} messageExternalId - The external ID of the message to check
 * @param {string} teamId - The team ID to check within
 * @param {Object} emailData - Optional email data for secondary check if external ID check fails
 * @returns {Promise} - Information about the existing message if found
 */
export async function checkMessageExists(messageExternalId, teamId, emailData = null) {
  try {
    // If no message ID or team ID provided, we can't check
    if (!messageExternalId && (!emailData || !teamId)) {
      return { exists: false };
    }

    console.log('Checking if message exists:', messageExternalId, 'for team:', teamId);
    
    // Get access token for authenticated requests
    const { accessToken } = await chrome.storage.sync.get(['accessToken']);
    
    if (!accessToken) {
      console.error('No access token found, cannot check if message exists');
      return { exists: false, error: 'Authentication required' };
    }
    
    const headers = {
      'Content-Type': 'application/json',
      'apikey': SUPABASE_KEY,
      'Authorization': `Bearer ${accessToken}`
    };
    
    // PRIMARY CHECK: Query for message with this external ID in this team
    if (messageExternalId) {
      // Now also retrieving handling information
      const messageResponse = await fetch(
        `${SUPABASE_URL}/rest/v1/messages?message_external_id=eq.${encodeURIComponent(messageExternalId)}&team_id=eq.${teamId}&select=id,created_at,user_id,summary,urgency,handled_at,handled_by,handling_note`,
        {
          method: 'GET',
          headers: headers
        }
      );
      
      if (!messageResponse.ok) {
        const error = await messageResponse.json();
        console.error('Error checking if message exists by ID:', error);
        // Continue to try the secondary check
      } else {
        const messages = await messageResponse.json();
        
        if (messages && messages.length > 0) {
          const messageData = messages[0];
          console.log('Message found by external ID:', messageData);
          
          // Initialize result object
          const result = await buildMessageResult(messageData, headers);
          return result;
        }
      }
    }
    
    // SECONDARY CHECK: If we have email data and the external ID check failed, try with sender and timestamp
    if (emailData && emailData.sender && emailData.sender.email && emailData.timestamp && teamId) {
      console.log('No match found by external ID, trying secondary check with sender and timestamp');
      
      try {
        // Standardize the timestamp to ISO format
        let timestamp;
        try {
          timestamp = new Date(emailData.timestamp).toISOString();
        } catch (e) {
          console.error('Error converting timestamp for secondary check:', e);
          timestamp = new Date().toISOString();
        }
        
        // Construct a query with a 2-minute window (±2 minutes) around the timestamp
        const twoMinutesBeforeTimestamp = new Date(new Date(timestamp).getTime() - 2 * 60 * 1000).toISOString();
        const twoMinutesAfterTimestamp = new Date(new Date(timestamp).getTime() + 2 * 60 * 1000).toISOString();
        
        console.log('Checking with time window:', {
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
          `&select=id,created_at,user_id,summary,urgency,handled_at,handled_by,handling_note,timestamp`;
        
        const timeWindowResponse = await fetch(
          `${SUPABASE_URL}/rest/v1/messages?${query}`,
          {
            method: 'GET',
            headers: headers
          }
        );
        
        if (!timeWindowResponse.ok) {
          const error = await timeWindowResponse.json();
          console.error('Error checking message by timestamp window:', error);
          return { exists: false, error };
        }
        
        const timeWindowMessages = await timeWindowResponse.json();
        
        if (timeWindowMessages && timeWindowMessages.length > 0) {
          console.log(`Found ${timeWindowMessages.length} messages matching sender and timestamp window`);
          
          // Sort messages by timestamp closest to the search timestamp
          timeWindowMessages.sort((a, b) => {
            const aDiff = Math.abs(new Date(a.timestamp) - new Date(timestamp));
            const bDiff = Math.abs(new Date(b.timestamp) - new Date(timestamp));
            return aDiff - bDiff;
          });
          
          // Use the closest match
          const closestMessage = timeWindowMessages[0];
          console.log('Using closest timestamp match:', {
            messageTimestamp: closestMessage.timestamp,
            searchTimestamp: timestamp,
            diffSeconds: Math.abs(new Date(closestMessage.timestamp) - new Date(timestamp)) / 1000
          });
          
          // Initialize result object with a note that it was found by secondary check
          const result = await buildMessageResult(closestMessage, headers);
          result.foundBySecondaryCheck = true;
          return result;
        } else {
          console.log('No matches found with timestamp window check');
        }
      } catch (secondaryCheckError) {
        console.error('Error during secondary message check:', secondaryCheckError);
      }
    }
    
    // Message doesn't exist by either check
    return { exists: false };
  } catch (err) {
    console.error('Exception when checking if message exists:', err);
    return { exists: false, error: err };
  }
}

/**
 * Helper function to build a result object from a message
 * @param {Object} messageData - The message data
 * @param {Object} headers - Headers for API requests
 * @returns {Promise<Object>} - The formatted result object
 */
async function buildMessageResult(messageData, headers) {
  // Initialize result object
  const result = { 
    exists: true, 
    message: messageData,
    savedAt: messageData.created_at,
    summary: messageData.summary,
    urgency: messageData.urgency
  };
  
  // Add handling info to the returned data if it exists
  if (messageData.handled_at) {
    const handlingInfo = {
      handledAt: messageData.handled_at,
      handledBy: null,
      handlingNote: messageData.handling_note
    };
    
    // If we have a user ID for who handled it, get user details
    if (messageData.handled_by) {
      try {
        // Fetch user details for who handled it
        const userResponse = await fetch(
          `${SUPABASE_URL}/rest/v1/users?id=eq.${messageData.handled_by}&select=first_name,last_name,email`,
          { method: 'GET', headers: headers }
        );
        
        if (userResponse.ok) {
          const users = await userResponse.json();
          if (users && users.length > 0) {
            const userData = users[0];
            const userName = `${userData.first_name || ''} ${userData.last_name || ''}`.trim() || userData.email || 'Unknown User';
            handlingInfo.handledBy = {
              name: userName,
              email: userData.email
            };
          }
        }
      } catch (userError) {
        console.error('Error fetching handler user details:', userError);
      }
    }
    
    // Add handling info to the return object
    result.handling = handlingInfo;
  }
  
  // If we have a user ID, get user details
  if (messageData.user_id) {
    try {
      // Fetch user details to get name
      const userResponse = await fetch(
        `${SUPABASE_URL}/rest/v1/users?id=eq.${messageData.user_id}&select=first_name,last_name,email`,
        {
          method: 'GET',
          headers: headers
        }
      );
      
      if (userResponse.ok) {
        const users = await userResponse.json();
        if (users && users.length > 0) {
          const userData = users[0];
          
          // Format the user's name
          let userName = 'Unknown User';
          if (userData.first_name || userData.last_name) {
            userName = `${userData.first_name || ''} ${userData.last_name || ''}`.trim();
          } else if (userData.email) {
            userName = userData.email;
          }
          
          result.user = {
            name: userName,
            email: userData.email
          };
        }
      }
    } catch (userError) {
      console.error('Error fetching user details:', userError);
    }
  }
  
  return result;
}

/**
 * Add a note to a specific message
 * @param {string} messageId - The ID of the message to add the note to
 * @param {string} userId - The ID of the user adding the note
 * @param {string} noteBody - The content of the note
 * @param {string} category - The category of the note (Action Required, Pending, Be Aware, etc.)
 * @returns {Promise} - Result of the insert operation
 */
export async function addNoteToMessage(messageId, userId, noteBody, category = null) {
  try {
    console.log('Adding note to message:', messageId, 'by user:', userId, 'category:', category);
    
    const { accessToken } = await chrome.storage.sync.get(['accessToken']);
    
    if (!accessToken || !userId || !messageId || !noteBody) {
      console.error('Missing required data for adding note to message');
      return { success: false, error: 'Missing required data (accessToken, userId, messageId, or noteBody)' };
    }
    
    const headers = {
      'Content-Type': 'application/json',
      'apikey': SUPABASE_KEY,
      'Authorization': `Bearer ${accessToken}`
    };
    
    const response = await fetch(`${SUPABASE_URL}/rest/v1/notes`, {
      method: 'POST',
      headers: headers,
      body: JSON.stringify({
        message_id: messageId,
        user_id: userId,
        note_body: noteBody,
        category: category,
        created_at: new Date().toISOString()
      })
    });
    
    if (!response.ok) {
      let errorData;
      try {
        errorData = await response.json();
      } catch (e) {
        // If parsing failed, use the status text
        errorData = { message: response.statusText };
      }
      console.error('Error adding note to message:', errorData);
      return { success: false, error: errorData };
    }
    
    // Check if there's content to parse
    const contentType = response.headers.get('content-type');
    let data = null;
    
    if (contentType && contentType.includes('application/json') && response.headers.get('content-length') !== '0') {
      try {
        // Only try to parse JSON if there's a JSON response
        data = await response.json();
      } catch (parseError) {
        console.log('No JSON response to parse, but operation was successful');
        // Continue as success even if parsing fails
      }
    } else {
      console.log('No content returned from Supabase, but operation was successful');
    }
    
    console.log('Note successfully added to message');
    return { success: true, data };
  } catch (err) {
    console.error('Exception when adding note to message:', err);
    return { success: false, error: err };
  }
}

/**
 * Get all notes for a specific message
 * @param {string} messageId - The ID of the message to get notes for
 * @returns {Promise} - The notes for the message
 */
export async function getMessageNotes(messageId) {
  try {
    console.log('Getting notes for message:', messageId);
    
    const { accessToken } = await chrome.storage.sync.get(['accessToken']);
    
    if (!accessToken || !messageId) {
      console.error('Missing required data for getting message notes');
      return { success: false, error: 'Missing required data (accessToken or messageId)' };
    }
    
    const headers = {
      'Content-Type': 'application/json',
      'apikey': SUPABASE_KEY,
      'Authorization': `Bearer ${accessToken}`
    };
    
    // Get notes for this message, ordered by creation date descending (newest first)
    const response = await fetch(
      `${SUPABASE_URL}/rest/v1/notes?message_id=eq.${messageId}&order=created_at.desc`,
      {
        method: 'GET',
        headers: headers
      }
    );
    
    if (!response.ok) {
      let errorData;
      try {
        errorData = await response.json();
      } catch (e) {
        // If parsing failed, use the status text
        errorData = { message: response.statusText };
      }
      console.error('Error getting notes for message:', errorData);
      return { success: false, error: errorData };
    }
    
    let notes = [];
    try {
      notes = await response.json();
      console.log(`Found ${notes.length} notes for message`);
    } catch (parseError) {
      console.error('Error parsing notes JSON:', parseError);
      return { success: true, data: [] }; // Return empty array on parse error
    }
    
    // If we found notes, get the user details for each note
    if (notes.length > 0) {
      try {
        // Get unique user IDs
        const userIds = [...new Set(notes.map(note => note.user_id))];
        
        // Format as a comma-separated list for the "in" operator
        const userIdList = userIds.map(id => `"${id}"`).join(',');
        
        const usersResponse = await fetch(`${SUPABASE_URL}/rest/v1/users?id=in.(${userIdList})&select=id,first_name,last_name,email`, {
          method: 'GET',
          headers: headers
        });
        
        if (usersResponse.ok) {
          let users = [];
          try {
            users = await usersResponse.json();
            
            // Create a map of user_id to user data
            const userMap = {};
            users.forEach(user => {
              userMap[user.id] = {
                name: `${user.first_name || ''} ${user.last_name || ''}`.trim() || user.email || 'Unknown User',
                email: user.email
              };
            });
            
            // Add user info to each note
            notes.forEach(note => {
              note.user = userMap[note.user_id] || { name: 'Unknown User', email: null };
            });
          } catch (userParseError) {
            console.error('Error parsing user data:', userParseError);
            // Continue with notes without user details
            notes.forEach(note => {
              note.user = { name: 'Unknown User', email: null };
            });
          }
        } else {
          // If user fetch fails, still continue with notes
          console.error('Failed to fetch user details for notes');
          notes.forEach(note => {
            note.user = { name: 'Unknown User', email: null };
          });
        }
      } catch (userError) {
        console.error('Exception when getting user details for notes:', userError);
        // Continue with notes without user details
        notes.forEach(note => {
          note.user = { name: 'Unknown User', email: null };
        });
      }
    }
    
    return { success: true, data: notes };
  } catch (err) {
    console.error('Exception when getting message notes:', err);
    return { success: false, error: err };
  }
}

export default { storeEmail, getThreads, getThreadMessages, getContactHistory, checkMessageExists, markMessageAsHandled, addNoteToMessage, getMessageNotes }; 