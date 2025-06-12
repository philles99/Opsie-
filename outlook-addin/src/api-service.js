/**
 * API Service for Opsie Email Assistant
 * This module handles all API communication between the add-in and the backend service
 */

// Configuration
const API_BASE_URL = 'https://vewnmfmnvumupdrcraay.supabase.co/rest/v1'; // Updated with the correct Supabase URL
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZld25tZm1udnVtdXBkcmNyYWF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDUyMTY2NDMsImV4cCI6MjA2MDc5MjY0M30.lX2oZscHBxlI9JDPDu4uoPgRNRZdV70ixYnOkENfPpc';
const API_TIMEOUT = 30000; // 30 seconds
const STORAGE_KEY_TOKEN = 'opsie_auth_token';
const STORAGE_KEY_REFRESH = 'opsie_refresh_token';
const STORAGE_KEY_OPENAI_API = 'openaiApiKey'; // Added OpenAI API key storage key

// Add caching for performance and to reduce API calls
const summaryCache = new Map();
const contactCache = new Map();
const replyCache = new Map();

// Create a global OpsieApi object
window.OpsieApi = {};

/**
 * General API request function with authentication
 * @param {string} endpoint - API endpoint path
 * @param {string} method - HTTP method (GET, POST, etc.)
 * @param {object} data - Request payload
 * @returns {Promise} - Promise with API response
 */
async function apiRequest(endpoint, method = 'GET', data = null) {
    try {
        // Get API token from storage
        const token = await getAuthToken();
        
        if (!token && !endpoint.includes('/auth/')) {
            // Show authentication error UI
            const authErrorContainer = document.getElementById('auth-error-container');
            if (authErrorContainer) {
                authErrorContainer.style.display = 'flex';
            }
            throw new Error('Authentication required');
        }

        // Make sure there's a slash between the base URL and endpoint
        const formattedEndpoint = endpoint.startsWith('/') ? endpoint : `/${endpoint}`;
        const url = `${API_BASE_URL}${formattedEndpoint}`;
        
        log('Making API request to:', 'info', { url, method });
        
        // Create abort controller for timeout
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), API_TIMEOUT);
        
        const options = {
            method: method,
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': token ? `Bearer ${token}` : `Bearer ${SUPABASE_ANON_KEY}`
            },
            // Set timeout signal
            signal: controller.signal
        };

        // Add request body for POST, PUT, PATCH requests
        if (data && (method === 'POST' || method === 'PUT' || method === 'PATCH')) {
            options.body = JSON.stringify(data);
        }

        // Make the request
        log('Making fetch request with options:', 'info', {
            url: url,
            method: options.method,
            headers: Object.fromEntries(Object.entries(options.headers).filter(([key]) => key !== 'Authorization')),
            hasBody: Boolean(options.body),
            bodyPreview: options.body ? truncateForLogging(options.body, 200) : null
        });

        const response = await fetch(url, options);
        
        // Clear the timeout since the request completed
        clearTimeout(timeoutId);
        
        // Log detailed response information for debugging
        log('API response received:', 'info', {
            url: url,
            method: method,
            status: response.status,
            statusText: response.statusText,
            headers: Object.fromEntries([...response.headers.entries()].map(entry => [entry[0], entry[1]]))
        });

        // Check for auth errors and attempt token refresh
        if (response.status === 401) {
            const refreshed = await refreshAuthToken();
            if (refreshed) {
                // Retry the request with the new token
                return apiRequest(endpoint, method, data);
            } else {
                // Show authentication error UI
                const authErrorContainer = document.getElementById('auth-error-container');
                if (authErrorContainer) {
                    authErrorContainer.style.display = 'flex';
                }
                throw new Error('Authentication failed');
            }
        }

        // Check for HTTP errors
        if (!response.ok) {
            let errorMessage = `API error: ${response.status} ${response.statusText}`;
            try {
                const errorData = await response.json();
                if (errorData && errorData.message) {
                    errorMessage = errorData.message;
                }
            } catch (parseError) {
                // If we can't parse the error response, just use the default message
                log('Error parsing error response:', 'warning', parseError);
            }
            
            showErrorNotification(errorMessage);
            throw new Error(errorMessage);
        }

        // Parse the response JSON
        let responseData;
        try {
            // Check if the response has content (based on headers or response text)
            const contentType = response.headers.get('content-type');
            const contentLength = response.headers.get('content-length');
            
            // First check if it's possible to have JSON content
            if (contentType && contentType.includes('application/json')) {
                // If content-length is 0 or null/undefined, or if it's a 204 No Content, handle appropriately
                if (contentLength === '0' || response.status === 204) {
                    log('Empty response (zero content length)', 'info');
                    responseData = null;
                } else {
                    // Check if there's actual content in the response
                    const responseText = await response.text();
                    if (!responseText || responseText.trim() === '') {
                        log('Empty response body', 'info');
                        responseData = null;
                    } else {
                        // Try to parse as JSON
                        responseData = JSON.parse(responseText);
                    }
                }
            } else {
                // Not JSON content type, just get the text
                const responseText = await response.text();
                log('Non-JSON response received:', 'info', {
                    contentType: contentType,
                    textLength: responseText ? responseText.length : 0
                });
                
                // For non-JSON responses (like text), just return the text in the data field
                responseData = responseText;
            }
        } catch (parseError) {
            log('Error parsing response:', 'error', parseError);
            // Return empty array for GET requests that should return collections
            if (method === 'GET') {
                return { success: true, data: [] };
            }
            
            // For POST/PUT operations, if the response was empty but the status is success (2xx),
            // we can consider it a success even without data
            if (response.status >= 200 && response.status < 300) {
                return { 
                    success: true, 
                    data: null,
                    status: response.status,
                    statusText: response.statusText
                };
            }
            
            throw new Error('Error parsing API response');
        }
        
        // Return a consistent response format
        const result = { 
            success: true, 
            data: responseData,
            status: response.status
        };
        
        // Add specific logging for qanda queries to debug source_filename issue
        if (endpoint.includes('qanda') && responseData && Array.isArray(responseData) && responseData.length > 0) {
            log('QandA query response data for debugging source_filename', 'info', {
                endpoint: endpoint,
                recordCount: responseData.length,
                firstRecord: responseData[0],
                sourceFilenameInFirstRecord: responseData[0].source_filename,
                sourceFilenameType: typeof responseData[0].source_filename,
                allFieldsInFirstRecord: Object.keys(responseData[0])
            });
        }
        
        return result;
    } catch (error) {
        // Handle different error types
        let errorMessage = error.message || 'An unexpected error occurred';
        
        if (error.name === 'AbortError') {
            errorMessage = 'Request timeout. Please try again later.';
        } else if (error.name === 'TypeError' && error.message.includes('Failed to fetch')) {
            errorMessage = 'Network error. Please check your internet connection.';
        }
        
        log('API request error:', 'error', { 
            message: errorMessage, 
            endpoint: endpoint,
            error: error
        });
        
        showErrorNotification(errorMessage);
        
        return {
            success: false,
            error: errorMessage
        };
    }
}

/**
 * Gets authentication token from storage or initiates login flow
 * @returns {Promise<string>} - JWT token
 */
async function getAuthToken() {
    try {
        log('Getting authentication token', 'info');
        
        // Try to get token from storage
        const token = localStorage.getItem(STORAGE_KEY_TOKEN);
        
        if (!token) {
            log('No token in localStorage', 'warning');
            return null;
        }
        
        log('Found token in localStorage', 'info');
        
        // Check if token exists and is valid
        if (token) {
            // Simple check if token is expired (assuming JWT format)
            try {
                const decoded = decodeJwtToken(token);
                if (!decoded) {
                    log('Token decode failed', 'warning');
                    return null;
                }
                
                if (decoded.exp && decoded.exp * 1000 < Date.now()) {
                    log('Token is expired, trying refresh', 'warning');
                    
                    // Try to refresh the token if we have a refresh token
                    const refreshed = await refreshAuthToken();
                    if (refreshed) {
                        log('Token refreshed successfully', 'info');
                        return localStorage.getItem(STORAGE_KEY_TOKEN);
                    } else {
                        log('Token refresh failed', 'error');
                        return null;
                    }
                }
                
                // Token is valid
                return token;
            } catch (tokenError) {
                log('Error checking token validity:', 'error', tokenError);
            }
        }
        
        // Token doesn't exist or is expired, and refresh failed
        return null;
    } catch (error) {
        log('Error in getAuthToken:', 'error', error);
        return null;
    }
}

/**
 * Attempts to refresh the authentication token
 * @returns {Promise<boolean>} - Whether the refresh was successful
 */
async function refreshAuthToken() {
    const refreshToken = localStorage.getItem(STORAGE_KEY_REFRESH);
    if (!refreshToken) {
        return false;
    }

    try {
        const response = await fetch(`https://vewnmfmnvumupdrcraay.supabase.co/auth/v1/token?grant_type=refresh_token`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY
            },
            body: JSON.stringify({ refresh_token: refreshToken })
        });

        if (!response.ok) {
            return false;
        }

        const data = await response.json();
        if (data.access_token) {
            localStorage.setItem(STORAGE_KEY_TOKEN, data.access_token);
            if (data.refresh_token) {
                localStorage.setItem(STORAGE_KEY_REFRESH, data.refresh_token);
            }
            return true;
        }
        return false;
    } catch (error) {
        console.error('Error refreshing token:', error);
        return false;
    }
}

/**
 * Displays an error notification
 * @param {string} message - Error message to display
 */
function showErrorNotification(message) {
    const notification = document.getElementById('error-notification');
    const notificationText = document.getElementById('error-notification-text');
    
    if (notification && notificationText) {
        notificationText.textContent = message;
        notification.style.display = 'flex';
        
        // Auto-hide after 5 seconds
        setTimeout(() => {
            notification.style.display = 'none';
        }, 5000);
    }
}

/**
 * Displays a success notification
 * @param {string} message - Success message to display
 */
function showNotification(message, type = 'success') {
    try {
        const notification = document.getElementById('notification');
        const notificationContent = document.querySelector('.notification-content');
        
        if (notification && notificationContent) {
            // Set notification content
            notificationContent.textContent = message;
            
            // Set notification type class
            notification.className = 'notification';
            notification.classList.add(`notification-${type}`);
            
            // Show the notification
            notification.style.display = 'flex';
            
            // Auto-hide after 5 seconds
            setTimeout(() => {
                notification.style.display = 'none';
            }, 5000);
            
            // Log the notification
            log(`Showing ${type} notification: ${message}`, 'info');
        } else {
            log('Could not show notification - elements not found', 'warning');
        }
    } catch (error) {
        log('Error showing notification: ' + error.message, 'error', error);
    }
}

/**
 * Safely decodes a JWT token to extract its payload
 * @param {string} token - The JWT token to decode
 * @returns {object|null} - The decoded token payload or null if invalid
 */
function decodeJwtToken(token) {
    if (!token) return null;
    
    try {
        // JWT tokens have three parts: header.payload.signature
        const parts = token.split('.');
        if (parts.length !== 3) {
            log('Invalid JWT token format', 'error');
            return null;
        }
        
        // Decode the payload (middle part)
        const payload = atob(parts[1].replace(/-/g, '+').replace(/_/g, '/'));
        
        // Parse the JSON payload
        return JSON.parse(payload);
    } catch (error) {
        log('Error decoding JWT token:', 'error', error);
        return null;
    }
}

/**
 * Generates an email summary
 * @param {object|string} emailContentOrData - The email content to summarize or an object containing email data
 * @param {string} [sender] - The email sender's name
 * @param {string} [subject] - The email subject
 * @param {Array} [threadHistory] - The email thread history
 * @returns {Promise<object>} - Promise with summary data
 */
async function generateEmailSummary(emailContentOrData, sender, subject, threadHistory = []) {
    // Handle case where a single object is passed (from Outlook add-in)
    let emailContent = emailContentOrData;
    
    if (typeof emailContentOrData === 'object' && emailContentOrData !== null) {
        log('Email data passed as object, extracting fields', 'info');
        // Extract fields from the object
        emailContent = emailContentOrData.body || '';
        sender = emailContentOrData.sender ? `${emailContentOrData.sender.name} <${emailContentOrData.sender.email}>` : '';
        subject = emailContentOrData.subject || '';
        threadHistory = emailContentOrData.threadHistory || [];
    }
    
    // Debug logs for input parameters
    log('===== EMAIL SUMMARY DEBUGGING =====', 'info');
    log('Input parameters received:', 'info', {
        emailContent: emailContent ? (typeof emailContent === 'string' ? emailContent.substring(0, 100) + '...' : 'Not a string') : 'null or undefined',
        emailContentType: typeof emailContent,
        emailContentLength: emailContent ? (typeof emailContent === 'string' ? emailContent.length : 'N/A') : 'N/A',
        sender: sender || 'null or undefined',
        senderType: typeof sender,
        subject: subject || 'null or undefined',
        subjectType: typeof subject,
        threadHistoryLength: threadHistory ? threadHistory.length : 'null or undefined',
        threadHistoryType: typeof threadHistory
    });

    // Set loading state
    window.OpsieApi.setLoading('summary', true);
    
    try {
        // Generate cache key based on content hash to avoid redundant API calls
        const cacheKey = `summary_${subject}_${sender}_${threadHistory.length}`;
        log('Using cache key:', 'info', cacheKey);
        
        // Check if we have a cached result
        const cachedResult = localStorage.getItem(cacheKey);
        if (cachedResult) {
            const parsedResult = JSON.parse(cachedResult);
            log('Using cached summary', 'info', parsedResult);
            window.OpsieApi.setLoading('summary', false);
            return parsedResult;
        }
        
        // Get OpenAI API key from storage
        const apiKey = localStorage.getItem('openaiApiKey');
        log('API key exists:', 'info', apiKey ? 'Yes (length: ' + apiKey.length + ')' : 'No');
        
        // If no API key, return default summary
        if (!apiKey) {
            log('No OpenAI API key found', 'warning');
            window.OpsieApi.setLoading('summary', false);
            return {
                summaryItems: [
                    "Please add your OpenAI API key in settings to generate summaries."
                ],
                urgencyScore: 5  // Changed from 0 to 5 to avoid database validation errors
            };
        }
        
        // Validate required fields
        if (!emailContent || !subject) {
            log('Missing required fields for email summary', 'error', {
                hasEmailContent: !!emailContent,
                hasSubject: !!subject
            });
            window.OpsieApi.setLoading('summary', false);
            return {
                summaryItems: [
                    "Error: Email is missing required content or subject.",
                    "Please make sure you have an email selected with a valid subject and body."
                ],
                urgencyScore: 5  // Changed from 0 to 5 to avoid database validation errors
            };
        }
        
        // Prepare thread history content
        let threadContent = "";
        if (threadHistory && threadHistory.length > 0) {
            threadContent = threadHistory.map(msg => 
                `From: ${msg.sender}\nTime: ${msg.time}\nContent: ${msg.content}`
            ).join('\n\n');
            log('Thread history processed:', 'info', {
                length: threadHistory.length,
                contentLength: threadContent.length,
                sampleThread: threadHistory.length > 0 ? 
                    { sender: threadHistory[0].sender, time: threadHistory[0].time } : 'None'
            });
        }
        
        // Create the API request body for debugging
        const requestBody = {
            model: "gpt-4o",
            messages: [
                {
                    role: "system",
                    content: `You are an AI assistant that summarizes emails. 
                    Your task is to extract the key points from the email and present them as a bullet point list.
                    Additionally, assign an urgency rating (high, medium, or low) based on the content and context.
                    Format your response as JSON with the following structure:
                    {
                        "summaryItems": [
                            "First key point",
                            "Second key point",
                            "Third key point"
                        ],
                        "urgencyScore": 7
                    }
                    Keep summaries concise and actionable. Identify any deadlines, requests, or important information.
                    The urgencyScore should be a number from 0 to 10, where 0 is lowest urgency and 10 is highest.`
                },
                {
                    role: "user",
                    content: `Summarize this email:
                    
                    Subject: ${subject}
                    From: ${sender}
                    
                    ${emailContent}
                    
                    ${threadHistory.length > 0 ? 'Previous thread history:\n' + threadContent : ''}
                    
                    Please provide your response as JSON only, with no additional text.`
                }
            ],
            temperature: 0.5
        };
        
        // Log the actual request payload that will be sent to OpenAI
        log('OpenAI API request payload:', 'info', {
            model: requestBody.model,
            systemContentLength: requestBody.messages[0].content.length,
            userContentPreview: requestBody.messages[1].content.substring(0, 100) + '...',
            userContentFull: requestBody.messages[1].content,
            temperature: requestBody.temperature
        });
        
        // Prepare the request to OpenAI API
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(requestBody)
        });
        
        // Check if the response is ok
        if (!response.ok) {
            const errorText = await response.text();
            log('OpenAI API error response:', 'error', {
                status: response.status,
                statusText: response.statusText,
                errorText: errorText
            });
            throw new Error(`OpenAI API error: ${response.status} - ${errorText}`);
        }
        
        // Parse the response
        const responseData = await response.json();
        log('OpenAI API raw response:', 'info', responseData);
        
        // Check if we have the expected response structure
        if (!responseData.choices || !responseData.choices[0] || !responseData.choices[0].message) {
            log('Unexpected API response format', 'error', responseData);
            throw new Error('Unexpected API response format');
        }
        
        // Extract the content from the response
        const aiResponse = responseData.choices[0].message.content;
        log('AI response content:', 'info', aiResponse);
        
        // Parse the JSON from the AI response
        let result;
        try {
            // Extract the JSON content from potential Markdown code blocks
            const jsonContent = extractContentFromCodeBlock(aiResponse);
            log('Extracted JSON content:', 'info', jsonContent);
            
            // Parse the JSON
            result = JSON.parse(jsonContent);
            log('Parsed JSON result:', 'info', result);
            
            // Basic validation
            if (!result.summaryItems || !Array.isArray(result.summaryItems) || result.urgencyScore === undefined) {
                log('AI response validation failed:', 'error', {
                    hasSummaryItems: !!result.summaryItems,
                    isSummaryItemsArray: Array.isArray(result.summaryItems),
                    hasUrgencyScore: result.urgencyScore !== undefined
                });
                throw new Error('AI response is missing required fields');
            }
            
            // Cache the result
            localStorage.setItem(cacheKey, JSON.stringify(result));
            log('Result cached with key:', 'info', cacheKey);
            
        } catch (parseError) {
            log('Error parsing AI response:', 'error', {
                error: parseError.message,
                aiResponse: aiResponse
            });
            // If parsing fails, create a fallback structure
            result = {
                summaryItems: [
                    "AI response format error. Please try again."
                ],
                urgencyScore: 0
            };
        }
        
        window.OpsieApi.setLoading('summary', false);
        log('Summary generation completed', 'info', result);
        return result;
        
    } catch (error) {
        log('Error generating summary: ' + error.message, 'error', error);
        window.OpsieApi.setLoading('summary', false);
        
        // Return a default summary on error
        return {
            summaryItems: [
                `Error: ${error.message}`,
                "Please check your API key and try again."
            ],
            urgencyScore: 0
        };
    }
}

/**
 * Generate contact history from the API
 * @param {string} emailAddress The email address to get contact history for
 * @returns {Promise<Object>} The contact history result
 */
async function generateContactHistory(emailAddress) {
    try {
        // Set loading state
        window.OpsieApi.setLoading('contact', true);
        
        log(`Generating contact history for email: ${emailAddress}`, 'info');
        
        // Generate a cache key
        const cacheKey = `contactHistory_${emailAddress}`;
        
        // Check if we have a cached result
        const cachedResult = localStorage.getItem(cacheKey);
        if (cachedResult) {
            try {
                const parsedResult = JSON.parse(cachedResult);
                log('Found cached contact history', 'info');
                window.OpsieApi.setLoading('contact', false);
                return parsedResult;
            } catch (error) {
                log('Error parsing cached contact history', 'error', error);
                // Continue with API call if parsing fails
            }
        }
        
        // Get team ID from local storage - check multiple possible keys
        let teamId = localStorage.getItem('currentTeamId');
        
        if (!teamId) {
            // Try alternative keys if primary key not found
            const alternateKeys = ['opsieTeamId', 'teamId', 'team_id'];
            
            for (const key of alternateKeys) {
                const altTeamId = localStorage.getItem(key);
                if (altTeamId) {
                    log(`Found team ID using alternate key: ${key}`, 'info');
                    teamId = altTeamId;
                    // Save it under the primary key for future use
                    localStorage.setItem('currentTeamId', teamId);
                    break;
                }
            }
        }
        
        if (!teamId) {
            log('No team ID found for contact history', 'warning');
            window.OpsieApi.setLoading('contact', false);
            return {
                summaryItems: [
                    "No team selected. Please select a team in settings."
                ]
            };
        }
        
        log(`Using team ID: ${teamId} for contact history`, 'info');
        
        // Get API key from localStorage
        const apiKey = localStorage.getItem('openaiApiKey');
        if (!apiKey) {
            log('No API key found for contact history', 'warning');
            window.OpsieApi.setLoading('contact', false);
            return {
                summaryItems: [
                    "Please add your OpenAI API key in settings to generate contact history."
                ]
            };
        }
        
        // Get authentication token for database requests
        const token = localStorage.getItem(STORAGE_KEY_TOKEN);
        if (!token) {
            log('No authentication token found for database access', 'warning');
            window.OpsieApi.setLoading('contact', false);
            return {
                summaryItems: [
                    "Authentication required to access contact history."
                ]
            };
        }
        
        // Step 1: Fetch contact history from database
        let contactHistory = [];
        try {
            log('Fetching previous messages from database', 'info');
            
            // Query the database for messages from this sender in this team
            const messagesResponse = await fetch(
                `https://vewnmfmnvumupdrcraay.supabase.co/rest/v1/messages?` +
                `sender_email=eq.${encodeURIComponent(emailAddress)}` +
                `&team_id=eq.${teamId}` +
                `&select=id,sender_name,sender_email,message_body,timestamp,created_at,summary,urgency,user_id` +
                `&order=timestamp.desc,created_at.desc` +
                `&limit=10`,
                {
                    method: 'GET',
                    headers: {
                        'Content-Type': 'application/json',
                        'apikey': SUPABASE_ANON_KEY,
                        'Authorization': `Bearer ${token}`
                    }
                }
            );
            
            if (!messagesResponse.ok) {
                const errorText = await messagesResponse.text();
                log(`Error fetching messages: ${errorText}`, 'error');
                throw new Error(`Database error: ${messagesResponse.status} ${errorText}`);
            }
            
            // Parse the messages
            contactHistory = await messagesResponse.json();
            log(`Retrieved ${contactHistory.length} previous messages from this contact`, 'info');
            
            // If we have user IDs, try to get user names
            const userIds = [...new Set(contactHistory.filter(msg => msg.user_id).map(msg => msg.user_id))];
            
            if (userIds.length > 0) {
                log(`Fetching user details for ${userIds.length} users`, 'info');
                
                // Fetch user details using RPC function to bypass RLS
                const usersResponse = await fetch(
                    `https://vewnmfmnvumupdrcraay.supabase.co/rest/v1/rpc/get_user_details`,
                    {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'apikey': SUPABASE_ANON_KEY,
                            'Authorization': `Bearer ${token}`
                        },
                        body: JSON.stringify({
                            user_ids: userIds
                        })
                    }
                );
                
                if (usersResponse.ok) {
                    const users = await usersResponse.json();
                    log(`Retrieved information for ${users.length} users`, 'info');
                    
                    // Create a map of user_id to user name
                    const userMap = new Map();
                    users.forEach(user => {
                        const userName = `${user.first_name || ''} ${user.last_name || ''}`.trim() || user.email || 'Unknown user';
                        userMap.set(user.id, userName);
                    });
                    
                    // Add user information to the contact history
                    contactHistory.forEach(msg => {
                        if (msg.user_id && userMap.has(msg.user_id)) {
                            msg.saved_by = userMap.get(msg.user_id);
                        } else {
                            msg.saved_by = "Unknown user";
                        }
                    });
                } else {
                    log('Failed to retrieve user information', 'warning');
                }
            }
        } catch (dbError) {
            log('Error retrieving contact history from database', 'error', dbError);
            // Continue with API call but with empty contact history
            contactHistory = [];
        }
        
        // Step 2: Format contact history for the API call
        const formattedHistory = contactHistory.map((email, index) => {
            // Use timestamp or created_at for date
            const dateStr = email.timestamp || email.created_at || 'Unknown date';
            const date = new Date(dateStr).toLocaleDateString();
            
            return `Email ${index + 1} (${date}):\n` +
                `From: ${email.sender_name || 'Unknown'} (${email.sender_email || 'No email'})\n` + 
                `Saved by: ${email.saved_by || 'Unknown user'}\n` +
                `Summary: ${email.summary || 'No summary available'}\n` +
                `Content: ${email.message_body ? (email.message_body.substring(0, 300) + '...') : 'No content available'}`;
        }).join('\n\n');
        
        // Create the message for the API call
        let userContent = `Generate a summary of my interaction history with the contact ${emailAddress}.`;
        
        // Add the history if we have any
        if (contactHistory.length > 0) {
            // Sort contact history to get the most recent interaction first
            const sortedHistory = [...contactHistory].sort((a, b) => {
                return new Date(b.timestamp || b.created_at || 0) - new Date(a.timestamp || a.created_at || 0);
            });
            
            // Get the date of the most recent interaction
            const mostRecentDate = sortedHistory[0].timestamp || sortedHistory[0].created_at;
            const formattedMostRecentDate = mostRecentDate ? new Date(mostRecentDate).toLocaleDateString() : 'unknown date';
            
            // Add information about the most recent interaction
            userContent += `\n\nThe most recent interaction was on ${formattedMostRecentDate}.`;
            userContent += `\n\nBased on these previous ${contactHistory.length} emails:\n\n${formattedHistory}`;
        } else {
            userContent += ' If there is no history, suggest that this might be a new contact.';
        }
        
        // Log the content being sent to the API (similar to browser extension)
        log('Contact history API request data:', 'info', {
            emailAddress,
            teamId,
            messageCount: contactHistory.length,
            requestContent: userContent.substring(0, 200) + '...' // Preview of the request content
        });
        
        // Prepare the API request
        const requestBody = {
            model: 'gpt-4o',
            messages: [
                {
                    role: 'system',
                    content: 'You are a helpful assistant that generates a concise summary of interaction history with a contact. Return only a JSON object with a "summaryItems" array containing 3-5 bullet points (as strings) about the contact. Each bullet point should be a key insight about past interactions. IMPORTANT: Do not include technical details like team IDs, database information, or implementation details in your summary - focus only on the business/personal relationship context that would be useful to the user. The LAST bullet point in your response MUST always be about when the last interaction with the contact occurred (e.g., "Last contacted on [specific date]" or "Most recent communication was 2 weeks ago about..."). Format your response as a valid JSON object.'
                },
                {
                    role: 'user',
                    content: userContent
                }
            ],
            temperature: 0.7
        };
        
        // Log the full request for debugging
        log('Full contact history request to OpenAI:', 'info', {
            model: requestBody.model,
            systemPromptLength: requestBody.messages[0].content.length,
            userContentLength: requestBody.messages[1].content.length,
            temperature: requestBody.temperature
        });
        
        // Make API request to get contact history
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(requestBody)
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            log(`API error for contact history: ${errorText}`, 'error');
            throw new Error(`API error: ${response.status} ${errorText}`);
        }
        
        // Parse the response JSON
        const responseData = await response.json();
        log('Received contact history response', 'info');
        
        // The newer OpenAI API format nests the content inside choices[0].message.content
        if (!responseData.choices || !responseData.choices[0] || !responseData.choices[0].message) {
            log('Unexpected API response format', 'error', responseData);
            throw new Error('Unexpected API response format');
        }
        
        // Get the actual content from the message
        const messageContent = responseData.choices[0].message.content;
        log('Extracted message content:', 'info', messageContent);
        
        // Extract the JSON content from potential Markdown code blocks
        const jsonData = extractContentFromCodeBlock(messageContent);
        log('JSON data extracted for parsing:', 'info', jsonData);
        
        // Parse the JSON data
        let result;
        try {
            result = JSON.parse(jsonData);
            
            // Validate that we have the expected structure
            if (!result || !Array.isArray(result.summaryItems)) {
                log('Invalid contact history response format', 'error', result);
                throw new Error('Invalid response format. Expected summaryItems array.');
            }
            
            // Add message count to the result
            result.messageCount = contactHistory.length;
            
            // Cache the result
            localStorage.setItem(cacheKey, JSON.stringify(result));
            
            log('Successfully processed contact history', 'info', result);
            window.OpsieApi.setLoading('contact', false);
            return result;
        } catch (error) {
            log('Error parsing contact history response', 'error', { error, data: jsonData });
            window.OpsieApi.setLoading('contact', false);
            return {
                summaryItems: [
                    "Error processing contact history. Please try again later."
                ]
            };
        }
    } catch (error) {
        log('Error generating contact history', 'error', error);
        window.OpsieApi.setLoading('contact', false);
        return {
            summaryItems: [
                "Error generating contact history: " + error.message
            ]
        };
    }
}

/**
 * Fetches contact history from Supabase
 * @param {string} emailAddress - The contact's email address
 * @param {string} teamId - The team ID
 * @returns {Promise<Array>} - Promise with contact history
 */
async function fetchContactHistory(emailAddress, teamId) {
    try {
        log('Fetching contact history for email', 'info', { email: emailAddress, teamId });
        
        if (!emailAddress) {
            throw new Error('Email address is required');
        }
        
        if (!teamId) {
            throw new Error('Team ID is required');
        }
        
        // Get access token
        const token = await getAuthToken();
        if (!token) {
            throw new Error('Not authenticated');
        }
        
        // Directly query the messages table in Supabase
        // Create query parameters for Supabase
        const params = new URLSearchParams();
        
        // Filter by team ID and sender email (using ILIKE for case-insensitive partial match)
        params.append('team_id', `eq.${teamId}`);
        
        // Fix the encoding issue - we need to encode the email properly without double-encoding
        // The URL parameter should be 'ilike.%email%' where email is the actual email address
        const cleanEmail = emailAddress.replace(/%/g, ''); // Remove any % that might already be there
        log('Using email for query', 'info', { 
            originalEmail: emailAddress, 
            cleanEmail: cleanEmail 
        });
        params.append('sender_email', `ilike.%${cleanEmail}%`);
        
        // Order by timestamp descending to get newest messages first
        params.append('order', 'timestamp.desc');
        
        // Select the fields we need - NOTE: 'subject' column doesn't exist in the messages table
        params.append('select', 'id,sender_name,sender_email,message_body,timestamp,thread_id,user_id,created_at');
        
        // Create the full query URL
        const queryUrl = `messages?${params.toString()}`;
        log('Full Supabase query URL', 'info', queryUrl);
        
        // Make the API request
        const response = await apiRequest(queryUrl, 'GET');
        
        log('Contact history API response', 'info', { 
            status: response.status,
            count: response.data ? response.data.length : 0
        });
        
        // Log the actual data received from Supabase for debugging
        if (response.data && response.data.length > 0) {
            // Log the first record's structure
            const firstRecord = response.data[0];
            log('Sample record from Supabase', 'info', {
                id: firstRecord.id,
                sender_name: firstRecord.sender_name,
                sender_email: firstRecord.sender_email,
                message_body_length: firstRecord.message_body ? firstRecord.message_body.length : 0,
                message_body_sample: firstRecord.message_body ? 
                    truncateForLogging(firstRecord.message_body, 100) : 'NULL OR EMPTY',
                timestamp: firstRecord.timestamp,
                created_at: firstRecord.created_at,
                user_id: firstRecord.user_id
            });
            
            // Check all records for message_body content
            const contentSummary = response.data.map((record, index) => ({
                index,
                has_content: Boolean(record.message_body),
                length: record.message_body ? record.message_body.length : 0,
                user_id: record.user_id
            }));
            log('Message body content summary for all records', 'info', contentSummary);
            
            // Create a map to store user information
            const userMap = new Map();
            
            // Extract all user IDs from contact history
            const userIds = response.data
                .map(message => message.user_id)
                .filter(id => id); // Filter out null/undefined
                
            // Only proceed if we have user IDs to look up
            if (userIds.length > 0) {
                // Remove duplicates
                const uniqueUserIds = [...new Set(userIds)];
                log(`Found ${uniqueUserIds.length} unique user IDs to look up`, 'info', uniqueUserIds);
                
                try {
                    // Query the users table using RPC function to bypass RLS
                    log('Fetching user details using RPC function', 'info', uniqueUserIds);
                    
                    const usersResponse = await apiRequest(`rpc/get_user_details`, 'POST', {
                        user_ids: uniqueUserIds
                    });
                    
                    if (usersResponse.success && usersResponse.data) {
                        log(`Retrieved information for ${usersResponse.data.length} users`, 'info', usersResponse.data);
                        
                        // Create a map of user_id to user name
                        usersResponse.data.forEach(user => {
                            const userName = `${user.first_name || ''} ${user.last_name || ''}`.trim() || user.email || 'Unknown user';
                            userMap.set(user.id, userName);
                        });
                        
                        log('User map created', 'info', Object.fromEntries([...userMap.entries()]));
                    } else {
                        log('Failed to retrieve user information', 'warning', usersResponse);
                    }
                } catch (userLookupError) {
                    log('Error retrieving user information', 'error', userLookupError);
                }
            }
            
            // Add user names to the response data
            response.data.forEach(record => {
                if (record.user_id && userMap.has(record.user_id)) {
                    record.saved_by_name = userMap.get(record.user_id);
                } else {
                    record.saved_by_name = 'Team member';
                }
            });
            
        } else {
            log('No records returned from Supabase query', 'warning', {
                email: emailAddress,
                teamId: teamId,
                queryUrl: queryUrl
            });
        }
        
        // Return the email history
        return response.data || [];
    } catch (error) {
        log('Error fetching contact history', 'error', error);
        return [];
    }
}

/**
 * Helper function to truncate text for logging
 * @param {string} text - The text to truncate
 * @param {number} maxLength - Maximum length before truncation
 * @returns {string} - Truncated text with indicator
 */
function truncateForLogging(text, maxLength = 1000) {
    if (!text) return '';
    if (text.length <= maxLength) return text;
    
    const halfLength = Math.floor(maxLength / 2);
    return `${text.substring(0, halfLength)}... [${text.length - maxLength} characters truncated] ...${text.substring(text.length - halfLength)}`;
}

/**
 * Searches through email history
 * @param {string} emailAddress - Contact's email address
 * @param {string} query - Search query
 * @returns {Promise<object>} - Promise with search results
 */
async function searchEmailHistory(emailAddress, query) {
    try {
        log('Searching email history with query:', 'info', { query, emailAddress });
        
        // Set loading state
        setLoading('search', true);
        
        if (!emailAddress) {
            throw new Error('No email address provided for search');
        }
        
        // Get the team ID
        const teamId = localStorage.getItem('currentTeamId');
        if (!teamId) {
            throw new Error('No team ID available');
        }
        
        log('Fetching contact history for search', 'info', { email: emailAddress, teamId });
        
        // Fetch contact history from Supabase
        const contactHistory = await fetchContactHistory(emailAddress, teamId);
        
        log('Contact history fetched for search', 'info', { 
            count: contactHistory ? contactHistory.length : 0 
        });
        
        // Get OpenAI API key
        const apiKey = await getOpenAIApiKey();
        
        if (!apiKey) {
            throw new Error('OpenAI API key not configured');
        }
        
        // Use local emailData from handleSearch if available, otherwise use minimal data
        // Try to get current email data from window if available
        let emailData = window.currentEmailData || {};
        
        // Prepare the current email data for searching
        const currentEmailFormatted = {
            id: emailData.messageId || 'current-email',
            sender: emailData.sender ? 
                `${emailData.sender.name || 'Unknown'} (${emailData.sender.email || 'No email'})` : 
                `Unknown sender (${emailAddress})`,
            subject: emailData.subject || 'Current email',
            date: emailData.date || emailData.timestamp || new Date().toISOString(),
            content: emailData.body || 'Email content not available',
            savedBy: (emailData.existingMessage && emailData.existingMessage.user) ? 
                emailData.existingMessage.user.name : 
                'Current user',
            savedAt: (emailData.existingMessage && emailData.existingMessage.savedAt) ? 
                new Date(emailData.existingMessage.savedAt).toLocaleString() : 
                new Date().toLocaleString()
        };
        
        // Format the saved emails from contact history
        let formattedHistory = [];
        
        if (contactHistory && contactHistory.length > 0) {
            // Sort by date descending (newest first)
            const sortedHistory = [...contactHistory].sort((a, b) => {
                return new Date(b.timestamp || b.created_at || 0) - new Date(a.timestamp || a.created_at || 0);
            });
            
            // Log the raw contact history before formatting
            log('Raw contact history before formatting', 'info', contactHistory.map(record => ({
                id: record.id, 
                message_body_length: record.message_body ? record.message_body.length : 0,
                has_message_body: Boolean(record.message_body)
            })));
            
            formattedHistory = sortedHistory.map((email, index) => {
                // Check if message body exists and log
                const hasContent = Boolean(email.message_body && email.message_body.trim());
                if (!hasContent) {
                    log('Missing content for email record', 'warning', {
                        id: email.id,
                        sender: email.sender_name || email.sender_email,
                        message_body: email.message_body || 'NULL/EMPTY'
                    });
                }
                
                return {
                    id: email.id || `history-${index}`,
                    sender: email.sender_name ? 
                        `${email.sender_name} (${email.sender_email || 'No email'})` : 
                        email.sender_email || 'Unknown sender',
                    subject: 'No subject', // Since 'subject' doesn't exist in the database, use a default value
                    date: email.timestamp || email.created_at || 'Unknown date',
                    content: email.message_body || 'No content available',
                    savedBy: email.saved_by_name || 'Team member',
                    savedAt: email.created_at ? 
                        new Date(email.created_at).toLocaleString() : 
                        'Unknown date'
                };
            });
            
            // Log the formatted history
            log('Formatted email history', 'info', formattedHistory.map(email => ({
                id: email.id,
                content_length: email.content ? email.content.length : 0,
                has_content: email.content !== 'No content available'
            })));
        }
        
        // Combine current email and history - if we have contact history, only use that
        const allEmails = contactHistory.length > 0 ? 
            formattedHistory :
            [currentEmailFormatted, ...formattedHistory];
        
        // Format the emails for the API call
        const emailsContent = allEmails.map((email, index) => {
            // Ensure we have some minimal content
            let contentToUse = email.content;
            
            // If empty content, use a more informative placeholder
            if (!contentToUse || contentToUse === 'No content available' || contentToUse === 'Email content not available') {
                contentToUse = 'This email has no available content. It may not have been saved or it may be a placeholder.';
            }
            
            return `
Email #${index + 1}:
Sender: ${email.sender}
Subject: ${email.subject}
Date: ${email.date}
Saved by: ${email.savedBy}
Saved at: ${email.savedAt}
Content: ${contentToUse.substring(0, 1000)}${contentToUse.length > 1000 ? '...(truncated)' : ''}
            `;
        }).join('\n\n');
        
        // Create the system message
        const systemMessage = `You are an AI assistant that searches through email content to find specific information. Your task is to:

1. Analyze the provided emails to find information relevant to the user's query
2. First provide a brief 1-2 sentence answer to the user's question
3. Then list up to 3 relevant references from the emails that support your answer
4. For each reference, provide a short quote with the relevant information and include who saved it and when

Format your response exactly like this:
[Your 1-2 sentence answer]

References:
1. "[Short relevant quote from email]" - Saved by [Name] on [Date]
2. "[Short relevant quote from email]" - Saved by [Name] on [Date]
3. "[Short relevant quote from email]" - Saved by [Name] on [Date]

If you can't find relevant information, say so clearly. Only include references that actually contain information relevant to the query. If fewer than 3 emails contain relevant information, only include those. Sort references with newest emails first.`;

        // Create the user message with the query and email content
        const userMessage = `Query: ${query}\n\nHere are the emails to search through:\n\n${emailsContent}`;
        
        log('Making OpenAI API call for search', 'info', { 
            emailCount: allEmails.length,
            contentLength: emailsContent.length
        });

        // Add detailed logging of the email content and API request
        log('Email content being sent to OpenAI', 'info', truncateForLogging(emailsContent, 2000));

        // Prepare the full API request payload
        const openAIPayload = {
            model: 'gpt-3.5-turbo',
            messages: [
                {
                    role: 'system',
                    content: systemMessage
                },
                {
                    role: 'user',
                    content: userMessage
                }
            ],
            temperature: 0.7,
            max_tokens: 500
        };

        // Log the full API request payload (exclude the full content for readability)
        const payloadForLogging = {
            ...openAIPayload,
            messages: [
                {
                    role: 'system',
                    content: truncateForLogging(systemMessage, 500)
                },
                {
                    role: 'user', 
                    content: `Query: ${query}\n\n[Email content truncated for logging]`
                }
            ]
        };
        log('OpenAI search API request payload', 'info', payloadForLogging);

        // Make the API call
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(openAIPayload)
        });

        // Get the response as text first for logging
        const responseText = await response.text();
        log('Raw OpenAI API response for search:', 'info', responseText);

        // Parse the JSON response
        let data;
        try {
            data = JSON.parse(responseText);
    } catch (error) {
            log('Error parsing OpenAI API response:', 'error', error);
            throw new Error('Failed to parse OpenAI API response');
        }

        if (data.error) {
            log('OpenAI API error:', 'error', data.error);
            throw new Error(data.error.message);
        }

        // Extract the response
        const answer = data.choices[0].message.content.trim();
        log('Search results:', 'info', answer);
        
        // Parse the answer to split it into the summary and references
        let mainAnswer = '';
        let references = [];
        
        // Try to extract the main answer (everything before 'References:')
        const referenceSplit = answer.split(/references:/i);
        if (referenceSplit.length > 1) {
            mainAnswer = referenceSplit[0].trim();
            
            // Extract references
            const referencesText = referenceSplit[1].trim();
            
            // Look for numbered items (1., 2., etc.) as reference starters
            const referenceLines = referencesText.split(/\n+/).filter(line => line.trim());
            
            // Process each reference line
            for (let i = 0; i < referenceLines.length; i++) {
                const line = referenceLines[i].trim();
                
                // Skip empty lines or lines that don't start with a number
                if (!line || !/^\d+\./.test(line)) {
                    continue;
                }
                
                // Extract the quote part (between quotes)
                const quoteMatch = line.match(/"([^"]+)"/);
                let quote = quoteMatch ? quoteMatch[1] : '';
                
                // Extract metadata (after the quote)
                let meta = '';
                if (quoteMatch) {
                    const afterQuote = line.slice(line.indexOf(quoteMatch[0]) + quoteMatch[0].length).trim();
                    // Check if the metadata starts with a dash
                    if (afterQuote.startsWith('-')) {
                        meta = afterQuote.substring(1).trim();
                    } else {
                        meta = afterQuote;
                    }
                } else {
                    // If no quotes were found, try to split by a dash
                    const parts = line.split(/\s+-\s+/);
                    if (parts.length > 1) {
                        // Remove the number from the first part
                        quote = parts[0].replace(/^\d+\.\s*/, '').trim();
                        meta = parts[1].trim();
                    } else {
                        // Just use the whole line without the number
                        quote = line.replace(/^\d+\.\s*/, '').trim();
                    }
                }
                
                // Add to references array
                references.push({ 
                    quote, 
                    meta: meta || 'No metadata available'
                });
            }
        } else {
            // If no "References:" section, take the whole answer
            mainAnswer = answer;
        }
        
        // Ensure we have valid references
        references = references.filter(ref => ref.quote);
        
        return {
            success: true,
            data: {
                answer: mainAnswer,
                references: references
            }
        };
    } catch (error) {
        log('Error searching emails:', 'error', error);
        setLoading('search', false);
        
        return {
            success: false,
            error: error.message
        };
    } finally {
        setLoading('search', false);
    }
}

/**
 * Generates an email reply suggestion
 * @param {object|string} emailContentOrData - The email content to reply to or an object containing email data
 * @param {object|string} senderOrOptions - The email sender's name/address or options object
 * @param {string} [subject] - The email subject
 * @param {Array} [threadHistory] - The email thread history
 * @param {Object} [options] - Options for the reply (tone, length, etc.)
 * @returns {Promise<object>} - Promise with reply text
 */
async function generateReplySuggestion(emailContentOrData, senderOrOptions, subject, threadHistory = [], options = {}) {
    // Handle case where a single object is passed (from Outlook add-in)
    let emailContent = emailContentOrData;
    
    // Check if first argument is an email data object
    if (typeof emailContentOrData === 'object' && emailContentOrData !== null) {
        log('Email data passed as object for reply generation, extracting fields', 'info');
        // Extract email fields from the object
        emailContent = emailContentOrData.body || '';
        subject = emailContentOrData.subject || '';
        threadHistory = emailContentOrData.threadHistory || [];
        
        // If sender is an object, it might be the options
        if (typeof senderOrOptions === 'object') {
            options = senderOrOptions; // The second parameter is the options object
            sender = emailContentOrData.sender ? `${emailContentOrData.sender.name} <${emailContentOrData.sender.email}>` : '';
        } else {
            sender = senderOrOptions;
        }
    } else {
        // Traditional parameter format where each parameter is passed separately
        sender = senderOrOptions;
    }
    
    // Log the parameters for debugging
    log('Reply generation parameters:', 'info', {
        emailContentPreview: typeof emailContent === 'string' ? emailContent.substring(0, 100) + '...' : 'Not a string',
        subject: subject || '(no subject)',
        sender: sender || '(no sender)',
        threadHistoryLength: threadHistory ? threadHistory.length : 0,
        options
    });
    
    // Set default options if not provided
    const tone = options.tone || 'professional';
    const length = options.length || 'medium';
    const language = options.language || 'English';
    const additionalContext = options.additionalContext || '';
    
    // Set loading state
    window.OpsieApi.setLoading('reply', true);
    
    try {
        // Generate cache key based on content hash and options to avoid redundant API calls
        // Include a simple hash of additional context to differentiate Q&A-enhanced replies
        const contextHash = additionalContext ? 
            additionalContext.length.toString() + '_' + additionalContext.substring(0, 50).replace(/\s+/g, '').substring(0, 20) : 
            'no_context';
        const cacheKey = `reply_${subject}_${sender}_${threadHistory.length}_${tone}_${length}_${language}_${contextHash}`;
        
        // Check if we have a cached result
        const cachedResult = localStorage.getItem(cacheKey);
        if (cachedResult) {
            const parsedResult = JSON.parse(cachedResult);
            log('Using cached reply suggestion', 'info', parsedResult);
            window.OpsieApi.setLoading('reply', false);
            return parsedResult;
        }
        
        // Get user name from storage for signature
        let userName = '';
        try {
            userName = localStorage.getItem('userName') || '';
        } catch (storageError) {
            log('Error retrieving user name from storage', 'warning', storageError);
        }
        
        // Get OpenAI API key from storage
        const apiKey = localStorage.getItem('openaiApiKey');
        
        // If no API key, return default reply
        if (!apiKey) {
            log('No OpenAI API key found', 'warning');
            window.OpsieApi.setLoading('reply', false);
            return {
                replyText: "Please add your OpenAI API key in settings to generate reply suggestions."
            };
        }
        
        // Validate required fields
        if (!emailContent || !subject) {
            log('Missing required fields for reply generation', 'error', {
                hasEmailContent: !!emailContent,
                hasSubject: !!subject
            });
            window.OpsieApi.setLoading('reply', false);
            return {
                replyText: "Error: Email is missing required content or subject. Please make sure you have an email selected with a valid subject and body."
            };
        }
        
        // Prepare thread history content
        let threadContent = "";
        if (threadHistory && threadHistory.length > 0) {
            threadContent = threadHistory.map(msg => 
                `From: ${msg.sender}\nTime: ${msg.time}\nContent: ${msg.content}`
            ).join('\n\n');
        }
        
        // Create system message based on tone, length, and language
        let systemMessage = `You are an AI assistant that helps draft professional email replies. 
        Write a reply to the email that is:
        - Tone: ${tone} (e.g., professional, friendly, formal, casual)
        - Length: ${length} (short: 2-3 sentences, medium: 4-6 sentences, long: 7+ sentences)
        - Language: ${language}
        
        DO NOT include any explanations or notes about your process. ONLY output the email reply text.
        DO NOT include the email headers like "To:", "Subject:" etc. ONLY the body of the reply.
        
        IMPORTANT: Include an appropriate closing and signature in the target language. Use the user's name "${userName}" in the signature.
        For example:
        - English: "Best regards, ${userName}"
        - German: "Mit freundlichen Grüßen, ${userName}" 
        - French: "Cordialement, ${userName}"
        - Swedish: "Med vänliga hälsningar, ${userName}"
        - Spanish: "Saludos cordiales, ${userName}"
        - Italian: "Cordiali saluti, ${userName}"
        - Portuguese: "Atenciosamente, ${userName}"
        
        If Questions & Answers are provided in the additional context, use them to:
        - Provide accurate, informed responses based on the knowledge base
        - Reference specific information when relevant to the email
        - Ensure consistency with verified answers
        - Address any questions or concerns raised in the email using the available knowledge
        
        If the user provided additional context, incorporate it naturally into your reply.`;
        
        // Add additional context instruction if provided
        if (additionalContext) {
            // Check if the additional context contains Q&A data
            const hasQAContext = additionalContext.includes('Relevant Questions & Answers');
            if (hasQAContext) {
                systemMessage += `\n\nIMPORTANT: The additional context below contains relevant Questions & Answers from the knowledge base. Use this information to provide accurate, well-informed responses. Pay special attention to verified answers and match confidence scores.`;
            }
            systemMessage += `\nAdditional context: ${additionalContext}`;
        }
        
        // Log system message
        log('System message for reply generation:', 'info', systemMessage);
        
        // Prepare the request to OpenAI API
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: "gpt-4o",
                messages: [
                    {
                        role: "system",
                        content: systemMessage
                    },
                    {
                        role: "user",
                        content: `Please write a reply to this email:
                        
                        Subject: ${subject}
                        From: ${sender}
                        
                        ${emailContent}
                        
                        ${threadHistory.length > 0 ? 'Previous thread history:\n' + threadContent : ''}
                        
                        Write only the reply text.`
                    }
                ],
                temperature: 0.7
            })
        });
        
        // Check if the response is ok
        if (!response.ok) {
            const errorText = await response.text();
            log('OpenAI API error response for reply generation:', 'error', {
                status: response.status,
                statusText: response.statusText,
                errorText: errorText
            });
            throw new Error(`OpenAI API error: ${response.status} - ${errorText}`);
        }
        
        // Parse the response
        const responseData = await response.json();
        log('OpenAI API raw response for reply:', 'info', responseData);
        
        // Check if we have the expected response structure
        if (!responseData.choices || !responseData.choices[0] || !responseData.choices[0].message) {
            log('Unexpected API response format for reply generation', 'error', responseData);
            throw new Error('Unexpected API response format');
        }
        
        // Extract the content from the response
        const replyText = responseData.choices[0].message.content.trim();
        log('AI reply content:', 'info', { previewLength: replyText.length, preview: replyText.substring(0, 100) + '...' });
        
        // Format the result (AI now handles the complete reply including signature)
        const result = {
            replyText: replyText
        };
        
        // Cache the result
        localStorage.setItem(cacheKey, JSON.stringify(result));
        
        window.OpsieApi.setLoading('reply', false);
        return result;
        
    } catch (error) {
        log('Error generating reply: ' + error.message, 'error', error);
        window.OpsieApi.setLoading('reply', false);
        
        // Return a default reply on error
        return {
            replyText: `Error generating reply: ${error.message}\n\nPlease check your API key and try again.`
        };
    }
}

/**
 * Saves an email to the backend
 * @param {object} emailData - The email data to save
 * @returns {Promise<object>} - Promise with save result
 */
async function saveEmail(emailData) {
    try {
        log('Saving email to database...', 'info', {
            subject: emailData.subject,
            sender: emailData.sender,
            hasBody: !!emailData.body,
            bodyLength: emailData.body ? emailData.body.length : 0,
            messageId: emailData.messageId || 'No message ID'
        });
        
        // Show loading state
        window.OpsieApi.setLoading('save', true);
        
        // Get API access token
        const token = await getAuthToken();
        if (!token) {
            log('No authentication token available for saving email', 'error');
            window.OpsieApi.setLoading('save', false);
            return {
                success: false,
                error: "Authentication required to save emails. Please log in."
            };
        }
        
        // Get team ID from local storage - check multiple possible keys
        let teamId = localStorage.getItem('currentTeamId');
        
        if (!teamId) {
            // Try alternative keys if primary key not found
            const alternateKeys = ['opsieTeamId', 'teamId', 'team_id'];
            
            for (const key of alternateKeys) {
                const altTeamId = localStorage.getItem(key);
                if (altTeamId) {
                    log(`Found team ID using alternate key: ${key}`, 'info');
                    teamId = altTeamId;
                    // Save it under the primary key for future use
                    localStorage.setItem('currentTeamId', teamId);
                    break;
                }
            }
        }
        
        if (!teamId) {
            log('No team ID found for saving email', 'error');
            window.OpsieApi.setLoading('save', false);
            return {
                success: false,
                error: "No team selected. Please select a team in settings."
            };
        }
        
        // Get user ID from token
        const tokenData = decodeJwtToken(token);
        if (!tokenData || !tokenData.sub) {
            log('Could not extract user ID from token', 'error');
            window.OpsieApi.setLoading('save', false);
            return {
                success: false,
                error: "Authentication error. Please log in again."
            };
        }
        
        const userId = tokenData.sub;
        log('Extracted user ID from token:', 'info', userId);
        
        // First, check if message already exists in database
        if (emailData.messageId) {
            try {
                log('Checking if email already exists in database...', 'info');
                
                const existingQuery = `message_external_id=eq.${encodeURIComponent(emailData.messageId)}&team_id=eq.${teamId}`;
                const checkResponse = await fetch(`${API_BASE_URL}/messages?${existingQuery}&select=id`, {
                    method: 'GET',
                    headers: {
                        'Content-Type': 'application/json',
                        'apikey': SUPABASE_ANON_KEY,
                        'Authorization': `Bearer ${token}`
                    }
                });
                
                if (checkResponse.ok) {
                    const existingMessages = await checkResponse.json();
                    if (existingMessages && existingMessages.length > 0) {
                        log('Email already exists in database', 'info', existingMessages[0]);
                        window.OpsieApi.setLoading('save', false);
                        return {
                            success: true,
                            message: "Email already exists in database",
                            data: {
                                message: existingMessages[0]
                            }
                        };
                    }
                    log('Email does not exist in database, proceeding with save', 'info');
                } else {
                    log('Error checking if email exists', 'warning');
                    // Continue with save attempt even if check fails
                }
            } catch (checkError) {
                log('Error checking if email exists', 'warning', checkError);
                // Continue with save attempt even if check fails
            }
        }
        
        // Prepare thread ID
        let threadId = null;
        // If we implemented thread creation logic, we would set threadId here
        
        // Prepare message data
        const messageData = {
            sender_name: emailData.sender.name,
            sender_email: emailData.sender.email,
            message_body: emailData.body,
            timestamp: emailData.timestamp,
            summary: emailData.summary,
            urgency: emailData.urgency,
            team_id: teamId,
            user_id: userId,
            message_external_id: emailData.messageId,
            thread_id: threadId
        };
        
        log('Prepared message data for saving:', 'info', messageData);
        
        // Send the save request
        const saveResponse = await fetch(`${API_BASE_URL}/messages`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${token}`,
                'Prefer': 'return=representation'
            },
            body: JSON.stringify(messageData)
        });
        
        if (!saveResponse.ok) {
            const errorText = await saveResponse.text();
            log('Error saving email:', 'error', {
                status: saveResponse.status,
                text: errorText
            });
            window.OpsieApi.setLoading('save', false);
            return {
                success: false,
                error: `Error saving email: ${saveResponse.status} - ${errorText}`
            };
        }
        
        // Parse the response to get the saved message data
        const savedMessage = await saveResponse.json();
        log('Email saved successfully:', 'info', savedMessage);
        
        window.OpsieApi.setLoading('save', false);
        return {
            success: true,
            message: "Email saved successfully",
            data: {
                message: savedMessage[0]  // Supabase returns an array with the inserted item
            }
        };
    } catch (error) {
        log('Exception saving email:', 'error', error);
        window.OpsieApi.setLoading('save', false);
        return {
            success: false,
            error: error.message || "Unknown error saving email"
        };
    }
}

/**
 * Login with email/password
 * @param {string} email - User email
 * @param {string} password - User password
 * @returns {Promise<object>} - Promise with login result
 */
async function login(email, password) {
    try {
        document.getElementById('login-loading').style.display = 'flex';
        
        // Use Supabase auth endpoint
        const response = await fetch(`https://vewnmfmnvumupdrcraay.supabase.co/auth/v1/token?grant_type=password`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY
            },
            body: JSON.stringify({ 
                email, 
                password 
            })
        });
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.message || `Login failed: ${response.status} ${response.statusText}`);
        }
        
        const result = await response.json();
        
        if (result.access_token) {
            localStorage.setItem(STORAGE_KEY_TOKEN, result.access_token);
            localStorage.setItem(STORAGE_KEY_REFRESH, result.refresh_token);
            
            const authErrorContainer = document.getElementById('auth-error-container');
            if (authErrorContainer) {
                authErrorContainer.style.display = 'none';
            }
        }
        
        document.getElementById('login-loading').style.display = 'none';
        return result;
    } catch (error) {
        document.getElementById('login-loading').style.display = 'none';
        showErrorNotification(error.message || 'Login failed');
        throw error;
    }
}

/**
 * Check if user is authenticated
 * @returns {Promise<boolean>} - Promise resolving to auth status
 */
async function isAuthenticated() {
    try {
        log('Checking authentication status', 'info');
        
        // First try to get token from our standard functions
        const token = await getAuthToken();
        
        if (token) {
            log('Authentication token found', 'info');
            
            // Basic validation - check if token is expired
            try {
                const decoded = decodeJwtToken(token);
                if (!decoded) {
                    log('Token decode failed', 'warning');
                    return false;
                }
                
                if (decoded.exp && decoded.exp * 1000 < Date.now()) {
                    log('Token is expired', 'warning');
                    return false;
                }
                
                log('Token is valid', 'info');
                return true;
            } catch (error) {
                log('Error validating token:', 'error', error);
                return false;
            }
        }
        
        log('No token found', 'warning');
        return false;
    } catch (error) {
        log('Error checking authentication:', 'error', error);
        return false;
    }
}

/**
 * Gets the OpenAI API key from storage
 * @returns {Promise<string>} - Promise resolving to API key
 */
async function getOpenAIApiKey() {
    try {
        return new Promise((resolve) => {
            // Get API key from local storage
            const apiKey = localStorage.getItem(STORAGE_KEY_OPENAI_API);
            console.log('API key retrieval - key exists:', apiKey ? 'Yes' : 'No');
            console.log('API key retrieval - key length:', apiKey ? apiKey.length : 0);
            
            // Check if the API key starts with "sk-" (OpenAI API keys format)
            const isValidFormat = apiKey && apiKey.startsWith('sk-');
            console.log('API key format valid:', isValidFormat ? 'Yes' : 'No');
            
            resolve(isValidFormat ? apiKey : null);
        });
    } catch (error) {
        console.error('Error getting OpenAI API key:', error);
        return null;
    }
}

/**
 * Initialize team and user information from localStorage or API
 * @param {Function} onTeamInfoReady - Optional callback function that will be called when team info is ready,
 *                                     receives teamInfo object with teamId and userId properties
 * @returns {Promise<Object|boolean>} - Object with teamId and userId if successful, or false
 */
async function initTeamAndUserInfo(onTeamInfoReady = null) {
    try {
        log('Initializing team and user information', 'info');
        
        // Check if we have userid in localStorage
        const existingUserId = localStorage.getItem('userId');
        
        // Check if team ID is already in localStorage
        const existingTeamId = localStorage.getItem('currentTeamId');
        
        if (existingTeamId && existingUserId) {
            log('Team ID and User ID already exist in localStorage:', 'info', {
                teamId: existingTeamId,
                userId: existingUserId
            });
            
            // If a callback was provided, execute it now since team info is already available
            if (typeof onTeamInfoReady === 'function') {
                try {
                    log('Executing onTeamInfoReady callback with cached team info', 'info');
                    onTeamInfoReady({
                        teamId: existingTeamId,
                        userId: existingUserId,
                        fromCache: true
                    });
                } catch (callbackError) {
                    log('Error in onTeamInfoReady callback', 'error', callbackError);
                }
            }
            
            return {
                teamId: existingTeamId,
                userId: existingUserId,
                fromCache: true
            };
        }
        
        // Try to get team ID from local storage with other possible keys
        const alternateTeamId = localStorage.getItem('opsieTeamId') || 
                              localStorage.getItem('teamId') || 
                              localStorage.getItem('team_id');
        
        if (alternateTeamId) {
            localStorage.setItem('currentTeamId', alternateTeamId);
            log('Found team ID using alternate key, saved with primary key', 'info', {
                teamId: alternateTeamId,
                source: 'alternate key'
            });
            
            // Still need to get user ID if we don't have it yet
            if (!existingUserId) {
                log('Team ID found but User ID missing, will fetch from API', 'info');
            } else {
                const result = {
                    teamId: alternateTeamId,
                    userId: existingUserId,
                    fromCache: true
                };
                
                // If a callback was provided, execute it now
                if (typeof onTeamInfoReady === 'function') {
                    try {
                        log('Executing onTeamInfoReady callback with alternate team info', 'info');
                        onTeamInfoReady(result);
                    } catch (callbackError) {
                        log('Error in onTeamInfoReady callback', 'error', callbackError);
                    }
                }
                
                return result;
            }
        }
        
        // If not found in localStorage, check if we have API access token
        const token = localStorage.getItem(STORAGE_KEY_TOKEN);
        if (!token) {
            log('No authentication token found, cannot initialize team ID', 'warning');
            return false;
        }
        
        // Get the user ID from the decoded JWT token
        const tokenData = decodeJwtToken(token);
        if (!tokenData) {
            log('Failed to decode auth token', 'error');
            return false;
        }
        
        // Extract user ID from token payload
        const userId = tokenData.sub || tokenData.user_id || tokenData.userId;
        if (!userId) {
            log('No user ID found in token data', 'error', tokenData);
            return false;
        }
        
        log('Extracted user ID from token:', 'info', userId);
        
        // Store the user ID in localStorage immediately
        localStorage.setItem('userId', userId);
        log('Saved user ID to localStorage', 'info', userId);
        
        // Make API call to get user's team information from the users table
        try {
            const response = await fetch(`https://vewnmfmnvumupdrcraay.supabase.co/rest/v1/users?id=eq.${userId}&select=team_id,role,first_name,last_name,email`, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'apikey': SUPABASE_ANON_KEY,
                    'Authorization': `Bearer ${token}`
                }
            });
            
            if (!response.ok) {
                log('API error when fetching user data:', 'error', {
                    status: response.status,
                    statusText: response.statusText
                });
                throw new Error(`API error: ${response.status}`);
            }
            
            const userData = await response.json();
            log('Retrieved user data:', 'info', userData);
            
            if (userData && userData.length > 0 && userData[0].team_id) {
                const teamId = userData[0].team_id;
                localStorage.setItem('currentTeamId', teamId);
                
                // Also save user's name if available for signatures
                if (userData[0].first_name || userData[0].last_name) {
                    const userName = `${userData[0].first_name || ''} ${userData[0].last_name || ''}`.trim();
                    localStorage.setItem('userName', userName);
                    log('Saved user name for signatures:', 'info', userName);
                }
                
                // If we have the team ID, let's also try to get the team name
                try {
                    const teamResponse = await fetch(`https://vewnmfmnvumupdrcraay.supabase.co/rest/v1/teams?id=eq.${teamId}&select=name,organization`, {
                        method: 'GET',
                        headers: {
                            'Content-Type': 'application/json',
                            'apikey': SUPABASE_ANON_KEY,
                            'Authorization': `Bearer ${token}`
                        }
                    });
                    
                    if (teamResponse.ok) {
                        const teamData = await teamResponse.json();
                        if (teamData && teamData.length > 0) {
                            // Save team name
                            localStorage.setItem('currentTeamName', teamData[0].name);
                            log('Saved team name:', 'info', teamData[0].name);
                        }
                    }
                } catch (teamError) {
                    log('Error fetching team details, but team ID was saved:', 'warning', teamError);
                    // Continue anyway since we have the team ID
                }
                
                log('User and team information initialized successfully', 'info', {
                    teamId: teamId,
                    userId: userId
                });
                
                const result = {
                    teamId: teamId,
                    userId: userId,
                    fromApi: true
                };
                
                // If a callback was provided, execute it with the result
                if (typeof onTeamInfoReady === 'function') {
                    try {
                        log('Executing onTeamInfoReady callback with API team info', 'info');
                        onTeamInfoReady(result);
                    } catch (callbackError) {
                        log('Error in onTeamInfoReady callback', 'error', callbackError);
                    }
                }
                
                return result;
            }
            
            log('User has no team assigned, but userId is saved', 'warning');
            
            const result = {
                userId: userId,
                fromApi: true
            };
            
            // Still execute callback even without team ID
            if (typeof onTeamInfoReady === 'function') {
                try {
                    log('Executing onTeamInfoReady callback with user info only (no team)', 'info');
                    onTeamInfoReady(result);
                } catch (callbackError) {
                    log('Error in onTeamInfoReady callback', 'error', callbackError);
                }
            }
            
            return result;
        } catch (apiError) {
            log('Error retrieving team information from API:', 'error', apiError);
            // If we at least have the userId, return it
            if (userId) {
                const result = {
                    userId: userId,
                    fromToken: true
                };
                
                // Execute callback with partial info
                if (typeof onTeamInfoReady === 'function') {
                    try {
                        log('Executing onTeamInfoReady callback with token info only', 'info');
                        onTeamInfoReady(result);
                    } catch (callbackError) {
                        log('Error in onTeamInfoReady callback', 'error', callbackError);
                    }
                }
                
                return result;
            }
            return false;
        }
    } catch (error) {
        log('Error initializing team and user information:', 'error', error);
        return false;
    }
}

// Add a utility logging function
function log(message, level = 'info', data = null) {
    try {
        const timestamp = new Date().toISOString();
        const logEntry = {
            timestamp,
            level,
            message,
            data
        };
        
        // Log to console with appropriate level
        switch (level) {
            case 'error':
                console.error(`[${timestamp}] ${message}`, data || '');
                break;
            case 'warning':
                console.warn(`[${timestamp}] ${message}`, data || '');
                break;
            case 'info':
            default:
                console.log(`[${timestamp}] ${message}`, data || '');
                break;
        }
        
        // Add to debug display if available
        if (typeof window.addDebugLogEntry === 'function') {
            window.addDebugLogEntry(logEntry);
        }
        
        return logEntry;
    } catch (error) {
        console.error('Error in logging function:', error);
        return null;
    }
}

// Add a loading state management function
function setLoading(section, isLoading) {
    try {
        log(`Setting loading state for ${section}: ${isLoading}`, 'info');
        
        // Handle loading states for different components
        switch (section) {
            case 'qa':
                const qaSpinner = document.getElementById('qa-loading-spinner');
                if (qaSpinner) {
                    qaSpinner.style.display = isLoading ? 'block' : 'none';
                }
                
                // If loading is complete, ensure the appropriate container is shown
                // based on whether questions were found
                if (!isLoading) {
                    const qaList = document.getElementById('qa-list');
                    const noQuestionsMessage = document.getElementById('no-questions-message');
                    
                    if (qaList && qaList.children.length > 0) {
                        document.getElementById('qa-container').style.display = 'block';
                        if (noQuestionsMessage) noQuestionsMessage.style.display = 'none';
                    } else {
                        if (noQuestionsMessage) noQuestionsMessage.style.display = 'block';
                    }
                }
                break;
                
            case 'summary':
                const summarySpinner = document.getElementById('summary-loading-spinner');
                if (summarySpinner) {
                    summarySpinner.style.display = isLoading ? 'block' : 'none';
                }
                break;
                
            // Add more components as needed
            
            default:
                log(`Unknown component for loading state: ${section}`, 'warning');
        }
    } catch (error) {
        log(`Error setting loading state for ${section}`, 'error', error);
    }
}

/**
 * Extracts content from Markdown code blocks
 * @param {string} markdownText - Text that might contain Markdown code blocks
 * @returns {string} - The extracted content without the code block markers
 */
function extractContentFromCodeBlock(markdownText) {
    if (!markdownText) return '';
    
    // Check if the text contains a code block (```...```)
    // Updated regex to better handle JSON code blocks with language identifiers
    const codeBlockRegex = /```(?:json|javascript|js)?(?:\s*)([\s\S]*?)```/;
    const match = markdownText.match(codeBlockRegex);
    
    if (match && match[1]) {
        return match[1].trim();
    }
    
    // If no code block, return the original text
    return markdownText;
}

/**
 * Internal utility to get the internal message ID from either a UUID or external ID
 * @param {string} messageId - The message ID (can be UUID or external ID)
 * @returns {Promise<object>} - Object with success flag and either id or error
 */
async function getInternalMessageId(messageId) {
    try {
        if (!messageId) {
            return { success: false, error: 'Message ID is required' };
        }
        
        // Check if the messageId is already a UUID
        const isUuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(messageId);
        
        // If it's already a UUID, return it directly
        if (isUuid) {
            return { success: true, id: messageId };
        }
        
        log('MessageId is not a UUID, looking up internal ID', 'info', { externalId: messageId });
        
        // Get the team ID from local storage - needed for the lookup query
        const teamId = localStorage.getItem('currentTeamId');
        if (!teamId) {
            return { success: false, error: 'Team ID is required for message lookup' };
        }
        
        // Query the messages table to find the internal ID
        const messagesResponse = await apiRequest(
            `messages?message_external_id=eq.${encodeURIComponent(messageId)}&team_id=eq.${teamId}&select=id`,
            'GET'
        );
        
        if (!messagesResponse.success || !messagesResponse.data || messagesResponse.data.length === 0) {
            log('Failed to find message with external ID', 'error', { 
                externalId: messageId,
                response: messagesResponse 
            });
            return {
                success: false,
                error: 'Message not found with the provided external ID'
            };
        }
        
        // Get the internal UUID from the response
        const internalMessageId = messagesResponse.data[0].id;
        log('Found internal message ID', 'info', { 
            externalId: messageId,
            internalId: internalMessageId 
        });
        
        return { success: true, id: internalMessageId };
    } catch (error) {
        log('Error resolving internal message ID', 'error', error);
        return { success: false, error: error.message };
    }
}

/**
 * Add a note to a message
 * @param {string} messageId - The ID of the message (can be internal UUID or external ID)
 * @param {string} userId - The ID of the user adding the note
 * @param {string} noteBody - The text content of the note
 * @param {string} category - The category of the note (optional)
 * @returns {Promise<object>} - The result of the operation
 */
async function addNoteToMessage(messageId, userId, noteBody, category = null) {
    try {
        log('Adding note to message', 'info', { messageId, userId, category });
        
        if (!messageId) {
            return { success: false, error: 'Message ID is required' };
        }
        
        if (!userId) {
            return { success: false, error: 'User ID is required' };
        }
        
        if (!noteBody || !noteBody.trim()) {
            return { success: false, error: 'Note content is required' };
        }
        
        // Make sure we have an auth token
        const token = await getAuthToken();
        if (!token) {
            return { success: false, error: 'Authentication required' };
        }
        
        // Get current user info to get team_id
        const userInfo = await getUserInfo(true); // Force refresh to get latest team info
        if (!userInfo || !userInfo.team_id) {
            return { success: false, error: 'User must be part of a team to add notes' };
        }
        
        // Get the internal message ID (resolve from external ID if needed)
        const messageIdResult = await getInternalMessageId(messageId);
        if (!messageIdResult.success) {
            return { success: false, error: messageIdResult.error };
        }
        
        // Prepare the note data with the internal message ID and team_id
        const noteData = {
            message_id: messageIdResult.id,
            user_id: userId,
            note_body: noteBody,
            category: category || 'Other',
            team_id: userInfo.team_id,
            created_at: new Date().toISOString()
        };
        
        // Add the note to the database
        const response = await apiRequest('notes', 'POST', noteData);
        
        if (!response.success) {
            log('Failed to add note', 'error', response.error);
            return {
                success: false,
                error: response.error || 'Failed to add note'
            };
        }
        
        // The API might not return data for POST operations (common with REST APIs)
        // As long as we get a success response, consider it successful
        log('Note added successfully', 'info', response.data || 'No data returned from API');
        
        // If we have data, return it. Otherwise, just return a success with the noteData we sent
        return {
            success: true,
            note: response.data ? (response.data[0] || response.data) : {
                ...noteData,
                id: null // We don't have an ID since the API didn't return one
            }
        };
    } catch (error) {
        log('Exception adding note', 'error', error);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * Get notes for a message
 * @param {string} messageId - The ID of the message (can be internal UUID or external ID)
 * @returns {Promise<object>} - The result of the operation
 */
async function getMessageNotes(messageId) {
    try {
        log('Getting notes for message', 'info', { messageId });
        
        if (!messageId) {
            return { success: false, error: 'Message ID is required' };
        }
        
        // Make sure we have an auth token
        const token = await getAuthToken();
        if (!token) {
            return { success: false, error: 'Authentication required' };
        }
        
        // Get current user info to get team_id
        const userInfo = await getUserInfo(true); // Force refresh to get latest team info
        if (!userInfo || !userInfo.team_id) {
            return { success: false, error: 'User must be part of a team to view notes' };
        }
        
        // Get the internal message ID (resolve from external ID if needed)
        const messageIdResult = await getInternalMessageId(messageId);
        if (!messageIdResult.success) {
            return { success: false, error: messageIdResult.error };
        }
        
        // Now use the internal message ID to fetch notes, filtered by team
        const response = await apiRequest(`notes?message_id=eq.${messageIdResult.id}&team_id=eq.${userInfo.team_id}&order=created_at.desc`, 'GET');
        
        if (!response.success) {
            log('Failed to get notes', 'error', response.error);
            return {
                success: false,
                error: response.error || 'Failed to get notes'
            };
        }
        
        const notes = response.data || [];
        log(`Found ${notes.length} notes for message`, 'info');
        
        // If we have notes, fetch user details for each note
        if (notes.length > 0) {
            // Get unique user IDs from notes
            const userIds = [...new Set(notes.map(note => note.user_id))];
            
            if (userIds.length > 0) {
                // Fetch user details using RPC function to bypass RLS
                const usersResponse = await apiRequest(`rpc/get_user_details`, 'POST', {
                    user_ids: userIds
                });
                
                if (usersResponse.success && usersResponse.data) {
                    // Create a map of user IDs to user details
                    const userMap = {};
                    usersResponse.data.forEach(user => {
                        userMap[user.id] = user;
                    });
                    
                    // Add user details to each note
                    notes.forEach(note => {
                        if (userMap[note.user_id]) {
                            const user = userMap[note.user_id];
                            note.user = {
                                name: `${user.first_name || ''} ${user.last_name || ''}`.trim() || 'Unknown User',
                                email: user.email || ''
                            };
                        } else {
                            note.user = { name: 'Unknown User', email: '' };
                        }
                    });
                }
            }
        }
        
        return {
            success: true,
            notes: notes
        };
    } catch (error) {
        log('Exception getting notes', 'error', error);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * Directly fetches user information from the API
 * @returns {Promise<Object|null>} - User data or null if not found
 */
async function getUserInfo(forceRefresh = false) {
    try {
        log('Fetching user information from API', 'info', { forceRefresh });
        
        // Get the authentication token
        const token = localStorage.getItem(STORAGE_KEY_TOKEN);
        if (!token) {
            log('No authentication token found, cannot fetch user info', 'warning');
            return null;
        }
        
        // Check for cached user data if not forcing refresh
        if (!forceRefresh) {
            const cachedUserData = localStorage.getItem('cachedUserData');
            if (cachedUserData) {
                try {
                    const parsedData = JSON.parse(cachedUserData);
                    const cacheTime = parsedData.timestamp || 0;
                    const now = Date.now();
                    
                    // Cache valid for 5 minutes (300000 ms)
                    if (now - cacheTime < 300000) {
                        log('Using cached user data', 'info', parsedData.data);
                        return parsedData.data;
                    } else {
                        log('Cached user data expired, fetching fresh data', 'info');
                    }
                } catch (cacheError) {
                    log('Error parsing cached user data', 'error', cacheError);
                }
            }
        } else {
            log('Forced refresh requested, bypassing cache', 'info');
        }
        
        // Get the user ID from the decoded JWT token
        const tokenData = decodeJwtToken(token);
        if (!tokenData) {
            log('Failed to decode auth token', 'error');
            return null;
        }
        
        // Extract user ID from token payload
        const userId = tokenData.sub || tokenData.user_id || tokenData.userId;
        if (!userId) {
            log('No user ID found in token data', 'error', tokenData);
            return null;
        }
        
        log('Using user ID from token:', 'info', userId);
        
        // Make API call to get user information from the users table
        const response = await fetch(`https://vewnmfmnvumupdrcraay.supabase.co/rest/v1/users?id=eq.${userId}&select=id,email,team_id,role,first_name,last_name,created_at`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (!response.ok) {
            log('API error when fetching user data:', 'error', {
                status: response.status,
                statusText: response.statusText
            });
            throw new Error(`API error: ${response.status}`);
        }
        
        const userData = await response.json();
        log('Retrieved user data:', 'info', userData);
        
        if (userData && userData.length > 0) {
            // Save useful user information to localStorage
            const user = userData[0];
            
            // Store user ID if we don't have it yet
            if (!localStorage.getItem('userId') || forceRefresh) {
                localStorage.setItem('userId', user.id);
                log('Saved user ID to localStorage', 'info', user.id);
            }
            
            // Store team ID if available
            if (user.team_id && (!localStorage.getItem('currentTeamId') || forceRefresh)) {
                localStorage.setItem('currentTeamId', user.team_id);
                log('Saved team ID to localStorage', 'info', user.team_id);
            }
            
            // Store user name if available
            if ((user.first_name || user.last_name) && (!localStorage.getItem('userName') || forceRefresh)) {
                const userName = `${user.first_name || ''} ${user.last_name || ''}`.trim();
                localStorage.setItem('userName', userName);
                log('Saved user name to localStorage', 'info', userName);
            }
            
            // Store user email
            if (user.email && (!localStorage.getItem('userEmail') || forceRefresh)) {
                localStorage.setItem('userEmail', user.email);
                log('Saved user email to localStorage', 'info', user.email);
            }
            
            // Cache user data for future quick access
            localStorage.setItem('cachedUserData', JSON.stringify({
                timestamp: Date.now(),
                data: user
            }));
            
            return user;
        }
        
        log('No user data found for the provided token', 'warning');
        return null;
    } catch (error) {
        log('Error fetching user information:', 'error', error);
        return null;
    }
}

// Initialize the API service
(function() {
    if (typeof window !== 'undefined') {
        // Initialize the object if it doesn't exist
        if (!window.OpsieApi) {
            window.OpsieApi = {};
        }

        // Add configuration
        window.OpsieApi.API_BASE_URL = API_BASE_URL;
        window.OpsieApi.SUPABASE_KEY = SUPABASE_ANON_KEY;
        window.OpsieApi.STORAGE_KEY_TOKEN = STORAGE_KEY_TOKEN;
        window.OpsieApi.STORAGE_KEY_REFRESH = STORAGE_KEY_REFRESH;
        
        // Add main API functions to the global object
        window.OpsieApi.apiRequest = apiRequest;
  window.OpsieApi.generateEmailSummary = generateEmailSummary;
  window.OpsieApi.generateReplySuggestion = generateReplySuggestion;
  window.OpsieApi.saveEmail = saveEmail;
  window.OpsieApi.login = login;
  window.OpsieApi.logout = logout;
  window.OpsieApi.isAuthenticated = isAuthenticated;
        window.OpsieApi.getUserInfo = getUserInfo;
        window.OpsieApi.getUserId = getUserId;
        window.OpsieApi.getTeamDetails = getTeamDetails;
        window.OpsieApi.getTeamMembers = getTeamMembers;
        window.OpsieApi.generateContactHistory = generateContactHistory;
        window.OpsieApi.searchEmailHistory = searchEmailHistory;
        window.OpsieApi.fetchContactHistory = fetchContactHistory;
  window.OpsieApi.getOpenAIApiKey = getOpenAIApiKey;
        window.OpsieApi.updateEmailSummary = updateEmailSummary;
        window.OpsieApi.initTeamAndUserInfo = initTeamAndUserInfo;
  window.OpsieApi.showNotification = showNotification;
        
        // Add authentication functions
        window.OpsieApi.signUp = signUp;
        window.OpsieApi.requestPasswordReset = requestPasswordReset;
        
        // Add team management functions
        window.OpsieApi.requestToJoinTeam = requestToJoinTeam;
        window.OpsieApi.createTeam = createTeam;
        window.OpsieApi.checkJoinRequestStatus = checkJoinRequestStatus;
        window.OpsieApi.getUserPendingJoinRequest = getUserPendingJoinRequest;
        window.OpsieApi.getTeamJoinRequests = getTeamJoinRequests;
window.OpsieApi.respondToJoinRequest = respondToJoinRequest;
window.OpsieApi.removeTeamMember = removeTeamMember;
window.OpsieApi.leaveTeam = leaveTeam;
window.OpsieApi.transferAdminRole = transferAdminRole;
window.OpsieApi.deleteTeam = deleteTeam;
        
        // Add utilities
        window.OpsieApi.log = log;
        window.OpsieApi.setLoading = setLoading;
        window.OpsieApi.getInternalMessageId = getInternalMessageId;
  window.OpsieApi.addNoteToMessage = addNoteToMessage;
  window.OpsieApi.getMessageNotes = getMessageNotes;
  window.OpsieApi.clearAllCaches = clearAllCaches;
    }
})();

/**
 * Get the user's ID from storage or token
 * @returns {Promise<string|null>} The user ID or null if not found
 */
async function getUserId() {
    try {
        // First try to get the ID from local storage
        const storedUserId = localStorage.getItem('userId');
        if (storedUserId) {
            return storedUserId;
        }
        
        // If not in storage, try to get it from the token
        const token = localStorage.getItem(STORAGE_KEY_TOKEN);
        if (!token) {
            return null;
        }
        
        const tokenData = decodeJwtToken(token);
        if (!tokenData) {
            return null;
        }
        
        // Extract user ID from token payload
        return tokenData.sub || tokenData.user_id || tokenData.userId || null;
    } catch (error) {
        log('Error getting user ID:', 'error', error);
        return null;
    }
}

/**
 * Get team details from the API
 * @param {string} teamId - The team ID to fetch details for
 * @param {boolean} forceRefresh - Whether to force a refresh of data from API
 * @returns {Promise<Object>} Result object with team details
 */
async function getTeamDetails(teamId, forceRefresh = false) {
    try {
        log('Fetching team details from API', 'info', { teamId, forceRefresh });
        
        if (!teamId) {
            log('No team ID provided, cannot fetch team details', 'warning');
            return { success: false, error: 'No team ID provided' };
        }
        
        // Check for cached team data if not forcing refresh
        if (!forceRefresh) {
            const cachedTeamData = localStorage.getItem(`cachedTeamData_${teamId}`);
            if (cachedTeamData) {
                try {
                    const parsedData = JSON.parse(cachedTeamData);
                    const cacheTime = parsedData.timestamp || 0;
                    const now = Date.now();
                    
                    // Cache valid for 5 minutes (300000 ms)
                    if (now - cacheTime < 300000) {
                        log('Using cached team data', 'info', parsedData.data);
                        return {
                            success: true,
                            data: parsedData.data
                        };
                    } else {
                        log('Cached team data expired, fetching fresh data', 'info');
                    }
                } catch (cacheError) {
                    log('Error parsing cached team data', 'error', cacheError);
                }
            }
        } else {
            log('Forced refresh requested, bypassing cache', 'info');
        }
        
        // Get the authentication token
        const token = localStorage.getItem(STORAGE_KEY_TOKEN);
        if (!token) {
            log('No authentication token found, cannot fetch team details', 'warning');
            return { success: false, error: 'Authentication required' };
        }
        
        // Make API call to get team details
        const response = await fetch(`https://vewnmfmnvumupdrcraay.supabase.co/rest/v1/teams?id=eq.${teamId}&select=*`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (!response.ok) {
            log('API error when fetching team details:', 'error', {
                status: response.status,
                statusText: response.statusText
            });
            return {
                success: false,
                error: `API error: ${response.status}`
            };
        }
        
        const teamData = await response.json();
        log('Retrieved team details:', 'info', teamData);
        
        if (teamData && teamData.length > 0) {
            // Save team name to localStorage for future use
            if (teamData[0].name) {
                localStorage.setItem('currentTeamName', teamData[0].name);
            }
            
            // Cache team data for future quick access
            localStorage.setItem(`cachedTeamData_${teamId}`, JSON.stringify({
                timestamp: Date.now(),
                data: teamData[0]
            }));
            
            return {
                success: true,
                data: teamData[0]
            };
        }
        
        log('No team data found for the provided team ID', 'warning');
        return {
            success: false,
            error: 'Team not found'
        };
    } catch (error) {
        log('Error fetching team details:', 'error', error);
        return {
            success: false,
            error: error.message || 'Unknown error'
        };
    }
}

/**
 * Get team members from the API
 * @param {string} teamId - The team ID to fetch members for
 * @param {boolean} forceRefresh - Whether to force a refresh of data from API
 * @returns {Promise<Object>} Result object with team members
 */
async function getTeamMembers(teamId, forceRefresh = false) {
    try {
        log('Fetching team members from API', 'info', { teamId, forceRefresh });
        
        if (!teamId) {
            log('No team ID provided, cannot fetch team members', 'warning');
            return { success: false, error: 'No team ID provided' };
        }
        
        // Check for cached team members data if not forcing refresh
        if (!forceRefresh) {
            const cachedMembersData = localStorage.getItem(`cachedTeamMembers_${teamId}`);
            if (cachedMembersData) {
                try {
                    const parsedData = JSON.parse(cachedMembersData);
                    const cacheTime = parsedData.timestamp || 0;
                    const now = Date.now();
                    
                    // Cache valid for 5 minutes (300000 ms)
                    if (now - cacheTime < 300000) {
                        log('Using cached team members data', 'info', parsedData.data);
                        return {
                            success: true,
                            data: parsedData.data
                        };
                    } else {
                        log('Cached team members data expired, fetching fresh data', 'info');
                    }
                } catch (cacheError) {
                    log('Error parsing cached team members data', 'error', cacheError);
                }
            }
        } else {
            log('Forced refresh requested, bypassing cache', 'info');
        }
        
        // Get the authentication token
        const token = localStorage.getItem(STORAGE_KEY_TOKEN);
        if (!token) {
            log('No authentication token found, cannot fetch team members', 'warning');
            return { success: false, error: 'Authentication required' };
        }
        
        // Make API call to get team members using RPC function to bypass RLS
        const response = await fetch(`https://vewnmfmnvumupdrcraay.supabase.co/rest/v1/rpc/get_team_members`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify({
                team_uuid: teamId
            })
        });
        
        if (!response.ok) {
            log('API error when fetching team members:', 'error', {
                status: response.status,
                statusText: response.statusText
            });
            return {
                success: false,
                error: `API error: ${response.status}`
            };
        }
        
        const memberData = await response.json();
        log('Retrieved team members:', 'info', memberData);
        
        // Cache team members data for future quick access
        localStorage.setItem(`cachedTeamMembers_${teamId}`, JSON.stringify({
            timestamp: Date.now(),
            data: memberData || []
        }));
        
        return {
            success: true,
            data: memberData || []
        };
    } catch (error) {
        log('Error fetching team members:', 'error', error);
        return {
            success: false,
            error: error.message || 'Unknown error'
        };
    }
}

/**
 * Clears all cached data from localStorage
 * @returns {number} Number of cache items cleared
 */
function clearAllCaches() {
    try {
        log('Clearing all cached data', 'info');
        
        // Get all keys from localStorage
        const allKeys = [];
        for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (key) {
                allKeys.push(key);
            }
        }
        
        // Identify cache keys based on naming patterns
        const cacheKeys = allKeys.filter(key => 
            key.startsWith('cached') || 
            key.includes('_cache_') || 
            key.includes('Cache') ||
            key.includes('summary_') ||
            key.includes('contactHistory_') ||
            key.includes('reply_')
        );
        
        // Remove all cache items
        let count = 0;
        cacheKeys.forEach(key => {
            localStorage.removeItem(key);
            count++;
        });
        
        // Clear in-memory caches if they exist
        if (typeof summaryCache !== 'undefined' && summaryCache instanceof Map) {
            summaryCache.clear();
        }
        
        if (typeof contactCache !== 'undefined' && contactCache instanceof Map) {
            contactCache.clear();
        }
        
        if (typeof replyCache !== 'undefined' && replyCache instanceof Map) {
            replyCache.clear();
        }
        
        log(`Cleared ${count} cached items`, 'info');
        return count;
    } catch (error) {
        log('Error clearing caches', 'error', error);
        return 0;
    }
}

/**
 * Log out the current user by clearing auth tokens and data
 * @returns {Promise<Object>} Result object indicating success or failure
 */
async function logout() {
    try {
        log('Logging out user', 'info');
        
        // Clear all caches first
        const clearedCount = clearAllCaches();
        log(`Cleared ${clearedCount} cached items during logout`, 'info');
        
        // Clear authentication token and related data
        localStorage.removeItem(STORAGE_KEY_TOKEN);
        localStorage.removeItem(STORAGE_KEY_REFRESH);
        localStorage.removeItem('refreshToken');
        localStorage.removeItem('userId');
        localStorage.removeItem('currentTeamId');
        localStorage.removeItem('currentTeamName');
        localStorage.removeItem('userName');
        localStorage.removeItem('userEmail');
        
        // Also try calling window.clearCaches if it exists (for backward compatibility)
        if (typeof window.clearCaches === 'function') {
            try {
                window.clearCaches();
            } catch (clearError) {
                log('Error calling window.clearCaches', 'warning', clearError);
            }
        }
        
        log('User logged out successfully', 'info');
        
        return {
            success: true,
            message: 'Logged out successfully'
        };
    } catch (error) {
        log('Error during logout', 'error', error);
        
        return {
            success: false,
            error: error.message || 'An error occurred during logout'
        };
    }
}

/**
 * Updates an existing email record with a summary and urgency score
 * @param {string} messageId - The ID of the message in the database
 * @param {string} summary - The summary text
 * @param {number} urgency - The urgency score
 * @returns {Promise<object>} - Promise with update result
 */
async function updateEmailSummary(messageId, summary, urgency) {
    try {
        log('Updating email summary in database...', 'info', {
            messageId,
            summary: summary ? summary.substring(0, 50) + '...' : 'No summary provided',
            urgency
        });
        
        // Show loading state
        window.OpsieApi.setLoading('summary', true);
        
        // Check if summary contains API key missing message
        if (summary && summary.includes("Please add your OpenAI API key in settings")) {
            log('Missing API key detected in summary', 'warning');
            window.OpsieApi.setLoading('summary', false);
            return {
                success: false,
                error: "OpenAI API key is required. Please add your API key in the settings panel to generate summaries."
            };
        }
        
        // Validate urgency score to avoid database constraints
        if (urgency === 0 || urgency === undefined || urgency === null) {
            log('Invalid urgency score detected', 'warning', { urgency });
            // Use default urgency to avoid database constraint errors
            urgency = 5;
            log('Using default urgency score instead', 'info', { urgency });
        }
        
        // Get API access token
        const token = await getAuthToken();
        if (!token) {
            log('No authentication token available for updating email', 'error');
            window.OpsieApi.setLoading('summary', false);
            return {
                success: false,
                error: "Authentication required to update emails. Please log in."
            };
        }
        
        // Get team ID from local storage
        const teamId = localStorage.getItem('currentTeamId');
        if (!teamId) {
            log('No team ID found for updating email', 'error');
            window.OpsieApi.setLoading('summary', false);
            return {
                success: false,
                error: "No team selected. Please select a team in settings."
            };
        }
        
        // Prepare update data
        const updateData = {
            summary,
            urgency
        };
        
        log('Prepared update data:', 'info', updateData);
        
        // Send the update request
        const updateResponse = await fetch(`${API_BASE_URL}/messages?id=eq.${messageId}`, {
            method: 'PATCH',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${token}`,
                'Prefer': 'return=representation'
            },
            body: JSON.stringify(updateData)
        });
        
        if (!updateResponse.ok) {
            const errorText = await updateResponse.text();
            log('Error updating email summary:', 'error', {
                status: updateResponse.status,
                text: errorText
            });
            
            // Try to determine if this is a database constraint error
            let userFriendlyError = `Error updating email summary: ${updateResponse.status}`;
            if (errorText.includes("messages_urgency_check")) {
                userFriendlyError = "Error updating summary: Urgency value is invalid. Please try again with a valid urgency score (1-10).";
            } else if (errorText.includes("violates check constraint")) {
                userFriendlyError = "Error updating summary: Database validation error. Please check your input values.";
            }
            
            window.OpsieApi.setLoading('summary', false);
            return {
                success: false,
                error: userFriendlyError
            };
        }
        
        // Parse the response to get the updated message data
        const updatedMessage = await updateResponse.json();
        log('Email summary updated successfully:', 'info', updatedMessage);
        
        window.OpsieApi.setLoading('summary', false);
        return {
            success: true,
            message: "Email summary updated successfully",
            data: {
                message: updatedMessage[0]  // Supabase returns an array with the updated item
            }
        };
    } catch (error) {
        log('Exception updating email summary:', 'error', error);
        window.OpsieApi.setLoading('summary', false);
        
        // Provide a more user-friendly error message
        let userFriendlyError = error.message || "Unknown error updating email summary";
        if (userFriendlyError.includes("violates check constraint")) {
            userFriendlyError = "Database validation error. Please check your summary and urgency values.";
        }
        
        return {
            success: false,
            error: userFriendlyError
        };
    }
}

/**
 * Extract questions from email content and try to find answers
 * @param {Object} emailData - The email data object
 * @returns {Promise<Object>} - Promise resolving to extracted questions and potential answers
 */
async function extractQuestionsAndAnswers(emailData) {
    try {
        log('Extracting questions from email', 'info', {
            subject: emailData.subject,
            sender: emailData.sender,
            bodyLength: emailData.body ? emailData.body.length : 0,
            hasBody: !!emailData.body,
            emailDataKeys: Object.keys(emailData)
        });
        
        // Set loading state
        window.OpsieApi.setLoading('qa', true);
        
        // Validate input
        if (!emailData || !emailData.body) {
            log('Missing email content for question extraction', 'error');
            window.OpsieApi.setLoading('qa', false);
            return {
                success: false,
                error: 'Email content is missing or incomplete'
            };
        }
        
        // Get OpenAI API key
        const apiKey = localStorage.getItem('openaiApiKey');
        if (!apiKey) {
            log('No OpenAI API key found for question extraction', 'warning');
            window.OpsieApi.setLoading('qa', false);
            return {
                success: false,
                error: 'OpenAI API key is required. Please add your API key in the settings panel.'
            };
        }
        
        // Extract questions from the email content
        const questionsResult = await extractQuestionsFromEmail(emailData, apiKey);
        if (!questionsResult.success) {
            log('Failed to extract questions', 'error', questionsResult.error);
            window.OpsieApi.setLoading('qa', false);
            return questionsResult;
        }
        
        const questions = questionsResult.questions;
        log(`Extracted ${questions.length} questions from email`, 'info', questions);
        
        // If no questions were found, return early
        if (questions.length === 0) {
            window.OpsieApi.setLoading('qa', false);
            return {
                success: true,
                questions: [],
                message: 'No questions were found in this email.'
            };
        }
        
        // Try to find answers for each question
        const teamId = localStorage.getItem('currentTeamId');
        
        // Process questions sequentially to avoid overwhelming the API
        const answeredQuestions = [];
        for (const question of questions) {
            try {
                log(`Searching for answer to question: ${question.text}`, 'info');
                
                // Log question details before checking for existing answers
                log('Processing question with details:', 'info', {
                    text: question.text,
                    type: question.type || 'unknown',
                    confidence: question.confidence || 0,
                    keywords: question.keywords || []
                });
                
                // First check if this question (or similar) exists in the qanda table
                log('Checking qanda table for existing answer...', 'info', {
                    questionText: question.text,
                    teamId,
                    timestamp: new Date().toISOString()
                });
                
                const existingAnswer = await findExistingAnswer(question.text, teamId);
                
                // Let's add this right before the section where we search for existing answers in extractQuestionsAndAnswers
                // The exact location is where we check if we've found an answer from the database
                log('Checking for existing answers in qanda table', 'info', {
                    questionCount: questions.length,
                    sampleQuestion: questions.length > 0 ? questions[0].text : 'No questions'
                });
                
                if (existingAnswer.found) {
                    log('Found existing answer in qanda table, processing answer object', 'info', {
                        questionText: question.text,
                        answerId: existingAnswer.id,
                        answerLength: existingAnswer.answer ? existingAnswer.answer.length : 0,
                        isVerified: existingAnswer.verified,
                        matchType: existingAnswer.matchType,
                        updatedAt: existingAnswer.updatedAt,
                        answerFields: Object.keys(existingAnswer)
                    });
                    
                    // Add to the section where a question is populated with an existing answer
                    question.answer = existingAnswer.answer;
                    question.source = 'database';
                    question.verified = existingAnswer.verified;
                    question.answerId = existingAnswer.id;
                    question.matchType = existingAnswer.matchType;

                    // If it's a semantic or fuzzy match, include the original question and similarity score
                    if (existingAnswer.matchType === 'semantic' || existingAnswer.matchType === 'fuzzy') {
                        question.originalQuestion = existingAnswer.originalQuestion;
                        question.similarityScore = existingAnswer.similarityScore;
                    }

                    // Add the updatedAt field from the database
                    question.updatedAt = existingAnswer.updatedAt;

                    // Log the complete question object after populating
                    log('Populated question with existing answer data', 'info', {
                        questionText: question.text,
                        hasUpdatedAt: !!question.updatedAt,
                        updatedAtValue: question.updatedAt,
                        questionKeys: Object.keys(question)
                    });
                    
                    question.references = existingAnswer.references || [];
                    
                    // Store translation info if available
                    if (question.original_text && question.original_language && question.original_language !== 'en') {
                        question.translatedFrom = {
                            originalText: question.original_text,
                            language: question.original_language
                        };
                    }
                } else {
                    log('No existing answer found in qanda table, searching emails instead', 'info', {
                        questionText: question.text
                    });
                    
                    // If no existing answer, search team emails for an answer
                    const searchResult = await searchForAnswer(question.text, teamId, apiKey);
                    
                    if (searchResult.success) {
                        log('Found answer from email search', 'info', {
                            answerType: searchResult.answerType,
                            confidence: searchResult.confidence,
                            referencesCount: searchResult.references?.length || 0
                        });
                        
                        question.answer = searchResult.answer;
                        question.answerType = searchResult.answerType;
                        question.confidence = searchResult.confidence;
                        question.references = searchResult.references || [];
                        question.source = 'search';
                        question.verified = false;
                        
                        // Store translation info if available
                        if (question.original_text && question.original_language && question.original_language !== 'en') {
                            question.translatedFrom = {
                                originalText: question.original_text,
                                language: question.original_language
                            };
                        }
                        
                        // Save this Q&A pair to the database for future use
                        try {
                            const saveResult = await saveQuestionAnswer(
                                question.text,
                                question.answer,
                                question.references,
                                teamId,
                                question.keywords || []
                            );
                            
                            if (saveResult.success) {
                                question.answerId = saveResult.id;
                                log('Saved new Q&A to database', 'info', { id: saveResult.id });
                            }
                        } catch (saveError) {
                            log('Error saving Q&A to database', 'error', saveError);
                            // Continue even if saving fails
                        }
                    } else {
                        log('No answer found for question', 'info', {
                            question: question.text,
                            error: searchResult.error,
                            keywords: searchResult.keywords || []
                        });
                        
                        question.noAnswerFound = true;
                        question.source = 'none';
                        
                        // Preserve keywords from search even if no answer found
                        if (searchResult.keywords && searchResult.keywords.length > 0) {
                            question.searchKeywords = searchResult.keywords;
                        }
                    }
                }
                
                answeredQuestions.push(question);
            } catch (questionError) {
                log('Error processing question', 'error', {
                    question: question.text,
                    error: questionError
                });
                
                // Add the question with error information
                question.error = questionError.message;
                answeredQuestions.push(question);
            }
        }
        
        window.OpsieApi.setLoading('qa', false);
        return {
            success: true,
            questions: answeredQuestions
        };
    } catch (error) {
        log('Exception extracting questions and answers', 'error', error);
        window.OpsieApi.setLoading('qa', false);
        return {
            success: false,
            error: error.message || 'Failed to extract questions and answers'
        };
    }
}

/**
 * Extract questions from email content
 * @param {Object} emailData - The email data
 * @param {string} apiKey - OpenAI API key
 * @returns {Promise<Object>} - Promise with extracted questions
 */
async function extractQuestionsFromEmail(emailData, apiKey) {
    try {
        log('Extracting questions from email content', 'info');
        
        // Format the email content
        const emailContent = `
From: ${emailData.sender.name} <${emailData.sender.email}>
Subject: ${emailData.subject}
Date: ${emailData.timestamp || new Date().toISOString()}

${emailData.body}
        `;
        
        // Debug: Log the email content being analyzed
        log('Email content being sent to AI for question extraction:', 'info', {
            emailContent: emailContent,
            bodyLength: emailData.body ? emailData.body.length : 0
        });
        
        // Create the API request
        const requestBody = {
            model: "gpt-4o",
            messages: [
                {
                    role: "system",
                    content: `You are an expert business email analyst that identifies questions and information requests from emails that would be valuable to answer and store in a knowledge base.

WHAT TO EXTRACT:
✅ Direct questions (marked with "?" or implied)
✅ Requests for information, clarification, or data
✅ Questions about processes, procedures, timelines
✅ Capability or resource availability questions
✅ Status or progress inquiries
✅ Questions about policies, requirements, or specifications

EXAMPLES OF WHAT TO EXTRACT:
- "Skulle du ha tid att titta på...?" → "Do you have time to look at...?"
- "Kan du vidarebefordra mejl?" → "Can you forward emails?"
- "Hur långt har vi kommit i processen?" → "How far have we progressed in the process?"
- "What is our delivery time?" → "What is our delivery time?"
- "Who is responsible for...?" → "Who is responsible for...?"

WHAT TO IGNORE:
❌ Pure greetings or pleasantries
❌ Questions already fully answered in the same email
❌ Simple yes/no confirmations without business value

CRITICAL REQUIREMENTS:
1. ALWAYS translate non-English questions to English
2. Make questions standalone and searchable
3. Be INCLUSIVE rather than overly restrictive
4. Extract questions that could help future team members

For each question found:
- text: Clear question in English
- original_text: Original text if not English
- original_language: Language code (sv, en, de, etc.)
- type: "direct_question", "information_request", or "clarification_needed"
- business_context: What area this relates to
- keywords: 3-5 search keywords
- priority: "high", "medium", or "low"

Response format:
{
  "questions": [
    {
      "text": "Do you have time to look at a solution for checking procurement emails?",
      "original_text": "Skulle du möjligtvis ha tid att titta på en lösning för att kontrollera de mejl vi får beträffande upphandlingar?",
      "original_language": "sv",
      "type": "information_request",
      "business_context": "procurement and email management",
      "keywords": ["procurement", "email", "solution", "checking"],
      "priority": "medium"
    }
  ]
}

If absolutely no questions exist, return empty array. But be generous in identifying questions!`
                },
                {
                    role: "user",
                    content: `Please extract actionable business questions from this email:\n\n${emailContent}`
                }
            ],
            temperature: 0.3
        };
        
        // Call the OpenAI API
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(requestBody)
        });
        
        // Check for API errors
        if (!response.ok) {
            const errorText = await response.text();
            log('OpenAI API error during question extraction', 'error', {
                status: response.status,
                text: errorText
            });
            throw new Error(`OpenAI API error: ${response.status} - ${errorText}`);
        }
        
        // Parse the response
        const responseData = await response.json();
        
        if (!responseData.choices || !responseData.choices[0] || !responseData.choices[0].message) {
            log('Unexpected API response format for question extraction', 'error', responseData);
            throw new Error('Unexpected API response format');
        }
        
        // Extract the content from the response
        const aiResponse = responseData.choices[0].message.content;
        log('AI response for question extraction:', 'info', aiResponse);
        
        // Parse the JSON response
        try {
            const jsonContent = extractContentFromCodeBlock(aiResponse);
            const parsedResponse = JSON.parse(jsonContent);
            
            // Validate response format
            if (!parsedResponse.questions || !Array.isArray(parsedResponse.questions)) {
                log('Invalid response format for question extraction', 'error', parsedResponse);
                throw new Error('Invalid response format: questions array missing');
            }
            
            return {
                success: true,
                questions: parsedResponse.questions
            };
        } catch (parseError) {
            log('Error parsing question extraction response', 'error', parseError);
            throw new Error(`Failed to parse response: ${parseError.message}`);
        }
    } catch (error) {
        log('Exception in extractQuestionsFromEmail', 'error', error);
        return {
            success: false,
            error: error.message || 'Failed to extract questions',
            questions: []
        };
    }
}

/**
 * Find existing answers in the QandA database
 * @param {string} questionText - The question text
 * @param {string} teamId - The team ID
 * @returns {Promise<Object>} - Promise with answer if found
 */
async function findExistingAnswer(questionText, teamId) {
    try {
        log('Searching for existing answer in qanda table', 'info', { 
            questionText, 
            teamId,
            searchType: 'semantic+exact+fuzzy'
        });
        
        // Get authentication token
        const token = await getAuthToken();
        if (!token) {
            log('No authentication token available for finding answers in qanda table', 'error');
            return { found: false };
        }
        
        // Step 1: Try an exact match search first (fastest method)
        log('Attempting exact match search in qanda table', 'info', { questionText });
        
        // Test query to debug source_filename retrieval
        log('Testing source_filename retrieval with a simple query', 'info');
        const testQuery = `qanda?team_id=eq.${teamId}&select=id,source_filename&limit=5`;
        const testResponse = await apiRequest(testQuery, 'GET');
        log('Test query results for source_filename', 'info', {
            success: testResponse.success,
            dataCount: testResponse.data ? testResponse.data.length : 0,
            sampleData: testResponse.data ? testResponse.data.slice(0, 2) : null
        });
        
        const exactMatchQuery = `qanda?team_id=eq.${teamId}&question_text=eq.${encodeURIComponent(questionText)}&select=id,question_text,answer_text,is_verified,keywords,updated_at,source_filename&order=created_at.desc`;
        
        // Log the exact query being sent
        log('Exact match query URL constructed', 'info', {
            query: exactMatchQuery,
            teamId: teamId,
            questionText: questionText,
            encodedQuestionText: encodeURIComponent(questionText)
        });
        
        const exactMatchResponse = await apiRequest(exactMatchQuery, 'GET');
        
        if (!exactMatchResponse.success) {
            log('Error during exact match search in qanda table', 'error', exactMatchResponse.error);
            return { found: false };
        }
        
        // Log the exact match results
        log(`Exact match search in qanda table returned ${exactMatchResponse.data?.length || 0} results`, 'info');
        
        // If we found an exact match, return it
        if (exactMatchResponse.data && exactMatchResponse.data.length > 0) {
            const matchedQA = exactMatchResponse.data[0];
            
            // Add detailed logging to debug source_filename issue
            log('Raw database response for exact match', 'info', {
                rawData: matchedQA,
                sourceFilenameRaw: matchedQA.source_filename,
                sourceFilenameType: typeof matchedQA.source_filename,
                sourceFilenameIsNull: matchedQA.source_filename === null,
                sourceFilenameIsUndefined: matchedQA.source_filename === undefined,
                sourceFilenameIsEmpty: matchedQA.source_filename === '',
                allKeys: Object.keys(matchedQA),
                allValues: Object.values(matchedQA)
            });
            
            log('Found exact match in qanda table', 'info', {
                qaId: matchedQA.id,
                questionText: matchedQA.question_text,
                answerText: matchedQA.answer_text ? `${matchedQA.answer_text.substring(0, 50)}...` : 'No answer',
                isVerified: matchedQA.is_verified,
                updatedAt: matchedQA.updated_at,
                sourceFilename: matchedQA.source_filename,
                availableFields: Object.keys(matchedQA)
            });
            
            const result = {
                found: true,
                id: matchedQA.id,
                answer: matchedQA.answer_text,
                verified: matchedQA.is_verified,
                references: [],  // References not currently stored in qanda table
                keywords: matchedQA.keywords || [],
                matchType: 'exact',
                updatedAt: matchedQA.updated_at,
                sourceFilename: matchedQA.source_filename
            };
            
            log('Returning exact match result with updatedAt and sourceFilename', 'info', {
                updatedAt: result.updatedAt,
                sourceFilename: result.sourceFilename,
                sourceFilenameInResult: result.sourceFilename,
                resultFields: Object.keys(result)
            });
            
            return result;
        }
        
        // Step 2: Get the OpenAI API key for embedding (needed for semantic search)
        const apiKey = localStorage.getItem('openaiApiKey');
        
        // Step 3: If we have an API key, try semantic search with embeddings
        if (apiKey) {
            log('Attempting semantic search with embeddings', 'info', { questionText });
            
            try {
                // First, fetch all Q&A pairs for this team (limited to 100 for performance)
                const allQAQuery = `qanda?team_id=eq.${teamId}&select=id,question_text,answer_text,is_verified,keywords,updated_at,source_filename&order=created_at.desc&limit=100`;
                const allQAResponse = await apiRequest(allQAQuery, 'GET');
                
                if (!allQAResponse.success || !allQAResponse.data || allQAResponse.data.length === 0) {
                    log('No Q&A data found for semantic search or error fetching data', 'info');
                } else {
                    const allQA = allQAResponse.data;
                    log(`Retrieved ${allQA.length} Q&A pairs for semantic search`, 'info');
                    
                    // Generate embedding for the query question
                    const queryEmbedding = await generateEmbedding(questionText, apiKey);
                    
                    if (!queryEmbedding) {
                        log('Failed to generate embedding for query question', 'error');
                    } else {
                        // For each Q&A pair, generate an embedding and calculate similarity
                        const similarities = [];
                        
                        // Process in batches to avoid rate limits and timeout
                        const batchSize = 5;
                        const batches = Math.ceil(allQA.length / batchSize);
                        
                        for (let batchIndex = 0; batchIndex < batches; batchIndex++) {
                            const batchStart = batchIndex * batchSize;
                            const batchEnd = Math.min(batchStart + batchSize, allQA.length);
                            const batch = allQA.slice(batchStart, batchEnd);
                            
                            // Process batch in parallel
                            const batchPromises = batch.map(async (qa) => {
                                const docEmbedding = await generateEmbedding(qa.question_text, apiKey);
                                if (!docEmbedding) return null;
                                
                                const similarity = calculateCosineSimilarity(queryEmbedding, docEmbedding);
                                return {
                                    qa,
                                    similarity
                                };
                            });
                            
                            const batchResults = await Promise.all(batchPromises);
                            batchResults.forEach(result => {
                                if (result) similarities.push(result);
                            });
                        }
                        
                        // Apply weights for recency and verification when scores are close
                        const SIMILARITY_THRESHOLD = 0.05; // Consider scores within 0.05 as close
                        const RECENCY_BOOST = 0.05; // Boost for recent answers (up to 30 days)
                        const VERIFICATION_BOOST = 0.05; // Boost for verified answers
                        const now = new Date();
                        const THIRTY_DAYS_MS = 30 * 24 * 60 * 60 * 1000;
                        
                        // Calculate adjusted scores with additional factors
                        similarities.forEach(item => {
                            // Start with base similarity score
                            let adjustedScore = item.similarity;
                            
                            // Add recency boost for answers updated within last 30 days
                            if (item.qa.updated_at) {
                                const updatedDate = new Date(item.qa.updated_at);
                                const ageMs = now - updatedDate;
                                
                                if (ageMs <= THIRTY_DAYS_MS) {
                                    // Recency boost: newer answers get up to 0.05 boost
                                    const recencyBoost = RECENCY_BOOST * (1 - (ageMs / THIRTY_DAYS_MS));
                                    adjustedScore += recencyBoost;
                                    
                                    log('Applied recency boost', 'info', {
                                        question: item.qa.question_text.substring(0, 50),
                                        updatedAt: item.qa.updated_at,
                                        ageInDays: Math.round(ageMs / (24 * 60 * 60 * 1000)),
                                        recencyBoost: recencyBoost.toFixed(3),
                                        beforeBoost: item.similarity.toFixed(3),
                                        afterBoost: adjustedScore.toFixed(3)
                                    });
                                }
                            }
                            
                            // Add verification boost
                            if (item.qa.is_verified) {
                                adjustedScore += VERIFICATION_BOOST; // Verified answers get a 0.05 boost
                                
                                log('Applied verification boost', 'info', {
                                    question: item.qa.question_text.substring(0, 50),
                                    verificationBoost: VERIFICATION_BOOST,
                                    beforeBoost: (adjustedScore - VERIFICATION_BOOST).toFixed(3),
                                    afterBoost: adjustedScore.toFixed(3)
                                });
                            }
                            
                            // Store the adjusted score
                            item.adjustedScore = adjustedScore;
                        });
                        
                        // Sort by adjusted score (highest first)
                        similarities.sort((a, b) => b.adjustedScore - a.adjustedScore);
                        
                        // Check if boosts affected the ranking when scores are close
                        if (similarities.length >= 2) {
                            const topMatch = similarities[0];
                            const runnerUp = similarities[1];
                            
                            if (Math.abs(topMatch.similarity - runnerUp.similarity) <= SIMILARITY_THRESHOLD &&
                                topMatch.adjustedScore > runnerUp.adjustedScore) {
                                log('Ranking changed due to recency/verification boosts', 'info', {
                                    originalTopMatch: {
                                        question: runnerUp.qa.question_text.substring(0, 50),
                                        similarity: runnerUp.similarity.toFixed(3),
                                        adjustedScore: runnerUp.adjustedScore.toFixed(3),
                                        isVerified: runnerUp.qa.is_verified,
                                        updatedAt: runnerUp.qa.updated_at
                                    },
                                    newTopMatch: {
                                        question: topMatch.qa.question_text.substring(0, 50),
                                        similarity: topMatch.similarity.toFixed(3),
                                        adjustedScore: topMatch.adjustedScore.toFixed(3),
                                        isVerified: topMatch.qa.is_verified,
                                        updatedAt: topMatch.qa.updated_at
                                    }
                                });
                            }
                        }
                        
                        // Log adjusted similarity scores for debugging
                        log('Adjusted similarity scores (with recency & verification boosts):', 'info', 
                            similarities.slice(0, 3).map(s => ({
                                question: s.qa.question_text,
                                baseSimilarity: s.similarity.toFixed(3),
                                adjustedScore: s.adjustedScore.toFixed(3),
                                isVerified: s.qa.is_verified,
                                updatedAt: s.qa.updated_at
                            }))
                        );
                        
                        // If we found a semantically similar question with high similarity
                        if (similarities.length > 0 && similarities[0].similarity > 0.35) {
                            const bestMatch = similarities[0];
                            
                            // Add detailed logging to debug source_filename issue
                            log('Raw database response for semantic match', 'info', {
                                rawData: bestMatch.qa,
                                sourceFilenameRaw: bestMatch.qa.source_filename,
                                sourceFilenameType: typeof bestMatch.qa.source_filename,
                                sourceFilenameIsNull: bestMatch.qa.source_filename === null,
                                sourceFilenameIsUndefined: bestMatch.qa.source_filename === undefined,
                                sourceFilenameIsEmpty: bestMatch.qa.source_filename === '',
                                allKeys: Object.keys(bestMatch.qa),
                                allValues: Object.values(bestMatch.qa)
                            });
                            
                            log('Found semantic match in qanda table', 'info', {
                                qaId: bestMatch.qa.id,
                                questionText: bestMatch.qa.question_text,
                                similarityScore: bestMatch.similarity,
                                answerText: bestMatch.qa.answer_text ? `${bestMatch.qa.answer_text.substring(0, 50)}...` : 'No answer',
                                isVerified: bestMatch.qa.is_verified,
                                updatedAt: bestMatch.qa.updated_at,
                                sourceFilename: bestMatch.qa.source_filename,
                                availableFields: Object.keys(bestMatch.qa)
                            });
                            
                            const result = {
                                found: true,
                                id: bestMatch.qa.id,
                                answer: bestMatch.qa.answer_text,
                                verified: bestMatch.qa.is_verified,
                                references: [],
                                keywords: bestMatch.qa.keywords || [],
                                matchType: 'semantic',
                                similarityScore: bestMatch.similarity,
                                adjustedScore: bestMatch.adjustedScore,
                                originalQuestion: bestMatch.qa.question_text,
                                updatedAt: bestMatch.qa.updated_at,
                                sourceFilename: bestMatch.qa.source_filename
                            };
                            
                            // Log additional information about why this answer was selected
                            if (bestMatch.adjustedScore > bestMatch.similarity) {
                                log('Answer selection was influenced by recency/verification boosts', 'info', {
                                    question: bestMatch.qa.question_text.substring(0, 50),
                                    baseSimilarity: bestMatch.similarity.toFixed(3),
                                    adjustedScore: bestMatch.adjustedScore.toFixed(3),
                                    boost: (bestMatch.adjustedScore - bestMatch.similarity).toFixed(3),
                                    isVerified: bestMatch.qa.is_verified,
                                    updatedAt: bestMatch.qa.updated_at,
                                    ageInDays: bestMatch.qa.updated_at ? 
                                        Math.round((now - new Date(bestMatch.qa.updated_at)) / (24 * 60 * 60 * 1000)) : 'unknown'
                                });
                            }
                            
                            log('Returning semantic match result with updatedAt', 'info', {
                                updatedAt: result.updatedAt,
                                resultFields: Object.keys(result)
                            });
                            
                            return result;
                        } else {
                            log('No high-confidence semantic match found', 'info', {
                                bestMatchScore: similarities.length > 0 ? similarities[0].similarity : 0
                            });
                        }
                    }
                }
            } catch (semanticError) {
                log('Error during semantic search', 'error', {
                    error: semanticError.message,
                    stack: semanticError.stack
                });
                // Continue to fuzzy search if semantic search fails
            }
        } else {
            log('No OpenAI API key available for semantic search, falling back to fuzzy search', 'warning');
        }
        
        // Step 4: Try a fuzzy match search (ilike operator) as a fallback
        log('No exact or semantic match found, attempting fuzzy search in qanda table', 'info', { questionText });
        
        // Normalize the question text for better matching
        const normalizedQuestion = questionText.toLowerCase().trim();
        
        // Extract keywords using a more sophisticated approach
        // First, remove common stop words
        const stopWords = ['a', 'an', 'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'with', 'about', 'is', 'are', 'was', 'were', 'be', 'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did', 'can', 'could', 'will', 'would', 'should', 'may', 'might', 'must', 'shall', 'please', 'thank', 'you', 'your', 'me', 'my', 'mine', 'i', 'we', 'us', 'our', 'ours', 'they', 'them', 'their', 'theirs'];
        
        // Split by non-alphanumeric characters and filter
        const wordCandidates = normalizedQuestion.split(/[^a-z0-9]+/)
            .filter(word => word.length > 2)  // Only keep words with 3+ characters
            .filter(word => !stopWords.includes(word)); // Remove stop words
        
        // Extract key business terms that are more likely to be relevant
        const businessTerms = extractBusinessTerms(normalizedQuestion);
        
        // Combine business terms with remaining filtered words
        const searchWords = [...new Set([...businessTerms, ...wordCandidates])];
        
        // Only perform fuzzy search if we have meaningful words to search for
        if (searchWords.length === 0) {
            log('No suitable search terms for fuzzy search in qanda table', 'warning');
            return { found: false };
        }
        
        // Create a query that searches for key terms using ILIKE
        // Use more terms for better matching
        const searchTerms = searchWords.slice(0, 4);
        
        // Log the search terms we're using
        log('Using the following terms for fuzzy search in qanda table', 'info', searchTerms);
        
        // Build the query conditions for each search term
        let fuzzyQuery = `qanda?team_id=eq.${teamId}&select=id,question_text,answer_text,is_verified,keywords,updated_at,source_filename`;
        
        // Add search conditions
        searchTerms.forEach((term, index) => {
            // For the first term, we start with "and="
            // For subsequent terms, we add more conditions
            const prefix = index === 0 ? '&question_text=ilike.' : ',';
            fuzzyQuery += `${prefix}%25${encodeURIComponent(term)}%25`;
        });
        
        fuzzyQuery += '&order=created_at.desc';
        
        log('Executing fuzzy search query on qanda table', 'info', { fuzzyQuery });
        const fuzzyMatchResponse = await apiRequest(fuzzyQuery, 'GET');
        
        if (!fuzzyMatchResponse.success) {
            log('Error during fuzzy search in qanda table', 'error', fuzzyMatchResponse.error);
            return { found: false };
        }
        
        // Log the fuzzy match results
        log(`Fuzzy search in qanda table returned ${fuzzyMatchResponse.data?.length || 0} results`, 'info');
        
        // If we found fuzzy matches, use the first (most recent) one
        if (fuzzyMatchResponse.data && fuzzyMatchResponse.data.length > 0) {
            const matchedQA = fuzzyMatchResponse.data[0];
            
            // Add detailed logging to debug source_filename issue
            log('Raw database response for fuzzy match', 'info', {
                rawData: matchedQA,
                sourceFilenameRaw: matchedQA.source_filename,
                sourceFilenameType: typeof matchedQA.source_filename,
                sourceFilenameIsNull: matchedQA.source_filename === null,
                sourceFilenameIsUndefined: matchedQA.source_filename === undefined,
                sourceFilenameIsEmpty: matchedQA.source_filename === '',
                allKeys: Object.keys(matchedQA),
                allValues: Object.values(matchedQA)
            });
            
            log('Found fuzzy match in qanda table', 'info', {
                qaId: matchedQA.id,
                questionText: matchedQA.question_text,
                answerText: matchedQA.answer_text ? `${matchedQA.answer_text.substring(0, 50)}...` : 'No answer',
                isVerified: matchedQA.is_verified,
                matchType: 'fuzzy',
                updatedAt: matchedQA.updated_at,
                sourceFilename: matchedQA.source_filename,
                availableFields: Object.keys(matchedQA)
            });
            
            const result = {
                found: true,
                id: matchedQA.id,
                answer: matchedQA.answer_text,
                verified: matchedQA.is_verified,
                references: [],  // References not currently stored in qanda table
                keywords: matchedQA.keywords || [],
                matchType: 'fuzzy',
                similarityScore: 0.6, // Fuzzy matches are given a moderate similarity score
                originalQuestion: matchedQA.question_text,
                updatedAt: matchedQA.updated_at,
                sourceFilename: matchedQA.source_filename
            };
            
            log('Returning fuzzy match result with updatedAt', 'info', {
                updatedAt: result.updatedAt,
                resultFields: Object.keys(result)
            });
            
            return result;
        }
        
        // No match found in any search method
        log('No matching answer found in qanda table (tried exact, semantic, and fuzzy search)', 'info');
        return { found: false };
    } catch (error) {
        log('Exception in findExistingAnswer when querying qanda table', 'error', {
            error: error.message,
            stack: error.stack
        });
        return { found: false };
    }
}

/**
 * Generate an embedding vector for a text using OpenAI API
 * @param {string} text - The text to generate embedding for
 * @param {string} apiKey - OpenAI API key
 * @returns {Promise<Array<number>|null>} - Promise resolving to embedding vector or null on error
 */
async function generateEmbedding(text, apiKey) {
    try {
        log('Generating embedding for text', 'info', { textLength: text.length });
        
        const response = await fetch('https://api.openai.com/v1/embeddings', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: "text-embedding-3-small",
                input: text
            })
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            log('OpenAI API error generating embedding', 'error', {
                status: response.status,
                text: errorText
            });
            return null;
        }
        
        const responseData = await response.json();
        
        if (!responseData.data || !responseData.data[0] || !responseData.data[0].embedding) {
            log('Unexpected API response format for embedding', 'error', responseData);
            return null;
        }
        
        return responseData.data[0].embedding;
    } catch (error) {
        log('Error generating embedding', 'error', error);
        return null;
    }
}

/**
 * Calculate cosine similarity between two embedding vectors
 * @param {Array<number>} vecA - First embedding vector
 * @param {Array<number>} vecB - Second embedding vector
 * @returns {number} - Similarity score between 0 and 1
 */
function calculateCosineSimilarity(vecA, vecB) {
    if (!vecA || !vecB || vecA.length !== vecB.length) {
        return 0;
    }
    
    let dotProduct = 0;
    let normA = 0;
    let normB = 0;
    
    for (let i = 0; i < vecA.length; i++) {
        dotProduct += vecA[i] * vecB[i];
        normA += vecA[i] * vecA[i];
        normB += vecB[i] * vecB[i];
    }
    
    normA = Math.sqrt(normA);
    normB = Math.sqrt(normB);
    
    if (normA === 0 || normB === 0) {
        return 0;
    }
    
    return dotProduct / (normA * normB);
}

/**
 * Extract business-related terms from text that are likely relevant for search
 * @param {string} text - The input text
 * @returns {Array<string>} - Array of business terms
 */
function extractBusinessTerms(text) {
    // Common business and e-commerce related terms
    const businessTermPatterns = [
        'delivery', 'shipping', 'order', 'arrival', 'arrive', 'receive',
        'timeline', 'schedule', 'time', 'date', 'when', 'expect',
        'price', 'cost', 'discount', 'offer', 'deal', 'promotion',
        'refund', 'return', 'cancel', 'warranty', 'guarantee',
        'product', 'item', 'purchase', 'buying', 'bought',
        'payment', 'invoice', 'billing', 'subscription',
        'account', 'login', 'password', 'profile',
        'support', 'help', 'service', 'contact',
        'availability', 'stock', 'inventory',
        'feature', 'specification', 'detail',
        'track', 'status', 'update',
        'download', 'upload', 'install',
        'version', 'model', 'size', 'color'
    ];
    
    const lowerText = text.toLowerCase();
    return businessTermPatterns.filter(term => lowerText.includes(term));
}

/**
 * Search team emails for an answer to a question
 * @param {string} questionText - The question text
 * @param {string} teamId - The team ID
 * @param {string} apiKey - OpenAI API key
 * @returns {Promise<Object>} - Promise with search results
 */
async function searchForAnswer(questionText, teamId, apiKey) {
    try {
        log('Searching team emails for answer', 'info', { questionText });
        
        // Get authentication token
        const token = await getAuthToken();
        if (!token) {
            log('No authentication token available for searching emails', 'error');
            return {
                success: false,
                error: 'Authentication required to search emails'
            };
        }
        
        // Fetch recent emails for the team (limited to 25 for now)
        // This will be enhanced to use better search once we set up proper indexing
        const messagesResponse = await apiRequest(
            `messages?team_id=eq.${teamId}&select=id,sender_name,sender_email,message_body,created_at&order=created_at.desc&limit=25`,
            'GET'
        );
        
        if (!messagesResponse.success || !messagesResponse.data) {
            log('Failed to fetch team emails for search', 'error', messagesResponse.error);
            return {
                success: false,
                error: 'Failed to retrieve email data for search'
            };
        }
        
        const emails = messagesResponse.data;
        log(`Retrieved ${emails.length} team emails for search`, 'info');
        
        if (emails.length === 0) {
            return {
                success: false,
                error: 'No emails available to search for answers'
            };
        }
        
        // Format emails for the API call
        const emailsContent = emails.map((email, index) => {
            let formattedContent = email.message_body || 'No content';
            
            // Truncate long content
            if (formattedContent.length > 1000) {
                formattedContent = formattedContent.substring(0, 1000) + '...(truncated)';
            }
            
            return `
Email #${index + 1}:
From: ${email.sender_name} <${email.sender_email}>
Date: ${new Date(email.created_at).toISOString()}
ID: ${email.id}
Content: ${formattedContent}
            `;
        }).join('\n\n');
        
        // Create the API request to search for an answer
        const requestBody = {
            model: "gpt-4o",
            messages: [
                {
                    role: "system",
                    content: `You are a strict business knowledge assistant. Your ONLY job is to find actual answers to questions - NOT to summarize or interpret the questions.

STRICT RULES - ONLY PROVIDE ANSWERS IF:
✅ The emails contain a direct response to the question
✅ The emails provide specific facts, data, or information that answers the question
✅ Someone explicitly answers the question in their reply
✅ The emails contain clear instructions, procedures, or specifications requested

❌ NEVER PROVIDE AN ANSWER IF:
❌ You're just restating what the question is asking
❌ You're summarizing the question or describing the request
❌ You're interpreting what someone "is asking for"
❌ The emails only contain the question but no answer
❌ You're speculating or inferring information

EXAMPLES:
Question: "What is our delivery time?"
✅ GOOD: "Standard delivery is 3-5 business days" (if this info exists in emails)
❌ BAD: "John is asking about delivery times" (this is just restating the question)

Question: "Can you send me the report?"
✅ GOOD: "Report sent via email attachment on Monday" (if someone responded)
❌ BAD: "Sarah is requesting the report" (this is just describing the request)

Question: "Do you have time to review emails?"
✅ GOOD: "Yes, I can review them this afternoon" (if someone replied this)
❌ BAD: "Marcus is asking if Philip has time to review" (this is just summarizing)

CRITICAL TEST: Before providing any answer, ask:
"Does this answer actually tell me something I didn't already know from reading the question?"

If the answer is NO, then return null.

Response format:
{
  "answer": null,  // ONLY provide actual answers, not summaries
  "confidence": 0,
  "answer_type": null,
  "references": [],
  "keywords": ["relevant", "search", "terms"]
}

BE EXTREMELY STRICT. When in doubt, return null. It's better to have no answer than a useless summary.`
                },
                {
                    role: "user",
                    content: `Find a specific, actionable answer to this question: ${questionText}\n\nSearch through these emails:\n\n${emailsContent}`
                }
            ],
            temperature: 0.2
        };
        
        // Call the OpenAI API
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify(requestBody)
        });
        
        // Check for API errors
        if (!response.ok) {
            const errorText = await response.text();
            log('OpenAI API error during answer search', 'error', {
                status: response.status,
                text: errorText
            });
            throw new Error(`OpenAI API error: ${response.status} - ${errorText}`);
        }
        
        // Parse the response
        const responseData = await response.json();
        
        if (!responseData.choices || !responseData.choices[0] || !responseData.choices[0].message) {
            log('Unexpected API response format for answer search', 'error', responseData);
            throw new Error('Unexpected API response format');
        }
        
        // Extract the content from the response
        const aiResponse = responseData.choices[0].message.content;
        log('AI response for answer search:', 'info', aiResponse);
        
        // Parse the JSON response
        try {
            const jsonContent = extractContentFromCodeBlock(aiResponse);
            const parsedResponse = JSON.parse(jsonContent);
            
            // Check if an answer was found
            if (parsedResponse.answer === null || !parsedResponse.answer) {
                log('No actionable answer found for question', 'info', { 
                    question: questionText,
                    keywords: parsedResponse.keywords || [],
                    confidence: parsedResponse.confidence
                });
                return {
                    success: false,
                    error: 'No actionable answer found in team emails',
                    keywords: parsedResponse.keywords || []
                };
            }
            
            // Log successful answer with details
            log('Found actionable answer for question', 'info', {
                question: questionText,
                answerType: parsedResponse.answer_type,
                confidence: parsedResponse.confidence,
                referencesCount: parsedResponse.references?.length || 0,
                answerPreview: parsedResponse.answer.substring(0, 100) + '...'
            });
            
            return {
                success: true,
                answer: parsedResponse.answer,
                answerType: parsedResponse.answer_type,
                confidence: parsedResponse.confidence,
                references: parsedResponse.references || [],
                keywords: parsedResponse.keywords || []
            };
        } catch (parseError) {
            log('Error parsing answer search response', 'error', parseError);
            throw new Error(`Failed to parse search response: ${parseError.message}`);
        }
    } catch (error) {
        log('Exception in searchForAnswer', 'error', error);
        return {
            success: false,
            error: error.message || 'Failed to search for answer'
        };
    }
}

/**
 * Save a question and answer to the database
 * @param {string} question - The question text
 * @param {string} answer - The answer text
 * @param {Array} references - References supporting the answer
 * @param {string} teamId - The team ID
 * @param {Array} keywords - Keywords for better search
 * @param {boolean} isUserVerified - Whether this is a user-verified answer (default: false)
 * @returns {Promise<Object>} - Promise with result of save operation
 */
async function saveQuestionAnswer(question, answer, references, teamId, keywords = [], isUserVerified = false, confidenceScore = null, sourceFilename = null) {
    try {
        const id = generateUUID();
        const userId = await getUserId();

        log('Saving question and answer to database', 'info', {
            question,
            answer,
            referencesCount: references?.length || 0,
            keywords,
            isUserVerified,
            sourceFilename
        });

        const response = await apiRequest('/qanda', 'POST', {
            id,
            question_text: question,
            answer_text: answer,
            keywords: keywords,
            confidence_score: confidenceScore,
            team_id: teamId,
            created_by: userId,
            last_updated_by: userId,
            is_verified: isUserVerified,
            source_filename: sourceFilename
        });

        if (!response.success) {
            throw new Error(response.error || 'Failed to save Q&A pair');
        }

        return {
            success: true,
            id: id
        };
    } catch (error) {
        log('Error saving Q&A pair', 'error', error);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * Generate a UUID v4
 * @returns {string} A UUID v4 string
 */
function generateUUID() {
    // Simple UUID v4 generation
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        const r = Math.random() * 16 | 0;
        const v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}

/**
 * Load Q&A data from the database
 * @param {string} teamId - The team ID
 * @param {string} search - Optional search term
 * @param {boolean} onlyVerified - Whether to only return verified answers
 * @returns {Promise<Object>} - Promise with the Q&A data
 */
async function loadQandAData(teamId, search = null, onlyVerified = false) {
    try {
        log('Loading Q&A data from database', 'info', { teamId, search, onlyVerified });
        
        // Get authentication token
        const token = await getAuthToken();
        if (!token) {
            log('No authentication token available for loading Q&A', 'error');
            return {
                success: false,
                error: 'Authentication required to load Q&A'
            };
        }
        
        // Prepare the query
        let endpoint = `qanda?team_id=eq.${teamId}`;
        
        // Add search filter if provided
        if (search && search.trim().length > 0) {
            const searchTerm = encodeURIComponent(search.trim());
            // Search in question text, answer text, and keywords
            endpoint += `&or=(question_text.ilike.%25${searchTerm}%25,answer_text.ilike.%25${searchTerm}%25)`;
        }
        
        // Add verification filter if requested
        if (onlyVerified) {
            endpoint += '&is_verified=eq.true';
        }
        
        // Order by most recent first
        endpoint += '&order=created_at.desc';
        
        // Make the API request
        const response = await apiRequest(endpoint);
        
        if (!response.success) {
            log('Failed to load Q&A data', 'error', response.error);
            return {
                success: false,
                error: response.error || 'Failed to load Q&A data'
            };
        }
        
        log(`Loaded ${response.data.length} Q&A items from database`, 'info');
        
        // Format the response data into a more usable structure
        const formattedData = response.data.map(item => ({
            id: item.id,
            question: item.question_text,
            answer: item.answer_text,
            keywords: item.keywords || [],
            isVerified: item.is_verified,
            confidence: item.confidence_score,
            createdAt: item.created_at,
            updatedAt: item.updated_at,
            createdBy: item.created_by,
            updatedBy: item.last_updated_by,
            sourceMessageId: item.source_message_id,
            sourceFilename: item.source_filename
        }));
        
        return {
            success: true,
            data: formattedData
        };
    } catch (error) {
        log('Error loading Q&A data', 'error', error);
        return {
            success: false,
            error: error.message || 'Failed to load Q&A data'
        };
    }
}

// Add functions to the global OpsieApi object
window.OpsieApi = window.OpsieApi || {};
Object.assign(window.OpsieApi, {
    // ... existing functions ...
    extractQuestionsAndAnswers,
    findExistingAnswer,
    searchForAnswer,
    saveQuestionAnswer,
    loadQandAData,
    generateUUID,
    generateEmbedding,
    calculateCosineSimilarity,
    extractBusinessTerms,
    analyzeQuestionRelationships,
    setLoading,
});

/**
 * Enhances question matching by checking if a new question is semantically related to existing ones
 * @param {string} questionText - The new question text
 * @param {string} teamId - The team ID
 * @param {string} apiKey - OpenAI API key
 * @returns {Promise<Object>} - Promise with semantic analysis results
 */
async function analyzeQuestionRelationships(questionText, teamId, apiKey) {
    try {
        log('Analyzing question for semantic relationships', 'info', { questionText });
        
        // Get all existing questions from the database
        const existingQAQuery = `qanda?team_id=eq.${teamId}&select=id,question_text,answer_text,is_verified&order=created_at.desc&limit=50`;
        const existingQAResponse = await apiRequest(existingQAQuery, 'GET');
        
        if (!existingQAResponse.success || !existingQAResponse.data || existingQAResponse.data.length === 0) {
            log('No existing Q&A data found for semantic analysis', 'info');
            return { 
                success: true,
                isRelated: false,
                message: 'No existing questions to compare with'
            };
        }
        
        const existingQA = existingQAResponse.data;
        log(`Retrieved ${existingQA.length} existing Q&A pairs for semantic analysis`, 'info');
        
        // Format the existing questions for the AI
        const existingQuestions = existingQA.map(qa => ({
            id: qa.id,
            question: qa.question_text,
            answer: qa.answer_text || 'No answer',
            isVerified: qa.is_verified
        }));
        
        // Use OpenAI to analyze if the new question is semantically related to any existing ones
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: "gpt-4o",
                messages: [
                    {
                        role: "system",
                        content: `You are an AI specialized in semantic understanding of questions. 
Your task is to determine if a new question is semantically similar to any existing questions in a database.

Analyze the new question and the list of existing questions. For each existing question, determine:
1. If it's asking for the same information despite different wording
2. The similarity score (0.0-1.0) indicating how closely they match semantically
3. If the existing question's answer would appropriately answer the new question

Output your analysis in JSON format:
{
  "isRelated": true/false,
  "relatedQuestions": [
    {
      "id": "existing-question-id",
      "question": "existing question text",
      "similarity": 0.92,
      "explanation": "Brief explanation of how they're related"
    }
  ],
  "recommendedAction": "use-existing" OR "create-new",
  "explanation": "Explanation of your recommendation"
}

Only include existing questions with similarity scores above 0.7 in the relatedQuestions array.
If no questions are sufficiently similar, set isRelated to false and return an empty relatedQuestions array.`
                    },
                    {
                        role: "user",
                        content: `New question: "${questionText}"

Existing questions:
${JSON.stringify(existingQuestions, null, 2)}`
                    }
                ],
                temperature: 0.3
            })
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            log('OpenAI API error during question relationship analysis', 'error', {
                status: response.status,
                text: errorText
            });
            return {
                success: false,
                error: `OpenAI API error: ${response.status} - ${errorText}`
            };
        }
        
        const responseData = await response.json();
        const content = responseData.choices[0].message.content;
        
        // Extract JSON from the response
        const jsonContent = extractContentFromCodeBlock(content);
        const analysis = JSON.parse(jsonContent);
        
        log('Question relationship analysis complete', 'info', analysis);
        
        return {
            success: true,
            ...analysis
        };
    } catch (error) {
        log('Error analyzing question relationships', 'error', error);
        return {
            success: false,
            error: error.message || 'Failed to analyze question relationships',
            isRelated: false
        };
    }
}

/**
 * Process a document and generate Q&A pairs
 * @param {File} file - The uploaded document file
 * @param {string} teamId - The team ID
 * @returns {Promise<Object>} - Promise resolving to processing results
 */
async function processDocumentForQA(file, teamId) {
    try {
        log('Processing document for Q&A generation', 'info', {
            fileName: file.name,
            fileSize: file.size,
            fileType: file.type,
            teamId: teamId
        });

        // Get OpenAI API key
        const apiKey = await getOpenAIApiKey();
        if (!apiKey) {
            throw new Error('OpenAI API key is required for document processing');
        }

        // Define the extraction prompt
        const extractionPrompt = `You are analyzing a business document that may contain technical specifications, product details, or other business-relevant information. 
        
        Please extract all relevant information, focusing on:
        1. Product specifications and measurements
        2. Technical details and capabilities
        3. Material information
        4. Pricing or cost-related details (if any)
        5. Compatibility or installation requirements
        6. Any other business-critical information
        
        Ignore any irrelevant metadata like drawing numbers, project IDs, or administrative details unless they are essential product identifiers.
        
        Format the information in a clear, structured way that emphasizes the practical, business-relevant details.
        
        If you see any measurements or technical specifications in the image, be sure to note them precisely with their units.`;

        let extractedContent;

        // Handle PDFs differently from images
        if (file.type === 'application/pdf') {
            log('Processing PDF file', 'info');
            try {
                // Convert file to ArrayBuffer for pdf.js
                const arrayBuffer = await file.arrayBuffer();
                
                // Load the PDF document using pdf.js
                const pdfjsLib = window['pdfjs-dist/build/pdf'];
                pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
                
                const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
                log('PDF loaded successfully', 'info', { numPages: pdf.numPages });

                // Create a canvas element for rendering
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');

                // Process each page
                let allContent = [];
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    
                    // Extract text
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    
                    // Render page to canvas
                    const viewport = page.getViewport({ scale: 1.5 }); // Increased scale for better quality
                    canvas.height = viewport.height;
                    canvas.width = viewport.width;
                    await page.render({ canvasContext: ctx, viewport: viewport }).promise;
                    
                    // Convert canvas to base64 image
                    const pageImage = canvas.toDataURL('image/jpeg', 0.95);
                    
                    // Extract base64 data from data URL
                    const base64Data = pageImage.split(',')[1];
                    
                    log('Processing page image', 'info', {
                        pageNumber: i,
                        imageSize: base64Data.length,
                        dimensions: `${canvas.width}x${canvas.height}`
                    });

                    // Use Vision API for the page image
                    log('Preparing Vision API request', 'info', {
                        pageNumber: i,
                        modelName: "o4-mini",
                        promptLength: extractionPrompt.length
                    });

                    const visionResponse = await fetch('https://api.openai.com/v1/chat/completions', {
                        method: 'POST',
                        headers: {
                            'Authorization': `Bearer ${apiKey}`,
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify({
                            model: "o4-mini",
                            messages: [
                                {
                                    role: "user",
                                    content: [
                                        { type: "text", text: extractionPrompt },
                                        {
                                            type: "image_url",
                                            image_url: {
                                                url: `data:image/jpeg;base64,${base64Data}`
                                            }
                                        }
                                    ]
                                }
                            ],
                            max_completion_tokens: 4000
                        })
                    });

                    if (!visionResponse.ok) {
                        const errorText = await visionResponse.text();
                        log('Vision API error response', 'error', {
                            status: visionResponse.status,
                            statusText: visionResponse.statusText,
                            errorText: errorText,
                            responseHeaders: Object.fromEntries([...visionResponse.headers.entries()])
                        });

                        try {
                            const errorJson = JSON.parse(errorText);
                            log('Parsed error details', 'error', {
                                errorType: errorJson.error?.type,
                                errorMessage: errorJson.error?.message,
                                errorCode: errorJson.error?.code
                            });
                            throw new Error(`Vision API Error: ${errorJson.error?.message || 'Unknown error'}`);
                        } catch (e) {
                            throw new Error(`Failed to process page ${i} (${visionResponse.status}): ${errorText}`);
                        }
                    }

                    const visionData = await visionResponse.json();
                    const visionContent = visionData.choices[0].message.content;

                    // Combine text and vision analysis
                    allContent.push(pageText, visionContent);
                }

                extractedContent = allContent.join('\n\n');
                log('PDF processed successfully', 'info', {
                    contentLength: extractedContent.length,
                    preview: extractedContent.substring(0, 100) + '...'
                });

            } catch (error) {
                log('Error processing PDF', 'error', error);
                throw new Error('Failed to process PDF: ' + error.message);
            }
        } else {
            // Handle non-PDF files (images) with Vision API directly
            log('Processing image file', 'info');
            const base64Content = await new Promise((resolve) => {
                const reader = new FileReader();
                reader.onload = () => {
                    const base64 = reader.result.split(',')[1];
                    resolve(base64);
                };
                reader.readAsDataURL(file);
            });

            const extractionResponse = await fetch('https://api.openai.com/v1/chat/completions', {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${apiKey}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    model: "o4-mini",
                    messages: [
                        {
                            role: "user",
                            content: [
                                { type: "text", text: extractionPrompt },
                                {
                                    type: "image_url",
                                    image_url: {
                                        url: `data:${file.type};base64,${base64Content}`
                                    }
                                }
                            ]
                        }
                    ],
                    max_completion_tokens: 4000
                })
            });

            if (!extractionResponse.ok) {
                throw new Error(`Failed to process image: ${extractionResponse.status}`);
            }

            const extractionData = await extractionResponse.json();
            extractedContent = extractionData.choices[0].message.content;
        }

        // Continue with Q&A generation using the extracted content
        log('Generating Q&A pairs from extracted content', 'info');
        
        const qaPrompt = `You are helping create a knowledge base for business employees who need quick, accurate answers about various business documents and information when communicating with stakeholders.

        Analyze the following content and generate question-answer pairs that would be valuable in business scenarios. The content could be from any type of business document (technical specifications, product descriptions, process documents, policies, reports, etc.).

        Focus on extracting information that would be useful for:
        - Answering stakeholder inquiries
        - Understanding key facts and figures
        - Clarifying important details and requirements
        - Supporting decision-making
        - Facilitating effective communication

        Rules:
        1. Only create Q&A pairs for information that is explicitly present in the content
        2. Make questions clear, specific, and directly relevant to business needs
        3. Provide precise answers that include all relevant details (numbers, dates, specifications, etc.)
        4. Focus on substantive information rather than formatting or metadata
        5. Skip creating questions if the information is unclear or incomplete
        6. Maintain any specific units, terminology, or formatting from the original content
        7. Adapt the detail level and technical depth to match the source material

        Format the output as JSON:
        {
            "qaPairs": [
                {
                    "question": "string: A clear, specific business question",
                    "answer": "string: A precise, detailed answer",
                    "keywords": ["array: relevant search terms"],
                    "confidence_score": "number: 0.0 to 1.0"
                }
            ]
        }

        Content to analyze:
        ${extractedContent}`;

        const qaResponse = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${apiKey}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                model: "gpt-4",
                messages: [
                    { role: "user", content: qaPrompt }
                ],
                max_tokens: 4000
            })
        });

        if (!qaResponse.ok) {
            const errorText = await qaResponse.text();
            log('Q&A generation error', 'error', {
                status: qaResponse.status,
                statusText: qaResponse.statusText,
                errorText: errorText
            });
            throw new Error(`Failed to generate Q&A pairs: ${qaResponse.status} - ${errorText}`);
        }

        const qaData = await qaResponse.json();
        const qaPairs = JSON.parse(qaData.choices[0].message.content);

        log('Successfully generated Q&A pairs', 'info', {
            numberOfPairs: qaPairs.qaPairs.length,
            sampleQuestion: qaPairs.qaPairs[0]?.question
        });

        // Save each Q&A pair to the database
        const savedPairs = [];
        for (const pair of qaPairs.qaPairs) {
            try {
                const result = await saveQuestionAnswer(
                    pair.question,
                    pair.answer,
                    [], // no references for document-generated Q&A
                    teamId,
                    pair.keywords,
                    false, // not verified
                    pair.confidence_score,
                    file.name // Add the filename here
                );
                if (result.success) {
                    savedPairs.push(result);
                }
                log('Saved Q&A pair', 'info', {
                    question: pair.question,
                    success: result.success,
                    id: result.id,
                    sourceFilename: file.name
                });
            } catch (error) {
                log('Error saving Q&A pair', 'error', error);
                // Continue with other pairs even if one fails
            }
        }

        return {
            success: true,
            message: `Successfully processed document and generated ${savedPairs.length} Q&A pairs`,
            data: {
                totalPairs: qaPairs.qaPairs.length,
                savedPairs: savedPairs.length
            }
        };

    } catch (error) {
        log('Error processing document', 'error', error);
        return {
            success: false,
            error: error.message || 'Failed to process document'
        };
    }
}

// Add to the window.OpsieApi object
window.OpsieApi.processDocumentForQA = processDocumentForQA;

/**
 * Check if an email address is already in use
 * @param {string} email - The email address to check
 * @returns {Promise} - Whether the email is already in use
 */
async function checkEmailExists(email) {
  try {
    log('Checking if email already exists:', 'info', email);
    
    // Try to search for users with this email using the REST API
    const response = await fetch(`${API_BASE_URL}/users?email=eq.${encodeURIComponent(email)}&select=id`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY
      }
    });
    
    if (!response.ok) {
      log('Error checking if email exists - status:', 'error', response.status);
      return { success: false, error: 'Failed to check email', exists: false };
    }
    
    const users = await response.json();
    log('Email exists check response:', 'info', users);
    
    // If we find any users with this email, it exists
    const exists = Array.isArray(users) && users.length > 0;
    
    return { 
      success: true, 
      exists: exists,
      data: users
    };
  } catch (err) {
    log('Exception when checking if email exists:', 'error', err);
    return { success: false, error: err.message, exists: false };
  }
}

/**
 * Normalize user data to handle different response formats from Supabase
 * @param {Object} data - The response data from Supabase
 * @returns {Object} - Normalized user data
 */
function normalizeUserData(data) {
  if (!data) return null;
  
  // If the data already has a user object, use it
  if (data.user) {
    return {
      id: data.user.id,
      email: data.user.email,
      created_at: data.user.created_at || data.user.created_at,
      // Add any other fields we need
      ...data.user
    };
  }
  
  // If the data has id and email at the root (Supabase Auth API format)
  if (data.id && data.email) {
    return {
      id: data.id,
      email: data.email,
      created_at: data.created_at,
      // Add any other fields we need
      ...data
    };
  }
  
  // If we can't find user data, return null
  return null;
}

/**
 * Create user record in the database
 * @param {string} userId - User ID from auth
 * @param {string} email - User's email
 * @param {string} firstName - User's first name
 * @param {string} lastName - User's last name
 * @returns {Promise} - Result of user record creation
 */
async function createUserRecord(userId, email, firstName = '', lastName = '') {
  try {
    log('Creating user record in database:', 'info', { userId, email, firstName, lastName });
    
    const response = await fetch(`${API_BASE_URL}/users`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY
      },
      body: JSON.stringify({
        id: userId,
        email: email,
        first_name: firstName,
        last_name: lastName,
        created_at: new Date().toISOString()
      })
    });
    
    if (!response.ok) {
      const errorText = await response.text();
      log('Error creating user record:', 'error', { status: response.status, error: errorText });
      return { success: false, error: errorText };
    }
    
    log('User record created successfully', 'info');
    return { success: true };
  } catch (err) {
    log('Exception creating user record:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Sign up a new user with email and password
 * @param {string} email - User's email
 * @param {string} password - User's password
 * @param {string} firstName - User's first name
 * @param {string} lastName - User's last name
 * @returns {Promise} - The result of the sign up operation
 */
async function signUp(email, password, firstName = '', lastName = '') {
  try {
    log('======= SIGNUP PROCESS STARTED =======', 'info');
    log('Signing up user with email:', 'info', email);
    log('Password length:', 'info', password ? password.length : 0);
    log('First name provided:', 'info', !!firstName);
    log('Last name provided:', 'info', !!lastName);
    
    // Capitalize first letter of first and last name
    if (firstName) {
      firstName = firstName.charAt(0).toUpperCase() + firstName.slice(1);
    }
    
    if (lastName) {
      lastName = lastName.charAt(0).toUpperCase() + lastName.slice(1);
    }
    
    // First, check if the email is already in use
    log('Checking if email is already in use', 'info');
    const emailCheck = await checkEmailExists(email);
    
    if (emailCheck.success && emailCheck.exists) {
      log('Email already in use:', 'error', email);
      return { success: false, error: 'Email address is already in use', code: 'email_in_use' };
    }
    
    // If email check failed, log but continue (we'll let Supabase handle it)
    if (!emailCheck.success) {
      log('Email existence check failed, proceeding with signup anyway:', 'warning', emailCheck.error);
    }
    
    // Prepare signup request
    const supabaseUrl = API_BASE_URL.replace('/rest/v1', '');
    const requestBody = JSON.stringify({
      email,
      password
    });
    
    log('Signup request details:', 'info', {
      url: `${supabaseUrl}/auth/v1/signup`,
      method: 'POST'
    });
    
    const response = await fetch(`${supabaseUrl}/auth/v1/signup`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY
      },
      body: requestBody
    });
    
    log('Signup response status:', 'info', response.status);
    
    const responseText = await response.text();
    log('Signup raw response text:', 'info', responseText);
    
    let data;
    try {
      data = JSON.parse(responseText);
      log('Parsed response data:', 'info', data);
    } catch (e) {
      log('Error parsing signup response:', 'error', e);
      return { success: false, error: 'Failed to parse response: ' + responseText };
    }
    
    if (!response.ok) {
      log('Error signing up - HTTP status:', 'error', response.status);
      log('Error response data:', 'error', data);
      return { success: false, error: data.error || data.msg || 'Failed to sign up' };
    }
    
    log('Auth signup successful. Full response data:', 'info', data);

    // Normalize the user data
    const userData = normalizeUserData(data);
    log('Normalized user data:', 'info', userData);

    if (userData) {
      log('User ID from signup:', 'info', userData.id);
      log('User email from signup:', 'info', userData.email);
    } else {
      log('CRITICAL ERROR: Could not normalize user data from response!', 'error');
      log('Response keys available:', 'error', Object.keys(data));
      return { success: false, error: 'Auth signup succeeded but no valid user data returned' };
    }
    
    // Check for confirmation URL - this indicates a need for email verification
    if (data.confirmation_url) {
      log('Confirmation URL present, suggesting email verification required:', 'info', data.confirmation_url);
    }
    
    // Check for access token in the response
    if (data.access_token) {
      log('Access token present in signup response', 'info');
    }
    
    // If we have a user ID, proceed with creating the user record
    const userId = userData.id;
    log('Attempting to create user record in database with ID:', 'info', userId);
    const userRecordResult = await createUserRecord(userId, userData.email, firstName, lastName);
    log('User record creation result:', 'info', userRecordResult);
    
    if (!userRecordResult.success) {
      log('WARNING: Auth signup succeeded but database record creation failed!', 'warning');
      log('Error creating user record:', 'error', userRecordResult.error);
      return { 
        success: true, 
        data,
        warning: 'Auth account created but database record failed. Error: ' + 
               (userRecordResult.error?.message || userRecordResult.error || 'Unknown error')
      };
    }
    
    log('======= SIGNUP PROCESS COMPLETED SUCCESSFULLY =======', 'info');
    return { success: true, data };
  } catch (err) {
    log('Exception during sign up process:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Request password reset for a user
 * @param {string} email - User's email address
 * @returns {Promise} - The result of the password reset request
 */
async function requestPasswordReset(email) {
  try {
    log('Requesting password reset for email:', 'info', email);
    
    const supabaseUrl = API_BASE_URL.replace('/rest/v1', '');
    const response = await fetch(`${supabaseUrl}/auth/v1/recover`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY
      },
      body: JSON.stringify({ email })
    });
    
    log('Password reset request response status:', 'info', response.status);
    
    if (!response.ok) {
      let errorData;
      try {
        errorData = await response.json();
      } catch (e) {
        errorData = { error: 'Unknown error occurred' };
      }
      log('Error requesting password reset:', 'error', errorData);
      return { success: false, error: errorData.error || 'Failed to request password reset' };
    }
    
    log('Password reset email sent successfully', 'info');
    return { success: true };
  } catch (err) {
    log('Exception during password reset request:', 'error', err);
    return { success: false, error: err.message };
  }
}

// Team Management Functions

/**
 * Request to join a team using access code
 */
async function requestToJoinTeam(accessCode) {
  try {
    log('Requesting to join team with access code:', 'info', { accessCode });
    
    const token = await getAuthToken();
    if (!token) {
      log('No authentication token found', 'error');
      return { success: false, error: 'Authentication required' };
    }
    
    const userId = await getUserId();
    if (!userId) {
      log('No user ID found', 'error');
      return { success: false, error: 'User ID not found' };
    }
    
    log('Making join request with:', 'info', { 
      accessCode: accessCode.toUpperCase(), 
      userId,
      url: `${API_BASE_URL}/rpc/request_to_join_team`
    });
    
    const response = await fetch(`${API_BASE_URL}/rpc/request_to_join_team`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        p_access_code: accessCode.toUpperCase(),
        p_user_id: userId
      })
    });
    
    log('Response status:', 'info', { 
      status: response.status, 
      statusText: response.statusText,
      ok: response.ok 
    });
    
    const result = await response.json();
    log('Response body:', 'info', result);
    
    if (!response.ok) {
      log('HTTP error in join request:', 'error', { 
        status: response.status, 
        statusText: response.statusText,
        result 
      });
      return { success: false, error: result.message || 'Failed to request team join' };
    }
    
    // Check if the RPC function returned an error
    if (result && result.success === false) {
      log('RPC function returned error:', 'error', result);
      return { success: false, error: result.error || 'Join request failed' };
    }
    
    log('Team join request submitted successfully', 'info', result);
    return { 
      success: true, 
      message: 'Join request submitted successfully!',
      data: result
    };
    
  } catch (err) {
    log('Exception during team join request:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Create a new team
 */
async function createTeam(teamName, options = {}) {
  try {
    log('Creating new team:', 'info', { teamName, options });
    
    const token = await getAuthToken();
    if (!token) {
      return { success: false, error: 'Authentication required' };
    }
    
    const userId = await getUserId();
    if (!userId) {
      return { success: false, error: 'User ID not found' };
    }
    
    const response = await fetch(`${API_BASE_URL}/rpc/create_team`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        p_team_name: teamName,
        p_organization: options.organization || null,
        p_invoice_email: options.invoiceEmail || null,
        p_billing_street: options.billingStreet || null,
        p_billing_city: options.billingCity || null,
        p_billing_region: options.billingRegion || null,
        p_billing_country: options.billingCountry || null,
        p_admin_user_id: userId
      })
    });
    
    const result = await response.json();
    
    if (!response.ok) {
      log('Error creating team:', 'error', result);
      return { success: false, error: result.message || 'Failed to create team' };
    }
    
    // Update local storage with new team ID
    if (result.team_id) {
      localStorage.setItem('currentTeamId', result.team_id);
      localStorage.setItem('currentTeamName', teamName);
      log('Team created and stored locally:', 'info', { teamId: result.team_id, teamName });
    }
    
    log('Team created successfully', 'info', result);
    return { 
      success: true, 
      message: `Team "${teamName}" created successfully!`,
      data: { teamId: result.team_id, ...result }
    };
    
  } catch (err) {
    log('Exception during team creation:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Check join request status
 */
async function checkJoinRequestStatus() {
  try {
    log('Checking join request status', 'info');
    
    const token = await getAuthToken();
    if (!token) {
      return { success: false, error: 'Authentication required' };
    }
    
    const userId = await getUserId();
    if (!userId) {
      return { success: false, error: 'User ID not found' };
    }
    
    log('Checking join requests for user:', 'info', { userId });
    
    const response = await fetch(`${API_BASE_URL}/join_requests?user_id=eq.${userId}&order=request_date.desc&limit=1`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      }
    });
    
    log('Join request status response:', 'info', { 
      status: response.status, 
      statusText: response.statusText,
      ok: response.ok 
    });
    
    const requests = await response.json();
    log('Join requests found:', 'info', requests);
    
    if (!response.ok) {
      log('Error checking join request status:', 'error', requests);
      return { success: false, error: 'Failed to check request status' };
    }
    
    if (!requests || requests.length === 0) {
      log('No join requests found for user', 'info');
      return { success: true, data: { status: 'none' } };
    }
    
    const latestRequest = requests[0];
    log('Latest join request:', 'info', latestRequest);
    
    if (latestRequest.status === 'approved') {
      // Request was approved, update local storage and return success
      if (latestRequest.team_id) {
        localStorage.setItem('currentTeamId', latestRequest.team_id);
        log('Join request approved, team ID stored:', 'info', latestRequest.team_id);
      }
      
      return { 
        success: true, 
        data: { 
          approved: true, 
          status: 'approved',
          teamId: latestRequest.team_id 
        } 
      };
    } else if (latestRequest.status === 'rejected') {
      log('Join request was rejected', 'info', latestRequest);
      return { 
        success: true, 
        data: { 
          approved: false, 
          status: 'rejected' 
        } 
      };
    } else {
      log('Join request is pending', 'info', latestRequest);
      return { 
        success: true, 
        data: { 
          approved: null, 
          status: 'pending' 
        } 
      };
    }
    
  } catch (err) {
    log('Exception during join request status check:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Get the current user's pending join requests for display in the UI
 * @returns {Promise<object>} - User's join request information
 */
async function getUserPendingJoinRequest() {
  try {
    log('Getting user pending join request for UI display', 'info');
    
    const token = await getAuthToken();
    if (!token) {
      return { success: false, error: 'Authentication required' };
    }
    
    const userId = await getUserId();
    if (!userId) {
      return { success: false, error: 'User ID not found' };
    }
    
    // Get pending join requests for this user
    const response = await fetch(`${API_BASE_URL}/join_requests?user_id=eq.${userId}&status=eq.pending&order=request_date.desc&limit=1`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      }
    });
    
    const requests = await response.json();
    
    if (!response.ok) {
      log('Error getting user pending join request:', 'error', requests);
      return { success: false, error: 'Failed to get pending request' };
    }
    
    if (!requests || requests.length === 0) {
      log('No pending join requests found for user', 'info');
      return { success: true, data: null };
    }
    
    const pendingRequest = requests[0];
    
    log('Found pending join request:', 'info', pendingRequest);
    
    return {
      success: true,
      data: {
        id: pendingRequest.id,
        teamId: pendingRequest.team_id,
        status: pendingRequest.status,
        requestDate: pendingRequest.request_date
      }
    };
    
  } catch (err) {
    log('Exception getting user pending join request:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Get pending join requests for a team (admin only)
 */
async function getTeamJoinRequests(teamId) {
  try {
    log('=== GET TEAM JOIN REQUESTS START ===', 'info', { teamId });
    
    if (!teamId) {
      log('No team ID provided to getTeamJoinRequests', 'error');
      return { success: false, error: 'Team ID is required' };
    }
    
    const token = await getAuthToken();
    if (!token) {
      log('No auth token available', 'error');
      return { success: false, error: 'Authentication required' };
    }
    
    const userId = await getUserId();
    if (!userId) {
      log('No user ID available', 'error');
      return { success: false, error: 'User ID not found' };
    }
    
    log('Making RPC call to get_team_join_requests', 'info', { 
      teamId, 
      userId,
      hasToken: !!token 
    });
    
    const response = await fetch(`${API_BASE_URL}/rpc/get_team_join_requests`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        p_team_id: teamId,
        p_admin_user_id: userId
      })
    });
    
    log('RPC response received:', 'info', { 
      status: response.status, 
      statusText: response.statusText,
      ok: response.ok 
    });
    
    const result = await response.json();
    log('RPC response body:', 'info', result);
    
    if (!response.ok) {
      log('HTTP error getting team join requests:', 'error', { 
        status: response.status, 
        statusText: response.statusText,
        result 
      });
      return { success: false, error: result.message || 'Failed to get join requests' };
    }
    
    if (result && result.success === false) {
      log('RPC function returned error:', 'error', result);
      return { success: false, error: result.error || 'Failed to get join requests' };
    }
    
    log('Team join requests retrieved successfully', 'info', result);
    log('Extracting requests from result:', 'info', { 
      hasRequests: !!result.requests,
      requestsType: typeof result.requests,
      requestsLength: result.requests ? result.requests.length : 'N/A',
      requests: result.requests 
    });
    
    return { 
      success: true, 
      data: result.requests || []
    };
    
  } catch (err) {
    log('Exception getting team join requests:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Approve or reject a join request
 */
async function respondToJoinRequest(requestId, approved) {
  try {
    log('Responding to join request:', 'info', { requestId, approved });
    
    const token = await getAuthToken();
    if (!token) {
      return { success: false, error: 'Authentication required' };
    }
    
    const userId = await getUserId();
    if (!userId) {
      return { success: false, error: 'User ID not found' };
    }
    
    const response = await fetch(`${API_BASE_URL}/rpc/respond_to_join_request`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        p_request_id: requestId,
        p_admin_user_id: userId,
        p_approved: approved
      })
    });
    
    log('Join request response:', 'info', { 
      status: response.status, 
      statusText: response.statusText,
      ok: response.ok 
    });
    
    const result = await response.json();
    log('Join request response result:', 'info', result);
    
    if (!response.ok) {
      log('HTTP error responding to join request:', 'error', { 
        status: response.status, 
        statusText: response.statusText,
        result 
      });
      return { success: false, error: result.message || 'Failed to respond to join request' };
    }
    
    if (result && result.success === false) {
      log('RPC function returned error:', 'error', result);
      return { success: false, error: result.error || 'Failed to respond to join request' };
    }
    
    log('Join request response submitted successfully', 'info', result);
    return { 
      success: true, 
      message: result.message || (approved ? 'Request approved' : 'Request rejected'),
      data: result
    };
    
  } catch (err) {
    log('Exception responding to join request:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Remove a team member (admin only)
 */
async function removeTeamMember(memberId) {
  try {
    log('Removing team member with ID:', 'info', { memberId });
    
    const token = await getAuthToken();
    if (!token) {
      return { success: false, error: 'Authentication required' };
    }
    
    const userId = await getUserId();
    if (!userId) {
      return { success: false, error: 'User ID not found' };
    }
    
    const currentTeamId = localStorage.getItem('currentTeamId');
    if (!currentTeamId) {
      return { success: false, error: 'No team ID found' };
    }
    
    // First, check if the current user is an admin
    const currentUserResponse = await fetch(`${API_BASE_URL}/users?id=eq.${userId}&select=role,team_id`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      }
    });
    
    if (!currentUserResponse.ok) {
      const error = await currentUserResponse.json();
      log('Error checking current user role:', 'error', error);
      return { success: false, error: 'Failed to verify admin status' };
    }
    
    const currentUserData = await currentUserResponse.json();
    
    if (currentUserData.length === 0) {
      log('Current user not found', 'error');
      return { success: false, error: 'User not found' };
    }
    
    const currentUser = currentUserData[0];
    
    // Verify the user is an admin
    if (currentUser.role !== 'admin') {
      log('Current user is not an admin, cannot remove team members', 'error');
      return { success: false, error: 'Admin privileges required to remove team members' };
    }
    
    // Verify the user to be removed exists and is part of the same team
    const memberResponse = await fetch(`${API_BASE_URL}/users?id=eq.${memberId}&select=email,role,team_id`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      }
    });
    
    if (!memberResponse.ok) {
      const error = await memberResponse.json();
      log('Error checking member:', 'error', error);
      return { success: false, error: 'Failed to verify member status' };
    }
    
    const memberData = await memberResponse.json();
    
    if (memberData.length === 0) {
      log('Member not found', 'error');
      return { success: false, error: 'Member not found' };
    }
    
    const member = memberData[0];
    
    // Prevent removing another admin
    if (member.role === 'admin') {
      log('Cannot remove an admin user', 'error');
      return { success: false, error: 'Cannot remove another admin. Transfer admin rights first.' };
    }
    
    // Make sure the member is part of the same team
    if (member.team_id !== currentTeamId) {
      log('Member is not part of the same team', 'error');
      return { success: false, error: 'Member is not part of your team' };
    }
    
    // Update the user to remove the team association and reset role
    const updateResponse = await fetch(`${API_BASE_URL}/users?id=eq.${memberId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Prefer': 'return=representation',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        team_id: null,
        role: 'member'
      })
    });
    
    if (!updateResponse.ok) {
      const error = await updateResponse.json();
      log('Error removing team member:', 'error', error);
      return { success: false, error: 'Failed to remove team member' };
    }
    
    log('Team member removed successfully', 'info');
    return { success: true };
  } catch (err) {
    log('Exception when removing team member:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Leave a team (member only)
 */
async function leaveTeam() {
  try {
    log('Leaving team', 'info');
    
    const token = await getAuthToken();
    if (!token) {
      return { success: false, error: 'Authentication required' };
    }
    
    const userId = await getUserId();
    if (!userId) {
      return { success: false, error: 'User ID not found' };
    }
    
    // Check if the current user is an admin
    const currentUserResponse = await fetch(`${API_BASE_URL}/users?id=eq.${userId}&select=role,team_id`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      }
    });
    
    if (!currentUserResponse.ok) {
      const error = await currentUserResponse.json();
      log('Error checking current user role:', 'error', error);
      return { success: false, error: 'Failed to verify user status' };
    }
    
    const currentUserData = await currentUserResponse.json();
    
    if (currentUserData.length === 0) {
      log('Current user not found', 'error');
      return { success: false, error: 'User not found' };
    }
    
    const currentUser = currentUserData[0];
    
    // Prevent admin from leaving if they are the only admin
    if (currentUser.role === 'admin') {
      // Check if there are other admins
      const adminResponse = await fetch(`${API_BASE_URL}/users?team_id=eq.${currentUser.team_id}&role=eq.admin&select=id`, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_ANON_KEY,
          'Authorization': `Bearer ${token}`
        }
      });
      
      if (!adminResponse.ok) {
        const error = await adminResponse.json();
        log('Error checking other admins:', 'error', error);
        return { success: false, error: 'Failed to verify admin status' };
      }
      
      const adminData = await adminResponse.json();
      
      if (adminData.length <= 1) {
        log('Cannot leave team as the only admin', 'error');
        return { success: false, error: 'Cannot leave team as the only admin. Transfer admin rights first or delete the team.' };
      }
    }
    
    // Update the user to remove the team association and reset role
    const updateResponse = await fetch(`${API_BASE_URL}/users?id=eq.${userId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Prefer': 'return=representation',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        team_id: null,
        role: 'member'
      })
    });
    
    if (!updateResponse.ok) {
      const error = await updateResponse.json();
      log('Error leaving team:', 'error', error);
      return { success: false, error: 'Failed to leave team' };
    }
    
    // Clear team-related data from localStorage
    localStorage.removeItem('currentTeamId');
    localStorage.removeItem('currentTeamName');
    localStorage.removeItem('teamInfo');
    
    log('Left team successfully and local storage cleared', 'info');
    return { success: true };
  } catch (err) {
    log('Exception when leaving team:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Transfer admin role to another team member (admin only)
 */
async function transferAdminRole(newAdminId) {
  try {
    log('Transferring admin role to user:', 'info', { newAdminId });
    
    const token = await getAuthToken();
    if (!token) {
      return { success: false, error: 'Authentication required' };
    }
    
    const userId = await getUserId();
    if (!userId) {
      return { success: false, error: 'User ID not found' };
    }
    
    // Check if the current user is an admin
    const currentUserResponse = await fetch(`${API_BASE_URL}/users?id=eq.${userId}&select=role,team_id`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      }
    });
    
    if (!currentUserResponse.ok) {
      const error = await currentUserResponse.json();
      log('Error checking current user role:', 'error', error);
      return { success: false, error: 'Failed to verify admin status' };
    }
    
    const currentUserData = await currentUserResponse.json();
    
    if (currentUserData.length === 0) {
      log('Current user not found', 'error');
      return { success: false, error: 'User not found' };
    }
    
    const currentUser = currentUserData[0];
    
    // Verify the user is an admin
    if (currentUser.role !== 'admin') {
      log('Current user is not an admin, cannot transfer admin rights', 'error');
      return { success: false, error: 'Admin privileges required to transfer admin rights' };
    }
    
    // Verify the new admin exists and is part of the same team
    const newAdminResponse = await fetch(`${API_BASE_URL}/users?id=eq.${newAdminId}&select=email,role,team_id`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      }
    });
    
    if (!newAdminResponse.ok) {
      const error = await newAdminResponse.json();
      log('Error checking new admin:', 'error', error);
      return { success: false, error: 'Failed to verify new admin status' };
    }
    
    const newAdminData = await newAdminResponse.json();
    
    if (newAdminData.length === 0) {
      log('New admin not found', 'error');
      return { success: false, error: 'New admin not found' };
    }
    
    const newAdmin = newAdminData[0];
    
    // Make sure the new admin is part of the same team
    if (newAdmin.team_id !== currentUser.team_id) {
      log('New admin is not part of the same team', 'error');
      return { success: false, error: 'New admin is not part of your team' };
    }
    
    // Update the new admin's role
    const updateNewAdminResponse = await fetch(`${API_BASE_URL}/users?id=eq.${newAdminId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Prefer': 'return=representation',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        role: 'admin'
      })
    });
    
    if (!updateNewAdminResponse.ok) {
      const error = await updateNewAdminResponse.json();
      log('Error updating new admin role:', 'error', error);
      return { success: false, error: 'Failed to update new admin role' };
    }
    
    // Update the current admin's role to member
    const updateCurrentUserResponse = await fetch(`${API_BASE_URL}/users?id=eq.${userId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Prefer': 'return=representation',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        role: 'member'
      })
    });
    
    if (!updateCurrentUserResponse.ok) {
      const error = await updateCurrentUserResponse.json();
      log('Error updating current user role:', 'error', error);
      return { success: false, error: 'Failed to update current user role' };
    }
    
    log('Admin role transferred successfully', 'info');
    return { success: true };
  } catch (err) {
    log('Exception when transferring admin role:', 'error', err);
    return { success: false, error: err.message };
  }
}

/**
 * Delete a team (admin only)
 */
async function deleteTeam() {
  try {
    log('Deleting team', 'info');
    
    const token = await getAuthToken();
    if (!token) {
      return { success: false, error: 'Authentication required' };
    }
    
    const userId = await getUserId();
    if (!userId) {
      return { success: false, error: 'User ID not found' };
    }
    
    // Check if the current user is an admin
    const currentUserResponse = await fetch(`${API_BASE_URL}/users?id=eq.${userId}&select=role,team_id`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      }
    });
    
    if (!currentUserResponse.ok) {
      const error = await currentUserResponse.json();
      log('Error checking current user role:', 'error', error);
      return { success: false, error: 'Failed to verify admin status' };
    }
    
    const currentUserData = await currentUserResponse.json();
    
    if (currentUserData.length === 0) {
      log('Current user not found', 'error');
      return { success: false, error: 'User not found' };
    }
    
    const currentUser = currentUserData[0];
    
    // Verify the user is an admin
    if (currentUser.role !== 'admin') {
      log('Current user is not an admin, cannot delete team', 'error');
      return { success: false, error: 'Admin privileges required to delete team' };
    }
    
    const teamId = currentUser.team_id;
    if (!teamId) {
      log('No team ID found', 'error');
      return { success: false, error: 'No team found' };
    }
    
    // First, remove all team members from the team and reset their roles
    const updateMembersResponse = await fetch(`${API_BASE_URL}/users?team_id=eq.${teamId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Prefer': 'return=representation',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        team_id: null,
        role: 'member'
      })
    });
    
    if (!updateMembersResponse.ok) {
      const error = await updateMembersResponse.json();
      log('Error removing team members:', 'error', error);
      return { success: false, error: 'Failed to remove team members' };
    }
    
    // Delete the team
    const deleteTeamResponse = await fetch(`${API_BASE_URL}/teams?id=eq.${teamId}`, {
      method: 'DELETE',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_ANON_KEY,
        'Authorization': `Bearer ${token}`
      }
    });
    
    if (!deleteTeamResponse.ok) {
      const error = await deleteTeamResponse.json();
      log('Error deleting team:', 'error', error);
      return { success: false, error: 'Failed to delete team' };
    }
    
    // Clear team-related data from localStorage
    localStorage.removeItem('currentTeamId');
    localStorage.removeItem('currentTeamName');
    localStorage.removeItem('teamInfo');
    
    log('Team deleted successfully and local storage cleared', 'info');
    return { success: true };
  } catch (err) {
    log('Exception when deleting team:', 'error', err);
    return { success: false, error: err.message };
  }
}