/**
 * Auth Service Utility
 * This file contains functions for authentication and user management with Supabase
 */

import { SUPABASE_URL, SUPABASE_KEY } from './supabase-config.js';

// Export Supabase config for other modules that import from auth-service
export { SUPABASE_URL, SUPABASE_KEY };

/**
 * Check if an email address is already in use
 * @param {string} email - The email address to check
 * @returns {Promise} - Whether the email is already in use
 */
async function checkEmailExists(email) {
  try {
    console.log('Checking if email already exists:', email);
    
    // Try to search for users with this email using the REST API instead
    // This is an alternative approach since the /auth/v1/user-exists endpoint is not available
    const response = await fetch(`${SUPABASE_URL}/rest/v1/users?email=eq.${encodeURIComponent(email)}&select=id`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY
      }
    });
    
    if (!response.ok) {
      console.error('Error checking if email exists - status:', response.status);
      return { success: false, error: 'Failed to check email', exists: false };
    }
    
    const users = await response.json();
    console.log('Email exists check response:', users);
    
    // If we find any users with this email, it exists
    const exists = Array.isArray(users) && users.length > 0;
    
    return { 
      success: true, 
      exists: exists,
      data: users
    };
  } catch (err) {
    console.error('Exception when checking if email exists:', err);
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
 * Sign up a new user with email and password
 * @param {string} email - User's email
 * @param {string} password - User's password
 * @param {string} firstName - User's first name
 * @param {string} lastName - User's last name
 * @returns {Promise} - The result of the sign up operation
 */
export async function signUp(email, password, firstName = '', lastName = '') {
  try {
    console.log('======= SIGNUP PROCESS STARTED =======');
    console.log('Signing up user with email:', email);
    console.log('Password length:', password ? password.length : 0);
    console.log('First name provided:', !!firstName);
    console.log('Last name provided:', !!lastName);
    
    // Capitalize first letter of first and last name
    if (firstName) {
      firstName = firstName.charAt(0).toUpperCase() + firstName.slice(1);
    }
    
    if (lastName) {
      lastName = lastName.charAt(0).toUpperCase() + lastName.slice(1);
    }
    
    // First, check if the email is already in use
    console.log('Checking if email is already in use');
    const emailCheck = await checkEmailExists(email);
    
    if (emailCheck.success && emailCheck.exists) {
      console.error('Email already in use:', email);
      return { success: false, error: 'Email address is already in use', code: 'email_in_use' };
    }
    
    // If email check failed, log but continue (we'll let Supabase handle it)
    if (!emailCheck.success) {
      console.warn('Email existence check failed, proceeding with signup anyway:', emailCheck.error);
    }
    
    // Log complete request details
    const requestBody = JSON.stringify({
      email,
      password,
      // Add a redirect URL that points back to the extension
      options: {
        redirect_to: chrome.runtime.getURL('popup/popup.html')
      }
    });
    
    console.log('Signup request details:');
    console.log('URL:', `${SUPABASE_URL}/auth/v1/signup`);
    console.log('Method: POST');
    console.log('Headers:', {
      'Content-Type': 'application/json',
      'apikey': SUPABASE_KEY ? 'Present (masked for security)' : 'Missing!'
    });
    console.log('Request body:', requestBody);
    
    const response = await fetch(`${SUPABASE_URL}/auth/v1/signup`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY
      },
      body: requestBody
    });
    
    console.log('Signup response status:', response.status);
    console.log('Response status text:', response.statusText);
    console.log('Response headers:', Object.fromEntries([...response.headers.entries()]));
    
    const responseText = await response.text();
    console.log('Signup raw response text:', responseText);
    
    let data;
    try {
      data = JSON.parse(responseText);
      console.log('Parsed response data:', JSON.stringify(data, null, 2));
    } catch (e) {
      console.error('Error parsing signup response:', e);
      return { success: false, error: 'Failed to parse response: ' + responseText };
    }
    
    if (!response.ok) {
      console.error('Error signing up - HTTP status:', response.status);
      console.error('Error response data:', data);
      return { success: false, error: data.error || data.msg || 'Failed to sign up' };
    }
    
    console.log('Auth signup successful. Full response data:', JSON.stringify(data, null, 2));

    // Normalize the user data
    const userData = normalizeUserData(data);
    console.log('Normalized user data:', userData);

    if (userData) {
      console.log('User ID from signup:', userData.id);
      console.log('User email from signup:', userData.email);
    } else {
      console.error('CRITICAL ERROR: Could not normalize user data from response!');
      console.error('Response keys available:', Object.keys(data));
      return { success: false, error: 'Auth signup succeeded but no valid user data returned' };
    }
    
    // Check for confirmation URL - this indicates a need for email verification
    if (data.confirmation_url) {
      console.log('Confirmation URL present, suggesting email verification required:', data.confirmation_url);
    }
    
    // Check for email confirmation in Supabase response
    if (userData.email) {
      console.log('Email address in response:', userData.email);
    }
    
    // Check for access token in the response
    if (data.access_token) {
      console.log('Access token present in signup response: (masked)');
    }
    
    // If we have a user ID, proceed with creating the user record
    const userId = userData.id;
    console.log('Attempting to create user record in database with ID:', userId);
    const userRecordResult = await createUserRecord(userId, userData.email, firstName, lastName);
    console.log('User record creation result:', userRecordResult);
    
    if (!userRecordResult.success) {
      console.error('WARNING: Auth signup succeeded but database record creation failed!');
      console.error('Error creating user record:', userRecordResult.error);
      return { 
        success: true, 
        data,
        warning: 'Auth account created but database record failed. Error: ' + 
               (userRecordResult.error?.message || userRecordResult.error || 'Unknown error')
      };
    }
    
    console.log('======= SIGNUP PROCESS COMPLETED SUCCESSFULLY =======');
    return { success: true, data };
  } catch (err) {
    console.error('Exception during sign up process:', err);
    console.error('Error details:', err.stack);
    
    // LAST RESORT FALLBACK: If we have a complete failure in the signup process,
    // but we have credentials, we can try to sign in directly
    console.log('ATTEMPTING EMERGENCY FALLBACK: Direct sign-in after signup failure');
    try {
      // Wait briefly before attempting sign-in
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      // Try to sign in with the provided credentials
      const fallbackSignInResult = await signIn(email, password);
      console.log('Emergency fallback sign-in result:', fallbackSignInResult);
      
      if (fallbackSignInResult.success) {
        console.log('Emergency fallback sign-in successful!');
        return { 
          success: true, 
          data: fallbackSignInResult.data,
          warning: 'Signup process failed but emergency sign-in succeeded. Original error: ' + err.message 
        };
      }
    } catch (fallbackErr) {
      console.error('Emergency fallback sign-in also failed:', fallbackErr);
    }
    
    // If we get here, all recovery attempts have failed
    return { success: false, error: err.message };
  }
}

/**
 * Sign in an existing user with email and password
 * @param {string} email - User's email
 * @param {string} password - User's password
 * @returns {Promise} - The result of the sign in operation
 */
export async function signIn(email, password) {
  try {
    console.log('Signing in user with email:', email);
    
    // Set up login request
    const loginResponse = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY
      },
      body: JSON.stringify({
        email,
        password
      })
    });
    
    const loginData = await loginResponse.text();
    console.log('Login response status:', loginResponse.status);
    
    if (!loginResponse.ok) {
      let errorMessage = 'Login failed';
      
      try {
        const errorData = JSON.parse(loginData);
        console.error('Login error details:', errorData);
        errorMessage = errorData.error_description || errorData.error || 'Authentication failed';
    } catch (e) {
        console.error('Could not parse error response:', loginData);
      }
      
      return { success: false, error: errorMessage };
    }
    
    // Parse the response data
    let data;
    try {
      data = JSON.parse(loginData);
    } catch (e) {
      console.error('Error parsing login response:', e);
      return { success: false, error: 'Invalid response from server' };
    }
    
    // Get the access token and user ID
    const accessToken = data.access_token;
    const refreshToken = data.refresh_token; 
    const userId = data.user.id;
    
    // Store the token and user ID
    console.log('Storing login data in Chrome storage');
    await chrome.storage.sync.set({
      accessToken,
      refreshToken,
      userId,
      userEmail: email // Store the user's email for reconnection
    });
    
    // Normalize the user data
    const normalizedUserData = normalizeUserData(data.user);
    
    // Get team information if available
    let userDetailsResult = null;
    
    try {
      console.log('Retrieving user details after login to get team information');
      userDetailsResult = await getUserDetails(userId, accessToken);
      
      if (userDetailsResult.success) {
        console.log('Retrieved user details after login');
        
        // Check if user has a team_id in the user details
        if (userDetailsResult.data && userDetailsResult.data.team_id) {
          console.log('User has a team assigned:', userDetailsResult.data.team_id);
          
          // Note: team ID will be stored by getUserDetails function directly
          console.log('Team ID has been stored in Chrome storage');
    } else {
          console.log('User does not have a team assigned, clearing any existing team data');
          await chrome.storage.sync.remove(['currentTeamId', 'userRole']);
        }
      } else {
        console.error('Failed to retrieve user details after login:', userDetailsResult.error);
      }
    } catch (e) {
      console.error('Error getting user details after login:', e);
    }
    
      return { 
        success: true, 
      user: normalizedUserData,
      userDetails: userDetailsResult?.success ? userDetailsResult.data : null
      };
  } catch (err) {
    console.error('Exception during login:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Sign out the current user
 * @returns {Promise} - The result of the sign out operation
 */
export async function signOut() {
  try {
    // Get the current access token
    const { accessToken } = await chrome.storage.sync.get(['accessToken']);
    
    if (!accessToken) {
      console.log('No user is signed in');
      return { success: true };
    }
    
    const response = await fetch(`${SUPABASE_URL}/auth/v1/logout`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    // Clear storage regardless of response
    chrome.storage.sync.remove([
      'accessToken', 
      'refreshToken', 
      'userId', 
      'currentTeamId', 
      'userRole',
      'userEmail'
    ]);
    
    if (!response.ok) {
      const data = await response.json();
      console.error('Error signing out:', data);
      return { success: false, error: data.error || 'Failed to sign out properly' };
    }
    
    console.log('User signed out successfully');
    return { success: true };
  } catch (err) {
    console.error('Exception during sign out:', err);
    // Still clear storage even if there was an error
    chrome.storage.sync.remove([
      'accessToken', 
      'refreshToken', 
      'userId', 
      'currentTeamId', 
      'userRole',
      'userEmail'
    ]);
    return { success: false, error: err.message };
  }
}

/**
 * Generate a random access code for a team
 * @returns {string} - An 8-character alphanumeric access code
 */
function generateAccessCode() {
  // Use characters that are easy to distinguish to avoid confusion
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  
  // Generate 8 random characters
  for (let i = 0; i < 8; i++) {
    const randomIndex = Math.floor(Math.random() * chars.length);
    code += chars[randomIndex];
  }
  
  return code;
}

/**
 * Create a new team
 * @param {string} teamName - The name of the team to create
 * @param {object} teamDetails - Additional team details
 * @param {string} teamDetails.organization - Organization name
 * @param {string} teamDetails.invoiceEmail - Email for receiving invoices
 * @param {string} teamDetails.billingStreet - Street address for billing
 * @param {string} teamDetails.billingCity - City for billing
 * @param {string} teamDetails.billingRegion - State/Province/Region for billing
 * @param {string} teamDetails.billingCountry - Country for billing
 * @returns {Promise} - The result of the create team operation
 */
export async function createTeam(teamName, teamDetails = {}) {
  try {
    console.log('Creating team with name:', teamName);
    console.log('Team details:', teamDetails);
    
    const { accessToken, userId } = await chrome.storage.sync.get(['accessToken', 'userId']);
    console.log('User ID for team creation:', userId);
    
    if (!accessToken || !userId) {
      console.error('User not authenticated for team creation');
      return { success: false, error: 'User is not authenticated' };
    }
    
    // Capitalize first letter of organization if provided
    let organization = teamDetails.organization || '';
    if (organization) {
      organization = organization.charAt(0).toUpperCase() + organization.slice(1);
    }
    
    // Generate a unique access code for the team
    const accessCode = generateAccessCode();
    console.log('Generated access code for team:', accessCode);
    
    // Create the team with all the new fields
    console.log('Sending team creation request');
    const response = await fetch(`${SUPABASE_URL}/rest/v1/teams`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`,
        'Prefer': 'return=representation'
      },
      body: JSON.stringify({
        name: teamName,
        organization: organization,
        invoice_email: teamDetails.invoiceEmail,
        billing_street: teamDetails.billingStreet,
        billing_city: teamDetails.billingCity,
        billing_region: teamDetails.billingRegion,
        billing_country: teamDetails.billingCountry,
        access_code: accessCode
      })
    });
    
    const responseText = await response.text();
    console.log('Team creation response:', responseText);
    
    if (!response.ok) {
      let error;
      try {
        error = JSON.parse(responseText);
      } catch (e) {
        error = { message: responseText || 'Unknown error' };
      }
      console.error('Error creating team:', error);
      return { success: false, error };
    }
    
    let teamData;
    try {
      teamData = JSON.parse(responseText);
    } catch (e) {
      console.error('Error parsing team data response:', e);
      return { success: false, error: 'Failed to parse team response data' };
    }
    
    const teamId = teamData[0].id;
    console.log('Team created with ID:', teamId);
    
    // Update the user record to link to the team and set as admin
    console.log('Updating user record to link to team and set role as admin');
    const userUpdateResponse = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${userId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`,
        'Prefer': 'return=representation'
      },
      body: JSON.stringify({
        team_id: teamId,
        role: 'admin'
      })
    });
    
    const userUpdateText = await userUpdateResponse.text();
    console.log('User update response:', userUpdateText);
    
    if (!userUpdateResponse.ok) {
      let error;
      try {
        error = JSON.parse(userUpdateText);
      } catch (e) {
        error = { message: userUpdateText || 'Unknown error' };
      }
      console.error('Error updating user:', error);
      return { success: false, error };
    }
    
    // Store the team ID and role in storage
    console.log('Storing team ID and role in Chrome storage');
    chrome.storage.sync.set({
      'currentTeamId': teamId,
      'userRole': 'admin'
    });
    
    return { 
      success: true, 
      data: {
        teamId,
        role: 'admin',
        teamName,
        ...teamData[0]
      }
    };
  } catch (err) {
    console.error('Exception when creating team:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Join an existing team
 * @param {string} teamId - The ID of the team to join
 * @returns {Promise} - The result of the join team operation
 */
export async function joinTeam(teamId) {
  try {
    console.log('Joining team with ID:', teamId);
    const { accessToken, userId } = await chrome.storage.sync.get(['accessToken', 'userId']);
    console.log('User ID for team joining:', userId);
    
    if (!accessToken || !userId) {
      console.error('User not authenticated for team joining');
      return { success: false, error: 'User is not authenticated' };
    }
    
    // First check if the team exists
    console.log('Checking if team exists');
    const teamCheckResponse = await fetch(`${SUPABASE_URL}/rest/v1/teams?id=eq.${teamId}`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    const teamCheckText = await teamCheckResponse.text();
    console.log('Team check response:', teamCheckText);
    
    if (!teamCheckResponse.ok) {
      let error;
      try {
        error = JSON.parse(teamCheckText);
      } catch (e) {
        error = { message: teamCheckText || 'Unknown error' };
      }
      console.error('Error checking team:', error);
      return { success: false, error };
    }
    
    let teamData;
    try {
      teamData = JSON.parse(teamCheckText);
    } catch (e) {
      console.error('Error parsing team check response:', e);
      return { success: false, error: 'Failed to parse team response data' };
    }
    
    if (teamData.length === 0) {
      console.error('Team not found with ID:', teamId);
      return { success: false, error: 'Team not found' };
    }
    
    const teamName = teamData[0].name;
    console.log('Found team:', teamName);
    
    // Update the user record to link to the team and set as member
    console.log('Updating user record to link to team and set role as member');
    const userUpdateResponse = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${userId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`,
        'Prefer': 'return=representation'
      },
      body: JSON.stringify({
        team_id: teamId,
        role: 'member'
      })
    });
    
    const userUpdateText = await userUpdateResponse.text();
    console.log('User update response:', userUpdateText);
    
    if (!userUpdateResponse.ok) {
      let error;
      try {
        error = JSON.parse(userUpdateText);
      } catch (e) {
        error = { message: userUpdateText || 'Unknown error' };
      }
      console.error('Error updating user:', error);
      return { success: false, error };
    }
    
    // Store the team ID and role in storage
    console.log('Storing team ID and role in Chrome storage');
    chrome.storage.sync.set({
      'currentTeamId': teamId,
      'userRole': 'member'
    });
    
    return { 
      success: true, 
      data: {
        teamId,
        role: 'member',
        teamName
      }
    };
  } catch (err) {
    console.error('Exception when joining team:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Get list of all available teams
 * @returns {Promise} - List of teams
 */
export async function getAllTeams() {
  try {
    console.log('Getting all available teams');
    const { accessToken } = await chrome.storage.sync.get(['accessToken']);
    console.log('Access token for getting teams:', accessToken ? 'Available' : 'Not available');
    
    if (!accessToken) {
      console.error('No access token available to fetch teams');
      return { success: false, error: 'User is not authenticated' };
    }
    
    console.log('Fetching teams from API');
    const response = await fetch(`${SUPABASE_URL}/rest/v1/teams?select=*`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    const responseText = await response.text();
    console.log('Teams API response:', responseText);
    
    if (!response.ok) {
      let error;
      try {
        error = JSON.parse(responseText);
      } catch (e) {
        error = { message: responseText || 'Unknown error' };
      }
      console.error('Error fetching teams:', error);
      return { success: false, error };
    }
    
    let teams;
    try {
      teams = JSON.parse(responseText);
    } catch (e) {
      console.error('Error parsing teams response:', e);
      return { success: false, error: 'Failed to parse teams data' };
    }
    
    console.log(`Found ${teams.length} teams`);
    return { success: true, data: teams };
  } catch (err) {
    console.error('Exception when fetching teams:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Get details for the current user
 * @param {string} userId - The user ID to get details for
 * @param {string} accessToken - The access token for authentication
 * @returns {Promise} - User details including team information
 */
export async function getUserDetails(userId, accessToken) {
  try {
    console.log('getUserDetails called for userId:', userId);
    const token = accessToken || (await chrome.storage.sync.get(['accessToken'])).accessToken;
    
    console.log('getUserDetails: Token check:', {
      hasToken: !!token,
      tokenLength: token ? token.length : 0,
      tokenFirstChars: token ? token.substring(0, 10) + '...' : 'none'
    });
    
    // Check if token is expired by parsing it
    if (token) {
      try {
        const parts = token.split('.');
        if (parts.length === 3) {
          const payload = JSON.parse(atob(parts[1]));
          const expTime = payload.exp ? new Date(payload.exp * 1000) : null;
          const currentTime = new Date();
          const isExpired = expTime ? (expTime < currentTime) : false;
          
          console.log('getUserDetails: Token analysis:', {
            exp: expTime ? expTime.toISOString() : 'not found',
            iat: payload.iat ? new Date(payload.iat * 1000).toISOString() : 'not found',
            currentTime: currentTime.toISOString(),
            isExpired: isExpired,
            timeRemaining: expTime ? ((expTime - currentTime) / 1000 / 60).toFixed(2) + ' minutes' : 'unknown'
          });
          
          if (isExpired) {
            console.error('getUserDetails: Token is expired, cannot proceed');
            return { success: false, error: 'Access token expired' };
          }
        }
      } catch (e) {
        console.error('getUserDetails: Error analyzing token:', e);
      }
    }
    
    if (!userId) {
      console.error('getUserDetails: Missing userId');
      return { success: false, error: 'User ID missing' };
    }
    
    if (!token) {
      console.error('getUserDetails: Missing access token');
      return { success: false, error: 'Access token missing' };
    }
    
    console.log(`Fetching user details from ${SUPABASE_URL}/rest/v1/users?id=eq.${userId}`);
    const response = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${userId}`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${token}`
      }
    });
    
    console.log('getUserDetails: Response status:', response.status);
    
    if (!response.ok) {
      const responseText = await response.text();
      console.error('Error fetching user details, raw response:', responseText);
      
      try {
        const error = JSON.parse(responseText);
        console.error('Error fetching user details, parsed:', error);
        
        // If token is expired, clear it from storage
        if (error.code === 'PGRST301' || responseText.includes('JWT expired')) {
          console.error('Token is expired, clearing from storage');
          chrome.storage.sync.remove(['accessToken']);
          return { success: false, error: 'Authentication token expired' };
        }
        
      return { success: false, error };
      } catch (e) {
        console.error('Failed to parse error response:', e);
        return { success: false, error: { message: responseText || 'Unknown error' } };
      }
    }
    
    const userData = await response.json();
    console.log('User data response:', userData);
    
    if (userData.length === 0) {
      console.error('User not found in database, attempting to create user record');
      
      // Get email from storage
      const { userEmail } = await chrome.storage.sync.get(['userEmail']);
      
      if (userEmail) {
        // Try to create the user record
        const createResult = await createUserRecord(userId, userEmail);
        console.log('User record creation result:', createResult);
        
        if (createResult.success) {
          return { 
            success: true, 
            data: createResult.data[0],
            teamDetails: null
          };
        } else {
          return { success: false, error: 'User not found and failed to create user record' };
        }
      } else {
        return { success: false, error: 'User not found and email not available' };
      }
    }
    
    // Get team details if user is part of a team
    let teamData = null;
    if (userData[0].team_id) {
      console.log('User has a team, fetching team details for ID:', userData[0].team_id);
      
      // IMPORTANT: Store the team_id in Chrome storage
      // This ensures the currentTeamId is always synced with what's in the database
      console.log('Storing team ID in Chrome storage:', userData[0].team_id);
      await chrome.storage.sync.set({
        'currentTeamId': userData[0].team_id,
        'userRole': userData[0].role || 'member'
      });
      
      const teamResponse = await fetch(`${SUPABASE_URL}/rest/v1/teams?id=eq.${userData[0].team_id}`, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': `Bearer ${token}`
        }
      });
      
      if (teamResponse.ok) {
        const teams = await teamResponse.json();
        console.log('Team data response:', teams);
        // Debug log for access code
        console.log('Access code in team data:', teams.length > 0 ? teams[0].access_code : 'No team data');
        if (teams.length > 0) {
          teamData = teams[0];
        }
      } else {
        console.error('Error fetching team details:', await teamResponse.text());
      }
    } else {
      console.log('User does not have a team assigned');
      // Clear any potentially stale team ID in storage
      await chrome.storage.sync.remove(['currentTeamId', 'userRole']);
    }
    
    // Debug log for final team data with access code
    if (teamData) {
      console.log('Final team data with access code:', {
        name: teamData.name,
        access_code: teamData.access_code,
        has_access_code: !!teamData.access_code
      });
      
      // Ensure the access_code is explicitly visible in the logs for debugging
      console.log('Access code value being returned:', teamData.access_code);
    }
    
    console.log('Returning user details with team info');
    return { 
      success: true, 
      data: {
        ...userData[0],
        teamDetails: teamData
      }
    };
  } catch (err) {
    console.error('Exception when fetching user details:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Create a user record in the database
 * @param {string} userId - The user ID from Supabase Auth
 * @param {string} email - The user's email
 * @param {string} firstName - The user's first name
 * @param {string} lastName - The user's last name
 * @returns {Promise} - Result of the operation
 */
async function createUserRecord(userId, email, firstName = '', lastName = '') {
  try {
    console.log('****** CREATING USER RECORD IN DATABASE ******');
    console.log('Creating user record for ID:', userId);
    console.log('Email:', email);
    console.log('First name:', firstName || '(none provided)');
    console.log('Last name:', lastName || '(none provided)');
    console.log('Supabase URL:', SUPABASE_URL);
    
    if (!userId) {
      console.error('ERROR: No user ID provided to createUserRecord function');
      return { success: false, error: 'No user ID provided' };
    }
    
    // First check if the Supabase database is accessible
    console.log('Checking database connectivity...');
    try {
      const connectivityCheck = await fetch(`${SUPABASE_URL}/rest/v1/users?limit=1`, {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY
        }
      });
      
      console.log('Database connectivity check status:', connectivityCheck.status);
      console.log('Database connectivity check headers:', Object.fromEntries([...connectivityCheck.headers.entries()]));
      
      if (!connectivityCheck.ok) {
        console.error('Database connectivity check failed:', await connectivityCheck.text());
        return { success: false, error: 'Database connectivity issue: ' + connectivityCheck.statusText };
      }
      
      console.log('Database connectivity check successful');
    } catch (connectErr) {
      console.error('Exception during database connectivity check:', connectErr);
      return { success: false, error: 'Failed to connect to database: ' + connectErr.message };
    }
    
    // First check if user already exists to avoid duplicate errors
    console.log('Checking if user already exists...');
    const checkUser = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${encodeURIComponent(userId)}`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY
      }
    });
    
    if (checkUser.ok) {
      const existingUserText = await checkUser.text();
      console.log('Existing user check response:', existingUserText);
      
      try {
        const existingUser = JSON.parse(existingUserText);
        if (existingUser && existingUser.length > 0) {
          console.log('User record already exists in database - returning success');
          
          // If the user exists but doesn't have first/last name, update it
          if (firstName || lastName) {
            if (!existingUser[0].first_name && !existingUser[0].last_name) {
              console.log('Existing user has no name - updating with provided names');
              
              const updateResult = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${encodeURIComponent(userId)}`, {
                method: 'PATCH',
                headers: {
                  'Content-Type': 'application/json',
                  'apikey': SUPABASE_KEY,
                  'Prefer': 'return=representation'
                },
                body: JSON.stringify({
                  first_name: firstName,
                  last_name: lastName
                })
              });
              
              if (updateResult.ok) {
                console.log('Successfully updated existing user with name information');
                return { success: true, data: { ...existingUser[0], first_name: firstName, last_name: lastName } };
              } else {
                console.warn('Failed to update existing user with name information:', await updateResult.text());
              }
            }
          }
          
          return { success: true, data: existingUser[0] };
        }
      } catch (e) {
        console.error('Error parsing existing user response:', e);
      }
    }
    
    // Log the request we're about to make
    const requestBody = JSON.stringify({
      id: userId,
      email: email,
      first_name: firstName,
      last_name: lastName,
      role: 'member'  // Default role
    });
    
    console.log('Request URL:', `${SUPABASE_URL}/rest/v1/users`);
    console.log('Request headers:', {
      'Content-Type': 'application/json',
      'apikey': SUPABASE_KEY ? 'Present (masked)' : 'Missing',
      'Prefer': 'return=representation'
    });
    console.log('Request body:', requestBody);
    
    const response = await fetch(`${SUPABASE_URL}/rest/v1/users`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Prefer': 'return=representation'
      },
      body: requestBody
    });
    
    console.log('User record creation response status:', response.status);
    console.log('Response headers:', Object.fromEntries([...response.headers.entries()]));
    
    const responseText = await response.text();
    console.log('User record creation response text:', responseText);
    
    if (!response.ok) {
      let error;
      try {
        error = JSON.parse(responseText);
      } catch (e) {
        error = { message: responseText || 'Unknown error' };
      }
      console.error('Error creating user record:', error);
      
      // Check if it might be a foreign key constraint issue
      if (responseText.includes('foreign key constraint')) {
        console.error('Possible foreign key constraint violation - check the users table schema');
      }
      
      // Check if it might be a duplicate key issue
      if (responseText.includes('duplicate key') || responseText.includes('unique constraint')) {
        console.log('Duplicate key detected - checking again if user exists...');
        const recheckUser = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${encodeURIComponent(userId)}`, {
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            'apikey': SUPABASE_KEY
          }
        });
        
        if (recheckUser.ok) {
          const existingUser = await recheckUser.json();
          console.log('Existing user recheck result:', existingUser);
          if (existingUser && existingUser.length > 0) {
            console.log('User record found on second check - returning success');
            return { success: true, data: existingUser[0] };
          }
        }
      }
      
      return { success: false, error };
    }
    
    let userData;
    try {
      userData = JSON.parse(responseText);
      
      // Make sure we got a valid response
      if (!userData || (Array.isArray(userData) && userData.length === 0)) {
        console.error('Empty or invalid response data from user creation');
        return { success: false, error: 'Empty response from database' };
      }
      
      // If it's an array, get the first item
      if (Array.isArray(userData)) {
        userData = userData[0];
      }
      
    } catch (e) {
      console.error('Error parsing user data response:', e);
      return { success: false, error: 'Failed to parse response data' };
    }
    
    console.log('User record created successfully:', userData);
    console.log('****** USER RECORD CREATION COMPLETED ******');
    return { success: true, data: userData };
  } catch (err) {
    console.error('Exception when creating user record:', err);
    console.log('****** USER RECORD CREATION FAILED ******');
    return { success: false, error: err.message };
  }
}

/**
 * Request a password reset email with token
 * @param {string} email - User's email
 * @returns {Promise} - The result of the password reset request
 */
export async function requestPasswordReset(email) {
  try {
    console.log('Requesting password reset for email:', email);
    
    const response = await fetch(`${SUPABASE_URL}/auth/v1/recover`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY
      },
      body: JSON.stringify({ email })
    });
    
    console.log('Password reset request response status:', response.status);
    
    if (!response.ok) {
      let errorData;
      try {
        errorData = await response.json();
      } catch (e) {
        errorData = { error: 'Unknown error occurred' };
      }
      console.error('Error requesting password reset:', errorData);
      return { success: false, error: errorData.error || 'Failed to request password reset' };
    }
    
    console.log('Password reset email sent successfully');
    return { success: true };
  } catch (err) {
    console.error('Exception during password reset request:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Reset password using token
 * @param {string} token - Reset password token from email
 * @param {string} newPassword - New password
 * @param {string} email - User's email address
 * @returns {Promise} - The result of the password reset operation
 */
export async function resetPassword(token, newPassword, email) {
  try {
    console.log('======= PASSWORD RESET PROCESS STARTED =======');
    console.log('Resetting password with token (first few chars):', token.substring(0, 5) + '...');
    console.log('New password length:', newPassword ? newPassword.length : 0);
    console.log('Email address:', email);
    
    // Validate required parameters
    if (!token || !newPassword || !email) {
      console.error('Missing required parameters for password reset');
      return {
        success: false,
        error: 'Missing required parameters: token, password, and email are all required'
      };
    }
    
    // STEP 1: Verify the recovery token first
    console.log('STEP 1: Verifying recovery token...');
    const verifyResponse = await fetch(`${SUPABASE_URL}/auth/v1/verify`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY
      },
      body: JSON.stringify({
        type: 'recovery',
        token: token,
        email: email
      })
    });
    
    console.log('Token verification response status:', verifyResponse.status);
    
    if (!verifyResponse.ok) {
      const verifyData = await verifyResponse.text();
      console.error('Token verification failed:', verifyData);
      try {
        const parsedError = JSON.parse(verifyData);
        return { 
          success: false, 
          error: parsedError.error || parsedError.msg || 'Invalid or expired token',
          status: verifyResponse.status
        };
      } catch (e) {
        return { success: false, error: 'Invalid or expired token', status: verifyResponse.status };
      }
    }
    
    // Parse the verification response to get access token
    const verifyData = await verifyResponse.json();
    console.log('Token successfully verified, received access token');
    
    if (!verifyData.access_token) {
      console.error('No access token received after verification');
      return { success: false, error: 'Verification succeeded but no access token received' };
    }
    
    // STEP 2: Update the password using the access token from verification
    console.log('STEP 2: Updating password with access token...');
    const updateResponse = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${verifyData.access_token}`
      },
      body: JSON.stringify({
        password: newPassword
      })
    });
    
    console.log('Password update response status:', updateResponse.status);
    
    if (!updateResponse.ok) {
      const updateErrorText = await updateResponse.text();
      console.error('Password update failed:', updateErrorText);
      try {
        const parsedError = JSON.parse(updateErrorText);
        return { 
          success: false, 
          error: parsedError.error || parsedError.msg || 'Failed to update password',
          status: updateResponse.status
        };
      } catch (e) {
        return { success: false, error: 'Failed to update password', status: updateResponse.status };
      }
    }
    
    // Password update successful
    console.log('Password update successful');
    console.log('======= PASSWORD RESET PROCESS COMPLETED =======');
    
    // Using the verify data as our success response since it contains the session info
    return { success: true, data: verifyData };
    
  } catch (err) {
    console.error('Exception during password reset:', err);
    console.error('Error details:', err.stack);
    console.log('======= PASSWORD RESET PROCESS FAILED =======');
    return { success: false, error: err.message };
  }
}

/**
 * Check if the user is authenticated
 * @returns {Promise<boolean>} - True if user is authenticated
 */
export async function isAuthenticated() {
  console.log('auth-service.js: Checking authentication status...');
  
  try {
    const { accessToken, userId } = await chrome.storage.sync.get(['accessToken', 'userId']);
    console.log('auth-service.js: Retrieved from storage:', { 
      hasAccessToken: !!accessToken, 
      hasUserId: !!userId,
      tokenLength: accessToken ? accessToken.length : 0
    });
    
    if (!accessToken || !userId) {
      console.log('auth-service.js: Missing token or user ID');
      return false;
    }
    
    // Check if the token is expired by parsing the JWT
    try {
      const parts = accessToken.split('.');
      if (parts.length === 3) {
        const payload = JSON.parse(atob(parts[1]));
        console.log('auth-service.js: Token payload:', {
          exp: payload.exp ? new Date(payload.exp * 1000).toISOString() : 'not found',
          iat: payload.iat ? new Date(payload.iat * 1000).toISOString() : 'not found',
          currentTime: new Date().toISOString()
        });
        
        if (payload.exp && payload.exp * 1000 < Date.now()) {
          console.log('auth-service.js: Token is expired');
          return false;
        }
      }
    } catch (e) {
      console.error('auth-service.js: Error checking token expiration:', e);
    }
    
    // Make a test request to verify the token is still valid with Supabase
    console.log('auth-service.js: Making test request to validate token');
    const response = await fetch(`${SUPABASE_URL}/rest/v1/users?limit=1`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    const isAuthenticated = response.ok;
    console.log('auth-service.js: Token validation response status:', response.status, isAuthenticated ? 'Valid' : 'Invalid');
    
    if (!isAuthenticated) {
      // If token is invalid, clear it from storage
      console.log('auth-service.js: Token is invalid, clearing from storage');
      chrome.storage.sync.remove(['accessToken']);
    }
    
    return isAuthenticated;
  } catch (error) {
    console.error('auth-service.js: Authentication check error:', error);
    return false;
  }
}

/**
 * Get all members of a team
 * @param {string} teamId - The ID of the team
 * @returns {Promise} - The result of the operation
 */
export async function getTeamMembers(teamId) {
  try {
    console.log('Getting team members for team ID:', teamId);
    
    // Get the access token from storage
    const { accessToken } = await chrome.storage.sync.get(['accessToken']);
    
    if (!accessToken) {
      console.error('No access token found in storage');
      return { success: false, error: 'Not authenticated' };
    }
    
    const response = await fetch(`${SUPABASE_URL}/rest/v1/users?team_id=eq.${teamId}&select=id,email,role`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    if (!response.ok) {
      const errorData = await response.json();
      console.error('Error getting team members:', errorData);
      return { success: false, error: errorData.error || 'Failed to get team members' };
    }
    
    const members = await response.json();
    console.log(`Found ${members.length} team members`);
    
    return { success: true, data: members };
  } catch (err) {
    console.error('Exception when getting team members:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Leave a team (for members)
 * @returns {Promise} - The result of the operation
 */
export async function leaveTeam() {
  try {
    console.log('Processing team leave request');
    
    // Check if the user is authenticated
    const { accessToken, userId, currentTeamId } = await chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId']);
    
    if (!accessToken || !userId) {
      console.error('Not authenticated. Cannot leave team.');
      return { success: false, error: 'Authentication required' };
    }
    
    if (!currentTeamId) {
      console.error('User is not in a team');
      return { success: false, error: 'You are not currently in a team' };
    }
    
    // Get user role to ensure they're not an admin trying to leave
    const userDetails = await getUserDetails(userId);
    if (!userDetails.success) {
      return { success: false, error: 'Failed to get user details' };
    }
    
    if (userDetails.data.role === 'admin') {
      return { 
        success: false, 
        error: 'Admins cannot leave a team directly. Transfer admin rights or delete the team instead.'
      };
    }
    
    // IMPORTANT: Clean up any existing join requests for this user and team
    // This prevents issues when they try to rejoin the team later
    console.log('Cleaning up any existing join requests before leaving team');
    try {
      const joinRequestsResponse = await fetch(
        `${SUPABASE_URL}/rest/v1/join_requests?user_id=eq.${userId}&team_id=eq.${currentTeamId}`,
        {
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            'apikey': SUPABASE_KEY,
            'Authorization': `Bearer ${accessToken}`
          }
        }
      );
      
      if (joinRequestsResponse.ok) {
        const existingRequests = await joinRequestsResponse.json();
        
        if (existingRequests && existingRequests.length > 0) {
          console.log(`Found ${existingRequests.length} existing join requests to clean up`);
          
          // Delete all existing requests
          const deleteResponse = await fetch(
            `${SUPABASE_URL}/rest/v1/join_requests?user_id=eq.${userId}&team_id=eq.${currentTeamId}`,
            {
              method: 'DELETE',
              headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_KEY,
                'Authorization': `Bearer ${accessToken}`
              }
            }
          );
          
          if (deleteResponse.ok) {
            console.log('Successfully cleaned up existing join requests');
          } else {
            console.error('Failed to clean up join requests:', await deleteResponse.text());
            // Continue with leaving team even if cleanup fails
          }
        } else {
          console.log('No existing join requests found to clean up');
        }
      } else {
        console.error('Error fetching join requests for cleanup:', await joinRequestsResponse.text());
        // Continue with leaving team even if the check fails
      }
    } catch (cleanupError) {
      console.error('Error during join request cleanup:', cleanupError);
      // Continue with leaving team even if cleanup fails
    }
    
    // Update the user's team_id to null
    const response = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${userId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`,
        'Prefer': 'return=representation'
      },
      body: JSON.stringify({
        team_id: null
      })
    });
    
    if (!response.ok) {
      const errorData = await response.json();
      console.error('Error leaving team:', errorData);
      return { success: false, error: errorData.error || 'Failed to leave team' };
    }
    
    // Clear the team ID from storage
    await chrome.storage.sync.remove(['currentTeamId']);
    
    console.log('Successfully left team');
    return { success: true };
  } catch (err) {
    console.error('Exception when leaving team:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Transfer admin role to another team member
 * @param {string} newAdminId - The ID of the user to transfer admin rights to
 * @returns {Promise} - The result of the operation
 */
export async function transferAdminRole(newAdminId) {
  try {
    console.log('Processing admin role transfer to user:', newAdminId);
    
    // Check if the user is authenticated
    const { accessToken, userId, currentTeamId } = await chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId']);
    
    if (!accessToken || !userId) {
      console.error('Not authenticated. Cannot transfer admin rights.');
      return { success: false, error: 'Authentication required' };
    }
    
    if (!currentTeamId) {
      console.error('User is not in a team');
      return { success: false, error: 'You are not currently in a team' };
    }
    
    if (!newAdminId) {
      return { success: false, error: 'New admin ID is required' };
    }
    
    // Get user role to ensure they're an admin
    const userDetails = await getUserDetails(userId);
    if (!userDetails.success) {
      return { success: false, error: 'Failed to get user details' };
    }
    
    if (userDetails.data.role !== 'admin') {
      return { success: false, error: 'Only admins can transfer admin rights' };
    }
    
    // Check if the new admin is in the same team
    const newAdminDetails = await getUserDetails(newAdminId);
    if (!newAdminDetails.success) {
      return { success: false, error: 'Failed to get new admin details' };
    }
    
    if (newAdminDetails.data.team_id !== currentTeamId) {
      return { success: false, error: 'The selected user is not in your team' };
    }
    
    // Start a transaction by using multiple requests
    
    // 1. Update current admin to be a member
    const currentUserResponse = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${userId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      },
      body: JSON.stringify({
        role: 'member'
      })
    });
    
    if (!currentUserResponse.ok) {
      const errorData = await currentUserResponse.json();
      console.error('Error updating current admin:', errorData);
      return { success: false, error: errorData.error || 'Failed to update your role' };
    }
    
    // 2. Update new admin role
    const newAdminResponse = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${newAdminId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      },
      body: JSON.stringify({
        role: 'admin'
      })
    });
    
    if (!newAdminResponse.ok) {
      // Rollback - restore admin rights to original user
      await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${userId}`, {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': `Bearer ${accessToken}`
        },
        body: JSON.stringify({
          role: 'admin'
        })
      });
      
      const errorData = await newAdminResponse.json();
      console.error('Error updating new admin:', errorData);
      return { success: false, error: errorData.error || 'Failed to transfer admin rights' };
    }
    
    console.log('Successfully transferred admin rights');
    return { success: true };
  } catch (err) {
    console.error('Exception when transferring admin role:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Delete a team (admin only)
 * @returns {Promise} - The result of the operation
 */
export async function deleteTeam() {
  try {
    console.log('Processing team deletion request');
    
    // Check if the user is authenticated
    const { accessToken, userId, currentTeamId } = await chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId']);
    
    if (!accessToken || !userId) {
      console.error('Not authenticated. Cannot delete team.');
      return { success: false, error: 'Authentication required' };
    }
    
    if (!currentTeamId) {
      console.error('User is not in a team');
      return { success: false, error: 'You are not currently in a team' };
    }
    
    // Get user role to ensure they're an admin
    const userDetails = await getUserDetails(userId);
    if (!userDetails.success) {
      return { success: false, error: 'Failed to get user details' };
    }
    
    if (userDetails.data.role !== 'admin') {
      return { success: false, error: 'Only admins can delete a team' };
    }
    
    // Delete the team
    const response = await fetch(`${SUPABASE_URL}/rest/v1/teams?id=eq.${currentTeamId}`, {
      method: 'DELETE',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    if (!response.ok) {
      const errorData = await response.json();
      console.error('Error deleting team:', errorData);
      return { success: false, error: errorData.error || 'Failed to delete team' };
    }
    
    // Clear the team ID from storage
    await chrome.storage.sync.remove(['currentTeamId']);
    
    console.log('Successfully deleted team');
    return { success: true };
  } catch (err) {
    console.error('Exception when deleting team:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Update team details (admin only)
 * @param {string} teamId - The ID of the team to update
 * @param {Object} teamDetails - The updated details for the team
 * @returns {Promise} - The result of the operation
 */
export async function updateTeamDetails(teamId, teamDetails) {
  try {
    console.log('Processing team details update for team:', teamId);
    
    // Check if the user is authenticated
    const { accessToken, userId, currentTeamId } = await chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId']);
    
    if (!accessToken || !userId) {
      console.error('Not authenticated. Cannot update team details.');
      return { success: false, error: 'Authentication required' };
    }
    
    if (!currentTeamId || currentTeamId !== teamId) {
      console.error('User is not in the specified team');
      return { success: false, error: 'You are not a member of this team' };
    }
    
    // Get user role to ensure they're an admin
    const userDetails = await getUserDetails(userId);
    if (!userDetails.success) {
      return { success: false, error: 'Failed to get user details' };
    }
    
    if (userDetails.data.role !== 'admin') {
      return { success: false, error: 'Only admins can update team details' };
    }
    
    // Prepare the update data
    const updateData = {};
    
    // Only include fields that are provided in the update
    if (teamDetails.organization !== undefined) updateData.organization = teamDetails.organization;
    if (teamDetails.invoice_email !== undefined) updateData.invoice_email = teamDetails.invoice_email;
    if (teamDetails.billing_street !== undefined) updateData.billing_street = teamDetails.billing_street;
    if (teamDetails.billing_city !== undefined) updateData.billing_city = teamDetails.billing_city;
    if (teamDetails.billing_region !== undefined) updateData.billing_region = teamDetails.billing_region;
    if (teamDetails.billing_country !== undefined) updateData.billing_country = teamDetails.billing_country;
    
    // Update the team
    const response = await fetch(`${SUPABASE_URL}/rest/v1/teams?id=eq.${teamId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`,
        'Prefer': 'return=representation'
      },
      body: JSON.stringify(updateData)
    });
    
    if (!response.ok) {
      const errorData = await response.json();
      console.error('Error updating team details:', errorData);
      return { success: false, error: errorData.error || 'Failed to update team details' };
    }
    
    const updatedTeam = await response.json();
    
    console.log('Successfully updated team details');
    return { success: true, data: updatedTeam[0] };
  } catch (err) {
    console.error('Exception when updating team details:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Request to join a team using an access code
 * @param {string} accessCode - The access code for the team
 * @returns {Promise} - The result of the operation
 */
export async function requestToJoinTeam(accessCode) {
  try {
    console.log('Processing request to join team with access code:', accessCode);
    
    // Check if the user is authenticated
    const { accessToken, userId } = await chrome.storage.sync.get(['accessToken', 'userId']);
    
    if (!accessToken || !userId) {
      console.error('Not authenticated. Cannot request to join team.');
      return { success: false, error: 'Authentication required' };
    }

    // Check if user is already in a team
    const userDetails = await getUserDetails(userId);
    if (!userDetails.success) {
      return { success: false, error: 'Failed to get user details' };
    }
    
    if (userDetails.data.team_id) {
      return { success: false, error: 'You are already a member of a team. Please leave your current team first.' };
    }
    
    // Find the team with the given access code
    const teamResponse = await fetch(`${SUPABASE_URL}/rest/v1/teams?access_code=eq.${accessCode}`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    if (!teamResponse.ok) {
      const errorData = await teamResponse.json();
      console.error('Error finding team with access code:', errorData);
      return { success: false, error: errorData.error || 'Failed to find team with this access code' };
    }
    
    const teams = await teamResponse.json();
    
    if (!teams || teams.length === 0) {
      return { success: false, error: 'Invalid access code. No team found with this code.' };
    }
    
    const team = teams[0];
    console.log('Found team with access code:', team.name);
    
    // Check if the user already has any request (pending, approved, or rejected) for this team
    // We don't filter by status to find ALL existing requests
    const existingRequestResponse = await fetch(
      `${SUPABASE_URL}/rest/v1/join_requests?user_id=eq.${userId}&team_id=eq.${team.id}`,
      {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': `Bearer ${accessToken}`
        }
      }
    );
    
    if (!existingRequestResponse.ok) {
      const errorData = await existingRequestResponse.json();
      console.error('Error checking existing requests:', errorData);
      return { success: false, error: errorData.error || 'Failed to check existing requests' };
    }
    
    const existingRequests = await existingRequestResponse.json();
    
    // Check for pending requests first
    const pendingRequests = existingRequests.filter(req => req.status === 'pending');
    if (pendingRequests && pendingRequests.length > 0) {
      return { 
        success: false, 
        error: 'You already have a pending request to join this team. Please wait for admin approval.' 
      };
    }
    
    // If there are existing requests with other statuses (approved/rejected), delete them first
    if (existingRequests && existingRequests.length > 0) {
      console.log(`Found ${existingRequests.length} existing requests with non-pending status. Cleaning up before creating a new request.`);
      
      for (const request of existingRequests) {
        const deleteResponse = await fetch(
          `${SUPABASE_URL}/rest/v1/join_requests?id=eq.${request.id}`,
          {
            method: 'DELETE',
            headers: {
              'Content-Type': 'application/json',
              'apikey': SUPABASE_KEY,
              'Authorization': `Bearer ${accessToken}`
            }
          }
        );
        
        if (!deleteResponse.ok) {
          console.error(`Failed to delete existing request ${request.id}:`, await deleteResponse.text());
          // Continue with the operation even if cleanup fails
        } else {
          console.log(`Successfully deleted existing request ${request.id}`);
        }
      }
    }
    
    // Create a new join request
    const requestResponse = await fetch(`${SUPABASE_URL}/rest/v1/join_requests`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`,
        'Prefer': 'return=representation'
      },
      body: JSON.stringify({
        user_id: userId,
        team_id: team.id,
        status: 'pending'
      })
    });
    
    if (!requestResponse.ok) {
      const errorData = await requestResponse.json();
      console.error('Error creating join request:', errorData);
      return { success: false, error: errorData.error || 'Failed to create join request' };
    }
    
    const requestData = await requestResponse.json();
    
    console.log('Successfully created join request:', requestData);
    return { 
      success: true, 
      data: {
        requestId: requestData[0].id,
        teamId: team.id,
        teamName: team.name
      },
      message: `Request to join team "${team.name}" has been sent. Please wait for admin approval.`
    };
  } catch (err) {
    console.error('Exception when requesting to join team:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Get pending join requests for a team (admin only)
 * @returns {Promise} - The result of the operation with the list of pending requests
 */
export async function getPendingJoinRequests() {
  try {
    console.log('Getting pending join requests for current team');
    
    // Check if the user is authenticated and an admin
    const { accessToken, userId, currentTeamId } = await chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId']);
    
    if (!accessToken || !userId) {
      console.error('Not authenticated. Cannot view join requests.');
      return { success: false, error: 'Authentication required' };
    }
    
    if (!currentTeamId) {
      console.error('User is not in a team');
      return { success: false, error: 'You are not currently in a team' };
    }
    
    // Get user role to ensure they're an admin
    const userDetails = await getUserDetails(userId);
    if (!userDetails.success) {
      return { success: false, error: 'Failed to get user details' };
    }
    
    if (userDetails.data.role !== 'admin') {
      return { success: false, error: 'Only admins can view join requests' };
    }
    
    // Get all pending requests for the team
    const response = await fetch(
      `${SUPABASE_URL}/rest/v1/join_requests?team_id=eq.${currentTeamId}&status=eq.pending&select=id,user_id,request_date`,
      {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': `Bearer ${accessToken}`
        }
      }
    );
    
    if (!response.ok) {
      const errorData = await response.json();
      console.error('Error getting join requests:', errorData);
      return { success: false, error: errorData.error || 'Failed to get join requests' };
    }
    
    let requests = await response.json();
    console.log(`Found ${requests.length} pending join requests`);
    
    // If there are requests, get the user details for each
    if (requests.length > 0) {
      const userIds = requests.map(req => req.user_id);
      
      // Get user details for all users in the requests
      const userDetailsResponse = await fetch(
        `${SUPABASE_URL}/rest/v1/users?id=in.(${userIds.join(',')})&select=id,email,first_name,last_name`,
        {
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            'apikey': SUPABASE_KEY,
            'Authorization': `Bearer ${accessToken}`
          }
        }
      );
      
      if (userDetailsResponse.ok) {
        const users = await userDetailsResponse.json();
        
        // Create a map of user IDs to user details
        const userMap = {};
        users.forEach(user => {
          userMap[user.id] = user;
        });
        
        // Add user details to each request
        requests = requests.map(req => ({
          ...req,
          user: userMap[req.user_id] || { email: 'Unknown' }
        }));
      } else {
        console.error('Error getting user details for requests:', await userDetailsResponse.text());
      }
    }
    
    return { success: true, data: requests };
  } catch (err) {
    console.error('Exception when getting join requests:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Respond to a join request (approve or reject)
 * @param {string} requestId - The ID of the join request
 * @param {boolean} approved - Whether to approve or reject the request
 * @returns {Promise} - The result of the operation
 */
export async function respondToJoinRequest(requestId, approved) {
  try {
    console.log('Responding to join request:', requestId, approved ? 'Approved' : 'Rejected');
    
    // Check if the user is authenticated and an admin
    const { accessToken, userId, currentTeamId } = await chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId']);
    
    if (!accessToken || !userId) {
      console.error('Not authenticated. Cannot respond to join request.');
      return { success: false, error: 'Authentication required' };
    }
    
    if (!currentTeamId) {
      console.error('User is not in a team');
      return { success: false, error: 'You are not currently in a team' };
    }
    
    // Get user role to ensure they're an admin
    const userDetails = await getUserDetails(userId);
    if (!userDetails.success) {
      return { success: false, error: 'Failed to get user details' };
    }
    
    if (userDetails.data.role !== 'admin') {
      return { success: false, error: 'Only admins can approve or reject join requests' };
    }
    
    // First get the request to make sure it exists and belongs to this team
    const requestResponse = await fetch(
      `${SUPABASE_URL}/rest/v1/join_requests?id=eq.${requestId}&select=*`,
      {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': `Bearer ${accessToken}`
        }
      }
    );
    
    if (!requestResponse.ok) {
      const errorData = await requestResponse.json();
      console.error('Error getting join request:', errorData);
      return { success: false, error: errorData.error || 'Failed to get join request' };
    }
    
    const requests = await requestResponse.json();
    
    if (!requests || requests.length === 0) {
      return { success: false, error: 'Join request not found' };
    }
    
    const request = requests[0];
    
    if (request.team_id !== currentTeamId) {
      return { success: false, error: 'This join request is for a different team' };
    }
    
    if (request.status !== 'pending') {
      return { success: false, error: 'This join request has already been processed' };
    }
    
    // The status we're trying to update to
    const targetStatus = approved ? 'approved' : 'rejected';
    
    // ALTERNATIVE APPROACH: Instead of trying to delete conflicting requests,
    // which appears problematic, let's try a different approach:
    console.log(`Using alternative approach to handle conflicting requests for status '${targetStatus}'`);
    
    try {
      // Step 1: Find all existing requests for this user and team
      const allRequestsResponse = await fetch(
        `${SUPABASE_URL}/rest/v1/join_requests?user_id=eq.${request.user_id}&team_id=eq.${request.team_id}`,
        {
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            'apikey': SUPABASE_KEY,
            'Authorization': `Bearer ${accessToken}`
          }
        }
      );
      
      if (allRequestsResponse.ok) {
        const allRequests = await allRequestsResponse.json();
        
        if (allRequests && allRequests.length > 0) {
          console.log(`Found ${allRequests.length} total requests for this user and team`);
          
          // Separate the current request from other requests
          const currentRequest = allRequests.find(req => req.id === requestId);
          const otherRequests = allRequests.filter(req => req.id !== requestId);
          
          // Log for debugging
          console.log(`Current request: ID=${currentRequest?.id}, Status=${currentRequest?.status}`);
          console.log(`Other requests: ${otherRequests.length} total`);
          otherRequests.forEach((req, index) => {
            console.log(`- Request ${index+1}: ID=${req.id}, Status=${req.status}`);
          });
          
          // Specifically check for any requests with the target status - these would cause constraint violations
          const conflictingRequests = otherRequests.filter(req => req.status === targetStatus);
          
          if (conflictingRequests.length > 0) {
            console.log(`Found ${conflictingRequests.length} conflicting requests with status '${targetStatus}'`);
            
            // Step 2: Instead of deleting conflicting requests, update them to a temporary status
            // to avoid unique constraint violations
            for (const conflictRequest of conflictingRequests) {
              console.log(`Updating conflicting request ${conflictRequest.id} from status '${conflictRequest.status}' to 'rejected'`);
              
              const tempStatus = 'rejected'; // Using 'rejected' which is an allowed value in the constraint
              
              // Update the status to a temporary value to avoid the unique constraint
              const updateTempResponse = await fetch(
                `${SUPABASE_URL}/rest/v1/join_requests?id=eq.${conflictRequest.id}`,
                {
                  method: 'PATCH',
                  headers: {
                    'Content-Type': 'application/json',
                    'apikey': SUPABASE_KEY,
                    'Authorization': `Bearer ${accessToken}`,
                    'Prefer': 'return=representation'
                  },
                  body: JSON.stringify({
                    status: tempStatus,
                    resolved_date: new Date().toISOString(),
                    resolved_by: userId
                  })
                }
              );
              
              if (!updateTempResponse.ok) {
                console.error(`Failed to update conflicting request ${conflictRequest.id} to temporary status:`, await updateTempResponse.text());
              } else {
                console.log(`Successfully updated conflicting request ${conflictRequest.id} to temporary status`);
              }
            }
            
            // Add a small delay to ensure database consistency
            console.log('Waiting for updates to process...');
            await new Promise(resolve => setTimeout(resolve, 500));
          }
        }
      } else {
        console.error('Error fetching all requests:', await allRequestsResponse.text());
      }
    } catch (cleanupError) {
      console.error('Error during request cleanup:', cleanupError);
    }
    
    // Step 3: Now update the current request to the target status
    console.log(`Now updating request ${requestId} to status '${targetStatus}'`);
    const updateResponse = await fetch(
      `${SUPABASE_URL}/rest/v1/join_requests?id=eq.${requestId}`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': `Bearer ${accessToken}`,
          'Prefer': 'return=representation'
        },
        body: JSON.stringify({
          status: targetStatus,
          resolved_date: new Date().toISOString(),
          resolved_by: userId
        })
      }
    );
    
    if (!updateResponse.ok) {
      const errorResponse = await updateResponse.text();
      console.error('Error updating join request:', errorResponse);
      
      try {
        let errorData;
        try {
          errorData = JSON.parse(errorResponse);
        } catch (e) {
          errorData = { message: errorResponse };
        }
        
        console.error('Parsed error updating join request:', errorData);
        
        // If we still hit a constraint violation despite our preparation steps
        if ((errorData.code === '23505' || errorResponse.includes('23505')) && 
            (errorData.message?.includes('unique_user_team_request') || errorResponse.includes('unique_user_team_request'))) {
          
          console.error('STILL HIT CONSTRAINT! Skipping status update and directly adding user to team');
          
          // Skip the join request update and directly add the user to the team if approving
          if (approved) {
            return await addUserToTeam(request.user_id, currentTeamId, accessToken);
          } else {
            // If rejecting, just consider it done
            return { 
              success: true, 
              data: { requestId, status: 'rejected' },
              message: 'Join request rejected successfully (bypassed update)'
            };
          }
        }
        
        return { success: false, error: errorData.message || errorData.error || 'Failed to update join request' };
      } catch (e) {
        return { success: false, error: 'Failed to update join request: ' + errorResponse };
      }
    }
    
    // Step 4: Now that we've successfully updated the current request, we can safely delete the old requests
    // since we've already completed the critical operation
    try {
      const oldRequestsResponse = await fetch(
        `${SUPABASE_URL}/rest/v1/join_requests?user_id=eq.${request.user_id}&team_id=eq.${request.team_id}&status=eq.rejected&id=neq.${requestId}`,
        {
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            'apikey': SUPABASE_KEY,
            'Authorization': `Bearer ${accessToken}`
          }
        }
      );
      
      if (oldRequestsResponse.ok) {
        const oldRequests = await oldRequestsResponse.json();
        
        if (oldRequests && oldRequests.length > 0) {
          console.log(`Found ${oldRequests.length} old requests with temporary status to clean up`);
          
          // Delete all old requests with the temporary status
          for (const oldRequest of oldRequests) {
            console.log(`Cleaning up old request ${oldRequest.id} with status 'rejected'`);
            
            // Delete the old request
            const deleteResponse = await fetch(
              `${SUPABASE_URL}/rest/v1/join_requests?id=eq.${oldRequest.id}`,
              {
                method: 'DELETE',
                headers: {
                  'Content-Type': 'application/json',
                  'apikey': SUPABASE_KEY,
                  'Authorization': `Bearer ${accessToken}`
                }
              }
            );
            
            if (!deleteResponse.ok) {
              console.error(`Failed to delete old request ${oldRequest.id}:`, await deleteResponse.text());
              // Continue anyway - this is just cleanup
            } else {
              console.log(`Successfully deleted old request ${oldRequest.id}`);
            }
          }
        } else {
          console.log('No old requests found to clean up');
        }
      }
    } catch (cleanupError) {
      console.error('Error during final cleanup:', cleanupError);
      // Continue anyway - the important operation has already succeeded
    }
    
    // If approved, add the user to the team
    if (approved) {
      return await addUserToTeam(request.user_id, currentTeamId, accessToken);
    }
    
    return { 
      success: true, 
      data: { requestId, status: targetStatus },
      message: `Join request ${approved ? 'approved' : 'rejected'} successfully`
    };
  } catch (err) {
    console.error('Exception when responding to join request:', err);
    return { success: false, error: err.message };
  }
}

// Helper function to add a user to a team
async function addUserToTeam(userId, teamId, accessToken) {
  console.log(`Adding user ${userId} to team ${teamId}`);
  
  try {
    const userUpdateResponse = await fetch(
      `${SUPABASE_URL}/rest/v1/users?id=eq.${userId}`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': `Bearer ${accessToken}`,
          'Prefer': 'return=representation'
        },
        body: JSON.stringify({
          team_id: teamId,
          role: 'member'
        })
      }
    );
    
    if (!userUpdateResponse.ok) {
      const errorData = await userUpdateResponse.json();
      console.error('Error adding user to team:', errorData);
      return { success: false, error: errorData.error || 'Failed to add user to team' };
    }
    
    console.log('Successfully added user to team');
    return { 
      success: true, 
      data: { userId, teamId, role: 'member' },
      message: 'User added to team successfully'
    };
  } catch (err) {
    console.error('Exception when adding user to team:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Check the status of user's join requests
 * @returns {Promise} - The result of the operation with the user's join requests
 */
export async function checkJoinRequestStatus() {
  try {
    console.log('Checking status of user join requests');
    
    // Check if the user is authenticated
    const { accessToken, userId } = await chrome.storage.sync.get(['accessToken', 'userId']);
    
    if (!accessToken || !userId) {
      console.error('Not authenticated. Cannot check join request status.');
      return { success: false, error: 'Authentication required' };
    }
    
    // Get all requests for the user
    const response = await fetch(
      `${SUPABASE_URL}/rest/v1/join_requests?user_id=eq.${userId}&select=id,team_id,status,request_date,resolved_date`,
      {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json',
          'apikey': SUPABASE_KEY,
          'Authorization': `Bearer ${accessToken}`
        }
      }
    );
    
    if (!response.ok) {
      const errorData = await response.json();
      console.error('Error getting user join requests:', errorData);
      return { success: false, error: errorData.error || 'Failed to get join requests' };
    }
    
    let requests = await response.json();
    console.log(`Found ${requests.length} join requests for user`);
    
    // If there are requests, get the team details for each
    if (requests.length > 0) {
      const teamIds = [...new Set(requests.map(req => req.team_id))];
      
      // Get team details for all teams in the requests
      const teamDetailsResponse = await fetch(
        `${SUPABASE_URL}/rest/v1/teams?id=in.(${teamIds.join(',')})&select=id,name`,
        {
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            'apikey': SUPABASE_KEY,
            'Authorization': `Bearer ${accessToken}`
          }
        }
      );
      
      if (teamDetailsResponse.ok) {
        const teams = await teamDetailsResponse.json();
        
        // Create a map of team IDs to team details
        const teamMap = {};
        teams.forEach(team => {
          teamMap[team.id] = team;
        });
        
        // Add team details to each request
        requests = requests.map(req => ({
          ...req,
          team: teamMap[req.team_id] || { name: 'Unknown' }
        }));
      } else {
        console.error('Error getting team details for requests:', await teamDetailsResponse.text());
      }
    }
    
    // Look specifically for approved requests
    const approvedRequest = requests.find(req => req.status === 'approved');
    
    return { 
      success: true, 
      data: {
        requests,
        hasApprovedRequest: !!approvedRequest,
        approvedTeamId: approvedRequest ? approvedRequest.team_id : null
      }
    };
  } catch (err) {
    console.error('Exception when checking join request status:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Remove a team member (admin only)
 * @param {string} memberId - The ID of the member to remove
 * @returns {Promise} - The result of the operation
 */
export async function removeTeamMember(memberId) {
  try {
    console.log('Removing team member with ID:', memberId);
    
    // Get authentication data
    const { accessToken, userId, currentTeamId } = await chrome.storage.sync.get(['accessToken', 'userId', 'currentTeamId']);
    
    if (!accessToken || !userId || !currentTeamId) {
      console.error('Authentication data missing, cannot remove team member');
      return { success: false, error: 'Authentication required' };
    }
    
    // First, check if the current user is an admin
    const currentUserResponse = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${userId}&select=role,team_id`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    if (!currentUserResponse.ok) {
      const error = await currentUserResponse.json();
      console.error('Error checking current user role:', error);
      return { success: false, error: 'Failed to verify admin status' };
    }
    
    const currentUserData = await currentUserResponse.json();
    
    if (currentUserData.length === 0) {
      console.error('Current user not found');
      return { success: false, error: 'User not found' };
    }
    
    const currentUser = currentUserData[0];
    
    // Verify the user is an admin
    if (currentUser.role !== 'admin') {
      console.error('Current user is not an admin, cannot remove team members');
      return { success: false, error: 'Admin privileges required to remove team members' };
    }
    
    // Verify the user to be removed exists and is part of the same team
    const memberResponse = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${memberId}&select=email,role,team_id`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    if (!memberResponse.ok) {
      const error = await memberResponse.json();
      console.error('Error checking member:', error);
      return { success: false, error: 'Failed to verify member status' };
    }
    
    const memberData = await memberResponse.json();
    
    if (memberData.length === 0) {
      console.error('Member not found');
      return { success: false, error: 'Member not found' };
    }
    
    const member = memberData[0];
    
    // Prevent removing another admin
    if (member.role === 'admin') {
      console.error('Cannot remove an admin user');
      return { success: false, error: 'Cannot remove another admin. Transfer admin rights first.' };
    }
    
    // Make sure the member is part of the same team
    if (member.team_id !== currentTeamId) {
      console.error('Member is not part of the same team');
      return { success: false, error: 'Member is not part of your team' };
    }
    
    // Update the user to remove the team association
    const updateResponse = await fetch(`${SUPABASE_URL}/rest/v1/users?id=eq.${memberId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY,
        'Prefer': 'return=representation',
        'Authorization': `Bearer ${accessToken}`
      },
      body: JSON.stringify({
        team_id: null
      })
    });
    
    if (!updateResponse.ok) {
      const error = await updateResponse.json();
      console.error('Error removing team member:', error);
      return { success: false, error: 'Failed to remove team member' };
    }
    
    console.log('Team member removed successfully');
    return { success: true };
  } catch (err) {
    console.error('Exception when removing team member:', err);
    return { success: false, error: err.message };
  }
}

/**
 * Refresh an expired access token using a refresh token
 * @param {string} refreshToken - The refresh token to use for obtaining a new access token
 * @returns {Promise} - The result of the token refresh operation
 */
export async function refreshAccessToken(refreshToken) {
  try {
    console.log('Attempting to refresh access token');
    
    if (!refreshToken) {
      console.error('No refresh token provided');
      return { success: false, error: 'No refresh token available' };
    }
    
    // Make the refresh token request to Supabase
    const refreshResponse = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=refresh_token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': SUPABASE_KEY
      },
      body: JSON.stringify({
        refresh_token: refreshToken
      })
    });
    
    const refreshData = await refreshResponse.text();
    console.log('Refresh token response status:', refreshResponse.status);
    
    if (!refreshResponse.ok) {
      let errorMessage = 'Token refresh failed';
      
      try {
        const errorData = JSON.parse(refreshData);
        console.error('Token refresh error details:', errorData);
        errorMessage = errorData.error_description || errorData.error || 'Authentication failed';
      } catch (e) {
        console.error('Could not parse error response:', refreshData);
      }
      
      return { success: false, error: errorMessage };
    }
    
    // Parse the response data
    let data;
    try {
      data = JSON.parse(refreshData);
    } catch (e) {
      console.error('Error parsing refresh token response:', e);
      return { success: false, error: 'Invalid response from server' };
    }
    
    // Extract the new tokens
    const newAccessToken = data.access_token;
    const newRefreshToken = data.refresh_token;
    
    if (!newAccessToken) {
      console.error('No access token received in refresh response');
      return { success: false, error: 'No access token in refresh response' };
    }
    
    // Store the new tokens in Chrome storage
    console.log('Storing refreshed tokens in Chrome storage');
    await chrome.storage.sync.set({
      accessToken: newAccessToken,
      refreshToken: newRefreshToken || refreshToken // Use new refresh token if provided, otherwise keep the old one
    });
    
    // Parse token expiration for logging
    try {
      const parts = newAccessToken.split('.');
      if (parts.length === 3) {
        const payload = JSON.parse(atob(parts[1]));
        const expTime = payload.exp ? new Date(payload.exp * 1000) : null;
        const currentTime = new Date();
        
        console.log('New token expiration:', {
          expTime: expTime ? expTime.toISOString() : 'unknown',
          currentTime: currentTime.toISOString(),
          timeRemaining: expTime ? ((expTime - currentTime) / 1000 / 60).toFixed(2) + ' minutes' : 'unknown'
        });
      }
    } catch (e) {
      console.error('Error checking new token expiration:', e);
    }
    
    // Notify about token refresh
    try {
      chrome.runtime.sendMessage({ 
        action: 'tokenRefreshed',
        success: true
      });
    } catch (e) {
      console.error('Error sending token refresh notification:', e);
    }
    
    return { 
      success: true, 
      accessToken: newAccessToken,
      refreshToken: newRefreshToken
    };
  } catch (err) {
    console.error('Exception during token refresh:', err);
    return { success: false, error: err.message };
  }
}

export default {
  signUp,
  signIn,
  signOut,
  createTeam,
  joinTeam,
  leaveTeam,
  transferAdminRole,
  deleteTeam,
  updateTeamDetails,
  requestToJoinTeam,
  getPendingJoinRequests,
  respondToJoinRequest,
  checkJoinRequestStatus,
  getTeamMembers,
  getAllTeams,
  getUserDetails,
  isAuthenticated,
  requestPasswordReset,
  resetPassword,
  removeTeamMember,
  refreshAccessToken
}; 