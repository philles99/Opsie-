/**
 * API Service Utility
 * This file contains functions for interacting with external APIs
 */

import { storeEmail, getThreads, getThreadMessages, getContactHistory } from './supabase-client.js';

// Add these variables at the top of your file
const summaryCache = new Map();
const contactCache = new Map();
const replyCache = new Map();
let isGeneratingSummary = false;
let isGeneratingContactSummary = false;

// Get API key from storage
async function getApiKey() {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({action: 'getApiKey'}, function(response) {
      console.log('API key retrieval response:', response ? 'Response received' : 'No response');
      console.log('API key exists:', response && response.apiKey ? 'Yes' : 'No');
      console.log('API key length:', response && response.apiKey ? response.apiKey.length : 0);
      
      // Check if the API key starts with "sk-" (OpenAI API keys format)
      const isValidFormat = response && response.apiKey && response.apiKey.startsWith('sk-');
      console.log('API key format valid:', isValidFormat ? 'Yes' : 'No');
      
      resolve(response?.apiKey);
    });
  });
}

// Function to access the summary cache
export function getSummaryCache() {
  return summaryCache;
}

// Function to access the contact cache
export function getContactCache() {
  return contactCache;
}

// Function to access the reply cache
export function getReplyCache() {
  return replyCache;
}

// Function to get the OpenAI API key from Chrome storage
async function getOpenAIApiKey() {
  return new Promise((resolve) => {
    chrome.storage.sync.get(['openaiApiKey'], (result) => {
      console.log('OpenAI API Key retrieval - key exists:', result.openaiApiKey ? 'Yes' : 'No');
      console.log('OpenAI API Key retrieval - key length:', result.openaiApiKey ? result.openaiApiKey.length : 0);
      resolve(result.openaiApiKey || null);
    });
  });
}

// Generate email summary using AI
async function generateEmailSummary(emailData) {
  try {
    // Generate a more reliable cache key that includes thread info
    const threadHistoryLength = emailData.threadHistory ? emailData.threadHistory.length : 0;
    const cacheKey = `${emailData.subject}-${emailData.sender.email}-${threadHistoryLength}`;
    
    // Log cache status
    console.log('Cache check for:', cacheKey);
    console.log('Cache hit:', summaryCache.has(cacheKey));
    
    // Check if we have a cached summary
    if (summaryCache.has(cacheKey)) {
      console.log('Using cached summary for:', emailData.subject);
      return summaryCache.get(cacheKey);
    }
    
    // Check if we're already generating a summary
    if (isGeneratingSummary) {
      console.log('Already generating a summary, returning default');
      const result = {
        error: 'Summary generation in progress',
        summaryItems: getDefaultSummaryItems(emailData),
        urgencyScore: 5 // Default middle urgency
      };
      return result;
    }
    
    isGeneratingSummary = true;
    
    const apiKey = await getOpenAIApiKey();
    
    if (!apiKey) {
      console.error('No OpenAI API key found');
      const result = {
        error: 'No API key provided',
        summaryItems: getDefaultSummaryItems(emailData),
        urgencyScore: 5 // Default middle urgency
      };
      summaryCache.set(cacheKey, result);
      return result;
    }
    
    console.log('Generating summary with OpenAI for:', emailData.subject);
    
    // Prepare thread history content if available
    let threadContent = '';
    if (emailData.threadHistory && emailData.threadHistory.length > 0) {
      threadContent = '\n\nThread History:\n';
      emailData.threadHistory.forEach((message, index) => {
        threadContent += `\n--- Previous Message ${index + 1} ---\n`;
        threadContent += `From: ${message.sender.name || 'Unknown'} (${message.sender.email || 'No email'})\n`;
        threadContent += `Date: ${message.timestamp || 'Unknown date'}\n`;
        threadContent += `Content: ${message.message || 'No content'}\n`;
      });
    }
    
    // Make a real API call to OpenAI's ChatGPT
    try {
      console.log('Making API request to OpenAI...');
      
      // Create the full payload
      const payload = {
        model: 'gpt-3.5-turbo',
        messages: [
          {
            role: 'system',
            content: 'You are an assistant that summarizes emails and assesses their urgency. Analyze the email and respond with JSON containing: 1) "summaryItems": an array of 3 key points from the email in bullet point format (without the bullet characters), and 2) "urgencyScore": an integer from 1-10 indicating how urgent the email is, where 1 is not urgent at all and 10 is extremely urgent. Base the urgency on factors like time-sensitivity, sender importance, action requirements, etc. Please keep in mind that these are emails from busy working professionals. Try to weigh the urgency to the importance when giving your urgency rating e.g ads with limited time-sensitivity are not urgent. In the summary, finish with "Action Required:" if there is a clear action to be done. If email thread history is provided, analyze the entire conversation for context.'
          },
          {
            role: 'user',
            content: `Please summarize this email and rate its urgency:
            From: ${emailData.sender.name || 'Unknown'} (${emailData.sender.email || 'No email'})
            Subject: ${emailData.subject || 'No subject'}
            Date: ${emailData.date || emailData.timestamp || 'Unknown date'}
            Body: ${emailData.message || emailData.body || 'No body content'}${threadContent}`
          }
        ],
        temperature: 0.7,
        max_tokens: 500
      };
      
      // Log the email content being sent to OpenAI
      console.log('Email content being sent to OpenAI:');
      console.log(`From: ${emailData.sender.name || 'Unknown'} (${emailData.sender.email || 'No email'})`);
      console.log(`Subject: ${emailData.subject || 'No subject'}`);
      console.log(`Body: ${(emailData.message || emailData.body || 'No body content').substring(0, 200)}...`);
      if (threadContent) {
        console.log(`Thread history included: ${emailData.threadHistory.length} previous messages`);
      }
      
      // Log the full payload being sent to OpenAI
      console.log('Full payload being sent to OpenAI:');
      console.log(JSON.stringify(payload, null, 2));
      
      const response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify(payload)
      });
      
      console.log('OpenAI API response status:', response.status);
      
      // Log the raw response for debugging
      const responseText = await response.text();
      console.log('OpenAI API raw response:', responseText);
      
      // Parse the JSON response
      const data = JSON.parse(responseText);
      
      if (data.error) {
        console.error('OpenAI API error:', data.error);
        const result = {
          error: data.error.message,
          summaryItems: getDefaultSummaryItems(emailData),
          urgencyScore: 5 // Default middle urgency
        };
        summaryCache.set(cacheKey, result);
        return result;
      }

      console.log('OpenAI response received:', data);

      // Parse the response content
      const aiContent = data.choices[0].message.content.trim();
      console.log('AI content:', aiContent);
      
      // Try to parse the JSON response from the AI
      let jsonResponse;
      try {
        // Extract JSON object if it's embedded in other text
        const jsonMatch = aiContent.match(/\{[\s\S]*\}/);
        const jsonString = jsonMatch ? jsonMatch[0] : aiContent;
        jsonResponse = JSON.parse(jsonString);
        console.log('Parsed JSON response:', jsonResponse);
      } catch (jsonError) {
        console.error('Error parsing JSON from AI response:', jsonError);
        
        // Fallback: Extract summary items and set default urgency
        const summaryItems = aiContent
          .split('\n')
          .filter(line => line.trim().startsWith('-') || line.trim().startsWith('•'))
          .map(line => line.replace(/^[•-]\s*/, '').trim());
        
        jsonResponse = {
          summaryItems: summaryItems.length > 0 ? summaryItems : aiContent.split('\n').filter(line => line.trim()).slice(0, 3),
          urgencyScore: 5
        };
        
        console.log('Created fallback JSON response:', jsonResponse);
      }
      
      // Ensure urgencyScore is within 1-10 range
      let urgencyScore = jsonResponse.urgencyScore;
      if (!urgencyScore || isNaN(urgencyScore) || urgencyScore < 1 || urgencyScore > 10) {
        urgencyScore = 5; // Default to middle urgency if invalid
      }
      
      // Ensure summaryItems is an array with content
      let summaryItems = jsonResponse.summaryItems;
      if (!Array.isArray(summaryItems) || summaryItems.length === 0) {
        summaryItems = getDefaultSummaryItems(emailData);
      }
      
      const result = { 
        summaryItems,
        urgencyScore: Math.round(urgencyScore) // Ensure it's an integer
      };
      
      // Before returning, release the lock
      isGeneratingSummary = false;
      
      // Cache and return the result
      summaryCache.set(cacheKey, result);
      return result;
    } catch (fetchError) {
      // Release the lock in case of error
      isGeneratingSummary = false;
      console.error('Fetch error when calling OpenAI API:', fetchError);
      const result = {
        error: 'API request failed: ' + fetchError.message,
        summaryItems: getDefaultSummaryItems(emailData),
        urgencyScore: 5 // Default middle urgency
      };
      summaryCache.set(cacheKey, result);
      return result;
    }
  } catch (error) {
    // Release the lock in case of error
    isGeneratingSummary = false;
    console.error('Error generating email summary:', error);
    const result = {
      error: 'Failed to generate summary: ' + error.message,
      summaryItems: getDefaultSummaryItems(emailData),
      urgencyScore: 5 // Default middle urgency
    };
    summaryCache.set(cacheKey, result);
    return result;
  }
}

// Generate contact summary using AI
async function generateContactSummary(contactData, contactHistory) {
  // If no contact data provided, return an error
  if (!contactData || (!contactData.name && !contactData.email)) {
    return {
      error: 'No contact information provided',
      summaryItems: ['Error: No contact information provided'],
      messageCount: 0
    };
  }
  
  // Generate a cache key using available contact information and contact history length
  const cacheKey = `${contactData.name || ''}-${contactData.email || ''}-${contactHistory.length}`;
  
  // Check if we already have a cached summary for this contact and number of messages
  if (contactCache.has(cacheKey)) {
    console.log('Using cached contact summary for:', cacheKey);
    return contactCache.get(cacheKey);
  }
  
  // Sort contact history by date (newest first)
  const sortedHistory = [...contactHistory].sort((a, b) => {
    return new Date(b.timestamp) - new Date(a.timestamp);
  });
  
  const messageCount = sortedHistory.length;
  
  try {
    // Get API key from Chrome storage
    const apiKey = await getOpenAIApiKey();
    
    if (!apiKey) {
      console.error('No OpenAI API key found');
      return {
        error: 'OpenAI API key not configured',
        summaryItems: ['Error: API key not configured. Please set up your OpenAI API key in the settings.'],
        messageCount
      };
    }
    
    // Create a summary of the contact history
    // Limit the number of emails to analyze to avoid token limits (max 5)
    const historyToAnalyze = sortedHistory.slice(0, 5);
    
    // Format the messages for the API
    const formattedHistory = historyToAnalyze.map((email, index) => {
      return `Email ${index + 1} (${new Date(email.timestamp).toLocaleDateString()}):\nSubject: ${email.subject || 'No Subject'}\nFrom: ${email.sender_name || 'Unknown'} (${email.sender_email || 'No email'})\nContent: ${email.content || email.message_body || 'No content'}`;
    }).join('\n\n');
    
    // Log the formatted history being sent to OpenAI
    console.log('Formatted contact history being sent to OpenAI:');
    console.log(formattedHistory);
    
    // Prepare the API request
    console.log('Making API request to OpenAI for contact summary...');
    console.log('API key for contact summary:', apiKey ? `${apiKey.substring(0, 5)}...${apiKey.substring(apiKey.length - 4)}` : 'No key');
    
    // Create the full payload to send
    const payload = {
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'system',
          content: 'You are an intelligent email assistant. Analyze the following email history with a contact and create a concise summary of the key points from past interactions. Format your response as a bulleted list of 1-4 key insights about past communications. Each bullet point should be a separate insight. Be concise and focus on practical information that would be useful for the user to know before responding to this contact, dont comment on e.g the contact email styles only focus on the "facts" of what has been written. You might get duplicate data, so make sure to remove / ignore duplicates, no need to comment on this if you recieved duplicate data. Dont force the summary points, only write what is important. On the lastbullet point, comment on when the last interaction was. '
        },
        {
          role: 'user',
          content: `Please analyze the following email history with ${contactData.name || contactData.email} and provide a summary of the key points:\n\n${formattedHistory}`
        }
      ],
      temperature: 0.5
    };
    
    // Log the full payload being sent to OpenAI
    console.log('Full payload being sent to OpenAI:');
    console.log(JSON.stringify(payload, null, 2));
    
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify(payload)
    });
    
    // Process the API response
    const data = await response.json();
    
    if (!response.ok) {
      console.error('OpenAI API error:', data);
      throw new Error(data.error?.message || 'Failed to generate contact summary');
    }
    
    // Extract the summary text from the response
    const summaryText = data.choices[0]?.message?.content || '';
    
    // Convert the summary text to a list of bullet points
    const summaryItems = summaryText
      .split('\n')
      .filter(line => line.trim().startsWith('•') || line.trim().startsWith('-') || line.trim().startsWith('*'))
      .map(line => line.trim().replace(/^[•\-*]\s*/, ''));
    
    // If no bullet points were found, try to split by newlines
    const finalSummaryItems = summaryItems.length > 0 ? 
      summaryItems : 
      summaryText.split('\n').filter(line => line.trim().length > 0);
    
    // Store the result in the cache
    const result = {
      summaryItems: finalSummaryItems,
      messageCount
    };
    
    contactCache.set(cacheKey, result);
    console.log('Contact summary cached with key:', cacheKey);
    
    return result;
  } catch (error) {
    console.error('Error generating contact summary:', error);
    
    return {
      error: error.message || 'Failed to generate contact summary',
      summaryItems: ['Error: Failed to generate contact summary. Please try again later.'],
      messageCount
    };
  }
}

// Generate reply suggestion using AI
async function generateReplySuggestion(emailData, options = {}) {
  try {
    // Set default options if not provided
    const tone = options.tone || 'professional';
    const length = options.length || 'standard';
    const language = options.language || 'english';
    const additionalContext = options.additionalContext || '';
    
    console.log('Reply options:', { tone, length, language, hasAdditionalContext: !!additionalContext });
    if (language === 'auto') {
      console.log('Auto-detect language mode is enabled - will instruct AI to match email language');
    }
    
    // Get user's name from storage
    const { firstName, lastName } = await chrome.storage.sync.get(['firstName', 'lastName']);
    const userSignature = (firstName || lastName) ? `${firstName || ''} ${lastName || ''}`.trim() : '';
    console.log('User signature for email:', userSignature);
    
    // Generate a cache key including the options and signature
    const threadHistoryLength = emailData.threadHistory ? emailData.threadHistory.length : 0;
    // Include a hash of additional context in cache key
    const additionalContextHash = additionalContext ? 
      (additionalContext.length.toString() + '-' + additionalContext.slice(0, 10).replace(/\s+/g, '')) : 
      'none';
    const cacheKey = `${emailData.subject}-${emailData.sender.email}-${threadHistoryLength}-${tone}-${length}-${language}-${userSignature}-${additionalContextHash}-reply`;
    
    // Check if we have this in the cache already
    if (replyCache.has(cacheKey)) {
      console.log('Using cached reply for:', cacheKey);
      return replyCache.get(cacheKey);
    }
    
    const apiKey = await getOpenAIApiKey();
    
    if (!apiKey) {
      console.error('No OpenAI API key found');
      return getDefaultReply(emailData, userSignature);
    }
    
    // Prepare thread history content if available
    let threadContent = '';
    if (emailData.threadHistory && emailData.threadHistory.length > 0) {
      threadContent = '\n\nThread History:\n';
      emailData.threadHistory.forEach((message, index) => {
        threadContent += `\n--- Previous Message ${index + 1} ---\n`;
        threadContent += `From: ${message.sender.name || 'Unknown'} (${message.sender.email || 'No email'})\n`;
        threadContent += `Date: ${message.timestamp || 'Unknown date'}\n`;
        threadContent += `Content: ${message.message || 'No content'}\n`;
      });
    }
    
    // Add additional context if provided
    let additionalContextSection = '';
    if (additionalContext) {
      additionalContextSection = '\n\nAdditional Context (not in the email but relevant to the reply):\n' + additionalContext;
    }
    
    // Customize the system message based on tone, length, and language
    let systemMessage = 'You are an assistant that drafts professional email replies.';
    
    // Adjust for tone
    switch (tone) {
      case 'friendly':
        systemMessage = 'You are an assistant that drafts friendly and approachable email replies. Use a warm, conversational tone while still being professional.';
        break;
      case 'formal':
        systemMessage = 'You are an assistant that drafts formal email replies. Use proper business language, avoid contractions, and maintain a respectful distance.';
        break;
      case 'casual':
        systemMessage = 'You are an assistant that drafts casual email replies. Use a relaxed, conversational tone with common expressions and informal language where appropriate.';
        break;
      case 'concise':
        systemMessage = 'You are an assistant that drafts concise email replies. Be direct and to the point while still addressing all necessary points.';
        break;
      default: // professional is default
        systemMessage = 'You are an assistant that drafts professional email replies. Keep the tone friendly but professional.';
    }
    
    // Adjust for length
    switch (length) {
      case 'brief':
        systemMessage += ' Keep responses brief and to the point, ideally 2-3 sentences.';
        break;
      case 'detailed':
        systemMessage += ' Provide detailed responses that thoroughly address all points raised in the email.';
        break;
      default: // standard is default
        systemMessage += ' Provide a balanced response with appropriate detail.';
    }
    
    // Adjust for language
    if (language === 'auto') {
      // Create an enhanced system message for language auto-detection
      systemMessage = `You are a language detection and email reply assistant. This is a two-step process:
1. FIRST: Identify the language of the original email. Think carefully about the text and determine what language it is written in.
2. SECOND: Draft a reply in EXACTLY the same language as the original email. Do not translate to English - use the original language.

For style: ${systemMessage.substring(systemMessage.indexOf('You are an assistant that drafts') + 'You are an assistant that drafts'.length)}`;
      
      console.log('Enhanced auto-detect language system message:', systemMessage);
    } else if (language !== 'english') {
      // Capitalize first letter of language
      const formattedLanguage = language.charAt(0).toUpperCase() + language.slice(1);
      systemMessage += ` Write your response entirely in ${formattedLanguage}.`;
    }
    
    // Add signature instructions if we have the user's name
    if (userSignature) {
      systemMessage += ` Always end your email with a signature using the name "${userSignature}" (e.g. "Regards, ${userSignature}" or "Best regards, ${userSignature}").`;
    } else {
      systemMessage += ' End with an appropriate signature (e.g. "Regards," or "Best regards,").';
    }
    
    // Add additional context instruction to system message
    if (additionalContext) {
      systemMessage += ' Pay special attention to the additional context provided and incorporate it into your reply appropriately.';
    }
    
    // Add general guidelines
    systemMessage += ' If email thread history is provided, incorporate that context into your reply.';
    
    // Prepare the user content with or without language detection instruction
    let userContent = '';
    
    if (language === 'auto') {
      userContent = `STEP 1: What language is this email written in?
STEP 2: Write your reply in EXACTLY that same language.

Email to reply to:
From: ${emailData.sender.name} (${emailData.sender.email})
Subject: ${emailData.subject}
Date: ${emailData.date || emailData.timestamp}
Body: ${emailData.message || emailData.body}${threadContent}${additionalContextSection}`;
    } else {
      userContent = `Please draft a reply to this email:
From: ${emailData.sender.name} (${emailData.sender.email})
Subject: ${emailData.subject}
Date: ${emailData.date || emailData.timestamp}
Body: ${emailData.message || emailData.body}${threadContent}${additionalContextSection}`;
    }
    
    if (language === 'auto') {
      console.log('Auto-detect language user prompt:', userContent.substring(0, 200) + '...');
    }
    
    // Make a real API call to OpenAI's ChatGPT
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: 'gpt-3.5-turbo',
        messages: [
          {
            role: 'system',
            content: systemMessage
          },
          {
            role: 'user',
            content: userContent
          }
        ],
        temperature: 0.7,
        max_tokens: length === 'detailed' ? 800 : (length === 'brief' ? 300 : 500)
      })
    });

    const data = await response.json();
    
    if (data.error) {
      console.error('OpenAI API error:', data.error);
      return {
        error: data.error.message,
        replyText: getDefaultReply(emailData, userSignature)
      };
    }

    // Extract the response 
    const replyText = data.choices[0].message.content.trim();
    
    // Clean up the response for auto-detect language mode
    let cleanedReplyText = replyText;
    if (language === 'auto') {
      // Log the first few lines for debugging
      const firstLines = replyText.split('\n').slice(0, 3).join('\n');
      console.log('First few lines of AI response:', firstLines);
      
      // Check for and remove language identification lines
      // Common patterns: 
      // "The email is written in [Language]."
      // "Language: [Language]"
      // "STEP 1: This email is in [Language]"
      // "This email is written in [Language]."
      const languageIdentificationRegexes = [
        /^(Language|STEP 1|The language|The email is written in|This email is written in|The email is in)[^\n]+\n+/i,
        /^[^\n]*language[^\n]*:\s*[^\n]+\n+/i,
        /^[^\n]*(?:written|composed) in[^\n]+\n+/i,
        /^STEP 1:[^\n]*\nSTEP 2:[^\n]*\n+/i
      ];
      
      // Try each regex and remove any matches
      for (const regex of languageIdentificationRegexes) {
        cleanedReplyText = cleanedReplyText.replace(regex, '');
      }
      
      // If we modified the text, log what was removed
      if (cleanedReplyText !== replyText) {
        console.log('Removed language identification prefix from response');
        console.log('Original first 100 chars:', replyText.substring(0, 100));
        console.log('Cleaned first 100 chars:', cleanedReplyText.substring(0, 100));
      }
    }

    // Cache the result before returning
    const result = {
      replyText: cleanedReplyText
    };
    
    replyCache.set(cacheKey, result);
    return result;
  } catch (error) {
    console.error('Error generating reply suggestion:', error);
    return {
      error: 'Failed to generate reply',
      replyText: getDefaultReply(emailData, userSignature)
    };
  }
}

// Get default summary items when no API key is available
function getDefaultSummaryItems(emailData) {
  return [
    'You replied to the proposal last Friday.',
    `${emailData.sender.name} requested closing details on April 18.`,
    'The contract was sent to her on March 27.'
  ];
}

// Get default contact summary when no API key is available
function getDefaultContactSummary(contactData) {
  return [
    `${contactData.name || contactData.email || 'This contact'} has previous emails in your records.`,
    'The most common topics include project updates and meeting requests.',
    'Typical response time is 1-2 business days.'
  ];
}

// Get default reply when no API key is available
function getDefaultReply(emailData, userSignature = '') {
  const signature = userSignature ? `Best regards,\n${userSignature}` : 'Best regards,';
  return `Hi ${emailData.sender.name},\n\nSure, I'll get those details over to you shortly.\n\n${signature}`;
}

// Function to store emails in Supabase
export async function saveEmailToSupabase(emailData) {
  return await storeEmail(emailData);
}

// Function to retrieve threads from Supabase
export async function getThreadsFromSupabase() {
  return await getThreads();
}

// Function to retrieve messages for a thread from Supabase
export async function getThreadMessagesFromSupabase(threadId) {
  return await getThreadMessages(threadId);
}

// Function to search for contact history
export async function getContactHistoryFromSupabase(contactData, teamId) {
  return await getContactHistory(contactData, teamId);
}

// Search through contact emails and current email for specific information
async function searchContactEmails(query, currentEmail, contactHistory) {
  try {
    if (!query || !currentEmail) {
      return { 
        success: false, 
        error: 'Missing query or email data'
      };
    }

    console.log('Searching emails with query:', query);
    console.log('Contact history length:', contactHistory ? contactHistory.length : 0);
    
    // Get API key
    const apiKey = await getOpenAIApiKey();
    
    if (!apiKey) {
      console.error('No OpenAI API key found');
      return { 
        success: false, 
        error: 'API key not configured' 
      };
    }
    
    // Get SUPABASE_URL and SUPABASE_KEY for user lookup
    let SUPABASE_URL, SUPABASE_KEY;
    try {
      const supabaseConfigModule = await import('./supabase-config.js');
      SUPABASE_URL = supabaseConfigModule.SUPABASE_URL;
      SUPABASE_KEY = supabaseConfigModule.SUPABASE_KEY;
    } catch (error) {
      console.error('Error importing Supabase config:', error);
    }
    
    // Prepare the current email data
    const currentEmailFormatted = {
      id: currentEmail.id || 'current-email',
      sender: currentEmail.sender ? 
        `${currentEmail.sender.name || 'Unknown'} (${currentEmail.sender.email || 'No email'})` : 
        'Unknown sender',
      subject: currentEmail.subject || 'No subject',
      date: currentEmail.date || currentEmail.timestamp || new Date().toISOString(),
      content: currentEmail.message || currentEmail.body || '',
      savedBy: currentEmail.existingMessage && currentEmail.existingMessage.user ? 
        currentEmail.existingMessage.user.name : 
        'Unknown user',
      savedAt: currentEmail.existingMessage && currentEmail.existingMessage.savedAt ? 
        new Date(currentEmail.existingMessage.savedAt).toLocaleString() : 
        'Unknown date'
    };
    
    // Log how the current email is formatted
    console.log('=== CURRENT EMAIL DATA PREPARATION ===');
    console.log('Current email existingMessage data:', currentEmail.existingMessage);
    console.log('savedBy calculation:', {
      hasExistingMessage: !!currentEmail.existingMessage,
      hasUser: !!(currentEmail.existingMessage && currentEmail.existingMessage.user),
      userName: currentEmail.existingMessage && currentEmail.existingMessage.user ? 
        currentEmail.existingMessage.user.name : 'N/A',
      finalSavedBy: currentEmailFormatted.savedBy
    });
    console.log('Formatted current email:', currentEmailFormatted);
    
    // Create a map to store user information
    const userMap = new Map();
    
    // If we have contact history, fetch user information for all messages
    if (contactHistory && contactHistory.length > 0 && SUPABASE_URL && SUPABASE_KEY) {
      try {
        // Get access token for authenticated requests
        const { accessToken } = await chrome.storage.sync.get(['accessToken']);
        
        if (!accessToken) {
          console.warn('No access token available for user information lookup');
        } else {
          // Extract all user IDs from contact history
          const userIds = contactHistory
            .map(message => message.user_id)
            .filter(id => id); // Filter out null/undefined
            
          // Only proceed if we have user IDs to look up
          if (userIds.length > 0) {
            // Remove duplicates
            const uniqueUserIds = [...new Set(userIds)];
            console.log(`Found ${uniqueUserIds.length} unique user IDs to look up:`, uniqueUserIds);
            
            // Format as a comma-separated list for the "in" operator
            const userIdList = uniqueUserIds.map(id => `"${id}"`).join(',');
            
            // Query the users table to get user information
            const usersQueryUrl = `${SUPABASE_URL}/rest/v1/users?id=in.(${userIdList})&select=id,first_name,last_name,email`;
            console.log('Users query URL:', usersQueryUrl);
            
            const usersResponse = await fetch(usersQueryUrl, {
              method: 'GET',
              headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_KEY,
                'Authorization': `Bearer ${accessToken}`
              }
            });
            
            if (usersResponse.ok) {
              const users = await usersResponse.json();
              console.log(`Retrieved information for ${users.length} users`);
              
              // Create a map of user_id to user name
              users.forEach(user => {
                const userName = `${user.first_name || ''} ${user.last_name || ''}`.trim() || user.email || 'Unknown user';
                userMap.set(user.id, userName);
              });
              
              console.log('User map created:', Object.fromEntries([...userMap.entries()]));
            } else {
              console.error('Failed to retrieve user information:', await usersResponse.text());
            }
          }
        }
      } catch (userLookupError) {
        console.error('Error retrieving user information:', userLookupError);
      }
    }
    
    // Format the saved emails from contact history
    let formattedHistory = [];
    
    if (contactHistory && contactHistory.length > 0) {
      // Sort by date descending (newest first)
      const sortedHistory = [...contactHistory].sort((a, b) => {
        return new Date(b.timestamp || b.created_at || 0) - new Date(a.timestamp || a.created_at || 0);
      });
      
      formattedHistory = sortedHistory.map((email, index) => {
        // Get user name from the map if available
        let savedByName = 'Unknown user';
        if (email.user_id && userMap.has(email.user_id)) {
          savedByName = userMap.get(email.user_id);
        }
        
        return {
          id: email.id || `history-${index}`,
          sender: email.sender_name ? 
            `${email.sender_name} (${email.sender_email || 'No email'})` : 
            email.sender_email || 'Unknown sender',
          subject: email.subject || 'No subject',
          date: email.timestamp || email.created_at || 'Unknown date',
          content: email.message_body || email.content || '',
          savedBy: savedByName,
          savedAt: email.created_at ? 
            new Date(email.created_at).toLocaleString() : 
            'Unknown date'
        };
      });
    }
    
    // Combine current email and history
    const allEmails = [currentEmailFormatted, ...formattedHistory];
    
    // Format the emails for the API call
    const emailsContent = allEmails.map((email, index) => {
      // Create the formatted email content for this email
      const formattedEmail = `
Email #${index + 1}:
Sender: ${email.sender}
Subject: ${email.subject}
Date: ${email.date}
Saved by: ${email.savedBy}
Saved at: ${email.savedAt}
Content: ${email.content.substring(0, 1000)}${email.content.length > 1000 ? '...(truncated)' : ''}
      `;
      
      // Log the exact format of each email going to the API
      console.log(`Formatted email #${index + 1} for API:`, {
        sender: email.sender,
        subject: email.subject,
        date: email.date,
        savedBy: email.savedBy,
        savedAt: email.savedAt
      });
      
      return formattedEmail;
    }).join('\n\n');
    
    // Add detailed log for each email being searched
    console.log('=== ALL EMAILS BEING SEARCHED ===');
    allEmails.forEach((email, index) => {
      console.log(`Email #${index + 1}:`);
      console.log(`  Sender: ${email.sender}`);
      console.log(`  Subject: ${email.subject}`);
      console.log(`  Date: ${email.date}`);
      console.log(`  Saved by: ${email.savedBy}`);
      console.log(`  Saved at: ${email.savedAt}`);
    });
    
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
    
    // Log the full API payload for debugging
    console.log('=== SEARCH EMAIL API CALL ===');
    console.log('Query:', query);
    console.log('System Message:', systemMessage);
    console.log('Email Content Length:', emailsContent.length, 'characters');
    console.log('First 200 chars of email content:', emailsContent.substring(0, 200) + '...');
    console.log('Number of emails included:', allEmails.length);
    console.log('========================');

    // Make the API call
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
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
      })
    });

    const data = await response.json();
    
    if (data.error) {
      console.error('OpenAI API error:', data.error);
      return {
        success: false,
        error: data.error.message
      };
    }

    // Extract the response
    const answer = data.choices[0].message.content.trim();
    console.log('Search results:', answer);
    
    // Parse the answer to split it into the summary and references
    let mainAnswer = '';
    let references = [];
    
    // Try to extract the main answer (everything before 'References:')
    const referenceSplit = answer.split(/references:/i);
    if (referenceSplit.length > 1) {
      mainAnswer = referenceSplit[0].trim();
      console.log('Extracted main answer:', mainAnswer);
      
      // Extract references
      const referencesText = referenceSplit[1].trim();
      console.log('Raw references text:', referencesText);
      
      // Improved reference extraction using a more robust approach
      // Look for numbered items (1., 2., etc.) as reference starters
      const referenceLines = referencesText.split(/\n+/).filter(line => line.trim());
      console.log('Reference lines after splitting:', referenceLines);
      
      // Process each reference line
      for (let i = 0; i < referenceLines.length; i++) {
        const line = referenceLines[i].trim();
        console.log(`Processing reference line ${i+1}:`, line);
        
        // Skip empty lines or lines that don't start with a number
        if (!line || !/^\d+\./.test(line)) {
          console.log(`Skipping line: "${line}" - not a numbered reference`);
          continue;
        }
        
        // Extract the number at the beginning
        const numberMatch = line.match(/^(\d+)\./);
        const refNumber = numberMatch ? numberMatch[1] : (i + 1).toString();
        console.log(`Reference number: ${refNumber}`);
        
        // Extract the quote part (between quotes)
        const quoteMatch = line.match(/"([^"]+)"/);
        let quote = quoteMatch ? quoteMatch[1] : '';
        console.log(`Quote match:`, quoteMatch ? quoteMatch[0] : 'No quoted text found');
        
        // Extract metadata (after the quote)
        let meta = '';
        if (quoteMatch) {
          const afterQuote = line.slice(line.indexOf(quoteMatch[0]) + quoteMatch[0].length).trim();
          console.log(`Text after quote: "${afterQuote}"`);
          // Check if the metadata starts with a dash
          if (afterQuote.startsWith('-')) {
            meta = afterQuote.substring(1).trim();
          } else {
            meta = afterQuote;
          }
        } else {
          // If no quotes were found, try to split by a dash
          const parts = line.split(/\s+-\s+/);
          console.log(`Split by dash:`, parts);
          if (parts.length > 1) {
            // Remove the number from the first part
            quote = parts[0].replace(/^\d+\.\s*/, '').trim();
            meta = parts[1].trim();
          } else {
            // Just use the whole line without the number
            quote = line.replace(/^\d+\.\s*/, '').trim();
          }
        }
        
        console.log(`Final extracted quote: "${quote}"`);
        console.log(`Final extracted meta: "${meta}"`);
        
        // Add to references array
        references.push({ 
          quote, 
          meta: meta || 'No metadata available'
        });
      }
    } else {
      // If no "References:" section, take the whole answer
      mainAnswer = answer;
      console.log('No references section found, using entire answer');
    }
    
    // Ensure we have valid references
    references = references.filter(ref => ref.quote);
    
    console.log('Parsed references:', references);
    
    return {
      success: true,
      mainAnswer,
      references
    };
  } catch (error) {
    console.error('Error searching emails:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// Export functions
export { 
  getApiKey, 
  generateEmailSummary, 
  generateReplySuggestion,
  generateContactSummary,
  searchContactEmails,
  getOpenAIApiKey
};