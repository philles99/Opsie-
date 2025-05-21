/**
 * Gmail Parser Utility
 * This file contains functions for parsing Gmail's DOM to extract email data
 */

// Extract email data from Gmail
function extractGmailEmailData(emailContainer) {
  try {
    // Get sender info
    const senderElement = emailContainer.querySelector('.gD');
    const senderName = senderElement ? senderElement.getAttribute('name') : 'Unknown';
    const senderEmail = senderElement ? senderElement.getAttribute('email') : 'unknown@email.com';
    
    // Get subject
    const subjectElement = document.querySelector('h2.hP');
    const subject = subjectElement ? subjectElement.textContent : 'No Subject';
    
    // Get timestamp
    const timestampElement = emailContainer.querySelector('.g3');
    const timestamp = timestampElement ? timestampElement.textContent : '';
    
    // Get message content
    const messageElement = emailContainer.querySelector('.a3s.aiL');
    const message = messageElement ? messageElement.textContent.trim() : '';
    
    // Get thread history if available
    const threadHistory = extractGmailThreadHistory();
    
    // Get message URL for external ID extraction
    const messageUrl = window.location.href;
    console.log('Captured Gmail message URL:', messageUrl);
    
    // Try to extract message ID directly from the URL if possible
    let messageId = null;
    try {
      // Gmail format: https://mail.google.com/mail/u/0/?hl=sv#inbox/FMfcgzQbdrTrrfCfrVdgqZwbpChKPdPC
      const gmailMatches = messageUrl.match(/[#/](?:inbox|sent|drafts|trash|spam|category\/\w+)\/([^/?#]+)/i);
      if (gmailMatches && gmailMatches[1]) {
        messageId = gmailMatches[1];
        console.log('Extracted Gmail message ID from URL:', messageId);
      }
    } catch (e) {
      console.error('Error extracting Gmail message ID from URL:', e);
    }
    
    return {
      sender: {
        name: senderName,
        email: senderEmail
      },
      subject: subject,
      timestamp: timestamp,
      message: message,
      threadHistory: threadHistory,
      messageUrl: messageUrl,
      messageId: messageId // Include the extracted message ID
    };
  } catch (error) {
    console.error('Error extracting Gmail email data:', error);
    return null;
  }
}

// Extract thread history from Gmail
function extractGmailThreadHistory() {
  try {
    const threadContainer = document.querySelectorAll('.adn.ads');
    const threadHistory = [];
    
    if (threadContainer && threadContainer.length > 0) {
      threadContainer.forEach(emailBlock => {
        // Skip if this is the currently opened email (usually the last one)
        if (emailBlock === document.querySelector('.adn.ads:last-child')) {
          return;
        }
        
        const senderElement = emailBlock.querySelector('.gD');
        const senderName = senderElement ? senderElement.getAttribute('name') : 'Unknown';
        const senderEmail = senderElement ? senderElement.getAttribute('email') : 'unknown@email.com';
        
        const timestampElement = emailBlock.querySelector('.g3');
        const timestamp = timestampElement ? timestampElement.textContent : '';
        
        const messageElement = emailBlock.querySelector('.a3s.aiL');
        const message = messageElement ? messageElement.textContent.trim() : '';
        
        threadHistory.push({
          sender: {
            name: senderName,
            email: senderEmail
          },
          timestamp: timestamp,
          message: message
        });
      });
    }
    
    return threadHistory;
  } catch (error) {
    console.error('Error extracting Gmail thread history:', error);
    return [];
  }
}

// Export functions
export { extractGmailEmailData, extractGmailThreadHistory }; 