/**
 * Authentication UI Functions
 * This file contains functions for handling authentication UI interactions
 */

// Authentication view switching functions
function showLoginView() {
    document.getElementById('login-view').style.display = 'block';
    document.getElementById('signup-view').style.display = 'none';
    document.getElementById('password-reset-view').style.display = 'none';
    
    // Clear any error messages
    const errorElement = document.getElementById('auth-error');
    if (errorElement) {
        errorElement.style.display = 'none';
    }
}

function showSignupView() {
    document.getElementById('login-view').style.display = 'none';
    document.getElementById('signup-view').style.display = 'block';
    document.getElementById('password-reset-view').style.display = 'none';
    
    // Clear any error messages
    const errorElement = document.getElementById('auth-error');
    if (errorElement) {
        errorElement.style.display = 'none';
    }
    
    // Focus on first name field
    const firstNameField = document.getElementById('signup-first-name');
    if (firstNameField) {
        firstNameField.focus();
    }
}

function showPasswordResetView() {
    document.getElementById('login-view').style.display = 'none';
    document.getElementById('signup-view').style.display = 'none';
    document.getElementById('password-reset-view').style.display = 'block';
    
    // Clear any error messages
    const errorElement = document.getElementById('auth-error');
    if (errorElement) {
        errorElement.style.display = 'none';
    }
    
    // Focus on email field
    const emailField = document.getElementById('reset-email');
    if (emailField) {
        emailField.focus();
    }
}

// Handle signup form submission
async function handleSignup() {
    try {
        // Disable signup controls
        disableSignupControls();
        
        // Show loading indicator
        const signupLoading = document.getElementById('signup-loading');
        if (signupLoading) {
            signupLoading.style.display = 'flex';
        }
        
        // Get form values
        const firstName = document.getElementById('signup-first-name').value.trim();
        const lastName = document.getElementById('signup-last-name').value.trim();
        const email = document.getElementById('signup-email').value.trim();
        const password = document.getElementById('signup-password').value;
        const confirmPassword = document.getElementById('signup-confirm-password').value;
        
        // Validate input
        if (!firstName || !lastName || !email || !password || !confirmPassword) {
            showAuthError('Please fill in all fields');
            return;
        }
        
        if (!isValidEmail(email)) {
            showAuthError('Please enter a valid email address');
            return;
        }
        
        if (password !== confirmPassword) {
            showAuthError('Passwords do not match');
            return;
        }
        
        if (password.length < 6) {
            showAuthError('Password must be at least 6 characters long');
            return;
        }
        
        console.log('Attempting signup with Supabase', { email, firstName, lastName });
        
        // Use the global signUp function from auth-service
        if (typeof window.OpsieApi === 'undefined' || typeof window.OpsieApi.signUp !== 'function') {
            showAuthError('Authentication service not available. Please refresh the page.');
            return;
        }
        
        const signupResult = await window.OpsieApi.signUp(email, password, firstName, lastName);
        
        if (signupResult.error) {
            showAuthError(signupResult.error || 'Signup failed');
            return;
        }
        
        console.log('Signup successful', signupResult);
        
        if (signupResult.success) {
            console.log('Signup successful');
            showAuthSuccess('Account created successfully! Please select or create a team.');
            
            // Clear form
            document.getElementById('signup-first-name').value = '';
            document.getElementById('signup-last-name').value = '';
            document.getElementById('signup-email').value = '';
            document.getElementById('signup-password').value = '';
            document.getElementById('signup-confirm-password').value = '';
            
            // Show team selection view after a short delay
            setTimeout(() => {
                if (window.showTeamSelectView) {
                    window.showTeamSelectView();
                } else {
                    console.error('showTeamSelectView function not available');
                    showAuthError('Team selection not available. Please refresh and try again.');
                }
            }, 1500);
        } else {
            showAuthError(signupResult.error || 'Signup failed');
        }
        
    } catch (error) {
        showAuthError(error.message || 'An error occurred during signup');
        console.error('Signup error:', error);
    } finally {
        enableSignupControls();
    }
}

// Handle password reset form submission
async function handlePasswordReset() {
    try {
        // Disable reset controls
        disableResetControls();
        
        // Show loading indicator
        const resetLoading = document.getElementById('reset-loading');
        if (resetLoading) {
            resetLoading.style.display = 'flex';
        }
        
        // Get email value
        const email = document.getElementById('reset-email').value.trim();
        
        // Validate input
        if (!email) {
            showAuthError('Please enter your email address');
            return;
        }
        
        if (!isValidEmail(email)) {
            showAuthError('Please enter a valid email address');
            return;
        }
        
        console.log('Attempting password reset', { email });
        
        // Use the global requestPasswordReset function from auth-service
        if (typeof window.OpsieApi === 'undefined' || typeof window.OpsieApi.requestPasswordReset !== 'function') {
            showAuthError('Authentication service not available. Please refresh the page.');
            return;
        }
        
        const resetResult = await window.OpsieApi.requestPasswordReset(email);
        
        if (resetResult.error) {
            showAuthError(resetResult.error || 'Password reset failed');
            return;
        }
        
        console.log('Password reset request successful', resetResult);
        
        // Show success message
        showAuthSuccess('Password reset email sent! Please check your email for instructions.');
        
        // Clear form
        document.getElementById('reset-email').value = '';
        
        // Switch to login view after a short delay
        setTimeout(() => {
            showLoginView();
        }, 2000);
        
    } catch (error) {
        showAuthError(error.message || 'An error occurred during password reset');
        console.error('Password reset error:', error);
    } finally {
        enableResetControls();
    }
}

// Control functions for signup form
function disableSignupControls() {
    const signupButton = document.getElementById('signup-button');
    const signupInputs = document.querySelectorAll('#signup-view input');
    
    if (signupButton) {
        signupButton.disabled = true;
        signupButton.textContent = 'Creating Account...';
    }
    
    signupInputs.forEach(input => {
        input.disabled = true;
    });
}

function enableSignupControls() {
    const signupButton = document.getElementById('signup-button');
    const signupInputs = document.querySelectorAll('#signup-view input');
    const signupLoading = document.getElementById('signup-loading');
    
    if (signupButton) {
        signupButton.disabled = false;
        signupButton.textContent = 'Create Account';
    }
    
    signupInputs.forEach(input => {
        input.disabled = false;
    });
    
    if (signupLoading) {
        signupLoading.style.display = 'none';
    }
}

// Control functions for password reset form
function disableResetControls() {
    const resetButton = document.getElementById('reset-button');
    const resetInputs = document.querySelectorAll('#password-reset-view input');
    
    if (resetButton) {
        resetButton.disabled = true;
        resetButton.textContent = 'Sending...';
    }
    
    resetInputs.forEach(input => {
        input.disabled = true;
    });
}

function enableResetControls() {
    const resetButton = document.getElementById('reset-button');
    const resetInputs = document.querySelectorAll('#password-reset-view input');
    const resetLoading = document.getElementById('reset-loading');
    
    if (resetButton) {
        resetButton.disabled = false;
        resetButton.textContent = 'Send Reset Email';
    }
    
    resetInputs.forEach(input => {
        input.disabled = false;
    });
    
    if (resetLoading) {
        resetLoading.style.display = 'none';
    }
}

// Show authentication error message
function showAuthError(message) {
    const errorElement = document.getElementById('auth-error');
    if (errorElement) {
        errorElement.textContent = message;
        errorElement.style.display = 'block';
        errorElement.className = 'auth-message error';
    }
}

// Show authentication success message
function showAuthSuccess(message) {
    const errorElement = document.getElementById('auth-error');
    if (errorElement) {
        errorElement.textContent = message;
        errorElement.style.display = 'block';
        errorElement.className = 'auth-message success';
    }
}

// Email validation function
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

// Make functions available globally
window.showLoginView = showLoginView;
window.showSignupView = showSignupView;
window.showPasswordResetView = showPasswordResetView;
window.handleSignup = handleSignup;
window.handlePasswordReset = handlePasswordReset; 