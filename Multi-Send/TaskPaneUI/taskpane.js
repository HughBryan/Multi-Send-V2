// Global recipients array

let recipients = [];



// Initialize the taskpane

document.addEventListener('DOMContentLoaded', function () {
    // Start with no recipients - users can add them manually or paste
    updateCountText();
    setupPasteHandler();
    renderRecipients(); // Render empty list
});



// Communication with C# backend

function sendMessageToCS(action, data = null) {

    try {

        const message = {

            action: action,

            data: data

        };

        window.chrome.webview.postMessage(message);

    } catch (error) {

        console.error('Error sending message to C#:', error);

        showStatus('Error communicating with add-in backend', 'error');

    }

}



// Listen for messages from C# backend

window.addEventListener('message', function (event) {

    try {

        const response = JSON.parse(event.data);

        handleBackendResponse(response);

    } catch (error) {

        console.error('Error handling backend response:', error);

    }

});



function handleBackendResponse(response) {

    if (response.type === 'success') {

        showStatus(response.message, 'success');

    } else if (response.type === 'error') {

        showStatus(response.message, 'error');

    } else if (response.type === 'progress') {

        updateProgress(response.current, response.total, response.message);

    } else if (response.type === 'emailData') {

        // Handle email data received from backend

        console.log('Received email data:', response.data);

    }

}



async function generateEmails() {

    try {

        const placeholder = document.getElementById("placeholder").value.trim();

        const filledRecipients = recipients.filter(r => r.email.trim() && r.name.trim());



        if (filledRecipients.length === 0) {

            showStatus("Please add at least one recipient with both email and name.", 'error');

            return;

        }



        if (!placeholder) {

            showStatus("Please enter a placeholder (e.g., {{name}}) that will be replaced with each person's name.", 'error');

            return;

        }



        // Show progress

        document.getElementById("progress").classList.add("show");

        document.getElementById("generateBtn").disabled = true;



        // Send duplication request to C# backend

        sendMessageToCS('duplicateEmail', {

            placeholder: placeholder,

            recipients: filledRecipients

        });



    } catch (error) {

        console.error("Error in generateEmails:", error);

        showStatus("Error: " + error.message, 'error');

        resetUI();

    }

}



function addRecipient() {

    const recipient = { email: '', name: '' };

    recipients.push(recipient);

    renderRecipients();

    updateCountText();

}



function removeRecipient(index) {

    recipients.splice(index, 1);

    renderRecipients();

    updateCountText();

}



function renderRecipients() {
    const emailList = document.getElementById('emailList');
    emailList.innerHTML = '';

    recipients.forEach((recipient, index) => {
        const row = document.createElement('div');
        row.className = 'email-input-row';

        row.innerHTML = `
            <input type="email" 
                   class="email-input" 
                   placeholder="email@example.com"
                   value="${recipient.email}"
                   oninput="updateRecipient(${index}, 'email', this.value); autoDetectName(${index});">
            <input type="text" 
                   class="name-input" 
                   placeholder="Name"
                   value="${recipient.name}"
                   oninput="updateRecipient(${index}, 'name', this.value)">
            <button class="suggest-btn" onclick="suggestName(${index})">Auto</button>
            <button class="remove-btn" onclick="removeRecipient(${index})">×</button>
        `;

        emailList.appendChild(row);
    });
}

// Add this at the top of your file with other global variables
let autoDetectTimeouts = {};




function updateRecipient(index, field, value) {

    if (recipients[index]) {

        recipients[index][field] = value;

        updateCountText();

    }

}

function suggestName(index) {
    suggestNameForIndex(index, true); // true = show status message for manual clicks
}

function suggestNameForIndex(index, shouldShowStatus = true) {
    const email = recipients[index].email;

    if (email && email.includes('@')) {
        const namePart = email.split('@')[0];

        const words = namePart.replace(/[._-]/g, ' ')
            .split(' ')
            .map(word => word.charAt(0).toUpperCase() + word.slice(1))
            .filter(word => word.length > 0);

        // Only use the FIRST name instead of joining all words
        const suggestedName = words.length > 0 ? words[0] : '';

        recipients[index].name = suggestedName;
        renderRecipients();
        updateCountText();

        // Only show status for manual button clicks, not auto-detection
        if (shouldShowStatus) {
            showStatus(`Auto-suggested name: ${suggestedName}`, 'success');
            setTimeout(hideStatus, 2000);
        }
    }
}



function clearAllAndReset() {

    recipients = [];

    addRecipient();

    updateCountText();

    hideStatus();

}



function updateCountText() {

    const filledCount = recipients.filter(r => r.email.trim() && r.name.trim()).length;

    const totalCount = recipients.length;

    document.getElementById('countText').textContent =

        `${filledCount}/${totalCount} recipients ready`;



    const clearBtn = document.getElementById('clearAllBtn');

    clearBtn.disabled = totalCount === 0;

}



function detectPlaceholder() {

    // Send request to C# to detect placeholder in current email

    sendMessageToCS('detectPlaceholder');

}



function setupPasteHandler() {
    document.addEventListener('paste', function (e) {
        const pastedText = e.clipboardData.getData('text');
        
        if (pastedText) {
            // Check if it contains emails or looks like structured data
            if (pastedText.includes('@') || pastedText.includes('\t') || pastedText.split('\n').length > 1) {
                showPasteHint();
                
                setTimeout(() => {
                    processPastedEmails(pastedText);
                    hidePasteHint();
                }, 100);
            }
        }
    });
}



function processPastedEmails(text) {
    const lines = text.split(/\r?\n/).filter(line => line.trim());
    const newRecipients = [];
    
    lines.forEach(line => {
        // Try to detect tab-separated or comma-separated values
        const parts = line.split(/[\t,]/).map(part => part.trim()).filter(part => part);
        
        if (parts.length >= 2) {
            // Two or more columns detected
            let email = '';
            let name = '';
            
            // Determine which part is email vs name
            parts.forEach(part => {
                if (part.includes('@') && isValidEmail(part)) {
                    email = part;
                } else if (!name) {
                    name = part;
                }
            });
            
            if (email) {
                newRecipients.push({ email: email, name: name || '' });
            }
        } else if (parts.length === 1) {
            // Single column - check if it's an email
            const part = parts[0];
            if (part.includes('@') && isValidEmail(part)) {
                newRecipients.push({ email: part, name: '' }); // Empty name for now
            }
        }
    });
    
    if (newRecipients.length > 0) {
        // Clear existing empty recipients
        recipients = recipients.filter(r => r.email.trim() || r.name.trim());
        
        // Add new recipients (avoid duplicates)
        newRecipients.forEach(newRecipient => {
            if (!recipients.some(r => r.email.toLowerCase() === newRecipient.email.toLowerCase())) {
                recipients.push(newRecipient);
            }
        });
        
        // Add new recipients (avoid duplicates)
        newRecipients.forEach(newRecipient => {
            if (!recipients.some(r => r.email.toLowerCase() === newRecipient.email.toLowerCase())) {
                recipients.push(newRecipient);
            }
        });

        // REMOVED: Auto-add empty row

        renderRecipients();
        updateCountText();
        
        renderRecipients();
        updateCountText();
        
        // NOW AUTO-DETECT NAMES FOR NEWLY ADDED EMAILS
        // Find recipients that have emails but no names and auto-suggest
        setTimeout(() => {
            recipients.forEach((recipient, index) => {
                if (recipient.email && !recipient.name) {
                    suggestNameForIndex(index, false); // Use your existing function!
                }
            });
        }, 100); // Small delay to ensure UI is rendered
        
        const emailCount = newRecipients.filter(r => r.email).length;
        const nameCount = newRecipients.filter(r => r.name).length;
        
        if (nameCount > 0) {
            showStatus(`Added ${emailCount} emails with ${nameCount} names from paste data`, 'success');
        } else {
            showStatus(`Added ${emailCount} email addresses with auto-detected names`, 'success');
        }
        
        setTimeout(hideStatus, 3000);
    }
}

// Add this helper function
function isValidEmail(email) {
    const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    return emailRegex.test(email);
}



function showPasteHint() {

    document.getElementById('pasteHint').classList.add('show');

}



function hidePasteHint() {

    document.getElementById('pasteHint').classList.remove('show');

}



function updateProgress(current, total, message) {

    const progressElement = document.getElementById("progress");

    const progressFill = document.getElementById("progressFill");

    const progressText = document.getElementById("progressText");



    progressElement.classList.add("show");



    const percentage = (current / total) * 100;

    progressFill.style.width = `${percentage}%`;

    progressText.textContent = message || `Processing ${current}/${total}...`;

}



function showStatus(message, type = 'info') {

    const statusElement = document.getElementById('status');

    statusElement.textContent = message;

    statusElement.className = `status show ${type}`;

}



function hideStatus() {

    const statusElement = document.getElementById('status');

    statusElement.classList.remove('show');

}



function resetUI() {

    document.getElementById("progress").classList.remove("show");

    document.getElementById("generateBtn").disabled = false;

    document.getElementById("progressFill").style.width = "0%";

}