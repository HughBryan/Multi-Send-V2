// Global recipients array

let recipients = [];



// Initialize the taskpane

document.addEventListener('DOMContentLoaded', function () {
    updateCountText();
    setupPasteHandler();
    renderRecipients();

    document.querySelector('.add-email-btn')?.addEventListener('click', addRecipient);
    document.getElementById('clearAllBtn')?.addEventListener('click', clearAllAndReset);
    document.getElementById('generateBtn')?.addEventListener('click', generateEmails);
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

window.chrome.webview.addEventListener('message', function (event) {
    try {
        console.log('Received message from C#:', event.data);
        const response = JSON.parse(event.data);
        handleBackendResponse(response);
    } catch (error) {
        console.error('Error handling backend response:', error, 'Raw data:', event.data);
    }
});


function handleBackendResponse(response) {
    // Clear safety timeout when we receive any response
    if (window.currentOperationTimeout) {
        clearTimeout(window.currentOperationTimeout);
        window.currentOperationTimeout = null;
    }

    if (response.type === 'success') {
        showStatus(response.message, 'success');
        resetUI();
    } else if (response.type === 'error') {
        showStatus(response.message, 'error');
        resetUI();
    } else if (response.type === 'progress') {
        updateProgress(response.current, response.total, response.message);
    } else if (response.type === 'info') {
        showStatus(response.message, 'info');
    } else if (response.type === 'placeholderWarning') {
        // Handle placeholder warning - show confirmation dialog
        handlePlaceholderWarning(response.message, response.data);
    } else if (response.type === 'emailData') {
        console.log('Received email data:', response.data);
    }
}
function handlePlaceholderWarning(message, data) {
    // Reset UI first
    resetUI();

    // Show the warning message
    showStatus(message, 'error');

    // Create confirmation buttons
    const statusElement = document.getElementById('status');

    // Add confirmation buttons to the status message
    const buttonContainer = document.createElement('div');
    buttonContainer.style.marginTop = '10px';

    const continueBtn = document.createElement('button');
    continueBtn.textContent = 'Continue Anyway';
    continueBtn.className = 'generate-btn';
    continueBtn.style.marginRight = '10px';
    continueBtn.style.fontSize = '12px';
    continueBtn.style.padding = '8px 16px';
    continueBtn.onclick = () => confirmPlaceholderWarning(data, true);

    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    cancelBtn.className = 'clear-all-btn';
    cancelBtn.style.fontSize = '12px';
    cancelBtn.style.padding = '8px 16px';
    cancelBtn.onclick = () => cancelPlaceholderWarning();

    buttonContainer.appendChild(continueBtn);
    buttonContainer.appendChild(cancelBtn);
    statusElement.appendChild(buttonContainer);
}

function confirmPlaceholderWarning(originalData, forceWithoutPlaceholder) {
    // Clear the warning message
    hideStatus();

    // Show progress again
    document.getElementById("progress").classList.add("show");
    document.getElementById("generateBtn").disabled = true;

    // Add safety timeout again
    const safetyTimeout = setTimeout(() => {
        console.warn("No response received within 30 seconds, resetting UI");
        showStatus("Operation timed out. Please try again.", 'error');
        resetUI();
    }, 30000);

    window.currentOperationTimeout = safetyTimeout;

    // Send the request again with forceWithoutPlaceholder flag
    sendMessageToCS('duplicateEmail', {
        placeholder: originalData.placeholder,
        recipients: originalData.recipients,
        autoSend: originalData.autoSend,
        forceWithoutPlaceholder: forceWithoutPlaceholder
    });
}

function cancelPlaceholderWarning() {
    // Just hide the warning and reset to normal state
    hideStatus();
}

async function generateEmails() {
    try {
        const placeholder = document.getElementById("placeholder").value.trim();
        const filledRecipients = recipients.filter(r => r.email.trim() && r.name.trim());
        const autoSend = document.getElementById("autoSendCheckbox").checked;

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

        // Add a safety timeout to reset UI if no response comes back
        const safetyTimeout = setTimeout(() => {
            console.warn("No response received within 30 seconds, resetting UI");
            showStatus("Operation timed out. Please try again.", 'error');
            resetUI();
        }, 30000); // 30 second timeout

        // Store timeout ID so we can clear it when we get a response
        window.currentOperationTimeout = safetyTimeout;

        // Send duplication request to C# backend
        sendMessageToCS('duplicateEmail', {
            placeholder: placeholder,
            recipients: filledRecipients,
            autoSend: autoSend  // Pass the auto-send flag to C#
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



// Replace renderRecipients() body with safe DOM creation
function renderRecipients() {
    const emailList = document.getElementById('emailList');
    emailList.innerHTML = '';
    recipients.forEach((r, i) => {
        const row = document.createElement('div');
        row.className = 'email-input-row';

        const email = document.createElement('input');
        email.type = 'email';
        email.className = 'email-input';
        email.placeholder = 'email@example.com';
        email.value = r.email || '';
        email.oninput = () => { updateRecipient(i, 'email', email.value); autoDetectName(i); };

        const name = document.createElement('input');
        name.type = 'text';
        name.className = 'name-input';
        name.placeholder = 'Name';
        name.value = r.name || '';
        name.oninput = () => updateRecipient(i, 'name', name.value);

        const autoBtn = document.createElement('button');
        autoBtn.className = 'suggest-btn';
        autoBtn.textContent = 'Auto';
        autoBtn.onclick = () => suggestName(i);

        const rm = document.createElement('button');
        rm.className = 'remove-btn';
        rm.textContent = '×';
        rm.onclick = () => removeRecipient(i);

        row.append(email, name, autoBtn, rm);
        emailList.appendChild(row);
    });
}

let autoDetectTimeouts = {};
function autoDetectName(index) {
    clearTimeout(autoDetectTimeouts[index]);
    autoDetectTimeouts[index] = setTimeout(() => {
        if (recipients[index]?.email && !recipients[index]?.name) {
            const email = recipients[index].email;
            // Only auto-detect if email looks complete (has @ and a dot after it)
            if (email.includes('@') && email.split('@')[1]?.includes('.')) {
                suggestNameForIndex(index, false);
            }
        }
    }, 1000); // Increased from 300ms to 1000ms
}

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
        
        // FIXED: Update the specific input field instead of re-rendering everything
        const nameInput = document.querySelector(`.email-input-row:nth-child(${index + 1}) .name-input`);
        if (nameInput) {
            nameInput.value = suggestedName;
        }
        
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
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    const parsed = [];

    lines.forEach(line => {
        const parts = line.split(/[\t,]/).map(p => p.trim()).filter(Boolean);

        if (parts.length >= 2) {
            let email = '';
            let name = '';
            parts.forEach(part => {
                if (!email && part.includes('@') && isValidEmail(part)) email = part;
                else if (!name) name = part;
            });
            if (email) parsed.push({ email, name: name || '' });
        } else if (parts.length === 1) {
            const part = parts[0];
            if (part.includes('@') && isValidEmail(part)) {
                parsed.push({ email: part, name: '' });
            }
        }
    });

    if (parsed.length === 0) return;

    // Keep only non-empty existing recipients
    recipients = recipients.filter(r => r.email.trim() || r.name.trim());

    // Add unique emails only
    const existing = new Set(recipients.map(r => (r.email || '').toLowerCase()));
    let added = 0;
    let addedWithNames = 0;

    parsed.forEach(nr => {
        const key = (nr.email || '').toLowerCase();
        if (key && !existing.has(key)) {
            recipients.push(nr);
            existing.add(key);
            added++;
            if (nr.name) addedWithNames++;
        }
    });

    renderRecipients();
    updateCountText();

    // Auto-suggest names where missing
    setTimeout(() => {
        recipients.forEach((r, i) => {
            if (r.email && !r.name) {
                suggestNameForIndex(i, false);
            }
        });
    }, 100);

    if (added > 0) {
        if (addedWithNames > 0) {
            showStatus(`Added ${added} recipients (${addedWithNames} with names)`, 'success');
        } else {
            showStatus(`Added ${added} recipients`, 'success');
        }
        setTimeout(hideStatus, 3000);
    }
}

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

    hideStatus();
}