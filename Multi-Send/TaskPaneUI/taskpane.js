// Global recipients array
let recipients = [];

// Initialize the taskpane
document.addEventListener('DOMContentLoaded', function () {
    console.log('=== DOM LOADED ===');

    // Add one default recipient
    recipients.push({ email: '', name: '' });

    updateCountText();
    setupPasteHandler();
    renderRecipients();

    // Attach event listeners
    document.querySelector('.add-email-btn')?.addEventListener('click', addRecipient);
    document.getElementById('clearAllBtn')?.addEventListener('click', clearAllAndReset);
    document.getElementById('generateBtn')?.addEventListener('click', generateEmails);
    document.getElementById('confirmSendBtn')?.addEventListener('click', proceedWithSend);
    document.getElementById('cancelSendBtn')?.addEventListener('click', hideSendConfirmation);

    console.log('=== INITIALIZATION COMPLETE ===');
});

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
    console.log('Handling response type:', response.type);

    if (response.type === 'success') {
        // Show success in dialog if it's open, otherwise use normal status
        const progressElement = document.getElementById('sendingProgress');
        if (progressElement && progressElement.style.display !== 'none') {
            document.getElementById('confirmProgressText').textContent = response.message;
            document.getElementById('confirmProgressFill').style.width = '100%';
            setTimeout(() => {
                hideSendConfirmation();
                showStatus(response.message, 'success');
                resetUI();
            }, 1500);
        } else {
            showStatus(response.message, 'success');
            resetUI();
            hideSendConfirmation();
        }

    } else if (response.type === 'error') {
        // Show error in dialog if it's open, otherwise use normal status
        const progressElement = document.getElementById('sendingProgress');
        if (progressElement && progressElement.style.display !== 'none') {
            document.getElementById('confirmProgressText').textContent = response.message;
            setTimeout(() => {
                hideSendConfirmation();
                showStatus(response.message, 'error');
                resetUI();
            }, 2000);
        } else {
            showStatus(response.message, 'error');
            resetUI();
            hideSendConfirmation();
        }

    } else if (response.type === 'progress') {
        // Try to show progress in dialog first, fallback to main progress bar
        const dialogProgress = document.getElementById('sendingProgress');
        const dialogText = document.getElementById('confirmProgressText');
        const dialogFill = document.getElementById('confirmProgressFill');

        if (dialogProgress && dialogText && dialogFill && dialogProgress.style.display !== 'none') {
            // Show in confirmation dialog
            dialogText.textContent = response.message;
            if (response.current && response.total) {
                const percentage = (response.current / response.total) * 100;
                dialogFill.style.width = `${percentage}%`;
            }
        } else {
            // Fallback to main progress bar
            updateProgress(response.current, response.total, response.message);
        }

    } else if (response.type === 'placeholderWarning') {
        hideSendConfirmation();
        handlePlaceholderWarning(response.message, response.data);

    } else if (response.type === 'attachmentCount') {
        const attachmentElement = document.getElementById('confirmAttachmentCount');
        if (attachmentElement) {
            const count = response.data?.count || 0;
            attachmentElement.textContent = count === 0 ? '0' : count.toString();
            showSendConfirmation();
        }
    } else if (response.type === 'emailSubject') {
        console.log('Received subject:', response.data);
        const subjectElement = document.getElementById('confirmSubject');
        if (subjectElement && response.data && response.data.subject) {
            subjectElement.textContent = response.data.subject;
        }
    }

}



function handlePlaceholderWarning(message, data) {
    resetUI();
    showStatus(message, 'error');

    const statusElement = document.getElementById('status');
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
    hideStatus();
    sendMessageToCS('duplicateEmail', {
        placeholder: originalData.placeholder,
        recipients: originalData.recipients,
        autoSend: originalData.autoSend,
        forceWithoutPlaceholder: forceWithoutPlaceholder
    });
}

function cancelPlaceholderWarning() {
    hideStatus();
}

// Main generate emails function
async function generateEmails() {
    console.log('=== GENERATE EMAILS FUNCTION CALLED ===');

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

        console.log('Showing confirmation dialog for:', filledRecipients.length, 'recipients');

        
        document.getElementById('confirmRecipientCount').textContent = filledRecipients.length;
        document.getElementById('confirmPlaceholder').textContent = placeholder;
        document.getElementById('confirmAttachmentCount').textContent = 'Loading...';
        document.getElementById('confirmSubject').textContent = 'Loading...';

        // Store data for later use
        window.pendingSendData = {
            placeholder: placeholder,
            recipients: filledRecipients,
            autoSend: true
        };

        // FIXED: Only request attachment count (not full send)
        sendMessageToCS('getAttachmentCount', window.pendingSendData);
        sendMessageToCS('getEmailSubject', {});

    } catch (error) {
        console.error("Error in generateEmails:", error);
        showStatus("Error: " + error.message, 'error');
    }
}

// Show confirmation dialog
function showSendConfirmation() {
    console.log('=== SHOWING CONFIRMATION DIALOG ===');
    const overlay = document.getElementById('confirmationOverlay');
    if (overlay) {
        overlay.classList.add('show');
        console.log('Confirmation dialog should now be visible');
    } else {
        console.error('Confirmation overlay element not found!');
    }
}

// Hide confirmation dialog
function hideSendConfirmation() {
    console.log('=== HIDING CONFIRMATION DIALOG ===');
    const overlay = document.getElementById('confirmationOverlay');
    if (overlay) {
        overlay.classList.remove('show');
    }
}

// User confirmed - proceed with sending
function proceedWithSend() {
    console.log('=== USER CONFIRMED SEND ===');
    if (window.pendingSendData) {
        // Try to show progress elements (gracefully handle if they don't exist)
        const buttons = document.querySelector('.confirmation-buttons');
        const progress = document.getElementById('sendingProgress');
        const progressText = document.getElementById('confirmProgressText');

        if (buttons) buttons.style.display = 'none';
        if (progress) progress.style.display = 'block';
        if (progressText) progressText.textContent = 'Starting to send emails...';

        sendMessageToCS('duplicateEmail', window.pendingSendData);
    }
}

// Communication with C# backend
function sendMessageToCS(action, data = null) {
    try {
        const message = {
            action: action,
            data: data
        };
        console.log('Sending to C#:', message);

        // FIX: Convert object to JSON string
        const jsonString = JSON.stringify(message);
        console.log('JSON string:', jsonString);

        window.chrome.webview.postMessage(jsonString);
    } catch (error) {
        console.error('Error sending message to C#:', error);
        showStatus('Error communicating with add-in backend', 'error');
    }
}

// Rest of your functions (keep your existing ones)
function addRecipient() {
    console.log('Adding recipient');
    recipients.push({ email: '', name: '' });
    renderRecipients();
    updateCountText();
}

function removeRecipient(index) {
    console.log('Removing recipient at index:', index);
    recipients.splice(index, 1);
    renderRecipients();
    updateCountText();
}

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
            if (email.includes('@') && email.split('@')[1]?.includes('.')) {
                suggestNameForIndex(index, false);
            }
        }
    }, 1000);
}

function updateRecipient(index, field, value) {
    if (recipients[index]) {
        recipients[index][field] = value;
        updateCountText();
    }
}

function suggestName(index) {
    suggestNameForIndex(index, true);
}

function suggestNameForIndex(index, shouldShowStatus = true) {
    const email = recipients[index].email;

    if (email && email.includes('@')) {
        const namePart = email.split('@')[0];
        const words = namePart.replace(/[._-]/g, ' ')
            .split(' ')
            .map(word => word.charAt(0).toUpperCase() + word.slice(1))
            .filter(word => word.length > 0);

        const suggestedName = words.length > 0 ? words[0] : '';
        recipients[index].name = suggestedName;

        const nameInput = document.querySelector(`.email-input-row:nth-child(${index + 1}) .name-input`);
        if (nameInput) {
            nameInput.value = suggestedName;
        }

        updateCountText();

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
    document.getElementById('countText').textContent = `${filledCount}/${totalCount} recipients ready`;
    const clearBtn = document.getElementById('clearAllBtn');
    clearBtn.disabled = totalCount === 0;
}

function setupPasteHandler() {
    document.addEventListener('paste', function (e) {
        const pastedText = e.clipboardData.getData('text');

        if (pastedText) {
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

    recipients = recipients.filter(r => r.email.trim() || r.name.trim());
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
    console.log('Showing status:', type, message);
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