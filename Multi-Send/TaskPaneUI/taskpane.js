// Global recipients array

let recipients = [];



// Initialize the taskpane

document.addEventListener('DOMContentLoaded', function () {

    addRecipient(); // Start with one empty recipient

    updateCountText();

    setupPasteHandler();

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

                   oninput="updateRecipient(${index}, 'email', this.value)">

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



function updateRecipient(index, field, value) {

    if (recipients[index]) {

        recipients[index][field] = value;

        updateCountText();

    }

}



function suggestName(index) {

    const email = recipients[index].email;

    if (email && email.includes('@')) {

        const namePart = email.split('@')[0];

        const suggestedName = namePart.replace(/[._-]/g, ' ')

            .split(' ')

            .map(word => word.charAt(0).toUpperCase() + word.slice(1))

            .join(' ');

        recipients[index].name = suggestedName;

        renderRecipients();

        updateCountText();

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

        if (pastedText && pastedText.includes('@')) {

            showPasteHint();

            setTimeout(() => {

                processPastedEmails(pastedText);

                hidePasteHint();

            }, 100);

        }

    });

}



function processPastedEmails(text) {

    const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

    const foundEmails = text.match(emailRegex);



    if (foundEmails && foundEmails.length > 0) {

        // Clear existing empty recipients

        recipients = recipients.filter(r => r.email.trim() || r.name.trim());



        // Add new recipients

        foundEmails.forEach(email => {

            if (!recipients.some(r => r.email === email)) {

                recipients.push({ email: email, name: '' });

            }

        });



        renderRecipients();

        updateCountText();



        showStatus(`Added ${foundEmails.length} email addresses`, 'success');

        setTimeout(hideStatus, 3000);

    }

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