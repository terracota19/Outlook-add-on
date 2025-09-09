Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        displayEmailDetails();
        // Assign the verifyEmail function to the button's click event
        document.getElementById("verifyButton").onclick = verifyEmail;
    }
});

/**
 * Displays the email details in the add-in's UI.
 * This function runs on startup to populate the email information.
 */
function displayEmailDetails() {
    // Get the current email item from the Outlook context
    const item = Office.context.mailbox.item;
    
    // Configure options for date formatting
    const dateOptions = { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric', 
        hour: '2-digit', 
        minute: '2-digit', 
        timeZone: 'UTC', 
        timeZoneName: 'short' 
    };
    
    // Format the email's creation date
    const formattedDate = item.dateTimeCreated.toLocaleString('en-GB', dateOptions);

    // Use innerHTML to display the email information in the HTML elements
    document.getElementById("from-display").innerHTML = `<b>From:</b> ${item.from.emailAddress}`;
    
    // Check if the 'to' property exists and is not empty
    const toRecipients = item.to && item.to.length > 0 
        ? item.to.map(recipient => recipient.emailAddress).join(", ")
        : "N/A";
    document.getElementById("to-display").innerHTML = `<b>To:</b> ${toRecipients}`;
    
    document.getElementById("subject-display").innerHTML = `<b>Subject:</b> ${item.subject || "No Subject"}`;
    document.getElementById("date-display").innerHTML = `<b>Date:</b> ${formattedDate}`;
}

/**
 * Sends the email content to the API for phishing analysis.
 */
function verifyEmail() {
    const item = Office.context.mailbox.item;
    const resultDiv = document.getElementById("result-message");

    // Show a loading message to the user
    resultDiv.style.display = "block";
    resultDiv.innerHTML = "Analyzing email...";
    resultDiv.className = ""; // Clear previous style classes

    // Get the email body asynchronously
    item.body.getAsync(Office.CoercionType.Text, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = asyncResult.value;
            const emailData = {
                From: item.from.emailAddress,
                To: item.to.map(recipient => recipient.emailAddress).join(", "),
                Subject: item.subject,
                Body: emailBody,
                Date: item.dateTimeCreated.toISOString(),
                // Attachments section has been removed as per request
            };
            
            // Send the data to your Google Cloud Run API
            fetch('https://securemail-api-44381375902.europe-west1.run.app', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(emailData),
            })
            .then(response => {
                if (!response.ok) {
                    // If the response is not successful (e.g., 404, 500), throw an error
                    throw new Error(`Server response error: ${response.status} ${response.statusText}`);
                }
                return response.json();
            })
            .then(result => {
                // Process the API response
                const prediction = result.predictions[0];
                const isPhishing = (prediction.model_prediction.label === "Phishing");
                const probability = prediction.model_prediction.probability;

                let verdictMessage;
                if (isPhishing) {
                    verdictMessage = `This email has been classified as <b>Phishing</b> with a probability of <b>${(probability * 100).toFixed(2)}%</b>. We recommend you do not interact with this email.`;
                    resultDiv.classList.add("result-phishing");
                } else {
                    verdictMessage = `This email has been classified as <b>Safe</b> with a probability of <b>${((1 - probability) * 100).toFixed(2)}%</b>. This email appears to be safe, but be cautious.`;
                    resultDiv.classList.add("result-safe");
                }
                
                resultDiv.innerHTML = verdictMessage;
                
                // If it's a phishing email, add the support button
                if (isPhishing) {
                    const contactButton = document.createElement("button");
                    contactButton.innerText = "Contact Support";
                    contactButton.onclick = () => window.open('https://support.google.com/', '_blank');
                    contactButton.style.backgroundColor = '#FF4D4D';
                    resultDiv.appendChild(contactButton);
                }
            })
            .catch(error => {
                // Handle network or server errors
                resultDiv.innerHTML = `Error connecting to the server: ${error.message}`;
                resultDiv.style.backgroundColor = "#fff0e6";
                resultDiv.style.color = "#d9534f";
            });
        } else {
            // Handle error if the email body could not be retrieved
            resultDiv.innerHTML = "Error getting the email body.";
            resultDiv.style.backgroundColor = "#fff0e6";
            resultDiv.style.color = "#d9534f";
        }
    });
}
