Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        displayEmailDetails();
        document.getElementById("verifyButton").onclick = verifyEmail;
    }
});

function displayEmailDetails() {
    const item = Office.context.mailbox.item;
    const dateOptions = { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'UTC', timeZoneName: 'short' };
    const formattedDate = item.dateTimeCreated.toLocaleString('en-GB', dateOptions);

    // Usamos innerHTML para incluir la etiqueta <b> directamente
    document.getElementById("from-display").innerHTML = `<b>From:</b> ${item.from.emailAddress}`;
    document.getElementById("to-display").innerHTML = `<b>To:</b> ${item.to.map(recipient => recipient.emailAddress).join(", ")}`;
    document.getElementById("subject-display").innerHTML = `<b>Subject:</b> ${item.subject || "Unknown"}`;
    document.getElementById("date-display").innerHTML = `<b>Date:</b> ${formattedDate}`;

    // Si decides volver a añadir attachments, aquí iría el código
    // document.getElementById("attachments-display").innerHTML = `<b>Attachments:</b> No Data`;
}

function verifyEmail() {
    const item = Office.context.mailbox.item;
    
    item.body.getAsync(Office.CoercionType.Text, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = asyncResult.value;
            const emailData = {
                From: item.from.emailAddress,
                To: item.to.map(recipient => recipient.emailAddress).join(", "),
                Subject: item.subject,
                Body: emailBody,
                Date: item.dateTimeCreated.toISOString(),
                Attachments: []
            };

            const resultDiv = document.getElementById("result-message");
            resultDiv.style.display = "block";
            resultDiv.innerHTML = "Analizando el correo...";
            resultDiv.classList.remove("result-safe", "result-phishing");

            fetch('https://securemail-api-44381375902.europe-west1.run.app', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(emailData),
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error('Network response was not ok');
                }
                return response.json();
            })
            .then(result => {
                const prediction = result.predictions[0];
                const isPhishing = (prediction.model_prediction.label === "Phishing");
                const probability = prediction.model_prediction.probability;

                let verdictMessage;
                if (isPhishing) {
                    verdictMessage = `Este correo ha sido clasificado como <b>Phishing</b> con una probabilidad de <b>${(probability * 100).toFixed(2)}%</b>. Recomendamos que no interactúes con este correo, ni abras ningún enlace o archivo adjunto.`;
                    resultDiv.classList.add("result-phishing");
                } else {
                    verdictMessage = `Este correo ha sido clasificado como <b>Seguro</b> con una probabilidad de <b>${((1 - probability) * 100).toFixed(2)}%</b>. Este correo parece ser seguro, pero siempre debes ser cauteloso.`;
                    resultDiv.classList.add("result-safe");
                }
                
                resultDiv.innerHTML = verdictMessage;
                
                if (isPhishing) {
                    const contactButton = document.createElement("button");
                    contactButton.innerText = "Contact Support";
                    contactButton.onclick = () => window.open('https://support.google.com/', '_blank');
                    contactButton.style.backgroundColor = '#FF4D4D';
                    contactButton.style.marginTop = '10px';
                    resultDiv.appendChild(contactButton);
                }
            })
            .catch(error => {
                resultDiv.innerHTML = "Error al conectar con el servidor: " + error.message;
                resultDiv.style.backgroundColor = "#fff0e6";
            });
        } else {
            document.getElementById("result-message").innerText = "Error al obtener el cuerpo del correo.";
        }
    });

}
