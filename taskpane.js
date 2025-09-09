Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        displayEmailDetails();
        document.getElementById("verifyButton").onclick = verifyEmail;
    }
});

/**
 * Muestra los detalles del correo en la interfaz.
 */
function displayEmailDetails() {
    const item = Office.context.mailbox.item;
    const dateOptions = { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'UTC', timeZoneName: 'short' };
    const formattedDate = item.dateTimeCreated.toLocaleString('en-GB', dateOptions);

    // Usa innerHTML para mostrar la información del correo
    document.getElementById("from-display").innerHTML = `<b>From:</b> ${item.from.emailAddress}`;
    document.getElementById("to-display").innerHTML = `<b>To:</b> ${item.to.map(recipient => recipient.emailAddress).join(", ")}`;
    document.getElementById("subject-display").innerHTML = `<b>Subject:</b> ${item.subject || "Unknown"}`;
    document.getElementById("date-display").innerHTML = `<b>Date:</b> ${formattedDate}`;

    // Nota: Por ahora, los adjuntos se marcan como "No Data" para simplificar la lógica
    document.getElementById("attachments-display").innerHTML = `<b>Attachments:</b> No Data`;
}

/**
 * Envía el contenido del correo a la API para su análisis.
 */
function verifyEmail() {
    const item = Office.context.mailbox.item;
    const resultDiv = document.getElementById("result-message");

    // Muestra un mensaje de carga
    resultDiv.style.display = "block";
    resultDiv.innerHTML = "Analizando el correo...";
    resultDiv.className = ""; // Limpia las clases de estilo anteriores

    // Obtiene el cuerpo del correo de forma asíncrona
    item.body.getAsync(Office.CoercionType.Text, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = asyncResult.value;
            const emailData = {
                From: item.from.emailAddress,
                To: item.to.map(recipient => recipient.emailAddress).join(", "),
                Subject: item.subject,
                Body: emailBody,
                Date: item.dateTimeCreated.toISOString(),
                Attachments: [] // Simplificado para no incluir archivos binarios
            };
            
            // Envía los datos a tu API
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
                    verdictMessage = `Este correo ha sido clasificado como <b>Phishing</b> con una probabilidad de <b>${(probability * 100).toFixed(2)}%</b>. Recomendamos que no interactúes con este correo.`;
                    resultDiv.classList.add("result-phishing");
                } else {
                    verdictMessage = `Este correo ha sido clasificado como <b>Seguro</b> con una probabilidad de <b>${((1 - probability) * 100).toFixed(2)}%</b>. Este correo parece ser seguro, pero sé cauteloso.`;
                    resultDiv.classList.add("result-safe");
                }
                
                resultDiv.innerHTML = verdictMessage;
                
                // Si es phishing, agrega el botón de soporte
                if (isPhishing) {
                    const contactButton = document.createElement("button");
                    contactButton.innerText = "Contactar Soporte";
                    contactButton.onclick = () => window.open('https://support.google.com/', '_blank');
                    contactButton.style.backgroundColor = '#FF4D4D';
                    resultDiv.appendChild(contactButton);
                }
            })
            .catch(error => {
                resultDiv.innerHTML = `Error al conectar con el servidor: ${error.message}`;
                resultDiv.style.backgroundColor = "#fff0e6";
                resultDiv.style.color = "#d9534f";
            });
        } else {
            resultDiv.innerHTML = "Error al obtener el cuerpo del correo.";
            resultDiv.style.backgroundColor = "#fff0e6";
            resultDiv.style.color = "#d9534f";
        }
    });
}
