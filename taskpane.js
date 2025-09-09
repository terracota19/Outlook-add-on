Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        displayEmailDetails();
        // Asigna la función verifyEmail al evento de clic del botón
        document.getElementById("verifyButton").onclick = verifyEmail;
    }
});

/**
 * Muestra los detalles del correo en la interfaz del complemento.
 * La función se ejecuta al inicio para rellenar la información del correo.
 */
function displayEmailDetails() {
    // Obtiene el objeto de correo actual del contexto de Outlook
    const item = Office.context.mailbox.item;
    
    // Configura las opciones para formatear la fecha
    const dateOptions = { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric', 
        hour: '2-digit', 
        minute: '2-digit', 
        timeZone: 'UTC', 
        timeZoneName: 'short' 
    };
    
    // Formatea la fecha de creación del correo
    const formattedDate = item.dateTimeCreated.toLocaleString('en-GB', dateOptions);

    // Usa innerHTML para mostrar la información del correo en los elementos HTML
    document.getElementById("from-display").innerHTML = `<b>De:</b> ${item.from.emailAddress}`;
    
    // Verifica si la propiedad 'to' existe y si no está vacía
    const toRecipients = item.to && item.to.length > 0 
        ? item.to.map(recipient => recipient.emailAddress).join(", ")
        : "N/A";
    document.getElementById("to-display").innerHTML = `<b>Para:</b> ${toRecipients}`;
    
    document.getElementById("subject-display").innerHTML = `<b>Asunto:</b> ${item.subject || "Sin Asunto"}`;
    document.getElementById("date-display").innerHTML = `<b>Fecha:</b> ${formattedDate}`;

    // Nota: El elemento de adjuntos debe existir en el HTML para evitar errores
    document.getElementById("attachments-display").innerHTML = `<b>Adjuntos:</b> No Data`;
}

/**
 * Envía el contenido del correo a la API para su análisis de phishing.
 */
function verifyEmail() {
    const item = Office.context.mailbox.item;
    const resultDiv = document.getElementById("result-message");

    // Muestra un mensaje de carga al usuario
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
                Attachments: [] // Simplificado para evitar problemas con datos binarios
            };
            
            // Envía los datos a tu API en Google Cloud Run
            fetch('https://securemail-api-44381375902.europe-west1.run.app', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(emailData),
            })
            .then(response => {
                if (!response.ok) {
                    // Si la respuesta no es exitosa (ej. 404, 500), lanza un error
                    throw new Error(`Error en la respuesta del servidor: ${response.status} ${response.statusText}`);
                }
                return response.json();
            })
            .then(result => {
                // Procesa la respuesta de la API
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
                // Maneja errores de red o del servidor
                resultDiv.innerHTML = `Error al conectar con el servidor: ${error.message}`;
                resultDiv.style.backgroundColor = "#fff0e6";
                resultDiv.style.color = "#d9534f";
            });
        } else {
            // Maneja el error si no se pudo obtener el cuerpo del correo
            resultDiv.innerHTML = "Error al obtener el cuerpo del correo.";
            resultDiv.style.backgroundColor = "#fff0e6";
            resultDiv.style.color = "#d9534f";
        }
    });
}
