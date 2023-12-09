// Initialize the Microsoft Teams SDK
microsoftTeams.initialize();

// Function to save the custom message
function saveCustomMessage() {
    const customMessage = document.getElementById('customMessage').value;

    // Save the custom message to Teams context (replace 'customMessage' with your property name)
    microsoftTeams.settings.setSettings({
        entityId: 'customMessageTab',
        contentUrl: window.location.origin + '/configurable-tab.html',
        suggestedDisplayName: 'Custom Message Tab',
        websiteUrl: window.location.origin + '/configurable-tab.html',
        customSettings: {
            message: customMessage,
        },
    });

    // Notify Teams that the settings have been saved
    microsoftTeams.settings.registerOnSaveHandler((saveEvent) => {
        saveEvent.notifySuccess();
    });
}

// Add a click event listener to the "Save" button
document.getElementById('saveButton').addEventListener('click', saveCustomMessage);
