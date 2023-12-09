// Initialize the Microsoft Teams SDK
microsoftTeams.initialize();

// Function to handle message extension query
function handleMessagingExtensionQuery(query) {
    // Fetch the custom message from Teams context
    const customMessage = microsoftTeams.settings.getSettings().customSettings.message;

    // Return the custom message as a messaging extension result
    const result = {
        type: 'message',
        attachmentLayout: 'list',
        attachments: [
            {
                content: {
                    title: `Custom Message:`,
                    text: customMessage,
                },
            },
        ],
    };

    microsoftTeams.tasks.submitTaskResults([result]);
}

// Register messaging extension query handler
microsoftTeams.tasks.registerOnQuery(context => {
    handleMessagingExtensionQuery(context);
});
