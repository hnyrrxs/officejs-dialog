Office.onReady(() => {
    document.getElementById("sendMessageButton").addEventListener("click", function () {
        if (Office.context.ui && Office.context.ui.messageParent) {
            Office.context.ui.messageParent("Hello from dialog!");
        } else {
            console.error("messageParent is not available. Ensure this is running in a dialog.");
        }
    });
});