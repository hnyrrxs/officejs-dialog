$("#open").on("click", () => tryCatch(openDialog));

function openDialog() {
  Office.context.ui.displayDialogAsync(
    window.location.origin + "/dialog.html", // Open a separate dialog page
    { height: 30, width: 30 },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Error opening dialog:", asyncResult.error.message);
      } else {
        let dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
          console.log("Message from dialog:", arg.message);
          dialog.close();
        });
      }
    }
  );
}
