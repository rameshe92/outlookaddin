Office.initialize = function () {
}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}

function LaunchGuide(event) { 
	window.open("https://www.sigsync.com/kb/deploy-email-signatures-add-in-for-outlook.html", "_blank");
	/*Office.context.ui.displayDialogAsync('https://www.sigsync.com/kb/deploy-email-signatures-add-in-for-outlook.html', {
                    height: 100,
                    width: 100,
                });*/
    try {
		if(Office.context.ui)
			Office.context.ui.openBrowserWindow('https://www.sigsync.com/kb/deploy-email-signatures-add-in-for-outlook.html');
	} catch(err) {
		
	}
	event.completed();
}
