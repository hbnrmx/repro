/* eslint-disable no-undef */

function onNewMessageComposeHandler(event) {
  console.log("openDialog was invoked.");

  Office.context.ui.displayDialogAsync("https://www.google.com", { asyncContext: event }, (asyncResult) => {
    console.log("this code is never reached");
    console.log("this code is never reached");
    console.log("this code is never reached");
    console.log(asyncResult);
  });
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
