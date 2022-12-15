/* eslint-disable no-undef */
function onNewMessageComposeHandler(event) {
  setSubject(event);
}

function setSubject(event) {
  Office.context.mailbox.item.subject.setAsync(
    "Set by an event-based add-in!",
    {
      asyncContext: event,
    },
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
      }

      // Call event.completed() after all work is done.
      asyncResult.asyncContext.completed();
    });
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
