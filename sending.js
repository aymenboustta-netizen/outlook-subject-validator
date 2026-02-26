
Office.onReady(() => {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  item.subject.getAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({
        allowEvent: false,
        errorMessage: "Kon het onderwerp niet uitlezen."
      });
      return;
    }

    const subject = (result.value || "").trim();

    const pattern =
      /^(?:(?:RE|FW|FWD|Antw|Doorst)\s*:\s*)?E\d{7}\s*-\s*\d{6}\s*-\s*BP\s*-\s*ZTM\d{3}\s*-\s*$/i;

    if (pattern.test(subject)) {
      event.completed({ allowEvent: true });
    } else {
      event.completed({
        allowEvent: false,
        errorMessage:
          "Onderwerp is niet correct opgesteld.\n\n" +
          "Gebruik dit format:\n" +
          "E2308387 - 914217 - BP - ZTM275 -"
      });
    }
  });
}
