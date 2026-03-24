Office.onReady(() => {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

function onMessageSendHandler(event) {
  const item = Office.context?.mailbox?.item;
  const HELP_EXAMPLE = "E2308387 - 914217 - BP - ZTM275 -";
  const ERROR_TEXT = "Onderwerp voldoet niet aan het vereiste format.\nBijv.: " + HELP_EXAMPLE;

  if (!item) {
    event.completed({ allowEvent: false, errorMessage: "Mailbox-item niet beschikbaar." });
    return;
  }

  item.subject.getAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: false, errorMessage: "Kon het onderwerp niet uitlezen." });
      return;
    }

    // Normaliseer subject
    const subject = (result.value || "")
      .replace(/\u00A0/g, " ")   // NBSP → spatie
      .replace(/\s+/g, " ")      // meerdere spaties → 1
      .trim();

    // Strip reply/forward prefixes (meertalig)
    const prefixRe = /^(?:\s*(RE|FW|FWD|Antw|Doorst)\s*:\s*)+/i;
    const core = subject.replace(prefixRe, "").trim();

    // Core pattern (trailing '-' is verplicht; maak '-?' als het optioneel mag)
    const corePattern = /^E\d{7}\s*-\s*\d{6}\s*-\s*BP\s*-\s*ZTM\d{3}\s*-\s*$/i;

    if (corePattern.test(core)) {
      event.completed({ allowEvent: true });
    } else {
      event.completed({ allowEvent: false, errorMessage: ERROR_TEXT });
    }
  });
}
