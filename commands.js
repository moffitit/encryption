// commands.js — On-send handler for IRM Encryption Check add-in
// Runs in the background function file (commands.html). Never shown to users.

Office.onReady(function () {
  if (Office.actions) {
    Office.actions.associate('onItemSend', onItemSend);
  }
});

/**
 * onItemSend
 * Fires synchronously before Outlook sends the message.
 *
 * Logic:
 *  1. If the user already reviewed the email via the task pane
 *     (X-Reviewed-By-Addin header is set), allow send.
 *  2. Otherwise, scan body + subject for sensitive patterns.
 *  3. If sensitive content found, block send with an actionable notification.
 *  4. If no sensitive content found, allow send.
 */
function onItemSend(event) {
  const item = Office.context.mailbox.item;

  item.internetHeaders.getAsync(['X-Reviewed-By-Addin'], function (headerResult) {
    if (headerResult.status === Office.AsyncResultStatus.Failed) {
      // Can't read headers — fail open to avoid disrupting send flow
      event.completed({ allowEvent: true });
      return;
    }

    const headers = headerResult.value || {};
    if (headers['X-Reviewed-By-Addin']) {
      // User completed the task pane review — allow send
      event.completed({ allowEvent: true });
      return;
    }

    // Not yet reviewed — scan for sensitive patterns
    item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
      if (bodyResult.status === Office.AsyncResultStatus.Failed) {
        event.completed({ allowEvent: true });
        return;
      }

      const subject = item.subject || '';
      const body    = bodyResult.value || '';
      const text    = subject + ' ' + body;

      const PATTERNS = [
        /\b(?:\d[ -]?){13,16}\b/,
        /\b\d{3}[ -]?\d{3}[ -]?\d{3}\b/,
        /\b(password|passphrase|passwd|credentials?)\b/i,
        /\b\d{3}-\d{2}-\d{4}\b/,
        /\b(bsb|account\s*no\.?|bank\s*account)\b/i,
        /\b(api[ _-]?key|bearer\s+[a-zA-Z0-9\-._~+/]{20,})\b/i,
        /\b(medicare|health\s*record|patient\s*id)\b/i,
        /\b(confidential|strictly\s+private|do\s+not\s+distribute|restricted)\b/i,
      ];

      const hasSensitive = PATTERNS.some(re => re.test(text));

      if (hasSensitive) {
        item.notificationMessages.addAsync('irmWarning', {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message:
            '⚠ Sensitive content detected. ' +
            'Please open the IRM Check panel (ribbon) to review and apply ' +
            'an IRM / Sensitivity Label before sending.',
          persistent: true,
        }, function () {
          event.completed({ allowEvent: false });
        });
      } else {
        event.completed({ allowEvent: true });
      }
    });
  });
}
