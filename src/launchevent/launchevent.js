// Domains that never count toward the warning (e.g. our own). Keep lowercase,
// since collected domains are lowercased before comparison.
var DOMAIN_IGNORE_LIST = ["hortiful.com"];

/**
 * Runs when the user clicks Send (registered as OnMessageSend in manifest.xml).
 * Collects the email domain of every recipient in To, CC and BCC, then drops
 * any in DOMAIN_IGNORE_LIST. If the remaining recipients don't all share one
 * domain, it shows a warning before the message goes out.
 */
function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  var fields = [item.to, item.cc, item.bcc];

  // Used as a set: each unique domain becomes a key, e.g. { "testing.com": true }.
  var domains = {};

  // getAsync is asynchronous, so count down as each field's callback returns.
  // When the count hits zero, all three fields have been read.
  var pending = fields.length;
  var failed = false;

  fields.forEach(function (field) {
    // Pass the event along via asyncContext so the callback can finish it.
    field.getAsync({ asyncContext: event }, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        failed = true;
      } else {
        result.value.forEach(function (recipient) {
          // Take everything after the last "@" as the domain. Lowercase it so
          // "Testing.com" and "testing.com" count as the same domain.
          // Recipients with no "@" (e.g. some contact groups) are skipped.
          var address = recipient.emailAddress || "";
          var at = address.lastIndexOf("@");
          if (at > -1) domains[address.slice(at + 1).toLowerCase()] = true;
        });
      }

      // Last field to finish: all recipients are collected, so filter and decide.
      if (--pending === 0) {
        DOMAIN_IGNORE_LIST.forEach(function (domain) {
          delete domains[domain];
        });
        finish(result.asyncContext, Object.keys(domains), failed);
      }
    });
  });
}

/**
 * Tells Outlook whether to send the message.
 * event.completed must be called exactly once on every path. Otherwise Outlook
 * waits on the add-in and eventually times out.
 */
function finish(event, domainList, failed) {
  // If any recipient field couldn't be read, prompt the user rather than guess.
  if (failed) {
    event.completed({
      allowEvent: false,
      errorMessage: "Couldn't read the recipients.",
      sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
    });
    return;
  }

  // Zero or one domain: nothing to warn about, so let the message go.
  if (domainList.length <= 1) {
    event.completed({ allowEvent: true });
    return;
  }

  // More than one domain: show the Smart Alerts dialog with the list.
  // PromptUser adds a "Send Anyway" button, so this is a warning the user
  // can override rather than a hard block.
  var msg = "Recipients span " + domainList.length + " domains: " + domainList.join(", ") + ".";
  event.completed({
    allowEvent: false,
    errorMessage: msg,
    sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
  });
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
