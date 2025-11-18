Office.initialize = () => {};

async function insertSignature() {
  let item = Office.context.mailbox.item;

  // Get recipients
  let toList = item.to.getAsync ? await getRecipients(item.to) : [];
  let ccList = item.cc.getAsync ? await getRecipients(item.cc) : [];
  let recipients = [...toList, ...ccList];

  // Check internal
  let isInternal = recipients.some(r =>
    r.emailAddress.toLowerCase().endsWith("@atul.co.in")
  );

  // Load correct signature HTML
  let signatureUrl = isInternal
    ? "https://shaileshpandire.github.io/signatureaddin/internal.html"
    : "https://shaileshpandire.github.io/signatureaddin/external.html";

  let signatureHTML = await fetch(signatureUrl).then(r => r.text());

  // Remove old signature
  let body = await item.body.getAsync(Office.CoercionType.Html);
  let cleanBody = body.value
    .replace(/<!--signature-start-->[\s\S]*?<!--signature-end-->/g, "");

  // Insert new signature at bottom
  let newBody = cleanBody + 
    "<!--signature-start-->" + signatureHTML + "<!--signature-end-->";

  await item.body.setAsync(newBody, { coercionType: Office.CoercionType.Html });
}

// Helper to read recipients
function getRecipients(recipObj) {
  return new Promise(resolve => {
    recipObj.getAsync((asyncResult) => {
      resolve(asyncResult.value || []);
    });
  });
}

Office.actions.associate("onMessageSend", insertSignature);
