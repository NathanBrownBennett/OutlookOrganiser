import { Client } from "@microsoft/microsoft-graph-client";
import { login } from "./ssoauth-helper";

async function createFoldersAndCopyEmails() {
  const accessToken = await login();
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });

  // Fetch emails from the inbox
  const emails = await client.api("/me/messages").top(10).get();

  emails.value.forEach(async (email) => {
    const senderDomain = email.sender.emailAddress.address.split('@')[1];
    const senderUser = email.sender.emailAddress.name;

    // Create folder in OneDrive
    const domainFolder = `/drive/root:/Emails/${senderDomain}`;
    await client.api(domainFolder).put();

    const userFolder = `${domainFolder}/${senderUser}`;
    await client.api(userFolder).put();

    // Save email and attachments to OneDrive
    const emailContent = `
    Subject: ${email.subject}
    From: ${email.sender.emailAddress.name} <${email.sender.emailAddress.address}>
    Date: ${new Date(email.receivedDateTime).toLocaleString()}
    
    ${email.body.content}
    `;

    // Save email as .eml or .txt
    await client.api(`${userFolder}/${email.id}.txt`).put(emailContent);

    // Save attachments
    if (email.hasAttachments) {
      const attachments = await client.api(`/me/messages/${email.id}/attachments`).get();
      attachments.value.forEach(async (attachment) => {
        await client.api(`${userFolder}/${attachment.name}`).put(attachment.contentBytes);
      });
    }
  });
}

// Call this function on an interval or as a background task
setInterval(createFoldersAndCopyEmails, 60000); // Runs every minute
