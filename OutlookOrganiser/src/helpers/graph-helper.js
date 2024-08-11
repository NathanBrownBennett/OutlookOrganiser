import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "./auth-helper";

export async function fetchEmails() {
  const accessToken = await getAccessToken();
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  try {
    const messages = await client.api("/me/mailFolders/inbox/messages").top(10).get();
    return messages.value;
  } catch (error) {
    console.error("Error fetching emails: ", error);
  }
}

export async function createOneDriveFolder(domain, user) {
  const accessToken = await getAccessToken();
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  const folderPath = `/drive/root:/Emails/${domain}/${user}`;
  try {
    await client.api(folderPath).put();
  } catch (error) {
    if (error.statusCode !== 409) {
      // 409 means the folder already exists, so ignore that error
      console.error("Error creating OneDrive folder: ", error);
    }
  }
  return folderPath;
}

export async function saveEmailToFolder(email, folderPath) {
  const accessToken = await getAccessToken();
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  const emailContent = `
  Subject: ${email.subject}
  From: ${email.sender.emailAddress.name} <${email.sender.emailAddress.address}>
  Date: ${new Date(email.receivedDateTime).toLocaleString()}
  
  ${email.body.content}
  `;

  try {
    await client.api(`${folderPath}/${email.id}.txt`).put(emailContent);

    // Save attachments
    if (email.hasAttachments) {
      const attachments = await client.api(`/me/messages/${email.id}/attachments`).get();
      attachments.value.forEach(async (attachment) => {
        await client.api(`${folderPath}/${attachment.name}`).put(attachment.contentBytes);
      });
    }
  } catch (error) {
    console.error("Error saving email to OneDrive: ", error);
  }
}
