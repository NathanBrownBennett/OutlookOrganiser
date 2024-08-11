import { createOneDriveFolder, saveEmailToFolder } from "./graph-helper";

export async function processEmails(emails) {
  for (const email of emails) {
    const senderDomain = email.sender.emailAddress.address.split("@")[1];
    const senderUser = email.sender.emailAddress.name;

    // Create the necessary folders in OneDrive
    const folderPath = await createOneDriveFolder(senderDomain, senderUser);

    // Save the email and its attachments to the created folder
    await saveEmailToFolder(email, folderPath);
  }
}
