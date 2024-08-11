import { fetchEmails } from "../helpers/graph-helper";
import { processEmails } from "../helpers/file-helper";

async function runBackgroundTask() {
  try {
    const emails = await fetchEmails();
    await processEmails(emails);
  } catch (error) {
    handleError(error);
  }
}

// Run the background task every minute
setInterval(runBackgroundTask, 60000);

// Initial run
runBackgroundTask();
