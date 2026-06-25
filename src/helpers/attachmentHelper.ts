/**
 * Promisified wrapper around getAttachmentsAsync.
 * Retrieves all attachments from the current compose item.
 */
export function getAttachmentsAsync(): Promise<Office.AttachmentDetailsCompose[]> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      return reject(new Error("No active mailbox item found."));
    }

    if (!item.getAttachmentsAsync) {
      // Some environments might not support attachments or have it disabled
      return resolve([]);
    }

    item.getAttachmentsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        reject(new Error(result.error?.message || "Failed to retrieve attachments."));
      }
    });
  });
}

/**
 * Promisified wrapper around getAttachmentContentAsync.
 * Retrieves the base64 content of a file attachment.
 */
export function getAttachmentContentAsync(attachmentId: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      return reject(new Error("No active mailbox item found."));
    }

    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const contentObj = result.value;
        if (contentObj.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          resolve(contentObj.content);
        } else {
          reject(new Error(`Unexpected content format: ${contentObj.format}`));
        }
      } else {
        reject(new Error(result.error?.message || "Failed to retrieve attachment content."));
      }
    });
  });
}

/**
 * Retrieves and filters attachments that match the supported Excel extensions.
 */
export async function getExcelAttachments(supportedExtensions: string[]): Promise<Office.AttachmentDetailsCompose[]> {
  const allAttachments = await getAttachmentsAsync();
  
  return allAttachments.filter((attachment) => {
    if (attachment.attachmentType !== Office.MailboxEnums.AttachmentType.File) {
      return false;
    }
    
    const fileName = attachment.name.toLowerCase();
    return supportedExtensions.some(ext => fileName.endsWith(ext.toLowerCase()));
  });
}
