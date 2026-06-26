/**
 * Helper to manage email recipients and identify external emails.
 */

/**
 * Checks if a single email address is external to the company domain.
 */
export function isExternalEmail(email: string, companyDomain: string): boolean {
  if (!email) return false;
  const domain = companyDomain.trim().toLowerCase();
  const emailLower = email.trim().toLowerCase();
  
  const suffix = domain.startsWith("@") ? domain : "@" + domain;
  return !emailLower.endsWith(suffix);
}

/**
 * Checks if there are any external email addresses in the provided list.
 */
export function hasExternalRecipients(emails: string[], companyDomain: string): boolean {
  return emails.some(email => isExternalEmail(email, companyDomain));
}

/**
 * Helper to wrap Office.js getAsync callback to Promise
 */
function getRecipientsFromFieldAsync(recipientsField: any): Promise<any[]> {
  return new Promise((resolve) => {
    try {
      if (!recipientsField || typeof recipientsField.getAsync !== "function") {
        resolve([]);
        return;
      }
      recipientsField.getAsync((asyncResult: any) => {
        if (asyncResult && (asyncResult.status === "succeeded" || (typeof Office !== "undefined" && Office.AsyncResultStatus && asyncResult.status === Office.AsyncResultStatus.Succeeded))) {
          resolve(asyncResult.value || []);
        } else {
          resolve([]);
        }
      });
    } catch (_) {
      resolve([]);
    }
  });
}

/**
 * Retrieves all recipient emails (To, CC, BCC) from the compose window item.
 */
export async function getRecipientEmails(item: any): Promise<string[]> {
  if (!item) return [];

  const toPromise = item.to ? getRecipientsFromFieldAsync(item.to) : Promise.resolve([]);
  const ccPromise = item.cc ? getRecipientsFromFieldAsync(item.cc) : Promise.resolve([]);
  const bccPromise = item.bcc ? getRecipientsFromFieldAsync(item.bcc) : Promise.resolve([]);

  try {
    const [toRecips, ccRecips, bccRecips] = await Promise.all([toPromise, ccPromise, bccPromise]);
    const allRecipients = [...toRecips, ...ccRecips, ...bccRecips];
    
    // Map recipient objects to email address strings.
    // In Outlook, recipient object is EmailAddressDetails: { emailAddress: string, displayName: string, recipientType: ... }
    return allRecipients
      .map(r => r && typeof r.emailAddress === "string" ? r.emailAddress : "")
      .filter(email => email !== "");
  } catch (_) {
    return [];
  }
}
