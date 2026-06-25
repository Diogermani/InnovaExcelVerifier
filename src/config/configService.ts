export interface AppConfig {
  companyDomain: string;
  supportedExtensions: string[];
  startColumn: string;
  language: string;
  validateAllRecipients: boolean;
  validateHiddenSheets: boolean;
  allowUserOverride: boolean;
  ignoreZipFiles: boolean;
}

export const DEFAULT_CONFIG: AppConfig = {
  companyDomain: "@aatb.com.br",
  supportedExtensions: [".xlsx", ".xlsm"],
  startColumn: "M",
  language: "pt-BR",
  validateAllRecipients: true,
  validateHiddenSheets: true,
  allowUserOverride: true,
  ignoreZipFiles: true
};

/**
 * Returns the configuration synchronously.
 * Using DEFAULT_CONFIG directly avoids a network fetch on the critical
 * OnMessageSend path, preventing timeouts in the Outlook event runtime.
 * To change settings, edit DEFAULT_CONFIG above and redeploy.
 */
export async function getConfig(): Promise<AppConfig> {
  return DEFAULT_CONFIG;
}
