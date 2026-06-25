import { getConfig } from "../config/configService";
import { getExcelAttachments, getAttachmentContentAsync } from "../helpers/attachmentHelper";
import { validateExcelWorkbook, summarizeValidation, AttachmentValidationResult, ValidationSummary } from "../helpers/excelValidator";

// ─── Diagnostic helper ───────────────────────────────────────────────────────

const T0 = Date.now();

function log(step: string, extra?: string): void {
  const elapsed = Date.now() - T0;
  const msg = `[InnovaVerifier +${elapsed}ms] ${step}${extra ? ` | ${extra}` : ""}`;
  console.log(msg);
}

/**
 * Shows a transient notification bar inside the compose window.
 * Useful for diagnosing execution progress without opening DevTools.
 */
function notify(key: string, message: string): void {
  try {
    const item = Office.context?.mailbox?.item;
    if (!item?.notificationMessages) return;
    item.notificationMessages.replaceAsync(key, {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "Icon.16x16",
      persistent: false
    });
  } catch (_) {
    // Notifications are best-effort — never let them throw
  }
}

function notifyError(key: string, message: string): void {
  try {
    const item = Office.context?.mailbox?.item;
    if (!item?.notificationMessages) return;
    item.notificationMessages.replaceAsync(key, {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message,
      persistent: true
    });
  } catch (_) { /* best-effort */ }
}

// ─── Message formatter ────────────────────────────────────────────────────────

export function formatAlertMessage(summary: ValidationSummary): string {
  const parts: string[] = [];

  if (summary.problematicFiles.length > 0) {
    if (summary.problematicFiles.length === 1) {
      const file = summary.problematicFiles[0];
      const sheetList = file.problemSheets
        .map((s) => `* ${s.sheetName}${s.firstProblemCell ? ` (célula ${s.firstProblemCell})` : ""}`)
        .join("\n");
      parts.push(`ATENÇÃO: o arquivo '${file.fileName}' possui conteúdo potencialmente confidencial nas seguintes abas:\n\n${sheetList}`);
    } else {
      parts.push("ATENÇÃO: os seguintes arquivos possuem conteúdo potencialmente confidencial:\n");
      for (const file of summary.problematicFiles) {
        const sheetNames = file.problemSheets
          .map((s) => `[${s.sheetName}]${s.firstProblemCell ? ` (${s.firstProblemCell})` : ""}`)
          .join(", ");
        parts.push(`* ${file.fileName}\n  Abas: ${sheetNames}\n`);
      }
    }
  }

  if (summary.unreadableFiles.length > 0) {
    if (parts.length > 0) {
      parts.push("\nAlém disso, não foi possível validar os seguintes arquivos (protegidos por senha ou corrompidos):\n");
    } else {
      parts.push("ATENÇÃO: não foi possível validar os seguintes arquivos (protegidos por senha ou corrompidos):\n");
    }
    for (const file of summary.unreadableFiles) {
      const errorDetail = file.error?.message ? ` (${file.error.message})` : "";
      parts.push(`* ${file.fileName}${errorDetail}`);
    }
  }

  parts.push("\nVerifique se essas informações podem ser enviadas antes de continuar.");
  return parts.join("\n");
}

// ─── Main event handler ───────────────────────────────────────────────────────

export async function onMessageSendHandler(event: Office.MailboxEvent): Promise<void> {
  log("HANDLER STARTED");
  notify("diag", "⏳ Innova Verifier: iniciando validação...");

  try {
    // 1. Config (synchronous — no network call)
    log("STEP 1: getConfig");
    const config = await getConfig();
    log("STEP 1: done", `startColumn=${config.startColumn}`);

    // 2. List attachments
    log("STEP 2: getExcelAttachments");
    notify("diag", "⏳ Innova Verifier: listando anexos...");
    const excelAttachments = await getExcelAttachments(config.supportedExtensions);
    log("STEP 2: done", `found ${excelAttachments.length} Excel file(s)`);

    if (excelAttachments.length === 0) {
      log("No Excel attachments — allowing send");
      notify("diag", "✅ Innova Verifier: sem anexos Excel — enviando.");
      event.completed({ allowEvent: true });
      return;
    }

    // 3. Download and validate each file
    log("STEP 3: downloading and validating attachments");
    notify("diag", `⏳ Innova Verifier: validando ${excelAttachments.length} arquivo(s)...`);

    const validationPromises = excelAttachments.map(async (attachment): Promise<AttachmentValidationResult> => {
      try {
        log(`  DOWNLOAD start: ${attachment.name}`);
        const base64Content = await getAttachmentContentAsync(attachment.id);
        log(`  DOWNLOAD done: ${attachment.name}`, `base64 length=${base64Content.length}`);

        log(`  PARSE start: ${attachment.name}`);
        const result = validateExcelWorkbook(base64Content, attachment.name, config);
        log(`  PARSE done: ${attachment.name}`, `hasProblem=${result.hasProblem}, sheets=${result.problemSheets.length}`);

        return result;
      } catch (err: any) {
        log(`  ERROR on ${attachment.name}`, err.message);
        return {
          fileName: attachment.name,
          hasProblem: true,
          problemSheets: [],
          error: {
            code: "ATTACHMENT_READ_FAILED",
            message: err.message || "Falha ao obter o conteúdo do anexo."
          }
        };
      }
    });

    const results = await Promise.all(validationPromises);
    const summary = summarizeValidation(results);
    log("STEP 3: done", `hasProblems=${summary.hasProblems}`);

    // 4. Decide outcome
    if (summary.hasProblems) {
      const alertMsg = formatAlertMessage(summary);
      log("BLOCKING SEND — problems found");
      event.completed({ allowEvent: false, errorMessage: alertMsg });
    } else {
      log("ALLOWING SEND — all clear");
      event.completed({ allowEvent: true });
    }

  } catch (error: any) {
    const errMsg = error?.message || String(error);
    log("FATAL ERROR", errMsg);
    notifyError("diag-err", `Innova Verifier: erro inesperado contate o TI — ${errMsg}`);
    event.completed({
      allowEvent: false,
      errorMessage: `Erro de segurança: Não foi possível validar os anexos do e-mail. Por favor, tente novamente.\nDetalhes: ${errMsg}`
    });
  }
}

// ─── Registration ─────────────────────────────────────────────────────────────

log("Script loaded — waiting for Office.onReady()");

Office.onReady(() => {
  log("Office.onReady fired — registering handler");
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  log("Handler registered ✓");
});
