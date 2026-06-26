import * as XLSX from "xlsx";
import { AppConfig } from "../config/configService";

export type ProblemSheet = {
  sheetName: string;
  firstProblemCell?: string;
};

export type AttachmentValidationResult = {
  fileName: string;
  hasProblem: boolean;
  problemSheets: ProblemSheet[];
  error?: {
    code: string;
    message: string;
  };
};

export type ValidationSummary = {
  hasProblems: boolean;
  problematicFiles: AttachmentValidationResult[];
  unreadableFiles: AttachmentValidationResult[];
};

/**
 * Checks if a SheetJS cell object has actual content (value, formula, or hyperlink).
 * Respects RF05 constraints.
 */
export function hasActualContent(cell: XLSX.CellObject): boolean {
  if (!cell) return false;

  const val = cell.v;
  const isStringEmpty = typeof val === "string" && val.trim() === "";
  const hasValue = val !== undefined && val !== null && !isStringEmpty;

  const hasFormula = cell.f !== undefined && cell.f !== null && cell.f !== "";
  const hasHyperlink = cell.l !== undefined && cell.l !== null;

  return hasValue || hasFormula || hasHyperlink;
}

/**
 * Returns the visibility of a sheet:
 * 0 = Visible
 * 1 = Hidden
 * 2 = Very Hidden
 */
export function getSheetVisibility(workbook: XLSX.WorkBook, sheetName: string): number {
  if (workbook.Workbook && workbook.Workbook.Sheets) {
    const sheetMeta = workbook.Workbook.Sheets.find(s => s.name === sheetName);
    if (sheetMeta && typeof sheetMeta.Hidden === "number") {
      return sheetMeta.Hidden;
    }
  }
  return 0; // Default to visible
}

// Exporting reader wrapper to enable test mocking in read-only ESM environments
export const xlsxReader = {
  read(base64Content: string, options: XLSX.ParsingOptions): XLSX.WorkBook {
    return XLSX.read(base64Content, options);
  }
};

/**
 * Validates a single Excel workbook provided as a Base64 string.
 */
export function validateExcelWorkbook(
  base64Content: string,
  fileName: string,
  config: AppConfig
): AttachmentValidationResult {
  const result: AttachmentValidationResult = {
    fileName,
    hasProblem: false,
    problemSheets: []
  };

  try {
    // Parse the Excel file. 
    // xlsxReader.read will throw errors if the file is encrypted/password protected or corrupted.
    const workbook = xlsxReader.read(base64Content, { type: "base64", cellFormula: true });
    
    // Check if the workbook has a "RESUMO" sheet (case-insensitive)
    const hasResumoSheet = workbook.SheetNames.some(
      (name) => name.toLowerCase() === "resumo"
    );
    if (!hasResumoSheet) {
      return result; // Treat as authorized, skip column checks
    }

    for (const sheetName of workbook.SheetNames) {
      const visibility = getSheetVisibility(workbook, sheetName);
      const isHidden = visibility === 1 || visibility === 2;

      // Skip hidden sheets if configured to do so
      if (isHidden && !config.validateHiddenSheets) {
        continue;
      }

      const sheet = workbook.Sheets[sheetName];
      if (!sheet) {
        continue;
      }

      // Determine start column for this sheet (fallback to config.startColumn)
      let startCol = config.startColumn;
      if (config.sheetStartColumns) {
        if (config.sheetStartColumns[sheetName] !== undefined) {
          startCol = config.sheetStartColumns[sheetName];
        } else {
          const lowerSheet = sheetName.toLowerCase();
          const matchKey = Object.keys(config.sheetStartColumns).find(
            (k) => k.toLowerCase() === lowerSheet
          );
          if (matchKey) {
            startCol = config.sheetStartColumns[matchKey];
          }
        }
      }
      const startColIndex = XLSX.utils.decode_col(startCol);

      let firstProblemCell: string | undefined = undefined;

      // Iterate over cells in sheet
      for (const key in sheet) {
        if (key.startsWith("!")) {
          continue;
        }

        const cellCoords = XLSX.utils.decode_cell(key);
        if (cellCoords.c >= startColIndex) {
          const cell = sheet[key];
          if (cell && hasActualContent(cell)) {
            firstProblemCell = key;
            break; // Stop scanning this sheet on first violation (optimization)
          }
        }
      }

      if (firstProblemCell) {
        result.problemSheets.push({
          sheetName,
          firstProblemCell
        });
        result.hasProblem = true;
      }
    }
  } catch (err: any) {
    const msg = (err.message || "").toLowerCase();
    let code = "UNKNOWN_ERROR";
    let message = "Falha ao ler o arquivo Excel.";

    if (
      msg.includes("password") ||
      msg.includes("decrypt") ||
      msg.includes("encrypt") ||
      msg.includes("crypto")
    ) {
      code = "PASSWORD_PROTECTED";
      message = "O arquivo está protegido por senha ou criptografado.";
    } else {
      code = "CORRUPT_OR_UNREADABLE";
      message = "O arquivo parece estar corrompido ou em formato inválido.";
    }

    result.hasProblem = true;
    result.error = { code, message };
  }

  return result;
}

/**
 * Summarizes the validation results for a list of Excel attachments.
 */
export function summarizeValidation(
  results: AttachmentValidationResult[]
): ValidationSummary {
  const problematicFiles: AttachmentValidationResult[] = [];
  const unreadableFiles: AttachmentValidationResult[] = [];

  for (const res of results) {
    if (res.hasProblem) {
      if (res.error) {
        unreadableFiles.push(res);
      } else if (res.problemSheets.length > 0) {
        problematicFiles.push(res);
      }
    }
  }

  return {
    hasProblems: problematicFiles.length > 0 || unreadableFiles.length > 0,
    problematicFiles,
    unreadableFiles
  };
}
