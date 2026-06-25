import * as XLSX from "xlsx";
import { validateExcelWorkbook, summarizeValidation, hasActualContent, AttachmentValidationResult, xlsxReader } from "./excelValidator";
import { AppConfig } from "../config/configService";

const mockConfig: AppConfig = {
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
 * Helper to build an in-memory workbook base64 string.
 */
function createMockWorkbookBase64(sheets: { [name: string]: any[][] }, hiddenSheets?: string[]): string {
  const wb = XLSX.utils.book_new();
  
  const sheetsMeta: any[] = [];

  for (const sheetName in sheets) {
    const ws = XLSX.utils.aoa_to_sheet(sheets[sheetName]);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    
    const isHidden = hiddenSheets?.includes(sheetName);
    sheetsMeta.push({
      name: sheetName,
      Hidden: isHidden ? 1 : 0
    });
  }

  wb.Workbook = { Sheets: sheetsMeta };
  return XLSX.write(wb, { type: "base64", bookType: "xlsx" });
}

describe("Excel Cell Content Verification (hasActualContent)", () => {
  test("should identify filled values as content", () => {
    expect(hasActualContent({ v: "hello", t: "s" })).toBe(true);
    expect(hasActualContent({ v: 123, t: "n" })).toBe(true);
    expect(hasActualContent({ v: true, t: "b" })).toBe(true);
  });

  test("should ignore null, undefined and empty strings", () => {
    expect(hasActualContent({ v: undefined } as any)).toBe(false);
    expect(hasActualContent({ v: null } as any)).toBe(false);
    expect(hasActualContent({ v: "", t: "s" })).toBe(false);
    expect(hasActualContent({ v: "   ", t: "s" })).toBe(false); // trimmed string is empty
  });

  test("should identify formulas as content, even with empty result", () => {
    expect(hasActualContent({ v: "", f: "A1+B1", t: "s" })).toBe(true);
    expect(hasActualContent({ v: null, f: "SUM(A1:A5)", t: "s" } as any)).toBe(true);
  });

  test("should identify cell hyperlinks as content", () => {
    expect(hasActualContent({ v: "", l: { Target: "http://example.com" }, t: "s" })).toBe(true);
  });
});

describe("validateExcelWorkbook", () => {
  test("should pass for Excel with content only in columns A to L", () => {
    // Column L is index 11. 12 columns total (index 0 to 11).
    // Let's populate columns A to L (indices 0 to 11)
    const row: string[] = Array(12).fill("ok"); 
    const base64 = createMockWorkbookBase64({
      "Planilha1": [row]
    });

    const result = validateExcelWorkbook(base64, "teste_ok.xlsx", mockConfig);
    expect(result.hasProblem).toBe(false);
    expect(result.problemSheets.length).toBe(0);
    expect(result.error).toBeUndefined();
  });

  test("should fail for Excel with content in column M", () => {
    // Column M is index 12. 13 columns (index 0 to 12).
    const row: any[] = Array(12).fill("");
    row.push("violation"); // Column index 12 (M)
    const base64 = createMockWorkbookBase64({
      "Planilha1": [row]
    });

    const result = validateExcelWorkbook(base64, "teste_m_fail.xlsx", mockConfig);
    expect(result.hasProblem).toBe(true);
    expect(result.problemSheets.length).toBe(1);
    expect(result.problemSheets[0].sheetName).toBe("Planilha1");
    expect(result.problemSheets[0].firstProblemCell).toBe("M1");
  });

  test("should fail for Excel with content in column Z", () => {
    // Column Z is index 25.
    const row: any[] = Array(25).fill("");
    row.push(999); // Column index 25 (Z)
    const base64 = createMockWorkbookBase64({
      "Planilha1": [row]
    });

    const result = validateExcelWorkbook(base64, "teste_z_fail.xlsx", mockConfig);
    expect(result.hasProblem).toBe(true);
    expect(result.problemSheets.length).toBe(1);
    expect(result.problemSheets[0].sheetName).toBe("Planilha1");
    expect(result.problemSheets[0].firstProblemCell).toBe("Z1");
  });

  test("should validate multiple sheets and list only problematic ones", () => {
    const rowOk: any[] = Array(12).fill("ok");
    const rowFail: any[] = Array(12).fill("");
    rowFail.push("bad"); // column M
    
    const base64 = createMockWorkbookBase64({
      "Aba1": [rowOk],
      "Aba2": [rowFail],
      "Aba3": [rowOk]
    });

    const result = validateExcelWorkbook(base64, "multi_sheets.xlsx", mockConfig);
    expect(result.hasProblem).toBe(true);
    expect(result.problemSheets.length).toBe(1);
    expect(result.problemSheets[0].sheetName).toBe("Aba2");
  });

  test("should validate hidden sheets if configured to do so", () => {
    const rowFail: any[] = Array(12).fill("");
    rowFail.push("bad"); // column M

    const base64 = createMockWorkbookBase64({
      "Visivel": [["ok"]],
      "Oculta": [rowFail]
    }, ["Oculta"]);

    // Test with validation of hidden sheets enabled
    const result1 = validateExcelWorkbook(base64, "hidden_enabled.xlsx", mockConfig);
    expect(result1.hasProblem).toBe(true);
    expect(result1.problemSheets.length).toBe(1);
    expect(result1.problemSheets[0].sheetName).toBe("Oculta");

    // Test with validation of hidden sheets disabled
    const result2 = validateExcelWorkbook(base64, "hidden_disabled.xlsx", {
      ...mockConfig,
      validateHiddenSheets: false
    });
    expect(result2.hasProblem).toBe(false);
    expect(result2.problemSheets.length).toBe(0);
  });

  test("should return error code for corrupted files (using null)", () => {
    const result = validateExcelWorkbook(null as any, "corrompido.xlsx", mockConfig);
    expect(result.hasProblem).toBe(true);
    expect(result.error).toBeDefined();
    expect(result.error?.code).toBe("CORRUPT_OR_UNREADABLE");
  });

  test("should return PASSWORD_PROTECTED error code if reading throws password error", () => {
    const readSpy = jest.spyOn(xlsxReader, "read").mockImplementation(() => {
      throw new Error("Password protected workbook");
    });

    const result = validateExcelWorkbook("dummy", "protegido.xlsx", mockConfig);
    expect(result.hasProblem).toBe(true);
    expect(result.error).toBeDefined();
    expect(result.error?.code).toBe("PASSWORD_PROTECTED");

    readSpy.mockRestore();
  });
});

describe("summarizeValidation", () => {
  test("should separate problematic, unreadable, and successful validations", () => {
    const results: AttachmentValidationResult[] = [
      { fileName: "ok.xlsx", hasProblem: false, problemSheets: [] },
      { fileName: "ruim.xlsx", hasProblem: true, problemSheets: [{ sheetName: "Plan1", firstProblemCell: "M2" }] },
      { fileName: "senha.xlsx", hasProblem: true, problemSheets: [], error: { code: "PASSWORD_PROTECTED", message: "protegido" } }
    ];

    const summary = summarizeValidation(results);
    expect(summary.hasProblems).toBe(true);
    expect(summary.problematicFiles.length).toBe(1);
    expect(summary.problematicFiles[0].fileName).toBe("ruim.xlsx");
    expect(summary.unreadableFiles.length).toBe(1);
    expect(summary.unreadableFiles[0].fileName).toBe("senha.xlsx");
  });
});
