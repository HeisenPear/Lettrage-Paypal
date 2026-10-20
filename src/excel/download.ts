import type ExcelJS from "exceljs";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** Genere le .xlsx en memoire et declenche son telechargement. */
export async function downloadWorkbook(
  wb: ExcelJS.Workbook,
  filename: string
): Promise<void> {
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: XLSX_MIME });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

/** "Duhallé.csv" -> "duhalle" : sans accents ni caracteres genants dans un nom de fichier. */
function slug(text: string): string {
  return text
    .replace(/\.[^.]+$/, "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-zA-Z0-9-_]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase();
}

/**
 * Construit un nom de fichier propre :
 * lettrage-<source>-<AAAAMMJJ>.xlsx, ou lettrage-<activite>-<source>-<AAAAMMJJ>.xlsx
 */
export function buildFilename(sourceName: string, activity?: string): string {
  const base = slug(sourceName) || "export";
  const d = new Date();
  const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(
    d.getDate()
  ).padStart(2, "0")}`;
  const scope = activity ? `${slug(activity.replace(/\./g, ""))}-` : "";
  return `lettrage-${scope}${base}-${stamp}.xlsx`;
}
