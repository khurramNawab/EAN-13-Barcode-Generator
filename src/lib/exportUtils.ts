import { jsPDF } from "jspdf";
import JsBarcode from "jsbarcode";

export interface Company {
  id: string;
  name: string;
  pan: string;
  createdAt?: string;
}

export interface Barcode {
  id: string;
  code: string;
  sequence: number;
  rotatingDigit: number;
  datePart: string;
  checkDigit: number;
  productSKU: string | null;
  productName: string | null;
  productDesc: string | null;
  companyId: string;
  company: Company;
  createdAt: string;
}

import * as XLSX from "xlsx";

/**
 * Export array of barcodes to a native Excel (.xlsx) file.
 */
export function exportToExcel(barcodes: Barcode[], filename: string = "barcodes_export.xlsx") {
  if (!barcodes || barcodes.length === 0) return;

  // Prepare data array of objects
  const data = barcodes.map((item) => ({
    "EAN-13 Barcode": String(item.code),
    "Company Name": item.company?.name || "",
    "PAN Number": item.company?.pan || "",
    "Product SKU": item.productSKU || "",
    "Product Name": item.productName || "",
    "Product Description": item.productDesc || "",
    "Sequence Number": String(item.sequence).padStart(2, "0"),
    "Rotating Digit": item.rotatingDigit,
    "Date Part (YYDD)": item.datePart,
    "Check Digit": item.checkDigit,
    "Created Date & Time": new Date(item.createdAt).toLocaleString(),
  }));

  // Create sheet
  const worksheet = XLSX.utils.json_to_sheet(data);

  // Force the Barcode column to be explicitly formatted as a text/string cell.
  // This completely prevents Microsoft Excel from converting 13-digit numbers to scientific notation.
  Object.keys(worksheet).forEach((key) => {
    // Column 'A' corresponds to the first column (EAN-13 Barcode)
    if (key.startsWith("A") && key !== "A1") {
      const cell = worksheet[key];
      if (cell) {
        cell.t = "s"; // Set cell type to String
        cell.z = "@"; // Excel format code for Text
        // Ensure value is a string representation of the code
        cell.v = String(cell.v);
      }
    }
  });

  // Create workbook
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Barcodes");

  // Write workbook as binary array
  const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });

  // MIME type for Excel OpenXML (.xlsx) files
  const blob = new Blob([excelBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8",
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;

  // Enforce .xlsx extension
  const targetFilename = filename.toLowerCase().endsWith(".xlsx")
    ? filename
    : `${filename.replace(/\.[^/.]+$/, "")}.xlsx`;

  link.download = targetFilename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Export array of barcodes to PDF grid document.
 */
export function exportToPDF(
  barcodes: Barcode[],
  filename: string = "barcodes_export.pdf",
  options: { cols?: number; rows?: number } = {}
) {
  if (!barcodes || barcodes.length === 0) return;

  const cols = options.cols || 3;
  const rows = options.rows || 8;
  const itemsPerPage = cols * rows;

  const doc = new jsPDF({
    orientation: "portrait",
    unit: "mm",
    format: "a4",
  });

  const pageWidth = 210;
  const pageHeight = 297;
  const marginX = 10;
  const marginY = 15;

  const colWidth = (pageWidth - marginX * 2) / cols;
  const rowHeight = (pageHeight - marginY * 2) / rows;

  barcodes.forEach((item, index) => {
    if (index > 0 && index % itemsPerPage === 0) {
      doc.addPage();
    }

    const itemInPageIndex = index % itemsPerPage;
    const colIndex = itemInPageIndex % cols;
    const rowIndex = Math.floor(itemInPageIndex / cols);

    const x = marginX + colIndex * colWidth;
    const y = marginY + rowIndex * rowHeight;

    // Render Barcode image on canvas
    const canvas = document.createElement("canvas");
    try {
      JsBarcode(canvas, item.code.substring(0, 12), {
        format: "EAN13",
        displayValue: true,
        font: "monospace",
        fontSize: 16,
        background: "#ffffff",
        lineColor: "#000000",
        width: 5,
        height: 150,
        margin: 8,
        textMargin: 4,
      });

      const imgData = canvas.toDataURL("image/png");

      // Draw bounding box
      doc.setDrawColor(220, 225, 230);
      doc.rect(x + 1, y + 1, colWidth - 2, rowHeight - 2);

      // Add Company & SKU Header (Small text)
      doc.setFontSize(7);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(50, 50, 50);
      const companyNameShort = (item.company?.name || "Company").substring(0, 24);
      doc.text(companyNameShort, x + colWidth / 2, y + 5, { align: "center" });

      if (item.productSKU || item.productName) {
        doc.setFontSize(6.5);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(100, 100, 100);
        const prodInfo = `${item.productSKU || ""} ${item.productName || ""}`.trim().substring(0, 28);
        doc.text(prodInfo, x + colWidth / 2, y + 8.5, { align: "center" });
      }

      // Draw barcode image inside cell
      const imgY = item.productSKU || item.productName ? y + 9.5 : y + 6;
      const imgHeight = rowHeight - (imgY - y) - 3;
      doc.addImage(imgData, "PNG", x + 3, imgY, colWidth - 6, Math.max(12, imgHeight));

    } catch (err) {
      console.error(`Failed to generate barcode PDF cell for ${item.code}`, err);
    }
  });

  const targetFilename = filename.toLowerCase().endsWith(".pdf")
    ? filename
    : `${filename.replace(/\.[^/.]+$/, "")}.pdf`;

  doc.save(targetFilename);
}

/**
 * Helper to load an image client-side.
 */
const loadImage = (src: string): Promise<HTMLImageElement> => {
  return new Promise((resolve, reject) => {
    if (typeof window === "undefined") {
      reject(new Error("loadImage is browser-only"));
      return;
    }
    const img = new Image();
    img.src = src;
    img.onload = () => resolve(img);
    img.onerror = (err) => reject(err);
  });
};

/**
 * Export selected barcodes as a premium landscape Certificate of Barcode Allocation.
 */
export async function exportToCertificate(
  barcodes: Barcode[],
  filename: string = "barcode_allocation_certificate.pdf"
) {
  if (!barcodes || barcodes.length === 0) return;

  const doc = new jsPDF({
    orientation: "landscape",
    unit: "mm",
    format: "a4",
  });

  const pageWidth = 297;
  const pageHeight = 210;
  const centerX = pageWidth / 2;

  // Load single certificate template image client-side
  let singleTemplateImg: HTMLImageElement | null = null;
  try {
    singleTemplateImg = await loadImage("/single_certificate_template.jpg");
  } catch (err) {
    console.error("Failed to load single certificate template image:", err);
  }

  for (let i = 0; i < barcodes.length; i++) {
    if (i > 0) {
      doc.addPage();
    }

    const item = barcodes[i];
    const companyName = (item.company?.name || "Your Organization").toUpperCase();
    const barcodeCode = item.code;

    // Use barcode creation date or today's date
    const certDate = new Date(item.createdAt);
    const dayStr = String(certDate.getDate()).padStart(2, "0");
    const monthStr = String(certDate.getMonth() + 1).padStart(2, "0");
    const yearStr = String(certDate.getFullYear());
    const dateFormatted = `${dayStr}/${monthStr}/${yearStr}`;

    // 1. Draw Template Image Background
    if (singleTemplateImg) {
      doc.addImage(singleTemplateImg, "JPEG", 0, 0, pageWidth, pageHeight);
    } else {
      // Fallback in case template fails to load (draw basic border and text structure)
      doc.setFillColor(252, 250, 246);
      doc.rect(0, 0, pageWidth, pageHeight, "F");
      doc.setDrawColor(29, 70, 61);
      doc.setLineWidth(2);
      doc.rect(6, 6, pageWidth - 12, pageHeight - 12, "D");
      doc.setFont("times", "bold");
      doc.setFontSize(36);
      doc.text("CERTIFICATE", centerX, 35, { align: "center" });
    }

    // 2. Draw Company Name (Centered on underline)
    doc.setFont("times", "bold");
    let companyFontSize = 24;
    if (companyName.length > 20) {
      companyFontSize = Math.max(14, 24 - (companyName.length - 20) * 0.4);
    }
    doc.setFontSize(companyFontSize);
    doc.setTextColor(22, 59, 49); // Swadesh Ghee Green (#163b31)
    doc.text(companyName, centerX, 98.0, { align: "center" });

    // 3. Draw Barcode Values
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13.5);
    doc.setTextColor(0, 0, 0);

    // Barcode Code (centered on single template underline)
    doc.text(barcodeCode, 115.0, 106.8, { align: "center" });

    // Quantity (01 for individual certificate)
    doc.text("01", 238.3, 106.8, { align: "center" });

    // 4. Draw Date
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12.5);
    doc.setTextColor(0, 0, 0);
    doc.text(dateFormatted, 158.3, 138.5, { align: "center" });
  }

  const targetFilename = filename.toLowerCase().endsWith(".pdf")
    ? filename
    : `${filename.replace(/\.[^/.]+$/, "")}.pdf`;

  doc.save(targetFilename);
}
