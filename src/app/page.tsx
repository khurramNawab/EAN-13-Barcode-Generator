"use client";

import { useState, useEffect, useRef } from "react";
import dynamic from "next/dynamic";
import {
  Copy,
  Printer,
  Download,
  Check,
  AlertCircle,
  Building,
  CreditCard,
  FileText,
  Barcode as BarcodeIcon,
  Database,
  Settings,
  Plus,
  Search,
  RefreshCw,
  Sliders,
  ChevronLeft,
  ChevronRight,
  Eye,
  Trash,
  Calendar,
  FileSpreadsheet,
  CheckSquare,
  Square,
  XCircle,
  Filter,
  Award,
  Sun,
  Moon,
} from "lucide-react";
import JsBarcode from "jsbarcode";
import { jsPDF } from "jspdf";
import { exportToExcel, exportToPDF, exportToCertificate } from "@/lib/exportUtils";
import * as XLSX from "xlsx";

// Define Barcode type matching DB schema
interface Company {
  id: string;
  name: string;
  pan: string;
  createdAt: string;
  state?: {
    currentSequence: number;
    overflowReset: boolean;
  } | null;
  _count?: {
    barcodes: number;
  };
}

interface Barcode {
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

function BarcodeGenerator() {
  // Form State
  const [companyName, setCompanyName] = useState("");
  const [pan, setPan] = useState("");
  const [productSKU, setProductSKU] = useState("");
  const [productName, setProductName] = useState("");
  const [productDesc, setProductDesc] = useState("");
  const [batchCount, setBatchCount] = useState<number | "">(1);

  // App UI State
  const [generatedBarcodes, setGeneratedBarcodes] = useState<Barcode[]>([]);
  const [selectedBarcodeIndex, setSelectedBarcodeIndex] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);
  const [theme, setTheme] = useState<"dark" | "light">("dark");

  useEffect(() => {
    const savedTheme = localStorage.getItem("theme") as "dark" | "light" | null;
    if (savedTheme) {
      setTheme(savedTheme);
      document.documentElement.classList.toggle("dark", savedTheme === "dark");
      document.documentElement.classList.toggle("light", savedTheme === "light");
    } else {
      setTheme("dark");
      document.documentElement.classList.add("dark");
      document.documentElement.classList.remove("light");
    }
  }, []);

  const toggleTheme = () => {
    const nextTheme = theme === "dark" ? "light" : "dark";
    setTheme(nextTheme);
    localStorage.setItem("theme", nextTheme);
    document.documentElement.classList.toggle("dark", nextTheme === "dark");
    document.documentElement.classList.toggle("light", nextTheme === "light");
  };
  const [copiedCode, setCopiedCode] = useState(false);
  const [copiedImage, setCopiedImage] = useState(false);

  // Resizable split panels state & logic
  const [leftWidth, setLeftWidth] = useState(58); // default to 58% (similar to col-span-7, which is 58.3%)
  const [isDragging, setIsDragging] = useState(false);
  const [isMobile, setIsMobile] = useState(true);
  const containerRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const checkMobile = () => {
      setIsMobile(window.innerWidth < 1024);
    };
    checkMobile();
    window.addEventListener("resize", checkMobile);
    return () => window.removeEventListener("resize", checkMobile);
  }, []);

  const startResizing = (mouseDownEvent: React.MouseEvent) => {
    mouseDownEvent.preventDefault();
    setIsDragging(true);
  };

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!isDragging || !containerRef.current) return;
      
      const containerRect = containerRef.current.getBoundingClientRect();
      const newWidth = ((e.clientX - containerRect.left) / containerRect.width) * 100;
      
      // Limit size between 30% and 75%
      if (newWidth >= 30 && newWidth <= 75) {
        setLeftWidth(newWidth);
      }
    };

    const handleMouseUp = () => {
      setIsDragging(false);
    };

    if (isDragging) {
      window.addEventListener("mousemove", handleMouseMove);
      window.addEventListener("mouseup", handleMouseUp);
    }

    return () => {
      window.removeEventListener("mousemove", handleMouseMove);
      window.removeEventListener("mouseup", handleMouseUp);
    };
  }, [isDragging]);

  // History State
  const [historyBarcodes, setHistoryBarcodes] = useState<Barcode[]>([]);
  const [historyTotal, setHistoryTotal] = useState(0);
  const [historyPage, setHistoryPage] = useState(1);
  const [historyLimit] = useState(10);
  const [historySearch, setHistorySearch] = useState("");
  const [historyCompanyFilter, setHistoryCompanyFilter] = useState("all");
  const [historyStartDate, setHistoryStartDate] = useState("");
  const [historyEndDate, setHistoryEndDate] = useState("");
  const [historyLoading, setHistoryLoading] = useState(false);

  // Multi-selection & Bulk State
  const [selectedBarcodeIds, setSelectedBarcodeIds] = useState<string[]>([]);
  const [bulkExportLoading, setBulkExportLoading] = useState(false);

  // Navigation / Tabs State
  const [activeTab, setActiveTab] = useState<"generate" | "import">("generate");

  // Import Form State
  const [importMode, setImportMode] = useState<"single" | "bulk">("bulk");
  const [importCompanySource, setImportCompanySource] = useState<"existing" | "new">("existing");
  const [importSelectedCompany, setImportSelectedCompany] = useState("");
  const [importCompanyName, setImportCompanyName] = useState("");
  const [importPan, setImportPan] = useState("");
  const [importBarcode, setImportBarcode] = useState("");
  const [importSKU, setImportSKU] = useState("");
  const [importProductName, setImportProductName] = useState("");
  const [importProductDesc, setImportProductDesc] = useState("");

  // Excel Bulk Import State
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [excelCompanyOption, setExcelCompanyOption] = useState<"excel" | "assign">("excel");
  const [excelAssignCompanyId, setExcelAssignCompanyId] = useState("");
  const [parsedRows, setParsedRows] = useState<any[]>([]);
  const [excelSheets, setExcelSheets] = useState<string[]>([]);
  const [selectedExcelSheet, setSelectedExcelSheet] = useState<string>("");
  const workbookRef = useRef<any>(null);
  const [importLoading, setImportLoading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Admin Config State
  const [companies, setCompanies] = useState<Company[]>([]);
  const [adminSelectedCompany, setAdminSelectedCompany] = useState("");
  const [adminOverflowReset, setAdminOverflowReset] = useState(false);
  const [adminLoading, setAdminLoading] = useState(false);

  // Output/Layout options
  const [printLayout, setPrintLayout] = useState<"single" | "sheet" | "custom">("single");
  const [customCols, setCustomCols] = useState(3);
  const [customRows, setCustomRows] = useState(8);
  const [downloadFormat, setDownloadFormat] = useState<"png" | "jpg" | "svg" | "pdf" | "excel">("png");

  // Barcode SVG Ref for rendering preview
  const barcodeSvgRef = useRef<SVGSVGElement | null>(null);

  // Fetch History and Companies on mount / filter changes
  useEffect(() => {
    fetchHistory();
    fetchCompanies();
  }, [historyPage, historyCompanyFilter, historyStartDate, historyEndDate]);

  // Handle Search Debounce
  useEffect(() => {
    const delayDebounce = setTimeout(() => {
      setHistoryPage(1);
      fetchHistory();
    }, 400);

    return () => clearTimeout(delayDebounce);
  }, [historySearch]);

  // Render Barcode preview in UI when selection changes
  useEffect(() => {
    if (generatedBarcodes.length > 0 && barcodeSvgRef.current) {
      const activeBarcode = generatedBarcodes[selectedBarcodeIndex];
      // Feed the first 12 digits to JsBarcode. It computes and appends the 13th automatically
      const code12 = activeBarcode.code.substring(0, 12);
      try {
        JsBarcode(barcodeSvgRef.current, code12, {
          format: "EAN13",
          displayValue: true,
          font: "var(--font-mono)",
          fontSize: 16,
          background: "transparent",
          lineColor: theme === "dark" ? "#f8fafc" : "#0f172a",
          width: 2.2,
          height: 85,
          margin: 10,
        });
      } catch (err) {
        console.error("JsBarcode failed to render", err);
      }
    }
  }, [generatedBarcodes, selectedBarcodeIndex, theme]);

  // Load Companies
  const fetchCompanies = async () => {
    try {
      const res = await fetch("/api/companies");
      if (res.ok) {
        const data = await res.json();
        setCompanies(data.companies || []);
      }
    } catch (err) {
      console.error("Error fetching companies:", err);
    }
  };

  // Load History
  const fetchHistory = async () => {
    setHistoryLoading(true);
    try {
      const query = new URLSearchParams({
        page: historyPage.toString(),
        limit: historyLimit.toString(),
        search: historySearch,
        companyId: historyCompanyFilter,
      });
      if (historyStartDate) query.append("startDate", historyStartDate);
      if (historyEndDate) query.append("endDate", historyEndDate);

      const res = await fetch(`/api/barcodes?${query}`);
      if (res.ok) {
        const data = await res.json();
        setHistoryBarcodes(data.barcodes || []);
        setHistoryTotal(data.pagination.total || 0);
      }
    } catch (err) {
      console.error("Error fetching history:", err);
    } finally {
      setHistoryLoading(false);
    }
  };

  // Multi-Selection Handlers
  const handleToggleSelectBarcode = (id: string) => {
    setSelectedBarcodeIds((prev) =>
      prev.includes(id) ? prev.filter((i) => i !== id) : [...prev, id]
    );
  };

  const isPageAllSelected =
    historyBarcodes.length > 0 &&
    historyBarcodes.every((b) => selectedBarcodeIds.includes(b.id));

  const handleToggleSelectPage = () => {
    if (isPageAllSelected) {
      const pageIds = new Set(historyBarcodes.map((b) => b.id));
      setSelectedBarcodeIds((prev) => prev.filter((id) => !pageIds.has(id)));
    } else {
      const pageIds = historyBarcodes.map((b) => b.id);
      setSelectedBarcodeIds((prev) => Array.from(new Set([...prev, ...pageIds])));
    }
  };

  const handleSelectAllFiltered = async () => {
    setBulkExportLoading(true);
    try {
      const query = new URLSearchParams({
        all: "true",
        search: historySearch,
        companyId: historyCompanyFilter,
      });
      if (historyStartDate) query.append("startDate", historyStartDate);
      if (historyEndDate) query.append("endDate", historyEndDate);

      const res = await fetch(`/api/barcodes?${query}`);
      if (res.ok) {
        const data = await res.json();
        const allIds = (data.barcodes || []).map((b: Barcode) => b.id);
        setSelectedBarcodeIds(allIds);
        setSuccessMsg(`Selected all ${allIds.length} barcodes matching current filter!`);
      }
    } catch (err) {
      console.error("Failed to select all filtered:", err);
      setError("Failed to select all filtered barcodes.");
    } finally {
      setBulkExportLoading(false);
    }
  };

  const handleDeselectAll = () => {
    setSelectedBarcodeIds([]);
  };

  // Retrieve full details of selected barcodes for export
  const getBarcodesToExport = async (): Promise<Barcode[]> => {
    if (selectedBarcodeIds.length === 0) return [];

    const currentMap = new Map(historyBarcodes.map((b) => [b.id, b]));
    const missingIds = selectedBarcodeIds.filter((id) => !currentMap.has(id));

    if (missingIds.length === 0) {
      return selectedBarcodeIds
        .map((id) => currentMap.get(id))
        .filter((b): b is Barcode => Boolean(b));
    }

    // Fetch all matching barcodes from server to ensure complete export
    const query = new URLSearchParams({
      all: "true",
      search: historySearch,
      companyId: historyCompanyFilter,
    });
    if (historyStartDate) query.append("startDate", historyStartDate);
    if (historyEndDate) query.append("endDate", historyEndDate);

    const res = await fetch(`/api/barcodes?${query}`);
    if (res.ok) {
      const data = await res.json();
      const fetched: Barcode[] = data.barcodes || [];
      const selectedSet = new Set(selectedBarcodeIds);
      return fetched.filter((b) => selectedSet.has(b.id));
    }
    return [];
  };

  const handleBulkExportExcel = async () => {
    if (selectedBarcodeIds.length === 0) return;
    setBulkExportLoading(true);
    try {
      const items = await getBarcodesToExport();
      if (items.length > 0) {
        exportToExcel(
          items,
          `EAN13_Barcodes_Export_${new Date().toISOString().slice(0, 10)}.xlsx`
        );
        setSuccessMsg(`Successfully exported ${items.length} barcode(s) to Excel (.xlsx)!`);
      }
    } catch (err) {
      console.error("Excel Export Error:", err);
      setError("Failed to export barcodes to Excel.");
    } finally {
      setBulkExportLoading(false);
    }
  };

  const handleBulkExportPDF = async () => {
    if (selectedBarcodeIds.length === 0) return;
    setBulkExportLoading(true);
    try {
      const items = await getBarcodesToExport();
      if (items.length > 0) {
        exportToPDF(
          items,
          `EAN13_Barcodes_Export_${new Date().toISOString().slice(0, 10)}.pdf`,
          {
            cols: printLayout === "sheet" ? 3 : customCols,
            rows: printLayout === "sheet" ? 8 : customRows,
          }
        );
        setSuccessMsg(`Successfully exported ${items.length} barcode(s) to PDF (.pdf)!`);
      }
    } catch (err) {
      console.error("PDF Export Error:", err);
      setError("Failed to export barcodes to PDF.");
    } finally {
      setBulkExportLoading(false);
    }
  };

  const handleBulkExportCertificate = async () => {
    if (selectedBarcodeIds.length === 0) return;
    setBulkExportLoading(true);
    try {
      const items = await getBarcodesToExport();
      if (items.length > 0) {
        await exportToCertificate(
          items,
          `Barcode_Allocation_Certificate_${new Date().toISOString().slice(0, 10)}.pdf`
        );
        setSuccessMsg(`Successfully generated certificate(s) for ${items.length} barcode(s)!`);
      }
    } catch (err) {
      console.error("Certificate Export Error:", err);
      setError("Failed to generate allocation certificate.");
    } finally {
      setBulkExportLoading(false);
    }
  };

  // Generate Barcode Handler
  const handleGenerate = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError(null);
    setSuccessMsg(null);

    try {
      const res = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          companyName,
          pan,
          productSKU: productSKU || undefined,
          productName: productName || undefined,
          productDesc: productDesc || undefined,
          batchCount: Number(batchCount) || 1,
        }),
      });

      const data = await res.json();

      if (!res.ok) {
        if (data.error === "capacity_exceeded") {
          throw new Error(data.message || "Sequence limit exceeded. Please contact admin to toggle Reset mode.");
        }
        throw new Error(data.error || data.message || "Failed to generate barcode.");
      }

      setGeneratedBarcodes(data.barcodes);
      setSelectedBarcodeIndex(0);
      setSuccessMsg(`Successfully generated ${data.barcodes.length} compliant EAN-13 barcode(s)!`);

      // Clear product specific details after generation
      setProductSKU("");
      setProductName("");
      setProductDesc("");

      // Refresh list tables
      fetchHistory();
      fetchCompanies();
    } catch (err: any) {
      setError(err.message || "An unexpected error occurred.");
    } finally {
      setLoading(false);
    }
  };

  // Revalidate Parsed Excel Rows reactively when options change
  useEffect(() => {
    if (parsedRows.length === 0) return;

    const panRegex = /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/i;
    const codeSeen = new Set<string>();

    const updatedRows = parsedRows.map((row) => {
      let status = "valid";
      let reason = "";

      // Validate Barcode
      if (!row.code) {
        status = "invalid";
        reason = "Missing barcode code";
      } else if (row.code.length !== 13 || !/^\d+$/.test(row.code)) {
        status = "invalid";
        reason = "Must be exactly 13 digits";
      } else if (codeSeen.has(row.code)) {
        status = "duplicate";
        reason = "Duplicate barcode in Excel sheet";
      } else {
        codeSeen.add(row.code);
      }

      // If mapping from Excel, validate PAN/Company Name
      if (status === "valid" && excelCompanyOption === "excel") {
        if (!row.companyName) {
          status = "invalid";
          reason = "Missing company name";
        } else if (!row.pan) {
          status = "invalid";
          reason = "Missing PAN number";
        } else if (!panRegex.test(row.pan)) {
          status = "invalid";
          reason = "Invalid PAN format (AAAAA9999A)";
        }
      } else if (status === "valid" && excelCompanyOption === "assign") {
        if (!excelAssignCompanyId) {
          status = "invalid";
          reason = "No target company selected";
        }
      }

      return { ...row, status, reason };
    });

    const hasChanged = JSON.stringify(updatedRows.map(r => ({s: r.status, re: r.reason}))) !== 
                      JSON.stringify(parsedRows.map(r => ({s: r.status, re: r.reason})));
                      
    if (hasChanged) {
      setParsedRows(updatedRows);
    }
  }, [excelCompanyOption, excelAssignCompanyId, parsedRows.length]);

  // Parser logic for a specific sheet's raw row objects
  const parseSheetRows = (rawRows: any[], sheetName: string) => {
    if (rawRows.length === 0) {
      throw new Error(`The selected Excel sheet "${sheetName}" appears to be empty.`);
    }

    // Dynamically normalize column headers
    const normalized = rawRows.map((row: any, idx: number) => {
      const getValue = (keys: string[]) => {
        const foundKey = Object.keys(row).find((k) =>
          keys.some(
            (key) =>
              k.toLowerCase().replace(/[^a-z0-9]/g, "") ===
              key.toLowerCase().replace(/[^a-z0-9]/g, "")
          )
        );
        return foundKey ? String(row[foundKey]).trim() : "";
      };

      const barcodeValue = getValue([
        "barcode", "ean13barcode", "ean13", "code", "barcodevalue", "barcodenumber", 
        "gtin", "gtin13", "gtinnumber", "productqrbarcodeno", "productqrbarcode", 
        "barcodeno", "qrbarcodeno", "productbarcode"
      ]);
      const compName = getValue(["companyname", "company", "customername", "customer", "org", "organization", "manufacturername", "manufacturer"]);
      let panVal = getValue(["pannumber", "pan", "pancard", "companypan"]);
      const skuVal = getValue(["productsku", "sku", "productid", "id", "skucode"]);
      const nameVal = getValue(["productname", "product", "name", "itemname"]);
      const descVal = getValue(["productdescription", "productdesc", "description", "desc", "itemdesc"]);

      // Fallback: Generate a deterministic PAN based on company name if missing
      if (!panVal && compName) {
        const cleanName = compName.replace(/[^A-Z]/gi, "").toUpperCase();
        const prefix = (cleanName + "XXXXX").substring(0, 5);
        let hash = 0;
        for (let i = 0; i < compName.length; i++) {
          hash = compName.charCodeAt(i) + ((hash << 5) - hash);
        }
        const digits = String(Math.abs(hash) % 10000).padStart(4, "0");
        panVal = `${prefix}${digits}Z`;
      }

      return {
        rowNumber: idx + 2, // Excel rows are 1-based, plus header row
        code: barcodeValue,
        companyName: compName,
        pan: panVal,
        productSKU: skuVal,
        productName: nameVal,
        productDesc: descVal,
        status: "pending",
        reason: "",
      };
    });

    // Run initial validation
    const panRegex = /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/i;
    const codeSeen = new Set<string>();

    const validatedRows = normalized.map((row) => {
      let status = "valid";
      let reason = "";

      // Validate Barcode
      if (!row.code) {
        status = "invalid";
        reason = "Missing barcode code";
      } else if (row.code.length !== 13 || !/^\d+$/.test(row.code)) {
        status = "invalid";
        reason = "Must be exactly 13 digits";
      } else if (codeSeen.has(row.code)) {
        status = "duplicate";
        reason = "Duplicate barcode in Excel sheet";
      } else {
        codeSeen.add(row.code);
      }

      // If mapping from Excel, validate PAN/Company Name
      if (status === "valid" && excelCompanyOption === "excel") {
        if (!row.companyName) {
          status = "invalid";
          reason = "Missing company name";
        } else if (!row.pan) {
          status = "invalid";
          reason = "Missing PAN number";
        } else if (!panRegex.test(row.pan)) {
          status = "invalid";
          reason = "Invalid PAN format (AAAAA9999A)";
        }
      } else if (status === "valid" && excelCompanyOption === "assign") {
        if (!excelAssignCompanyId) {
          status = "invalid";
          reason = "No target company selected";
        }
      }

      return { ...row, status, reason };
    });

    setParsedRows(validatedRows);
    setSuccessMsg(`Successfully parsed ${validatedRows.length} rows from Excel sheet "${sheetName}".`);
  };

  const handleSheetChange = (sheetName: string) => {
    setSelectedExcelSheet(sheetName);
    setError(null);
    setSuccessMsg(null);
    if (!workbookRef.current) return;
    
    try {
      const worksheet = workbookRef.current.Sheets[sheetName];
      const rawRows = XLSX.utils.sheet_to_json<any>(worksheet);
      parseSheetRows(rawRows, sheetName);
    } catch (err: any) {
      console.error("Excel sheet parse error:", err);
      setError(err.message || "Failed to parse selected sheet.");
      setParsedRows([]);
    }
  };

  // Client-side Excel Parser
  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setExcelFile(file);
    setError(null);
    setSuccessMsg(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        workbookRef.current = workbook;

        const sheetNames = workbook.SheetNames;
        setExcelSheets(sheetNames);

        // Find best sheet containing barcode columns
        let bestSheet = sheetNames[0];
        for (const name of sheetNames) {
          const ws = workbook.Sheets[name];
          const sheetData = XLSX.utils.sheet_to_json<any>(ws, { header: 1 });
          if (sheetData && sheetData.length > 0) {
            const headers = sheetData[0].map((h: any) => String(h).toLowerCase().replace(/[^a-z0-9]/g, ""));
            const hasBarcode = headers.some((h: string) =>
              ["barcode", "ean13barcode", "ean13", "code", "barcodevalue", "barcodenumber", "gtin", "gtin13", "gtinnumber", "productqrbarcodeno", "productqrbarcode", "barcodeno", "qrbarcodeno", "productbarcode"].some(
                (key) => h === key.replace(/[^a-z0-9]/g, "")
              )
            );
            if (hasBarcode) {
              bestSheet = name;
              break;
            }
          }
        }

        setSelectedExcelSheet(bestSheet);
        const worksheet = workbook.Sheets[bestSheet];
        const rawRows = XLSX.utils.sheet_to_json<any>(worksheet);
        parseSheetRows(rawRows, bestSheet);
      } catch (err: any) {
        console.error("Excel parse error:", err);
        setError(err.message || "Failed to parse Excel file. Make sure it is a valid .xlsx or .xls file.");
        setParsedRows([]);
        setExcelSheets([]);
        setSelectedExcelSheet("");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Import Submission Handler
  const handleImportSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setImportLoading(true);
    setError(null);
    setSuccessMsg(null);

    try {
      let itemsToImport: any[] = [];

      if (importMode === "single") {
        // Validate single barcode fields
        const panRegex = /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/i;
        let finalCompanyName = "";
        let finalPan = "";

        if (importCompanySource === "existing") {
          if (!importSelectedCompany) {
            throw new Error("Please select an existing company.");
          }
          const comp = companies.find((c) => c.id === importSelectedCompany);
          if (!comp) throw new Error("Selected company not found.");
          finalCompanyName = comp.name;
          finalPan = comp.pan;
        } else {
          if (!importCompanyName.trim()) {
            throw new Error("Company name is required.");
          }
          if (!importPan.trim()) {
            throw new Error("PAN number is required.");
          }
          if (!panRegex.test(importPan.toUpperCase().trim())) {
            throw new Error("Invalid PAN format (AAAAA9999A).");
          }
          finalCompanyName = importCompanyName.trim();
          finalPan = importPan.toUpperCase().trim();
        }

        const cleanBarcode = importBarcode.trim();
        if (!cleanBarcode) {
          throw new Error("Barcode number is required.");
        }
        if (cleanBarcode.length !== 13 || !/^\d+$/.test(cleanBarcode)) {
          throw new Error("Barcode must be exactly 13 digits.");
        }

        itemsToImport = [
          {
            code: cleanBarcode,
            companyName: finalCompanyName,
            pan: finalPan,
            productSKU: importSKU || undefined,
            productName: importProductName || undefined,
            productDesc: importProductDesc || undefined,
          },
        ];
      } else {
        // Bulk import from parsedRows
        const validRows = parsedRows.filter((r) => r.status === "valid");
        if (validRows.length === 0) {
          throw new Error("There are no valid rows to import. Please resolve errors in the Excel sheet first.");
        }

        if (excelCompanyOption === "assign") {
          if (!excelAssignCompanyId) {
            throw new Error("Please select a target company to assign the barcodes to.");
          }
          const targetComp = companies.find((c) => c.id === excelAssignCompanyId);
          if (!targetComp) throw new Error("Target company not found.");

          itemsToImport = validRows.map((r) => ({
            code: r.code,
            companyName: targetComp.name,
            pan: targetComp.pan,
            productSKU: r.productSKU,
            productName: r.productName,
            productDesc: r.productDesc,
          }));
        } else {
          itemsToImport = validRows.map((r) => ({
            code: r.code,
            companyName: r.companyName,
            pan: r.pan,
            productSKU: r.productSKU,
            productName: r.productName,
            productDesc: r.productDesc,
          }));
        }
      }

      // Send payload to backend API
      const res = await fetch("/api/barcodes/import", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items: itemsToImport }),
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.message || "Failed to import barcodes.");
      }

      setSuccessMsg(
        `Import completed: ${data.importedCount} barcode(s) successfully imported, ${data.skippedCount} skipped as duplicate or invalid.`
      );

      // If single import, clear inputs
      if (importMode === "single") {
        setImportBarcode("");
        setImportSKU("");
        setImportProductName("");
        setImportProductDesc("");
      } else {
        // Clear Excel states
        setExcelFile(null);
        setParsedRows([]);
        if (fileInputRef.current) fileInputRef.current.value = "";
      }

      // Refresh data tables
      fetchHistory();
      fetchCompanies();
    } catch (err: any) {
      setError(err.message || "An unexpected error occurred during import.");
    } finally {
      setImportLoading(false);
    }
  };

  // Admin Toggle Save
  const handleSaveAdminConfig = async () => {
    if (!adminSelectedCompany) return;
    setAdminLoading(true);
    setError(null);
    setSuccessMsg(null);

    try {
      const res = await fetch("/api/admin/config", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          companyId: adminSelectedCompany,
          overflowReset: adminOverflowReset,
        }),
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || "Failed to update configuration.");
      }

      setSuccessMsg("Configuration updated successfully!");
      fetchCompanies();
    } catch (err: any) {
      setError(err.message);
    } finally {
      setAdminLoading(false);
    }
  };

  // Select company from admin list to populate state
  useEffect(() => {
    if (adminSelectedCompany) {
      const comp = companies.find((c) => c.id === adminSelectedCompany);
      if (comp && comp.state) {
        setAdminOverflowReset(comp.state.overflowReset);
      } else {
        setAdminOverflowReset(false);
      }
    }
  }, [adminSelectedCompany, companies]);

  // Copy 13-digit code text
  const handleCopyCode = () => {
    if (generatedBarcodes.length === 0) return;
    const code = generatedBarcodes[selectedBarcodeIndex].code;
    navigator.clipboard.writeText(code).then(() => {
      setCopiedCode(true);
      setTimeout(() => setCopiedCode(false), 2000);
    });
  };

  // Copy Barcode Image to Clipboard (High Definition)
  const handleCopyImage = () => {
    if (generatedBarcodes.length === 0) return;
    const activeBarcode = generatedBarcodes[selectedBarcodeIndex];

    // Render EAN-13 to an offscreen canvas
    const canvas = document.createElement("canvas");
    try {
      JsBarcode(canvas, activeBarcode.code.substring(0, 12), {
        format: "EAN13",
        displayValue: true,
        font: "monospace",
        fontSize: 18,
        background: "#ffffff",
        lineColor: "#000000",
        width: 6, // High resolution width multiplier
        height: 200, // High resolution height
        margin: 10,
        textMargin: 6,
      });

      canvas.toBlob((blob) => {
        if (blob) {
          navigator.clipboard.write([
            new ClipboardItem({ "image/png": blob })
          ]).then(() => {
            setCopiedImage(true);
            setTimeout(() => setCopiedImage(false), 2000);
          });
        }
      });
    } catch (err) {
      console.error("Failed to copy image:", err);
    }
  };

  // Download Trigger
  const handleDownload = () => {
    if (generatedBarcodes.length === 0) return;

    if (downloadFormat === "svg") {
      downloadSVG();
    } else if (downloadFormat === "png") {
      downloadPNG();
    } else if (downloadFormat === "jpg") {
      downloadJPEG();
    } else if (downloadFormat === "pdf") {
      downloadPDF();
    } else if (downloadFormat === "excel") {
      const company = generatedBarcodes[0]?.company?.name || "barcodes";
      exportToExcel(
        generatedBarcodes,
        `EAN13_${company.replace(/\s+/g, "_")}_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    }
  };

  // Download SVG
  const downloadSVG = () => {
    const activeBarcode = generatedBarcodes[selectedBarcodeIndex];
    const svgElement = document.createElementNS("http://www.w3.org/2000/svg", "svg");

    try {
      JsBarcode(svgElement, activeBarcode.code.substring(0, 12), {
        format: "EAN13",
        displayValue: true,
        fontSize: 16,
        background: "#ffffff",
        lineColor: "#000000",
      });

      const serializer = new XMLSerializer();
      const svgString = serializer.serializeToString(svgElement);
      const blob = new Blob([svgString], { type: "image/svg+xml;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `barcode_${activeBarcode.code}.svg`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("SVG Download failed:", err);
    }
  };

  // Download PNG (High Definition)
  const downloadPNG = () => {
    const activeBarcode = generatedBarcodes[selectedBarcodeIndex];
    const canvas = document.createElement("canvas");
    try {
      JsBarcode(canvas, activeBarcode.code.substring(0, 12), {
        format: "EAN13",
        displayValue: true,
        font: "monospace",
        fontSize: 18,
        background: "#ffffff",
        lineColor: "#000000",
        width: 6, // High resolution width multiplier
        height: 200, // High resolution height
        margin: 10,
        textMargin: 6,
      });

      const url = canvas.toDataURL("image/png");
      const link = document.createElement("a");
      link.href = url;
      link.download = `barcode_${activeBarcode.code}.png`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error("PNG Download failed:", err);
    }
  };

  // Download JPEG (High Definition)
  const downloadJPEG = () => {
    const activeBarcode = generatedBarcodes[selectedBarcodeIndex];
    const canvas = document.createElement("canvas");
    try {
      JsBarcode(canvas, activeBarcode.code.substring(0, 12), {
        format: "EAN13",
        displayValue: true,
        font: "monospace",
        fontSize: 18,
        background: "#ffffff",
        lineColor: "#000000",
        width: 6, // High resolution width multiplier
        height: 200, // High resolution height
        margin: 10,
        textMargin: 6,
      });

      const url = canvas.toDataURL("image/jpeg", 0.95); // High quality JPEG
      const link = document.createElement("a");
      link.href = url;
      link.download = `barcode_${activeBarcode.code}.jpg`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error("JPEG Download failed:", err);
    }
  };

  // Download PDF Sheets or Grid (High Definition)
  const downloadPDF = () => {
    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
    });

    const isGrid = printLayout === "sheet" || printLayout === "custom";
    const cols = printLayout === "sheet" ? 3 : customCols;
    const rows = printLayout === "sheet" ? 8 : customRows;
    const maxPerPage = cols * rows;

    if (!isGrid) {
      // Single centered barcode PDF
      const activeBarcode = generatedBarcodes[selectedBarcodeIndex];
      const canvas = document.createElement("canvas");
      JsBarcode(canvas, activeBarcode.code.substring(0, 12), {
        format: "EAN13",
        displayValue: true,
        font: "monospace",
        fontSize: 18,
        background: "#ffffff",
        lineColor: "#000000",
        width: 6, // High resolution width multiplier
        height: 200, // High resolution height
        margin: 10,
        textMargin: 6,
      });
      const imgData = canvas.toDataURL("image/png");

      // Page dimensions: 210 x 297 mm
      doc.addImage(imgData, "PNG", 35, 113.5, 140, 70);

      doc.save(`barcode_${activeBarcode.code}.pdf`);
    } else {
      // Grid Sheet PDF
      const colWidth = (210 - 20) / cols;
      const rowHeight = (297 - 30) / rows;
      const startX = 10;
      const startY = 15;

      // Determine barcodes list: repeat selected barcode if batch is only 1, otherwise use batch
      const barcodesToPrint = generatedBarcodes.length === 1
        ? Array(maxPerPage).fill(generatedBarcodes[0])
        : generatedBarcodes;

      let index = 0;
      barcodesToPrint.forEach((barcodeItem) => {
        if (index > 0 && index % maxPerPage === 0) {
          doc.addPage();
        }

        const colIndex = index % cols;
        const rowIndex = Math.floor(index / cols) % rows;

        const x = startX + colIndex * colWidth;
        const y = startY + rowIndex * rowHeight;

        // Render barcode to canvas
        const canvas = document.createElement("canvas");
        JsBarcode(canvas, barcodeItem.code.substring(0, 12), {
          format: "EAN13",
          displayValue: true,
          font: "monospace",
          fontSize: 18,
          background: "#ffffff",
          lineColor: "#000000",
          width: 5, // High resolution width multiplier for grid cells
          height: 160, // High resolution height for grid cells
          margin: 10,
          textMargin: 6,
        });
        const imgData = canvas.toDataURL("image/png");

        // Fit image inside cell margins
        doc.addImage(imgData, "PNG", x + 2, y + 2, colWidth - 4, rowHeight - 4);

        index++;
      });

      doc.save(`barcode_sheet_${generatedBarcodes[0].company.name.replace(/\s+/g, "_")}.pdf`);
    }
  };

  // Scope print to barcode only using browser print
  const handlePrint = () => {
    window.print();
  };

  // Helper: Populate form from history item
  const handleSelectFromHistory = (item: Barcode) => {
    setCompanyName(item.company.name);
    setPan(item.company.pan);
    setGeneratedBarcodes([item]);
    setSelectedBarcodeIndex(0);
    setSuccessMsg("Loaded barcode from audit history!");
  };

  return (
    <main className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-8 flex flex-col gap-8 no-print">
      {/* Header Banner */}
      <header className="flex flex-col md:flex-row items-center justify-between gap-4 pb-6 border-b border-slate-200 dark:border-slate-800">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-gradient-to-tr from-blue-600 to-indigo-600 rounded-xl shadow-lg shadow-blue-500/20 animate-pulse">
            <BarcodeIcon className="h-8 w-8 text-white" />
          </div>
          <div>
            <h1 className="text-2xl md:text-3xl font-display font-bold tracking-tight text-slate-900 dark:text-white">
              EAN-13 Barcode Generator
            </h1>
            <p className="text-sm text-slate-500 dark:text-slate-400">
              Compliant Enterprise Generation Console
            </p>
          </div>
        </div>
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-full text-xs text-slate-600 dark:text-slate-400">
            <span className="h-2 w-2 rounded-full bg-emerald-500 animate-ping"></span>
            System Live
          </div>
          <button
            type="button"
            onClick={toggleTheme}
            className="p-2 bg-slate-100 dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-full text-slate-600 dark:text-slate-400 hover:text-slate-950 dark:hover:text-white transition active:scale-95 flex items-center justify-center"
            title={theme === "dark" ? "Switch to Light Mode" : "Switch to Dark Mode"}
          >
            {theme === "dark" ? <Sun className="h-4 w-4" /> : <Moon className="h-4 w-4" />}
          </button>
        </div>
      </header>

      {/* Message Area */}
      {(error || successMsg) && (
        <div className="flex flex-col gap-3">
          {error && (
            <div className="flex items-start gap-3 p-4 bg-red-950/50 border border-red-900/60 rounded-xl text-red-200">
              <AlertCircle className="h-5 w-5 text-red-400 shrink-0 mt-0.5" />
              <div>
                <h4 className="font-semibold text-sm">Action Failed</h4>
                <p className="text-xs text-red-300/90 mt-1">{error}</p>
              </div>
            </div>
          )}
          {successMsg && (
            <div className="flex items-start gap-3 p-4 bg-emerald-950/40 border border-emerald-900/60 rounded-xl text-emerald-200">
              <Check className="h-5 w-5 text-emerald-400 shrink-0 mt-0.5" />
              <div>
                <h4 className="font-semibold text-sm">Success</h4>
                <p className="text-xs text-emerald-300/90 mt-1">{successMsg}</p>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Main Grid: Form & Output */}
      <div 
        ref={containerRef}
        className={`flex flex-col lg:flex-row gap-8 relative ${isDragging ? "select-none" : ""}`}
      >
        {/* Left Side: Input Form */}
        <section 
          className="flex flex-col gap-8 shrink-0 w-full"
          style={{ width: isMobile ? "100%" : `${leftWidth}%` }}
        >
          <div className="bg-white/80 dark:bg-slate-900/60 backdrop-blur-md border border-slate-200 dark:border-slate-850 rounded-2xl p-6 shadow-xl shadow-slate-100 dark:shadow-black/20 flex flex-col gap-6">
            {/* Tab Navigation header */}
            <div className="flex bg-slate-100 dark:bg-slate-950 p-1 rounded-xl border border-slate-200 dark:border-slate-800">
              <button
                type="button"
                onClick={() => {
                  setActiveTab("generate");
                  setError(null);
                  setSuccessMsg(null);
                }}
                className={`flex-1 flex items-center justify-center gap-2 py-2.5 rounded-lg text-xs font-semibold transition-all ${
                  activeTab === "generate"
                    ? "bg-blue-600 text-white shadow-md shadow-blue-900/20"
                    : "text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200"
                }`}
              >
                <Sliders className="h-3.5 w-3.5" />
                Generate New
              </button>
              <button
                type="button"
                onClick={() => {
                  setActiveTab("import");
                  setError(null);
                  setSuccessMsg(null);
                }}
                className={`flex-1 flex items-center justify-center gap-2 py-2.5 rounded-lg text-xs font-semibold transition-all ${
                  activeTab === "import"
                    ? "bg-blue-600 text-white shadow-md shadow-blue-900/20"
                    : "text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200"
                }`}
              >
                <Database className="h-3.5 w-3.5" />
                Import Existing
              </button>
            </div>

            {activeTab === "generate" ? (
              // --- TAB 1: GENERATE NEW BARCODES ---
              <form onSubmit={handleGenerate} className="flex flex-col gap-5">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                  {/* Company Name */}
                  <div className="flex flex-col gap-2">
                    <label htmlFor="companyName" className="text-xs font-semibold text-slate-500 dark:text-slate-400 flex items-center gap-1.5">
                      <Building className="h-3.5 w-3.5" />
                      Company / Customer Name *
                    </label>
                    <input
                      type="text"
                      id="companyName"
                      value={companyName}
                      onChange={(e) => setCompanyName(e.target.value)}
                      placeholder="e.g. Acme Industries Ltd"
                      className="w-full bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-3 text-sm text-slate-800 dark:text-slate-200 transition"
                      required
                    />
                  </div>

                  {/* PAN Number */}
                  <div className="flex flex-col gap-2">
                    <label htmlFor="pan" className="text-xs font-semibold text-slate-500 dark:text-slate-400 flex items-center gap-1.5">
                      <CreditCard className="h-3.5 w-3.5" />
                      PAN Number (AAAAA9999A) *
                    </label>
                    <input
                      type="text"
                      id="pan"
                      value={pan}
                      onChange={(e) => setPan(e.target.value)}
                      placeholder="e.g. ABCDE1234F"
                      className="w-full bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-3 text-sm text-slate-800 dark:text-slate-200 transition uppercase"
                      required
                    />
                  </div>
                </div>

                <div className="border-t border-slate-200 dark:border-slate-800/60 my-2"></div>

                {/* Product Details (Optional) */}
                <div className="flex flex-col gap-4">
                  <div className="flex items-center gap-1 text-slate-500 dark:text-slate-400">
                    <FileText className="h-4 w-4" />
                    <span className="text-xs font-bold uppercase tracking-wider">Product Details (Reference Only)</span>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                    <div className="flex flex-col gap-2">
                      <label htmlFor="productSKU" className="text-xs text-slate-500 dark:text-slate-400">
                        Product SKU / ID
                      </label>
                      <input
                        type="text"
                        id="productSKU"
                        value={productSKU}
                        onChange={(e) => setProductSKU(e.target.value)}
                        placeholder="e.g. PROD-1002"
                        className="w-full bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-2.5 text-sm text-slate-800 dark:text-slate-200 transition"
                      />
                    </div>

                    <div className="flex flex-col gap-2">
                      <label htmlFor="productName" className="text-xs text-slate-500 dark:text-slate-400">
                        Product Name
                      </label>
                      <input
                        type="text"
                        id="productName"
                        value={productName}
                        onChange={(e) => setProductName(e.target.value)}
                        placeholder="e.g. Premium Cotton Shirt"
                        className="w-full bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-2.5 text-sm text-slate-800 dark:text-slate-200 transition"
                      />
                    </div>
                  </div>

                  <div className="flex flex-col gap-2">
                    <label htmlFor="productDesc" className="text-xs text-slate-500 dark:text-slate-400">
                      Product Description
                    </label>
                    <textarea
                      id="productDesc"
                      value={productDesc}
                      onChange={(e) => setProductDesc(e.target.value)}
                      placeholder="Enter short description or SKU properties..."
                      rows={2}
                      className="w-full bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 focus:border-blue-500 focus:ring-1 focus:ring-blue-500 rounded-xl px-4 py-2.5 text-sm text-slate-800 dark:text-slate-200 transition"
                    />
                  </div>
                </div>

                <div className="border-t border-slate-200 dark:border-slate-800/60 my-2"></div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-5 items-end">
                  {/* Batch Count */}
                  <div className="flex flex-col gap-2">
                    <label htmlFor="batchCount" className="text-xs font-semibold text-slate-500 dark:text-slate-400">
                      Batch Generation Size (1 - 50 labels)
                    </label>
                    <div className="flex items-center gap-3">
                      <input
                        type="range"
                        min="1"
                        max="50"
                        value={batchCount || 1}
                        onChange={(e) => setBatchCount(parseInt(e.target.value))}
                        className="flex-1 h-1.5 bg-slate-800 rounded-lg appearance-none cursor-pointer accent-blue-500"
                      />
                      <input
                        type="number"
                        min="1"
                        max="50"
                        value={batchCount}
                        onChange={(e) => {
                          const val = e.target.value;
                          if (val === "") {
                            setBatchCount("");
                          } else {
                            const parsed = parseInt(val);
                            if (!isNaN(parsed)) {
                              setBatchCount(parsed);
                            }
                          }
                        }}
                        onBlur={() => {
                          if (batchCount !== "" && batchCount > 50) {
                            setBatchCount(50);
                          }
                        }}
                        className="w-16 bg-slate-950 border border-slate-800 text-center rounded-xl py-1.5 text-sm text-slate-200"
                      />
                    </div>
                  </div>

                  {/* Submit Button */}
                  <button
                    type="submit"
                    disabled={loading}
                    className="w-full bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 disabled:opacity-50 text-white font-semibold py-3 px-6 rounded-xl transition shadow-lg shadow-blue-900/20 active:scale-[0.98] flex items-center justify-center gap-2 text-sm"
                  >
                    {loading ? (
                      <>
                        <RefreshCw className="h-4 w-4 animate-spin" />
                        Generating...
                      </>
                    ) : (
                      <>
                        <Plus className="h-4.5 w-4.5" />
                        Generate Barcode
                      </>
                    )}
                  </button>
                </div>
              </form>
            ) : (
              // --- TAB 2: IMPORT EXISTING CLIENT BARCODES ---
              <div className="flex flex-col gap-6">
                {/* Mode Selector */}
                <div className="flex justify-center gap-4 border-b border-slate-200 dark:border-slate-800 pb-3">
                  <button
                    type="button"
                    onClick={() => {
                      setImportMode("bulk");
                      setError(null);
                      setSuccessMsg(null);
                    }}
                    className={`px-4 py-2 rounded-lg text-xs font-semibold border transition ${
                      importMode === "bulk"
                        ? "bg-blue-600 text-white border-blue-600 shadow-sm shadow-blue-600/20"
                        : "bg-transparent text-slate-500 dark:text-slate-400 border-transparent hover:text-slate-800 dark:hover:text-slate-200"
                    }`}
                  >
                    Excel File Upload (Bulk)
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setImportMode("single");
                      setError(null);
                      setSuccessMsg(null);
                    }}
                    className={`px-4 py-2 rounded-lg text-xs font-semibold border transition ${
                      importMode === "single"
                        ? "bg-blue-600 text-white border-blue-600 shadow-sm shadow-blue-600/20"
                        : "bg-transparent text-slate-500 dark:text-slate-400 border-transparent hover:text-slate-800 dark:hover:text-slate-200"
                    }`}
                  >
                    Manual Entry (Single)
                  </button>
                </div>

                {importMode === "single" ? (
                  /* MANUAL ENTRY FORM */
                  <form onSubmit={handleImportSubmit} className="flex flex-col gap-5">
                    {/* Company Source */}
                    <div className="flex flex-col gap-2">
                      <label className="text-xs font-semibold text-slate-500 dark:text-slate-400 flex items-center gap-1">
                        <Building className="h-3.5 w-3.5" /> Company Source
                      </label>
                      <div className="grid grid-cols-2 gap-3 bg-slate-100 dark:bg-slate-950 p-1 rounded-xl border border-slate-200 dark:border-slate-850">
                        <button
                          type="button"
                          onClick={() => setImportCompanySource("existing")}
                          className={`py-1.5 rounded-lg text-xs font-semibold transition ${
                            importCompanySource === "existing"
                              ? "bg-blue-600 text-white shadow-sm"
                              : "text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200"
                          }`}
                        >
                          Select Existing
                        </button>
                        <button
                          type="button"
                          onClick={() => setImportCompanySource("new")}
                          className={`py-1.5 rounded-lg text-xs font-semibold transition ${
                            importCompanySource === "new"
                              ? "bg-blue-600 text-white shadow-sm"
                              : "text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200"
                          }`}
                        >
                          Add New Company
                        </button>
                      </div>
                    </div>

                    {/* Company fields based on selection */}
                    {importCompanySource === "existing" ? (
                      <div className="flex flex-col gap-2">
                        <label htmlFor="importSelectedCompany" className="text-xs font-semibold text-slate-500 dark:text-slate-400">
                          Select Company *
                        </label>
                        <select
                          id="importSelectedCompany"
                          value={importSelectedCompany}
                          onChange={(e) => setImportSelectedCompany(e.target.value)}
                          className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-3 text-xs text-slate-800 dark:text-slate-200 w-full focus:outline-none focus:border-blue-500"
                          required
                        >
                          <option value="">-- Choose Company --</option>
                          {companies.map((c) => (
                            <option key={c.id} value={c.id}>
                              {c.name} (PAN: {c.pan})
                            </option>
                          ))}
                        </select>
                      </div>
                    ) : (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="flex flex-col gap-2">
                          <label htmlFor="importCompanyName" className="text-xs font-semibold text-slate-500 dark:text-slate-400">
                            Company Name *
                          </label>
                          <input
                            type="text"
                            id="importCompanyName"
                            value={importCompanyName}
                            onChange={(e) => setImportCompanyName(e.target.value)}
                            placeholder="e.g. Acme Industries Ltd"
                            className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-2.5 text-xs text-slate-800 dark:text-slate-200 focus:outline-none focus:border-blue-500"
                            required
                          />
                        </div>
                        <div className="flex flex-col gap-2">
                          <label htmlFor="importPan" className="text-xs font-semibold text-slate-500 dark:text-slate-400">
                            PAN Number *
                          </label>
                          <input
                            type="text"
                            id="importPan"
                            value={importPan}
                            onChange={(e) => setImportPan(e.target.value)}
                            placeholder="e.g. ABCDE1234F"
                            className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-2.5 text-xs text-slate-800 dark:text-slate-200 uppercase focus:outline-none focus:border-blue-500"
                            required
                          />
                        </div>
                      </div>
                    )}

                    <div className="border-t border-slate-200 dark:border-slate-800/60 my-1"></div>

                    {/* Barcode Number */}
                    <div className="flex flex-col gap-2">
                      <label htmlFor="importBarcode" className="text-xs font-semibold text-slate-500 dark:text-slate-400 flex items-center gap-1.5">
                        <BarcodeIcon className="h-3.5 w-3.5" />
                        EAN-13 Barcode *
                      </label>
                      <div className="relative">
                        <input
                          type="text"
                          id="importBarcode"
                          value={importBarcode}
                          onChange={(e) => setImportBarcode(e.target.value.replace(/\D/g, "").substring(0, 13))}
                          placeholder="e.g. 8530125123412"
                          maxLength={13}
                          className="w-full bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 focus:border-blue-500 rounded-xl px-4 py-3 text-sm text-slate-800 dark:text-slate-200 font-mono tracking-widest"
                          required
                        />
                        <span className="absolute right-3 top-3 text-[10px] font-mono text-slate-500">
                          {importBarcode.length}/13 digits
                        </span>
                      </div>
                    </div>

                    <div className="border-t border-slate-200 dark:border-slate-800/60 my-1"></div>

                    {/* Optional Product Info */}
                    <div className="flex flex-col gap-4">
                      <div className="flex items-center gap-1 text-slate-400">
                        <FileText className="h-3.5 w-3.5" />
                        <span className="text-[10px] font-bold uppercase tracking-wider">Product Details (Optional)</span>
                      </div>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="flex flex-col gap-2">
                          <label htmlFor="importSKU" className="text-xs text-slate-500 dark:text-slate-400">Product SKU / ID</label>
                          <input
                            type="text"
                            id="importSKU"
                            value={importSKU}
                            onChange={(e) => setImportSKU(e.target.value)}
                            placeholder="e.g. SKU-990"
                            className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-800 dark:text-slate-200 focus:outline-none"
                          />
                        </div>
                        <div className="flex flex-col gap-2">
                          <label htmlFor="importProductName" className="text-xs text-slate-500 dark:text-slate-400">Product Name</label>
                          <input
                            type="text"
                            id="importProductName"
                            value={importProductName}
                            onChange={(e) => setImportProductName(e.target.value)}
                            placeholder="e.g. Premium Basmati Rice"
                            className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-800 dark:text-slate-200 focus:outline-none"
                          />
                        </div>
                      </div>
                      <div className="flex flex-col gap-2">
                        <label htmlFor="importProductDesc" className="text-xs text-slate-500 dark:text-slate-400">Product Description</label>
                        <textarea
                          id="importProductDesc"
                          value={importProductDesc}
                          onChange={(e) => setImportProductDesc(e.target.value)}
                          placeholder="e.g. 10kg premium bag..."
                          rows={2}
                          className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-800 dark:text-slate-200 focus:outline-none"
                        />
                      </div>
                    </div>

                    <button
                      type="submit"
                      disabled={importLoading || importBarcode.length !== 13}
                      className="w-full bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 disabled:opacity-50 text-white font-semibold py-3 px-6 rounded-xl transition shadow-lg shadow-emerald-950/20 flex items-center justify-center gap-2 text-sm mt-2"
                    >
                      {importLoading ? (
                        <>
                          <RefreshCw className="h-4 w-4 animate-spin" />
                          Importing Barcode...
                        </>
                      ) : (
                        <>
                          <Database className="h-4 w-4" />
                          Save Barcode to Database
                        </>
                      )}
                    </button>
                  </form>
                ) : (
                  <div className="flex flex-col gap-5">
                    {/* Format instructions */}
                    <div className="bg-slate-100/80 dark:bg-slate-950 border border-slate-200 dark:border-slate-850 p-4 rounded-xl text-slate-600 dark:text-slate-400 text-[11px] leading-relaxed flex flex-col gap-1">
                      <span className="font-bold text-slate-800 dark:text-slate-300">Excel Format Requirements:</span>
                      <span>The sheet must contain a column for <strong className="text-slate-900 dark:text-white">Barcode</strong> (13 digits).</span>
                      <span>If mapping dynamically, include columns for <strong className="text-slate-900 dark:text-white">Company Name</strong> and <strong className="text-slate-900 dark:text-white">PAN</strong>.</span>
                      <span>Columns are case-insensitive and match common synonyms (e.g. <i>sku</i>, <i>desc</i>).</span>
                    </div>

                    {/* Company Setting for Excel Upload */}
                    <div className="flex flex-col gap-2">
                      <label className="text-xs font-semibold text-slate-500 dark:text-slate-400">Company Assignment</label>
                      <div className="grid grid-cols-2 gap-3 bg-slate-100 dark:bg-slate-950 p-1 rounded-xl border border-slate-200 dark:border-slate-850">
                        <button
                          type="button"
                          onClick={() => setExcelCompanyOption("excel")}
                          className={`py-1.5 rounded-lg text-xs font-semibold transition ${
                            excelCompanyOption === "excel"
                              ? "bg-blue-600 text-white shadow-sm"
                              : "text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200"
                          }`}
                        >
                          Map from Excel Rows
                        </button>
                        <button
                          type="button"
                          onClick={() => setExcelCompanyOption("assign")}
                          className={`py-1.5 rounded-lg text-xs font-semibold transition ${
                            excelCompanyOption === "assign"
                              ? "bg-blue-600 text-white shadow-sm"
                              : "text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200"
                          }`}
                        >
                          Assign to One Company
                        </button>
                      </div>
                    </div>

                    {/* Assign Single Company Select */}
                    {excelCompanyOption === "assign" && (
                      <div className="flex flex-col gap-2">
                        <label htmlFor="excelAssignCompanyId" className="text-xs font-semibold text-slate-500 dark:text-slate-400">
                          Select Company *
                        </label>
                        <select
                          id="excelAssignCompanyId"
                          value={excelAssignCompanyId}
                          onChange={(e) => setExcelAssignCompanyId(e.target.value)}
                          className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-3 text-xs text-slate-800 dark:text-slate-200 w-full focus:outline-none focus:border-blue-500"
                          required
                        >
                          <option value="">-- Choose Company --</option>
                          {companies.map((c) => (
                            <option key={c.id} value={c.id}>
                              {c.name} (PAN: {c.pan})
                            </option>
                          ))}
                        </select>
                      </div>
                    )}

                    {/* File Dropzone Area */}
                    <div className="flex flex-col gap-2">
                      <label className="text-xs font-semibold text-slate-500 dark:text-slate-400">Select File</label>
                      <div className="border-2 border-dashed border-slate-200 dark:border-slate-800 hover:border-blue-500/50 bg-slate-50 dark:bg-slate-950/40 hover:bg-slate-100 dark:hover:bg-slate-950/80 rounded-xl p-6 transition flex flex-col items-center justify-center gap-3 text-center cursor-pointer relative">
                        <input
                          type="file"
                          ref={fileInputRef}
                          accept=".xlsx,.xls"
                          onChange={handleExcelUpload}
                          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                        />
                        <FileSpreadsheet className="h-10 w-10 text-emerald-500" />
                        <div className="flex flex-col gap-1">
                          <span className="text-xs font-semibold text-slate-800 dark:text-white">
                            {excelFile ? excelFile.name : "Drag Excel file here or click to browse"}
                          </span>
                          <span className="text-[10px] text-slate-500">
                            {excelFile ? `${(excelFile.size / 1024).toFixed(1)} KB` : "Supports .xlsx and .xls formats"}
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* Sheet Selector (if multiple sheets found) */}
                    {excelSheets.length > 1 && (
                      <div className="flex flex-col gap-2">
                        <div className="flex items-center justify-between">
                          <label htmlFor="excelSheetSelect" className="text-xs font-semibold text-slate-500 dark:text-slate-400">
                            Select Excel Sheet *
                          </label>
                          <span className="text-[10px] text-blue-500 font-medium bg-blue-950/20 border border-blue-900/30 px-2 py-0.5 rounded-full">
                            Auto-detected sheet containing barcodes
                          </span>
                        </div>
                        <select
                          id="excelSheetSelect"
                          value={selectedExcelSheet}
                          onChange={(e) => handleSheetChange(e.target.value)}
                          className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-3 text-xs text-slate-800 dark:text-slate-200 w-full focus:outline-none focus:border-blue-500"
                        >
                          {excelSheets.map((name) => (
                            <option key={name} value={name}>
                              {name}
                            </option>
                          ))}
                        </select>
                      </div>
                    )}

                    {/* Preview Table */}
                    {parsedRows.length > 0 && (
                      <div className="flex flex-col gap-2.5">
                        <div className="flex items-center justify-between text-xs">
                          <span className="font-semibold text-slate-500 dark:text-slate-400">
                            Preview ({parsedRows.length} rows found)
                          </span>
                          <span className="text-[10px] text-slate-500">
                            Valid: {parsedRows.filter(r => r.status === "valid").length} | 
                            Errors: {parsedRows.filter(r => r.status !== "valid").length}
                          </span>
                        </div>

                        <div className="max-h-[220px] overflow-y-auto border border-slate-200 dark:border-slate-850 rounded-xl bg-slate-50 dark:bg-slate-950/60">
                          <table className="w-full text-left text-[11px] border-collapse">
                            <thead>
                              <tr className="border-b border-slate-200 dark:border-slate-850 text-slate-500 font-bold bg-slate-100 dark:bg-slate-950 py-2 px-3">
                                <th className="py-2 px-2.5 text-center w-8">Row</th>
                                <th className="py-2 px-2">Barcode</th>
                                <th className="py-2 px-2">Company Name (PAN)</th>
                                <th className="py-2 px-2">SKU / Product</th>
                                <th className="py-2 px-2 text-right">Status</th>
                              </tr>
                            </thead>
                            <tbody>
                              {parsedRows.map((row, idx) => (
                                <tr key={idx} className="border-b border-slate-200 dark:border-slate-850/50 hover:bg-slate-100 dark:hover:bg-slate-900/40 text-slate-700 dark:text-slate-300">
                                  <td className="py-2 px-2.5 text-center text-slate-500 font-mono">{row.rowNumber}</td>
                                  <td className="py-2 px-2 font-mono font-semibold text-slate-900 dark:text-white">{row.code || "---"}</td>
                                  <td className="py-2 px-2">
                                    {excelCompanyOption === "assign" ? (
                                      <span className="text-slate-400 italic">Assigned Company</span>
                                    ) : (
                                      <>
                                        <div>{row.companyName || "---"}</div>
                                        <div className="text-[9px] text-slate-500 font-mono">{row.pan || "---"}</div>
                                      </>
                                    )}
                                  </td>
                                  <td className="py-2 px-2">
                                    {row.productSKU ? (
                                      <span className="bg-slate-200 dark:bg-slate-800 text-slate-700 dark:text-slate-200 px-1 py-0.5 rounded text-[9px] font-mono mr-1">
                                        {row.productSKU}
                                      </span>
                                    ) : null}
                                    <span>{row.productName}</span>
                                  </td>
                                  <td className="py-2 px-2 text-right">
                                    {row.status === "valid" ? (
                                      <span className="inline-flex items-center px-1.5 py-0.5 rounded-full text-[9px] font-medium bg-emerald-50 dark:bg-emerald-950/80 text-emerald-700 dark:text-emerald-400 border border-emerald-250 dark:border-emerald-900">
                                        Valid
                                      </span>
                                    ) : row.status === "duplicate" ? (
                                      <span className="inline-flex items-center px-1.5 py-0.5 rounded-full text-[9px] font-medium bg-amber-50 dark:bg-amber-950/80 text-amber-700 dark:text-amber-400 border border-amber-250 dark:border-amber-900" title={row.reason}>
                                        Duplicate
                                      </span>
                                    ) : (
                                      <span className="inline-flex items-center px-1.5 py-0.5 rounded-full text-[9px] font-medium bg-red-50 dark:bg-red-950/80 text-red-700 dark:text-red-400 border border-red-250 dark:border-red-900" title={row.reason}>
                                        Error
                                      </span>
                                    )}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    {/* Import Button */}
                    <button
                      type="button"
                      onClick={handleImportSubmit}
                      disabled={importLoading || parsedRows.filter(r => r.status === "valid").length === 0}
                      className="w-full bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 disabled:opacity-50 text-white font-semibold py-3 px-6 rounded-xl transition shadow-lg shadow-emerald-950/20 flex items-center justify-center gap-2 text-sm"
                    >
                      {importLoading ? (
                        <>
                          <RefreshCw className="h-4 w-4 animate-spin" />
                          Importing Excel Rows...
                        </>
                      ) : (
                        <>
                          <Database className="h-4 w-4" />
                          Import {parsedRows.filter(r => r.status === "valid").length} Valid Row(s)
                        </>
                      )}
                    </button>
                  </div>
                )}
              </div>
            )}
          </div>
        </section>

        {/* Resizer Handle */}
        <div
          onMouseDown={startResizing}
          className={`hidden lg:flex items-center justify-center cursor-col-resize w-2 hover:w-3 bg-slate-200/50 hover:bg-blue-500 dark:bg-slate-800/50 dark:hover:bg-blue-600 rounded-full transition-all shrink-0 self-stretch my-2 relative group ${
            isDragging ? "bg-blue-600 dark:bg-blue-600 w-3" : ""
          }`}
          title="Drag to resize panels"
        >
          <div className="w-1 h-8 bg-slate-400 dark:bg-slate-600 rounded-full group-hover:bg-white transition-colors"></div>
        </div>

        {/* Right Side: Output Panel */}
        <section className="flex flex-col gap-8 flex-1">
          <div className="bg-white/80 dark:bg-slate-900/60 backdrop-blur-md border border-slate-200 dark:border-slate-850 rounded-2xl p-6 shadow-xl shadow-slate-100 dark:shadow-black/20 flex flex-col gap-6">
            <div className="flex items-center justify-between border-b border-slate-200 dark:border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <BarcodeIcon className="h-5 w-5 text-indigo-400" />
                <h2 className="font-display font-semibold text-lg text-slate-900 dark:text-white">
                  Output Panel
                </h2>
              </div>
              {generatedBarcodes.length > 1 && (
                <span className="text-xs bg-indigo-900/50 border border-indigo-800 text-indigo-200 px-2 py-0.5 rounded-full font-medium">
                  {selectedBarcodeIndex + 1} of {generatedBarcodes.length} labels
                </span>
              )}
            </div>

            {generatedBarcodes.length === 0 ? (
              <div className="flex flex-col items-center justify-center py-20 px-4 border border-dashed border-slate-800 rounded-xl text-slate-500 gap-3">
                <BarcodeIcon className="h-12 w-12 stroke-[1.2]" />
                <p className="text-xs text-center max-w-[240px]">
                  Fill in the generation parameters and click "Generate Barcode" to preview standard EAN-13 barcodes here.
                </p>
              </div>
            ) : (
              <div className="flex flex-col gap-6">
                {/* Barcode Preview Card */}
                <div className="relative group bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl p-6 flex flex-col items-center justify-center shadow-inner">
                  {/* Scannable SVG barcode */}
                  <svg ref={barcodeSvgRef} id="barcode" className="max-w-full text-slate-100"></svg>

                  {/* Overlay metadata */}
                  <div className="w-full flex items-center justify-between mt-4 text-[10px] text-slate-500 font-mono border-t border-slate-200 dark:border-slate-900 pt-3">
                    <span>Pan: {generatedBarcodes[selectedBarcodeIndex].company.pan}</span>
                    <span>Seq: {String(generatedBarcodes[selectedBarcodeIndex].sequence).padStart(2, "0")}</span>
                  </div>
                </div>

                {/* Batch Navigation */}
                {generatedBarcodes.length > 1 && (
                  <div className="flex items-center justify-between bg-slate-50 dark:bg-slate-950/40 p-2 border border-slate-200 dark:border-slate-800 rounded-xl">
                    <button
                      onClick={() => setSelectedBarcodeIndex(prev => Math.max(0, prev - 1))}
                      disabled={selectedBarcodeIndex === 0}
                      className="p-1.5 rounded-lg hover:bg-slate-800 text-slate-400 disabled:opacity-30 disabled:hover:bg-transparent"
                    >
                      <ChevronLeft className="h-5 w-5" />
                    </button>
                    <span className="text-xs text-slate-700 dark:text-slate-700 dark:text-slate-300 font-mono">
                      Previewing Code: <span className="font-bold text-white">{generatedBarcodes[selectedBarcodeIndex].code}</span>
                    </span>
                    <button
                      onClick={() => setSelectedBarcodeIndex(prev => Math.min(generatedBarcodes.length - 1, prev + 1))}
                      disabled={selectedBarcodeIndex === generatedBarcodes.length - 1}
                      className="p-1.5 rounded-lg hover:bg-slate-800 text-slate-400 disabled:opacity-30 disabled:hover:bg-transparent"
                    >
                      <ChevronRight className="h-5 w-5" />
                    </button>
                  </div>
                )}

                {/* Actions Toolbar */}
                <div className="grid grid-cols-2 gap-3">
                  <button
                    onClick={handleCopyCode}
                    className="bg-white dark:bg-slate-950 hover:bg-slate-50 dark:hover:bg-slate-800 border border-slate-200 dark:border-slate-800 text-slate-800 dark:text-slate-200 rounded-xl py-2.5 text-xs font-semibold flex items-center justify-center gap-1.5 transition active:scale-[0.98]"
                  >
                    {copiedCode ? (
                      <>
                        <Check className="h-4 w-4 text-emerald-400" />
                        Copied Text
                      </>
                    ) : (
                      <>
                        <Copy className="h-4 w-4" />
                        Copy Number
                      </>
                    )}
                  </button>

                  <button
                    onClick={handleCopyImage}
                    className="bg-white dark:bg-slate-950 hover:bg-slate-50 dark:hover:bg-slate-800 border border-slate-200 dark:border-slate-800 text-slate-800 dark:text-slate-200 rounded-xl py-2.5 text-xs font-semibold flex items-center justify-center gap-1.5 transition active:scale-[0.98]"
                  >
                    {copiedImage ? (
                      <>
                        <Check className="h-4 w-4 text-emerald-400" />
                        Copied Image
                      </>
                    ) : (
                      <>
                        <Copy className="h-4 w-4" />
                        Copy Image
                      </>
                    )}
                  </button>
                </div>

                <div className="border-t border-slate-200 dark:border-slate-800/60 my-1"></div>

                {/* Print Settings Layout */}
                <div className="flex flex-col gap-4">
                  <div className="grid grid-cols-2 gap-4">
                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] uppercase font-bold tracking-wider text-slate-400">Print Layout</label>
                      <select
                        value={printLayout}
                        onChange={(e: any) => setPrintLayout(e.target.value)}
                        className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-800 dark:text-slate-200"
                      >
                        <option value="single">Single Label</option>
                        <option value="sheet">Sheet / Grid (3x8)</option>
                        <option value="custom">Custom Grid</option>
                      </select>
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] uppercase font-bold tracking-wider text-slate-400">Download Format</label>
                      <select
                        value={downloadFormat}
                        onChange={(e: any) => setDownloadFormat(e.target.value)}
                        className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-800 dark:text-slate-200"
                      >
                        <option value="png">PNG (Raster Image)</option>
                        <option value="jpg">JPEG (Raster Image)</option>
                        <option value="svg">SVG (Vector XML)</option>
                        <option value="pdf">PDF (Document Grid)</option>
                        <option value="excel">Excel (.csv / .xlsx)</option>
                      </select>
                    </div>
                  </div>

                  {/* Custom Row/Col Config */}
                  {printLayout === "custom" && (
                    <div className="grid grid-cols-2 gap-4 bg-slate-50 dark:bg-slate-950/40 p-3 border border-slate-200 dark:border-slate-800 rounded-xl">
                      <div className="flex flex-col gap-1">
                        <label className="text-[10px] text-slate-400">Columns</label>
                        <input
                          type="number"
                          min="1"
                          max="10"
                          value={customCols}
                          onChange={(e) => setCustomCols(Math.max(1, parseInt(e.target.value) || 1))}
                          className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1 text-center text-xs"
                        />
                      </div>
                      <div className="flex flex-col gap-1">
                        <label className="text-[10px] text-slate-400">Rows</label>
                        <input
                          type="number"
                          min="1"
                          max="20"
                          value={customRows}
                          onChange={(e) => setCustomRows(Math.max(1, parseInt(e.target.value) || 1))}
                          className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1 text-center text-xs"
                        />
                      </div>
                    </div>
                  )}

                  <div className="grid grid-cols-2 gap-3 mt-1">
                    <button
                      onClick={handlePrint}
                      className="bg-white dark:bg-slate-950 hover:bg-slate-50 dark:hover:bg-slate-800 border border-slate-200 dark:border-slate-800 text-slate-800 dark:text-slate-200 rounded-xl py-3 text-xs font-semibold flex items-center justify-center gap-1.5 transition active:scale-[0.98]"
                    >
                      <Printer className="h-4.5 w-4.5 text-slate-400" />
                      Print Barcode
                    </button>

                    <button
                      onClick={handleDownload}
                      className="bg-blue-600 hover:bg-blue-500 text-white rounded-xl py-3 text-xs font-semibold flex items-center justify-center gap-1.5 transition active:scale-[0.98] shadow-lg shadow-blue-900/10"
                    >
                      <Download className="h-4.5 w-4.5" />
                      Download file
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        </section>
      </div>

      {/* Database History Section & Admin Config Row */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
        {/* Audit Log / History list */}
        <section className="lg:col-span-8 flex flex-col gap-4">
          <div className="bg-white/80 dark:bg-slate-900/60 backdrop-blur-md border border-slate-200 dark:border-slate-850 rounded-2xl p-6 shadow-xl shadow-slate-100 dark:shadow-black/20 flex flex-col gap-6">
            <div className="flex flex-col gap-4 border-b border-slate-200 dark:border-slate-800 pb-4">
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div className="flex items-center gap-2">
                  <Database className="h-5 w-5 text-indigo-400" />
                  <h2 className="font-display font-semibold text-lg text-slate-900 dark:text-white">
                    Audit History Log
                  </h2>
                </div>

                {/* Filters */}
                <div className="flex flex-wrap items-center gap-3">
                  {/* Search Bar */}
                  <div className="relative">
                    <Search className="absolute left-3 top-2.5 h-4 w-4 text-slate-500" />
                    <input
                      type="text"
                      value={historySearch}
                      onChange={(e) => setHistorySearch(e.target.value)}
                      placeholder="Search by Code/SKU..."
                      className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 pl-9 pr-4 py-2 rounded-xl text-xs text-slate-800 dark:text-slate-200 w-44 md:w-52 focus:outline-none focus:border-blue-500"
                    />
                  </div>

                  {/* Company filter */}
                  <select
                    value={historyCompanyFilter}
                    onChange={(e) => {
                      setHistoryCompanyFilter(e.target.value);
                      setHistoryPage(1);
                    }}
                    className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-800 dark:text-slate-200 focus:outline-none"
                  >
                    <option value="all">All Companies</option>
                    {companies.map((c) => (
                      <option key={c.id} value={c.id}>
                        {c.name} ({c.pan.substring(5, 9)})
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Date Range & Bulk Selection Row */}
              <div className="flex flex-wrap items-center justify-between gap-3 pt-2">
                {/* Date Filter Inputs */}
                <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500 dark:text-slate-400">
                  <Calendar className="h-4 w-4 text-blue-400 shrink-0" />
                  <span className="text-[11px] font-medium">Date Range:</span>
                  <input
                    type="date"
                    value={historyStartDate}
                    onChange={(e) => {
                      setHistoryStartDate(e.target.value);
                      setHistoryPage(1);
                    }}
                    onClick={(e) => e.currentTarget.showPicker()}
                    className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-800 dark:text-slate-200 focus:outline-none focus:border-blue-500 cursor-pointer"
                  />
                  <span>to</span>
                  <input
                    type="date"
                    value={historyEndDate}
                    onChange={(e) => {
                      setHistoryEndDate(e.target.value);
                      setHistoryPage(1);
                    }}
                    onClick={(e) => e.currentTarget.showPicker()}
                    className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-800 dark:text-slate-200 focus:outline-none focus:border-blue-500 cursor-pointer"
                  />
                  {(historyStartDate || historyEndDate) && (
                    <button
                      onClick={() => {
                        setHistoryStartDate("");
                        setHistoryEndDate("");
                        setHistoryPage(1);
                      }}
                      className="p-1 text-slate-500 hover:text-slate-800 dark:hover:text-slate-200 transition"
                      title="Clear Date Filters"
                    >
                      <XCircle className="h-4 w-4" />
                    </button>
                  )}
                </div>

                {/* Bulk Select Actions */}
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={handleSelectAllFiltered}
                    disabled={bulkExportLoading}
                    className="bg-slate-100 dark:bg-slate-950 hover:bg-slate-200 dark:hover:bg-slate-800 border border-slate-300 dark:border-slate-800 text-slate-800 dark:text-slate-300 text-[11px] px-2.5 py-1.5 rounded-lg transition flex items-center gap-1 font-medium"
                  >
                    <Filter className="h-3 w-3 text-blue-400" />
                    Select All Filtered ({historyTotal})
                  </button>

                  {selectedBarcodeIds.length > 0 && (
                    <button
                      onClick={handleDeselectAll}
                      className="text-slate-500 hover:text-slate-700 dark:text-slate-300 text-[11px] px-2 py-1 transition"
                    >
                      Deselect All
                    </button>
                  )}
                </div>
              </div>
            </div>

            {/* Selected Items Bulk Export Bar */}
            {selectedBarcodeIds.length > 0 && (
              <div className="flex flex-wrap items-center justify-between gap-3 bg-blue-950/40 border border-blue-900/60 p-3 rounded-xl">
                <div className="flex items-center gap-2">
                  <span className="bg-blue-600 text-white font-bold text-xs px-2.5 py-0.5 rounded-full">
                    {selectedBarcodeIds.length}
                  </span>
                  <span className="text-xs text-blue-200 font-semibold">
                    Barcode(s) Selected for Bulk Download
                  </span>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={handleBulkExportExcel}
                    disabled={bulkExportLoading}
                    className="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold px-3 py-1.5 rounded-lg shadow transition flex items-center gap-1.5 active:scale-95 disabled:opacity-50"
                  >
                    {bulkExportLoading ? (
                      <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                    ) : (
                      <FileSpreadsheet className="h-3.5 w-3.5" />
                    )}
                    Export to Excel (.xlsx)
                  </button>

                  <button
                    onClick={handleBulkExportPDF}
                    disabled={bulkExportLoading}
                    className="bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold px-3 py-1.5 rounded-lg shadow transition flex items-center gap-1.5 active:scale-95 disabled:opacity-50"
                  >
                    {bulkExportLoading ? (
                      <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                    ) : (
                      <Download className="h-3.5 w-3.5" />
                    )}
                    Export to PDF (.pdf)
                  </button>

                  <button
                    onClick={handleBulkExportCertificate}
                    disabled={bulkExportLoading}
                    className="bg-amber-600 hover:bg-amber-500 text-white text-xs font-semibold px-3 py-1.5 rounded-lg shadow transition flex items-center gap-1.5 active:scale-95 disabled:opacity-50"
                  >
                    {bulkExportLoading ? (
                      <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                    ) : (
                      <Award className="h-3.5 w-3.5" />
                    )}
                    Generate Certificate
                  </button>
                </div>
              </div>
            )}

            {/* History Table */}
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="border-b border-slate-200 dark:border-slate-800 text-slate-500 dark:text-slate-400 font-bold uppercase tracking-wider">
                    <th className="py-3 px-3 w-8">
                      <input
                        type="checkbox"
                        checked={isPageAllSelected}
                        onChange={handleToggleSelectPage}
                        className="rounded border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-950 text-blue-600 focus:ring-blue-500 cursor-pointer h-4 w-4"
                        title="Select/Deselect All on Current Page"
                      />
                    </th>
                    <th className="py-3 px-3">EAN-13 Barcode</th>
                    <th className="py-3 px-3">Company (PAN)</th>
                    <th className="py-3 px-3">SKU & Product</th>
                    <th className="py-3 px-3">Seq / Rot</th>
                    <th className="py-3 px-3">Generated At</th>
                    <th className="py-3 px-3 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {historyLoading ? (
                    <tr>
                      <td colSpan={7} className="text-center py-10 text-slate-500 font-mono">
                        <RefreshCw className="h-5 w-5 animate-spin mx-auto mb-2 text-slate-400" />
                        Loading database records...
                      </td>
                    </tr>
                  ) : historyBarcodes.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="text-center py-10 text-slate-500 font-mono">
                        No matching database entries found.
                      </td>
                    </tr>
                  ) : (
                    historyBarcodes.map((item) => {
                      const isSelected = selectedBarcodeIds.includes(item.id);
                      return (
                        <tr
                          key={item.id}
                          className={`border-b border-slate-800/50 transition ${
                            isSelected ? "bg-blue-950/20 text-white" : "hover:bg-slate-850/20 text-slate-700 dark:text-slate-300"
                          }`}
                        >
                          <td className="py-3 px-3">
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={() => handleToggleSelectBarcode(item.id)}
                              className="rounded border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-950 text-blue-600 focus:ring-blue-500 cursor-pointer h-4 w-4"
                            />
                          </td>
                          <td className="py-3 px-3 font-mono font-bold text-slate-800 dark:text-white select-all">
                            {item.code}
                          </td>
                          <td className="py-3 px-3">
                            <div className="font-semibold text-slate-800 dark:text-slate-200">{item.company.name}</div>
                            <div className="text-[10px] text-slate-600 dark:text-slate-400">{item.company.pan}</div>
                          </td>
                          <td className="py-3 px-3">
                            {item.productSKU ? (
                              <>
                                <span className="bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-200 border border-slate-200 dark:border-transparent px-1.5 py-0.5 rounded font-mono text-[10px] font-bold mr-1">
                                  {item.productSKU}
                                </span>
                                <span className="text-slate-700 dark:text-slate-300">{item.productName || "No Name"}</span>
                              </>
                            ) : (
                              <span className="text-slate-500 italic">No details</span>
                            )}
                          </td>
                          <td className="py-3 px-3 font-mono">
                            {String(item.sequence).padStart(2, "0")} / {item.rotatingDigit}
                          </td>
                          <td className="py-3 px-3 text-slate-600 dark:text-slate-400">
                            {new Date(item.createdAt).toLocaleString()}
                          </td>
                          <td className="py-3 px-3 text-right">
                            <div className="flex items-center justify-end gap-1.5">
                              <button
                                onClick={() => handleSelectFromHistory(item)}
                                className="bg-slate-100 dark:bg-slate-950 border border-slate-300 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-800 hover:text-slate-900 dark:hover:text-white px-2.5 py-1.5 rounded-lg text-[10px] font-semibold transition active:scale-95 inline-flex items-center gap-1"
                                title="Preview & load barcode into generator"
                              >
                                <Eye className="h-3 w-3" /> Load
                              </button>

                              <button
                                onClick={() => exportToExcel([item], `barcode_${item.code}.xlsx`)}
                                className="bg-slate-100 dark:bg-slate-950 border border-slate-300 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:bg-emerald-100 dark:hover:bg-emerald-950 hover:text-emerald-700 dark:hover:text-emerald-300 hover:border-emerald-350 dark:hover:border-emerald-800 px-2 py-1.5 rounded-lg text-[10px] font-semibold transition active:scale-95 inline-flex items-center gap-1"
                                title="Download single barcode as Excel (.xlsx)"
                              >
                                <FileSpreadsheet className="h-3 w-3 text-emerald-400" />
                              </button>

                              <button
                                onClick={() =>
                                  exportToPDF([item], `barcode_${item.code}.pdf`, {
                                    cols: 1,
                                    rows: 1,
                                  })
                                }
                                className="bg-slate-100 dark:bg-slate-950 border border-slate-300 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:bg-blue-100 dark:hover:bg-blue-950 hover:text-blue-700 dark:hover:text-blue-300 hover:border-blue-350 dark:hover:border-blue-800 px-2 py-1.5 rounded-lg text-[10px] font-semibold transition active:scale-95 inline-flex items-center gap-1"
                                title="Download single barcode as PDF"
                              >
                                <Download className="h-3 w-3 text-blue-400" />
                              </button>

                              <button
                                onClick={() => exportToCertificate([item], `barcode_${item.code}_certificate.pdf`)}
                                className="bg-slate-100 dark:bg-slate-950 border border-slate-300 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:bg-amber-100 dark:hover:bg-amber-950 hover:text-amber-700 dark:hover:text-amber-300 hover:border-amber-350 dark:hover:border-amber-800 px-2 py-1.5 rounded-lg text-[10px] font-semibold transition active:scale-95 inline-flex items-center gap-1"
                                title="Generate Allocation Certificate for this barcode"
                              >
                                <Award className="h-3 w-3 text-amber-400" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination */}
            {historyTotal > historyLimit && (
              <div className="flex items-center justify-between border-t border-slate-200 dark:border-slate-800 pt-4 text-xs">
                <span className="text-slate-600 dark:text-slate-400">
                  Showing <span className="text-slate-800 dark:text-white font-semibold">{((historyPage - 1) * historyLimit) + 1}</span> to{" "}
                  <span className="text-slate-800 dark:text-white font-semibold">
                    {Math.min(historyPage * historyLimit, historyTotal)}
                  </span>{" "}
                  of <span className="text-slate-800 dark:text-white font-semibold">{historyTotal}</span> items
                </span>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setHistoryPage((p) => Math.max(1, p - 1))}
                    disabled={historyPage === 1}
                    className="p-1.5 bg-white dark:bg-slate-950 border border-slate-250 dark:border-slate-800 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-800 disabled:opacity-30 transition"
                  >
                    <ChevronLeft className="h-4 w-4" />
                  </button>
                  <span className="text-slate-700 dark:text-slate-300 font-semibold px-2">
                    Page {historyPage} of {Math.ceil(historyTotal / historyLimit)}
                  </span>
                  <button
                    onClick={() => setHistoryPage((p) => p + 1)}
                    disabled={historyPage >= Math.ceil(historyTotal / historyLimit)}
                    className="p-1.5 bg-white dark:bg-slate-950 border border-slate-250 dark:border-slate-800 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-800 disabled:opacity-30 transition"
                  >
                    <ChevronRight className="h-4 w-4" />
                  </button>
                </div>
              </div>
            )}
          </div>
        </section>

        {/* Right Side: Admin Panel */}
        <section className="lg:col-span-4 flex flex-col gap-4">
          <div className="bg-white/80 dark:bg-slate-900/60 backdrop-blur-md border border-slate-200 dark:border-slate-850 rounded-2xl p-6 shadow-xl shadow-slate-100 dark:shadow-black/20 flex flex-col gap-5 h-full">
            <div className="flex items-center gap-2 border-b border-slate-200 dark:border-slate-800 pb-3">
              <Settings className="h-5 w-5 text-blue-400" />
              <h2 className="font-display font-semibold text-lg text-slate-900 dark:text-white">
                Admin Console
              </h2>
            </div>

            <div className="flex-1 flex flex-col gap-4">
              <p className="text-xs text-slate-500 dark:text-slate-400 leading-relaxed">
                Define the sequence overflow wrap-around behavior per company. Toggling "Reset/Wrap-around" allows restarting the sequence counter from 01 when the 990 limit is exhausted.
              </p>

              {/* Company Picker */}
              <div className="flex flex-col gap-2">
                <label className="text-xs font-semibold text-slate-500 dark:text-slate-400">Select Company to Configure</label>
                <select
                  value={adminSelectedCompany}
                  onChange={(e) => setAdminSelectedCompany(e.target.value)}
                  className="bg-white dark:bg-slate-950 border border-slate-300 dark:border-slate-800 rounded-xl px-3 py-3 text-xs text-slate-800 dark:text-slate-200 w-full focus:outline-none"
                >
                  <option value="">-- Choose Company --</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.name} (PAN: {c.pan})
                    </option>
                  ))}
                </select>
              </div>

              {/* Toggle switch */}
              {adminSelectedCompany && (
                <div className="flex items-center justify-between bg-slate-50 dark:bg-slate-950 p-4 border border-slate-200 dark:border-slate-800 rounded-xl mt-2">
                  <div className="flex flex-col">
                    <span className="text-xs font-semibold text-slate-900 dark:text-white">Reset on Overflow</span>
                    <span className="text-[10px] text-slate-500 dark:text-slate-400 mt-0.5">
                      {adminOverflowReset ? "Wrap-around enabled (01-99)" : "Show UI Error on limit"}
                    </span>
                  </div>
                  <button
                    onClick={() => setAdminOverflowReset((prev) => !prev)}
                    className={`relative inline-flex h-6 w-11 shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none ${
                        adminOverflowReset ? "bg-blue-600" : "bg-slate-200 dark:bg-slate-700"
                      }`}
                  >
                    <span
                      className={`pointer-events-none inline-block h-5 w-5 transform rounded-full bg-white shadow ring-0 transition duration-200 ease-in-out ${adminOverflowReset ? "translate-x-5" : "translate-x-0"
                        }`}
                    />
                  </button>
                </div>
              )}
            </div>

            {/* Save Config Button */}
            <button
              onClick={handleSaveAdminConfig}
              disabled={adminLoading || !adminSelectedCompany}
              className="w-full bg-white dark:bg-slate-950 hover:bg-slate-50 dark:hover:bg-slate-800 disabled:opacity-50 text-slate-800 dark:text-white font-semibold py-2.5 px-4 rounded-xl border border-slate-250 dark:border-slate-800 transition active:scale-[0.98] text-xs flex items-center justify-center gap-2 mt-auto"
            >
              {adminLoading ? (
                <>
                  <RefreshCw className="h-3 w-3 animate-spin" />
                  Saving Configuration...
                </>
              ) : (
                "Save Configuration"
              )}
            </button>
          </div>
        </section>
      </div>

      {/* Hidden print block for scoped browser window.print() */}
      <div id="print-area" className="hidden print:block print-area">
        {generatedBarcodes.length > 0 && (
          <div
            className={`print-grid`}
            style={{
              display: "grid",
              gridTemplateColumns:
                printLayout === "single"
                  ? "1fr"
                  : `repeat(${printLayout === "sheet" ? 3 : customCols}, minmax(0, 1fr))`,
              gap: "15px",
            }}
          >
            {/* If Single Layout: print one large barcode */}
            {printLayout === "single" ? (
              <div className="print-card" style={{ width: "100%", textAlign: "center" }}>
                <h2 style={{ fontSize: "16px", fontWeight: "bold", margin: "10px 0" }}>
                  GS1 EAN-13 BARCODE LABEL
                </h2>
                <div style={{ fontSize: "12px", margin: "5px 0" }}>
                  Company: {generatedBarcodes[selectedBarcodeIndex].company.name}
                </div>
                {/* Dynamically render barcode for printing */}
                <svg
                  ref={(el) => {
                    if (el) {
                      try {
                        JsBarcode(el, generatedBarcodes[selectedBarcodeIndex].code.substring(0, 12), {
                          format: "EAN13",
                          displayValue: true,
                          fontSize: 16,
                          width: 2.2,
                          height: 90,
                          lineColor: "#000000",
                        });
                      } catch (err) {
                        console.error(err);
                      }
                    }
                  }}
                ></svg>
                {generatedBarcodes[selectedBarcodeIndex].productSKU && (
                  <div style={{ fontSize: "10px", marginTop: "10px" }}>
                    SKU: {generatedBarcodes[selectedBarcodeIndex].productSKU} -{" "}
                    {generatedBarcodes[selectedBarcodeIndex].productName}
                  </div>
                )}
              </div>
            ) : (
              /* Grid Layout (Sheet or Custom Grid) */
              (generatedBarcodes.length === 1
                ? Array((printLayout === "sheet" ? 24 : customCols * customRows)).fill(
                  generatedBarcodes[0]
                )
                : generatedBarcodes
              ).map((barcodeItem, idx) => (
                <div key={idx} className="print-card" style={{ border: "1px solid #000" }}>
                  <svg
                    ref={(el) => {
                      if (el) {
                        try {
                          JsBarcode(el, barcodeItem.code.substring(0, 12), {
                            format: "EAN13",
                            displayValue: true,
                            fontSize: 12,
                            width: 1.3,
                            height: 50,
                            margin: 2,
                            lineColor: "#000000",
                          });
                        } catch (err) {
                          console.error(err);
                        }
                      }
                    }}
                  ></svg>
                  <div style={{ fontSize: "7px", marginTop: "4px", textAlign: "center" }}>
                    {barcodeItem.productSKU || "GTIN"}: {barcodeItem.company.name.substring(0, 15)}
                  </div>
                </div>
              ))
            )}
          </div>
        )}
      </div>
    </main>
  );
}

const Home = dynamic(() => Promise.resolve(BarcodeGenerator), {
  ssr: false,
  loading: () => (
    <div className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-8 flex flex-col gap-8">
      {/* Header Banner Skeleton */}
      <header className="flex flex-col md:flex-row items-center justify-between gap-4 pb-6 border-b border-slate-800">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-slate-900 border border-slate-850 rounded-xl flex items-center justify-center">
            <div className="h-8 w-8 bg-slate-800 rounded animate-pulse" />
          </div>
          <div className="flex flex-col gap-2">
            <div className="h-6 w-48 bg-slate-900 rounded animate-pulse" />
            <div className="h-4 w-72 bg-slate-900 rounded animate-pulse" />
          </div>
        </div>
      </header>
      <div className="flex-1 flex items-center justify-center min-h-[400px]">
        <div className="flex flex-col items-center gap-3">
          <div className="h-8 w-8 rounded-full border-2 border-slate-800 border-t-blue-500 animate-spin" />
          <p className="text-xs text-slate-500 font-mono">Initializing Console...</p>
        </div>
      </div>
    </div>
  )
});

export default Home;

