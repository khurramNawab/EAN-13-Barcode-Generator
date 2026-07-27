import { NextRequest, NextResponse } from "next/server";
import { db } from "@/lib/db";
import { extractPANLast4 } from "@/lib/barcodeEngine";

interface ImportRow {
  code: string;
  companyName: string;
  pan: string;
  productSKU?: string;
  productName?: string;
  productDesc?: string;
  companyId?: string; // Optional, if selected from dropdown
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { items } = body as { items: ImportRow[] };

    if (!items || !Array.isArray(items) || items.length === 0) {
      return NextResponse.json(
        { error: "invalid_input", message: "No data rows provided for import." },
        { status: 400 }
      );
    }

    // 1. Resolve Companies
    // Map company PAN to Company record
    const companyCache = new Map<string, any>();
    const panRegex = /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/i;

    // Prefetch all companies from DB to avoid N+1 queries during resolution
    const dbCompanies = await db.company.findMany({
      include: { state: true },
    });
    for (const comp of dbCompanies) {
      companyCache.set(comp.pan.toUpperCase(), comp);
    }

    // Process unique PANs from payload to avoid repeating company check queries
    const uniqueCompanyPayloads = new Map<string, { pan: string; name: string }>();
    for (const item of items) {
      if (!item.pan || typeof item.pan !== "string") continue;
      const cleanPAN = item.pan.toUpperCase().trim();
      if (!panRegex.test(cleanPAN)) continue;
      const cleanName = (item.companyName || "").trim();
      if (!cleanName) continue;

      uniqueCompanyPayloads.set(cleanPAN, { pan: cleanPAN, name: cleanName });
    }

    // Now resolve/create/update only for the unique companies
    for (const [cleanPAN, data] of uniqueCompanyPayloads.entries()) {
      let company = companyCache.get(cleanPAN);

      if (!company) {
        // Create new company
        company = await db.company.create({
          data: {
            name: data.name,
            pan: cleanPAN,
          },
        });
        // Cache it
        companyCache.set(cleanPAN, company);
      } else if (company.name !== data.name) {
        // Update company name if different
        company = await db.company.update({
          where: { id: company.id },
          data: { name: data.name },
        });
        companyCache.set(cleanPAN, company);
      }
    }

    // 2. Filter out duplicates from the payload itself and from DB
    const barcodesToProcess: ImportRow[] = [];
    const payloadCodes = new Set<string>();

    for (const item of items) {
      if (!item.code || typeof item.code !== "string") continue;
      const cleanCode = item.code.trim();

      // Validate EAN-13 length and format (13 digits)
      if (cleanCode.length !== 13 || !/^\d+$/.test(cleanCode)) {
        continue; // Skip invalid barcode values
      }

      // Check for duplicates within the uploaded file/payload itself
      if (payloadCodes.has(cleanCode)) {
        continue; // Skip duplicate inside the file
      }
      payloadCodes.add(cleanCode);

      // Verify if PAN is valid
      const cleanPAN = (item.pan || "").toUpperCase().trim();
      if (!panRegex.test(cleanPAN)) {
        continue; // Skip row with invalid PAN
      }

      // Verify if company exists in our resolved cache
      const company = companyCache.get(cleanPAN);
      if (!company) {
        continue; // Skip if we could not resolve company
      }

      barcodesToProcess.push({
        ...item,
        code: cleanCode,
        companyId: company.id,
      });
    }

    if (barcodesToProcess.length === 0) {
      return NextResponse.json({
        success: true,
        importedCount: 0,
        skippedCount: items.length,
        message: "No valid and unique barcodes were processed.",
      });
    }

    // Find which barcodes already exist in DB
    const uniqueCodes = barcodesToProcess.map((item) => item.code);
    const existingBarcodes = await db.barcode.findMany({
      where: { code: { in: uniqueCodes } },
      select: { code: true },
    });
    const dbCodesSet = new Set(existingBarcodes.map((b) => b.code));

    // Filter out existing DB codes
    const newBarcodesToInsert = barcodesToProcess.filter(
      (item) => !dbCodesSet.has(item.code)
    );

    const skippedCount = items.length - newBarcodesToInsert.length;

    // 3. Parse and Insert new barcodes
    const insertData = newBarcodesToInsert.map((item) => {
      const code = item.code;
      const company = companyCache.get(item.pan.toUpperCase().trim());

      // Parse standard barcode parts if it starts with 853 and matches expected sequence structure
      let sequence = 0;
      let datePart = "00";
      let rotatingDigit = 0;
      const checkDigit = parseInt(code.substring(12, 13), 10) || 0;

      if (code.startsWith("853")) {
        const seqPart = code.substring(3, 5);
        const dtPart = code.substring(5, 7);
        const panPart = code.substring(7, 11);
        const rotPart = code.substring(11, 12);

        // Optional safety: Verify if PAN part matches company's PAN
        const companyPanPart = extractPANLast4(company.pan);
        if (panPart === companyPanPart) {
          sequence = parseInt(seqPart, 10) || 0;
          datePart = dtPart;
          rotatingDigit = parseInt(rotPart, 10) || 0;
        }
      }

      return {
        code,
        sequence,
        rotatingDigit,
        datePart,
        checkDigit,
        productSKU: item.productSKU?.trim() || null,
        productName: item.productName?.trim() || null,
        productDesc: item.productDesc?.trim() || null,
        companyId: company.id,
      };
    });

    if (insertData.length > 0) {
      // Execute as a fast batch insert to avoid timeout
      await db.barcode.createMany({
        data: insertData,
      });
    }

    return NextResponse.json({
      success: true,
      importedCount: insertData.length,
      skippedCount,
      message: `Successfully imported ${insertData.length} barcode(s). ${skippedCount} barcode(s) were skipped (either invalid or already exist).`,
    });
  } catch (error: any) {
    console.error("Error importing barcodes:", error);
    return NextResponse.json(
      {
        error: "server_error",
        message: error.message || "An error occurred during barcode import.",
      },
      { status: 500 }
    );
  }
}
