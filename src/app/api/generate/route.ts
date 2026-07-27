import { NextRequest, NextResponse } from "next/server";
import { db } from "@/lib/db";
import {
  generateSingleBarcode,
  CapacityExceededError,
  DuplicateCollisionError,
} from "@/lib/barcodeEngine";

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { companyName, pan, productSKU, productName, productDesc, batchCount = 1 } = body;

    // 1. Validation
    if (!companyName || typeof companyName !== "string" || companyName.trim().length === 0) {
      return NextResponse.json({ error: "Company name is required." }, { status: 400 });
    }

    if (!pan || typeof pan !== "string") {
      return NextResponse.json({ error: "PAN number is required." }, { status: 400 });
    }

    const panRegex = /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/i;
    if (!panRegex.test(pan.trim())) {
      return NextResponse.json(
        { error: "Invalid PAN number format. Must be like AAAAA9999A." },
        { status: 400 }
      );
    }

    const cleanPAN = pan.toUpperCase().trim();
    const count = parseInt(batchCount, 10);
    if (isNaN(count) || count < 1 || count > 50) {
      return NextResponse.json(
        { error: "Batch count must be an integer between 1 and 50." },
        { status: 400 }
      );
    }

    // 2. Find or Create Company
    // If company with this PAN exists, we use it. If name changed, we update the name.
    let company = await db.company.findUnique({
      where: { pan: cleanPAN },
    });

    if (!company) {
      company = await db.company.create({
        data: {
          name: companyName.trim(),
          pan: cleanPAN,
        },
      });
    } else if (company.name !== companyName.trim()) {
      company = await db.company.update({
        where: { id: company.id },
        data: { name: companyName.trim() },
      });
    }

    // 3. Generate barcode(s) in a single transaction
    try {
      const barcodes = await db.$transaction(async (tx) => {
        const generated = [];
        for (let i = 0; i < count; i++) {
          const barcode = await generateSingleBarcode(tx, company!.id, {
            sku: productSKU,
            name: productName,
            desc: productDesc,
          });
          generated.push(barcode);
        }
        return generated;
      });

      return NextResponse.json({ barcodes });
    } catch (err: any) {
      if (err instanceof CapacityExceededError) {
        return NextResponse.json(
          { error: "capacity_exceeded", message: err.message },
          { status: 400 }
        );
      }
      if (err instanceof DuplicateCollisionError) {
        return NextResponse.json(
          { error: "duplicate_collision", message: err.message },
          { status: 409 }
        );
      }
      throw err;
    }
  } catch (error: any) {
    console.error("Barcode generation API error:", error);
    return NextResponse.json(
      { error: "server_error", message: error.message || "Internal server error" },
      { status: 500 }
    );
  }
}
