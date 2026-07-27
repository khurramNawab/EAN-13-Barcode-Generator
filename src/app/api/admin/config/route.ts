import { NextRequest, NextResponse } from "next/server";
import { db } from "@/lib/db";
import { shuffle } from "@/lib/barcodeEngine";

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { companyId, overflowReset } = body;

    if (!companyId || typeof companyId !== "string") {
      return NextResponse.json({ error: "Company ID is required." }, { status: 400 });
    }

    if (typeof overflowReset !== "boolean") {
      return NextResponse.json({ error: "overflowReset must be a boolean." }, { status: 400 });
    }

    // Check if company exists
    const company = await db.company.findUnique({
      where: { id: companyId },
    });

    if (!company) {
      return NextResponse.json({ error: "Company not found." }, { status: 404 });
    }

    // Upsert BarcodeState
    const state = await db.barcodeState.upsert({
      where: { companyId },
      update: { overflowReset },
      create: {
        companyId,
        overflowReset,
        currentSequence: 1,
        currentShuffle: JSON.stringify(shuffle([0, 1, 2, 3, 4, 5, 6, 7, 8, 9])),
      },
    });

    return NextResponse.json({ state });
  } catch (error: any) {
    console.error("API Error updating admin config:", error);
    return NextResponse.json(
      { error: "server_error", message: "Failed to update configuration." },
      { status: 500 }
    );
  }
}
