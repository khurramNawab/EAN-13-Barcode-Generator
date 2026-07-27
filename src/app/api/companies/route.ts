import { NextResponse } from "next/server";
import { db } from "@/lib/db";

export async function GET() {
  try {
    const companies = await db.company.findMany({
      include: {
        state: true,
        _count: {
          select: { barcodes: true },
        },
      },
      orderBy: { name: "asc" },
    });

    return NextResponse.json({ companies });
  } catch (error: any) {
    console.error("API Error listing companies:", error);
    return NextResponse.json(
      { error: "server_error", message: "Failed to fetch companies." },
      { status: 500 }
    );
  }
}
