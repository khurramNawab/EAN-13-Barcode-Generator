import { NextRequest, NextResponse } from "next/server";
import { db } from "@/lib/db";

export async function GET(req: NextRequest) {
  try {
    const { searchParams } = new URL(req.url);
    const companyId = searchParams.get("companyId");
    const search = searchParams.get("search");
    const startDate = searchParams.get("startDate");
    const endDate = searchParams.get("endDate");
    const fetchAll = searchParams.get("all") === "true" || searchParams.get("fetchAll") === "true";
    const limit = parseInt(searchParams.get("limit") || "50", 10);
    const page = parseInt(searchParams.get("page") || "1", 10);
    const offset = (page - 1) * limit;

    const where: any = {};

    if (companyId && companyId !== "all") {
      where.companyId = companyId;
    }

    if (startDate || endDate) {
      where.createdAt = {};
      if (startDate) {
        const [yr, mo, dy] = startDate.split("-").map(Number);
        const start = new Date(yr, mo - 1, dy, 0, 0, 0, 0);
        where.createdAt.gte = start;
      }
      if (endDate) {
        const [yr, mo, dy] = endDate.split("-").map(Number);
        const end = new Date(yr, mo - 1, dy, 23, 59, 59, 999);
        where.createdAt.lte = end;
      }
    }

    if (search && search.trim() !== "") {
      const trimmedSearch = search.trim();
      where.OR = [
        { code: { contains: trimmedSearch } },
        { productSKU: { contains: trimmedSearch } },
        { productName: { contains: trimmedSearch } },
        { productDesc: { contains: trimmedSearch } },
        { company: { name: { contains: trimmedSearch } } },
        { company: { pan: { contains: trimmedSearch } } },
      ];
    }

    const queryOptions: any = {
      where,
      include: { company: true },
      orderBy: { createdAt: "desc" },
    };

    if (!fetchAll) {
      queryOptions.take = limit;
      queryOptions.skip = offset;
    }

    const [barcodes, total] = await Promise.all([
      db.barcode.findMany(queryOptions),
      db.barcode.count({ where }),
    ]);

    return NextResponse.json({
      barcodes,
      pagination: {
        total,
        page: fetchAll ? 1 : page,
        limit: fetchAll ? total : limit,
        pages: fetchAll ? 1 : Math.ceil(total / limit),
      },
    });
  } catch (error: any) {
    console.error("API Error listing barcodes:", error);
    return NextResponse.json(
      { error: "server_error", message: "Failed to fetch barcode history." },
      { status: 500 }
    );
  }
}
