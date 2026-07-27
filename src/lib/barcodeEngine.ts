import { PrismaClient } from "@prisma/client";

// Fisher-Yates Shuffle
export function shuffle(array: number[]): number[] {
  const arr = [...array];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// GS1 Mod 10 checksum
export function computeGS1CheckDigit(first12: string): number {
  if (first12.length !== 12 || !/^\d+$/.test(first12)) {
    throw new Error("Input must be exactly 12 digits");
  }
  let sum = 0;
  for (let i = 0; i < 12; i++) {
    const digit = parseInt(first12[i], 10);
    // Position is 1-based (odd 1, 3, 5... are index 0, 2, 4...)
    // and even positions (2, 4, 6... are index 1, 3, 5...)
    if ((i + 1) % 2 === 0) {
      sum += digit * 3;
    } else {
      sum += digit * 1;
    }
  }
  const checkDigit = (10 - (sum % 10)) % 10;
  return checkDigit;
}

// Extract 4 numeric digits from PAN (format AAAAA9999A)
// Standard PAN format has exactly 4 numbers at indices 5,6,7,8 (positions 6-9)
export function extractPANLast4(pan: string): string {
  const cleaned = pan.toUpperCase().trim();
  const match = cleaned.match(/[A-Z]{5}(\d{4})[A-Z]{1}/);
  if (match && match[1]) {
    return match[1];
  }
  // Fallback: search for first 4 consecutive digits
  const fallbackMatch = cleaned.match(/\d{4}/);
  if (fallbackMatch) {
    return fallbackMatch[0];
  }
  throw new Error("Invalid PAN format. Must be in the format AAAAA9999A (e.g. ABCDE1234F).");
}

export class DuplicateCollisionError extends Error {
  constructor(code: string) {
    super(`Global collision guard triggered: barcode ${code} already exists in database.`);
    this.name = "DuplicateCollisionError";
  }
}

export class CapacityExceededError extends Error {
  constructor(companyName: string) {
    super(`Sequence limit reached for company "${companyName}" — contact admin.`);
    this.name = "CapacityExceededError";
  }
}

/**
 * Generates a single barcode for a company within a transaction context.
 */
export async function generateSingleBarcode(
  tx: any, // Prisma transaction delegate
  companyId: string,
  productDetails?: { sku?: string; name?: string; desc?: string },
  customDate?: Date
) {
  // Load company and state
  const company = await tx.company.findUnique({
    where: { id: companyId },
    include: { state: true },
  });

  if (!company) {
    throw new Error(`Company with ID ${companyId} not found.`);
  }

  let state = company.state;
  if (!state) {
    // Initialize state if not exists
    const initialShuffle = shuffle([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    state = await tx.barcodeState.create({
      data: {
        companyId: companyId,
        currentSequence: 1,
        currentShuffle: JSON.stringify(initialShuffle),
        overflowReset: false,
      },
    });
  }

  let shuffleArr: number[] = JSON.parse(state.currentShuffle);
  let currentSequence = state.currentSequence;

  // If the shuffle is empty, advance the sequence
  if (shuffleArr.length === 0) {
    currentSequence += 1;
    if (currentSequence > 99) {
      if (state.overflowReset) {
        currentSequence = 1;
      } else {
        throw new CapacityExceededError(company.name);
      }
    }
    shuffleArr = shuffle([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
  }

  // Pop one digit off the shuffle
  const rotatingDigit = shuffleArr.pop()!;

  // Construct first 12 digits
  const seqPart = String(currentSequence).padStart(2, "0");
  const today = customDate || new Date();
  const datePart = String(today.getDate()).padStart(2, "0");
  const panPart = extractPANLast4(company.pan);

  const first12 = `853${seqPart}${datePart}${panPart}${rotatingDigit}`;
  const checkDigit = computeGS1CheckDigit(first12);
  const code = `${first12}${checkDigit}`;

  // Unique constraint safety check
  const existing = await tx.barcode.findUnique({
    where: { code },
  });

  if (existing) {
    throw new DuplicateCollisionError(code);
  }

  // Create barcode
  const barcode = await tx.barcode.create({
    data: {
      code,
      sequence: currentSequence,
      rotatingDigit,
      datePart,
      checkDigit,
      productSKU: productDetails?.sku || null,
      productName: productDetails?.name || null,
      productDesc: productDetails?.desc || null,
      companyId: company.id,
    },
    include: {
      company: true,
    },
  });

  // Prepare states for saving
  let nextSequence = currentSequence;
  let nextShuffle = [...shuffleArr];

  // If empty, pre-advance sequence so next iteration sees correct state
  if (nextShuffle.length === 0) {
    nextSequence += 1;
    if (nextSequence > 99) {
      if (state.overflowReset) {
        nextSequence = 1;
        nextShuffle = shuffle([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
      } else {
        // Leave empty, let next request throw the error
      }
    } else {
      nextShuffle = shuffle([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    }
  }

  // Update State in DB
  await tx.barcodeState.update({
    where: { companyId: company.id },
    data: {
      currentSequence: nextSequence,
      currentShuffle: JSON.stringify(nextShuffle),
    },
  });

  return barcode;
}
