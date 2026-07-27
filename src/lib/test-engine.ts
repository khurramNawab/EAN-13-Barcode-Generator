import { db } from "./db";
import {
  computeGS1CheckDigit,
  extractPANLast4,
  generateSingleBarcode,
  shuffle,
  CapacityExceededError,
} from "./barcodeEngine";

async function runTests() {
  console.log("=== STARTING BARCODE ENGINE TESTS ===");

  // 1. Test Check Digit Computation
  console.log("\n1. Testing GS1 Check Digit Computation...");
  const test12Digit = "853012045217";
  const expectedCheckDigit = 2;
  const actualCheckDigit = computeGS1CheckDigit(test12Digit);
  if (actualCheckDigit === expectedCheckDigit) {
    console.log(`✓ Success: Check digit for ${test12Digit} is ${actualCheckDigit}`);
  } else {
    console.error(`✗ Failed: Expected ${expectedCheckDigit}, got ${actualCheckDigit}`);
    process.exit(1);
  }

  // 2. Test PAN Last 4 digits extraction
  console.log("\n2. Testing PAN Extraction...");
  const pans = [
    { input: "ABCDE1234F", expected: "1234" },
    { input: "abcde5678g", expected: "5678" },
    { input: "  xyzpq9876z  ", expected: "9876" },
  ];

  for (const { input, expected } of pans) {
    const ext = extractPANLast4(input);
    if (ext === expected) {
      console.log(`✓ Success: Extracted ${ext} from ${input}`);
    } else {
      console.error(`✗ Failed: Expected ${expected} from ${input}, got ${ext}`);
      process.exit(1);
    }
  }

  // 3. Setup a test company in database
  console.log("\n3. Setting up test company in DB...");
  const testPAN = "TESTP1111T";
  // Delete existing test company if any
  const existingCompany = await db.company.findUnique({ where: { pan: testPAN } });
  if (existingCompany) {
    await db.company.delete({ where: { pan: testPAN } });
  }

  const testCompany = await db.company.create({
    data: {
      name: "Test Barcode Company Inc",
      pan: testPAN,
    },
  });
  console.log(`✓ Test company created with ID: ${testCompany.id}`);

  // 4. Generate 990 barcodes (entire capacity limit of 99 sequences * 10 shuffles)
  console.log("\n4. Generating 990 barcodes sequentially (Capacity limit)...");
  const codes = new Set<string>();
  let count = 0;

  for (let i = 0; i < 990; i++) {
    const barcode = await db.$transaction(async (tx) => {
      return await generateSingleBarcode(tx, testCompany.id, {
        sku: `SKU-${i}`,
        name: `Product ${i}`,
      });
    });

    // Check code length
    if (barcode.code.length !== 13) {
      console.error(`✗ Failed: Barcode ${barcode.code} is not 13 digits long`);
      process.exit(1);
    }

    // Check code format prefix
    if (!barcode.code.startsWith("853")) {
      console.error(`✗ Failed: Barcode ${barcode.code} does not start with 853`);
      process.exit(1);
    }

    // Check Check Digit
    const calcDigit = computeGS1CheckDigit(barcode.code.slice(0, 12));
    if (calcDigit !== barcode.checkDigit || barcode.code[12] !== String(calcDigit)) {
      console.error(`✗ Failed: Incorrect check digit for ${barcode.code}. Calc: ${calcDigit}, DB: ${barcode.checkDigit}`);
      process.exit(1);
    }

    // Check uniqueness in local set
    if (codes.has(barcode.code)) {
      console.error(`✗ Failed: Duplicate barcode generated: ${barcode.code}`);
      process.exit(1);
    }
    codes.add(barcode.code);
    count++;
  }

  console.log(`✓ Success: Generated ${count} barcodes without any duplicates.`);

  // 5. Test capacity overflow error (991st barcode should fail)
  console.log("\n5. Testing capacity overflow (error mode)...");
  try {
    await db.$transaction(async (tx) => {
      return await generateSingleBarcode(tx, testCompany.id);
    });
    console.error("✗ Failed: Expected CapacityExceededError but no error was thrown.");
    process.exit(1);
  } catch (err: any) {
    if (err instanceof CapacityExceededError) {
      console.log(`✓ Success: Threw expected CapacityExceededError: "${err.message}"`);
    } else {
      console.error("✗ Failed: Threw unexpected error: ", err);
      process.exit(1);
    }
  }

  // 6. Test sequence reset admin override
  console.log("\n6. Testing sequence reset config override...");
  // Enable overflow reset
  await db.barcodeState.update({
    where: { companyId: testCompany.id },
    data: { overflowReset: true },
  });

  // Generate 991st barcode (should succeed now and reset sequence to 01, using tomorrow's date to avoid collision)
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const wrapBarcode = await db.$transaction(async (tx) => {
    return await generateSingleBarcode(tx, testCompany.id, undefined, tomorrow);
  });

  console.log(`✓ Success: Reset enabled. Generated barcode after limit: ${wrapBarcode.code} (Sequence: ${wrapBarcode.sequence})`);

  // Clean up
  await db.company.delete({ where: { pan: testPAN } });
  console.log("\n✓ Cleaned up test database records.");

  console.log("\n=== ALL TESTS PASSED SUCCESSFULLY ===");
}

runTests()
  .catch((err) => {
    console.error("Test execution failed with error: ", err);
    process.exit(1);
  })
  .finally(async () => {
    await db.$disconnect();
  });
