-- CreateTable
CREATE TABLE "Company" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "name" TEXT NOT NULL,
    "pan" TEXT NOT NULL,
    "createdAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- CreateTable
CREATE TABLE "BarcodeState" (
    "companyId" TEXT NOT NULL PRIMARY KEY,
    "currentSequence" INTEGER NOT NULL DEFAULT 1,
    "currentShuffle" TEXT NOT NULL,
    "overflowReset" BOOLEAN NOT NULL DEFAULT false,
    CONSTRAINT "BarcodeState_companyId_fkey" FOREIGN KEY ("companyId") REFERENCES "Company" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- CreateTable
CREATE TABLE "Barcode" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "code" TEXT NOT NULL,
    "sequence" INTEGER NOT NULL,
    "rotatingDigit" INTEGER NOT NULL,
    "datePart" TEXT NOT NULL,
    "checkDigit" INTEGER NOT NULL,
    "productSKU" TEXT,
    "productName" TEXT,
    "productDesc" TEXT,
    "companyId" TEXT NOT NULL,
    "createdAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT "Barcode_companyId_fkey" FOREIGN KEY ("companyId") REFERENCES "Company" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- CreateIndex
CREATE UNIQUE INDEX "Company_pan_key" ON "Company"("pan");

-- CreateIndex
CREATE UNIQUE INDEX "Barcode_code_key" ON "Barcode"("code");
