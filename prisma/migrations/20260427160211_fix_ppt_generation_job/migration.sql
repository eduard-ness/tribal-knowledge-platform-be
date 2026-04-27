-- AlterTable
ALTER TABLE "PptGenerationJob" ADD COLUMN     "sharePointUrl" TEXT;

-- CreateTable
CREATE TABLE "IngestDocument" (
    "id" TEXT NOT NULL,
    "sourceType" TEXT NOT NULL,
    "title" TEXT,
    "sourceUrl" TEXT,
    "content" TEXT NOT NULL,
    "ingestJobId" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "IngestDocument_pkey" PRIMARY KEY ("id")
);

-- AddForeignKey
ALTER TABLE "IngestDocument" ADD CONSTRAINT "IngestDocument_ingestJobId_fkey" FOREIGN KEY ("ingestJobId") REFERENCES "IngestJob"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
