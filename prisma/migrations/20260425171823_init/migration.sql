-- CreateTable
CREATE TABLE "IngestJob" (
    "id" TEXT NOT NULL,
    "status" TEXT NOT NULL,
    "progress" INTEGER NOT NULL DEFAULT 0,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "IngestJob_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "UploadedFile" (
    "id" TEXT NOT NULL,
    "fileName" TEXT NOT NULL,
    "contentType" TEXT,
    "size" INTEGER,
    "ingestJobId" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "UploadedFile_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "QueryLog" (
    "id" TEXT NOT NULL,
    "project" TEXT,
    "question" TEXT NOT NULL,
    "answer" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "QueryLog_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PptGenerationJob" (
    "id" TEXT NOT NULL,
    "project" TEXT NOT NULL,
    "audience" TEXT NOT NULL,
    "fileName" TEXT,
    "status" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "PptGenerationJob_pkey" PRIMARY KEY ("id")
);

-- AddForeignKey
ALTER TABLE "UploadedFile" ADD CONSTRAINT "UploadedFile_ingestJobId_fkey" FOREIGN KEY ("ingestJobId") REFERENCES "IngestJob"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
