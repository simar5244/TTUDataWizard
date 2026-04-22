-- AlterTable
ALTER TABLE "Dashboard" ADD COLUMN     "linkedMappings" JSONB NOT NULL DEFAULT '[]';

-- CreateTable
CREATE TABLE "MappingRun" (
    "id" TEXT NOT NULL,
    "mappingId" TEXT NOT NULL,
    "mappingVersionId" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "direction" TEXT NOT NULL DEFAULT 'excel_to_excel',
    "inputFileName" TEXT,
    "inputData" JSONB,
    "outputData" JSONB NOT NULL,
    "changeSet" JSONB,
    "changedColumns" JSONB,
    "changedCellCount" INTEGER DEFAULT 0,
    "changedRowCount" INTEGER DEFAULT 0,
    "rowCount" INTEGER NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "MappingRun_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "MappingRun_mappingId_idx" ON "MappingRun"("mappingId");

-- CreateIndex
CREATE INDEX "MappingRun_createdAt_idx" ON "MappingRun"("createdAt");

-- AddForeignKey
ALTER TABLE "MappingRun" ADD CONSTRAINT "MappingRun_mappingId_fkey" FOREIGN KEY ("mappingId") REFERENCES "Mapping"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "MappingRun" ADD CONSTRAINT "MappingRun_mappingVersionId_fkey" FOREIGN KEY ("mappingVersionId") REFERENCES "MappingVersion"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "MappingRun" ADD CONSTRAINT "MappingRun_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
