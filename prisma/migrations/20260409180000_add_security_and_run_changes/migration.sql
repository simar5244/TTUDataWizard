-- Add security controls on user
ALTER TABLE "User"
  ADD COLUMN IF NOT EXISTS "allowEdits" BOOLEAN NOT NULL DEFAULT true,
  ADD COLUMN IF NOT EXISTS "allowDeletes" BOOLEAN NOT NULL DEFAULT true;

-- Add run-level change tracking
ALTER TABLE IF EXISTS "MappingRun"
  ADD COLUMN IF NOT EXISTS "direction" TEXT NOT NULL DEFAULT 'excel_to_excel',
  ADD COLUMN IF NOT EXISTS "inputData" JSONB,
  ADD COLUMN IF NOT EXISTS "changeSet" JSONB,
  ADD COLUMN IF NOT EXISTS "changedColumns" JSONB,
  ADD COLUMN IF NOT EXISTS "changedCellCount" INTEGER DEFAULT 0,
  ADD COLUMN IF NOT EXISTS "changedRowCount" INTEGER DEFAULT 0;
