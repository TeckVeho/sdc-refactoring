"use client";

import { Download } from "lucide-react";
import * as XLSX from "xlsx";

import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

interface ExcelExportButtonProps {
  data: Record<string, unknown>[];
  fileName: string;
  sheetName?: string;
  className?: string;
  children?: React.ReactNode;
}

export function ExcelExportButton({
  data,
  fileName,
  sheetName = "Sheet1",
  className,
  children,
}: ExcelExportButtonProps) {
  const handleExport = () => {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    const outputFileName = fileName.endsWith(".xlsx") ? fileName : `${fileName}.xlsx`;
    XLSX.writeFile(workbook, outputFileName);
  };

  return (
    <Button type="button" variant="outline" onClick={handleExport} className={cn("gap-2", className)}>
      <Download className="h-4 w-4" />
      {children ?? "Excel出力"}
    </Button>
  );
}
