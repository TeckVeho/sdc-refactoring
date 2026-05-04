"use client";

import { Printer } from "lucide-react";

import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

interface PrintButtonProps {
  onPrint?: () => void;
  className?: string;
  children?: React.ReactNode;
}

export function PrintButton({ onPrint, className, children }: PrintButtonProps) {
  const handleClick = () => {
    if (onPrint) {
      onPrint();
    } else {
      window.print();
    }
  };

  return (
    <Button type="button" variant="outline" onClick={handleClick} className={cn("gap-2", className)}>
      <Printer className="h-4 w-4" />
      {children ?? "印刷"}
    </Button>
  );
}
