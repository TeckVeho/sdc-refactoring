import type { Metadata } from "next";
import { Inter } from "next/font/google";

import { QueryProvider } from "@/components/shared/query-provider";
import "./globals.css";

const inter = Inter({ subsets: ["latin"], variable: "--font-inter" });

export const metadata: Metadata = {
  title: "SDC 照射管理",
  description: "照射管理システム Web版",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="ja" className={inter.variable}>
      <body className="min-h-screen font-sans">
        <QueryProvider>{children}</QueryProvider>
      </body>
    </html>
  );
}
