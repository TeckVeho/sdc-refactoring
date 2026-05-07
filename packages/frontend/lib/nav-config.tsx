import type { LucideIcon } from "lucide-react";
import {
  Database,
  Factory,
  FileText,
  FlaskConical,
  LayoutDashboard,
  ScanLine,
  Settings,
  Truck,
} from "lucide-react";

import type { SectionTab } from "@/components/shared/section-tab-nav";

/** サイドバー直上の8項目 — フラット一覧 */
export type PrimaryNavItem = {
  label: string;
  href: string;
  icon: LucideIcon;
  /** アクティブ判定: pathname がいずれかで始まる */
  prefixes: readonly string[];
};

export const PRIMARY_NAV: PrimaryNavItem[] = [
  {
    label: "ダッシュボード",
    href: "/dashboard",
    icon: LayoutDashboard,
    prefixes: ["/dashboard"],
  },
  {
    label: "1号機",
    href: "/machine1/monitor",
    icon: ScanLine,
    prefixes: ["/machine1"],
  },
  {
    label: "設備・生産",
    href: "/operations/machine-status",
    icon: Factory,
    prefixes: ["/operations"],
  },
  {
    label: "入出荷",
    href: "/inventory/shipment-summary",
    icon: Truck,
    prefixes: ["/inventory"],
  },
  {
    label: "線量管理",
    href: "/dosimetry/oven-management",
    icon: FlaskConical,
    prefixes: ["/dosimetry"],
  },
  {
    label: "報告書",
    href: "/reports/ric2-dose-shortage",
    icon: FileText,
    prefixes: ["/reports"],
  },
  {
    label: "マスタ管理",
    href: "/admin/calendar",
    icon: Settings,
    prefixes: ["/admin"],
  },
  {
    label: "DBブラウザ",
    href: "/db-browser",
    icon: Database,
    prefixes: ["/db-browser"],
  },
];

export function isNavActiveForPath(pathname: string, item: PrimaryNavItem): boolean {
  return item.prefixes.some((p) => pathname === p || pathname.startsWith(`${p}/`));
}

export const MACHINE1_TABS: SectionTab[] = [
  { href: "/machine1/monitor", label: "照射情報" },
  { href: "/machine1/plan", label: "照射計画" },
  { href: "/machine1/work-order", label: "作業指図書" },
];

export const OPERATIONS_TABS: SectionTab[] = [
  { href: "/operations/machine-status", label: "装置運転状況" },
  { href: "/operations/production", label: "生産情報一覧" },
];

export const INVENTORY_TABS: SectionTab[] = [
  { href: "/inventory/shipment-summary", label: "入出荷集計" },
  { href: "/inventory/company-arrival", label: "会社別入荷" },
  { href: "/inventory/month-end", label: "月末在庫" },
  { href: "/inventory/customer-stock", label: "顧客在庫" },
  { href: "/inventory/ric3-repack", label: "Ric3詰替" },
];

export const DOSIMETRY_TABS: SectionTab[] = [
  { href: "/dosimetry/oven-management", label: "R3オーブン" },
  { href: "/dosimetry/dose-search", label: "線量検索" },
  { href: "/dosimetry/irradiation-results", label: "照射実績" },
  { href: "/dosimetry/jmm60", label: "JMM60φ" },
  { href: "/dosimetry/jmm90", label: "JMM90φ" },
  { href: "/dosimetry/dose-rate-calc", label: "線量率計算" },
];

export const REPORTS_TABS: SectionTab[] = [
  { href: "/reports/ric2-dose-shortage", label: "RIC2 線量不足" },
  { href: "/reports/ric3-dose-shortage", label: "RIC3 線量不足" },
  { href: "/reports/ric3-dose-shortage-sumiden", label: "RIC3（住電）" },
  { href: "/reports/shipment-method", label: "出荷方法報告" },
  { href: "/reports/shipment-method-verify", label: "出荷方法（検証）" },
  { href: "/reports/shipment-method-ss", label: "出荷方法（SS）" },
  { href: "/reports/unissued-search", label: "未発行検索" },
  { href: "/reports/gamma-results", label: "ガンマ実績" },
];

export const ADMIN_TABS: SectionTab[] = [
  { href: "/admin/calendar", label: "カレンダー" },
  { href: "/admin/shift-schedule", label: "勤務表" },
  { href: "/admin/price-search", label: "単価検索" },
  { href: "/admin/company-code-convert", label: "会社コード変換" },
];
