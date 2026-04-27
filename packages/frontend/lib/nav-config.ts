export type NavItem = { href: string; label: string };

export type NavGroup = { title: string; items: NavItem[] };

/** 全画面ナビ（サイドバーとホームメニューで共有） */
export const NAV: NavGroup[] = [
  {
    title: "ダッシュボード",
    items: [{ href: "/dashboard", label: "ホームメニュー" }],
  },
  {
    title: "照射管理",
    items: [
      { href: "/irradiation/machine1/monitor", label: "1号機照射情報" },
      { href: "/irradiation/machine1/plan", label: "1号機照射計画" },
      { href: "/irradiation/machine1/work-order", label: "1号機作業指図書" },
      { href: "/irradiation/machine-status", label: "装置運転状況" },
      { href: "/irradiation/production", label: "生産情報一覧" },
    ],
  },
  {
    title: "入出荷",
    items: [
      { href: "/inventory/shipment-summary", label: "入出荷集計" },
      { href: "/inventory/company-arrival", label: "会社別入荷集計" },
      { href: "/inventory/month-end", label: "月末在庫集計" },
      { href: "/inventory/customer-stock", label: "顧客在庫報告" },
      { href: "/inventory/ric3-repack", label: "Ric3詰替作業" },
    ],
  },
  {
    title: "線量管理",
    items: [
      { href: "/dosimetry/oven-management", label: "R3オーブン管理" },
      { href: "/dosimetry/dose-search", label: "線量検索" },
      { href: "/dosimetry/irradiation-results", label: "照射実績表示" },
      { href: "/dosimetry/jmm60", label: "JMM60φ記入用紙" },
      { href: "/dosimetry/jmm90", label: "JMM90φ記入用紙" },
      { href: "/dosimetry/dose-rate-calc", label: "線量率計算" },
    ],
  },
  {
    title: "報告書",
    items: [
      { href: "/reports/ric2-dose-shortage", label: "RIC2線量不足報告書" },
      { href: "/reports/ric3-dose-shortage", label: "RIC3線量不足報告書" },
      { href: "/reports/ric3-dose-shortage-sumiden", label: "RIC3（住電）" },
      { href: "/reports/shipment-method", label: "出荷方法報告書" },
      { href: "/reports/unissued-search", label: "報告書未発行検索" },
      { href: "/reports/gamma-results", label: "ガンマ照射課実績集計" },
    ],
  },
  {
    title: "管理・マスタ",
    items: [
      { href: "/admin/calendar", label: "カレンダー" },
      { href: "/admin/shift-schedule", label: "勤務表" },
      { href: "/admin/price-search", label: "単価検索" },
      { href: "/admin/db-browser", label: "DBブラウザ" },
      { href: "/admin/company-code-convert", label: "GM/EB会社コード変換" },
      { href: "/admin/shipment-method-verify", label: "出荷方法（検証）" },
      { href: "/admin/shipment-method-ss", label: "出荷方法（SS検証）" },
    ],
  },
];
