import { SectionTabNav } from "@/components/shared/section-tab-nav";
import { REPORTS_TABS } from "@/lib/nav-config";

export default function ReportsSectionLayout({ children }: { children: React.ReactNode }) {
  return (
    <>
      <SectionTabNav tabs={REPORTS_TABS} ariaLabel="報告書" />
      {children}
    </>
  );
}
