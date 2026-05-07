import { SectionTabNav } from "@/components/shared/section-tab-nav";
import { DOSIMETRY_TABS } from "@/lib/nav-config";

export default function DosimetrySectionLayout({ children }: { children: React.ReactNode }) {
  return (
    <>
      <SectionTabNav tabs={DOSIMETRY_TABS} ariaLabel="線量管理" />
      {children}
    </>
  );
}
