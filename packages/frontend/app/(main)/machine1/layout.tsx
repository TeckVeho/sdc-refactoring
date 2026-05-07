import { SectionTabNav } from "@/components/shared/section-tab-nav";
import { MACHINE1_TABS } from "@/lib/nav-config";

export default function Machine1SectionLayout({ children }: { children: React.ReactNode }) {
  return (
    <>
      <SectionTabNav tabs={MACHINE1_TABS} ariaLabel="1号機" />
      {children}
    </>
  );
}
