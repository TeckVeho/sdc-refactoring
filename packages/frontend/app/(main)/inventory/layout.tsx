import { SectionTabNav } from "@/components/shared/section-tab-nav";
import { INVENTORY_TABS } from "@/lib/nav-config";

export default function InventorySectionLayout({ children }: { children: React.ReactNode }) {
  return (
    <>
      <SectionTabNav tabs={INVENTORY_TABS} ariaLabel="入出荷" />
      {children}
    </>
  );
}
