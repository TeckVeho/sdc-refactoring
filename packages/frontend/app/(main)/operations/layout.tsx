import { SectionTabNav } from "@/components/shared/section-tab-nav";
import { OPERATIONS_TABS } from "@/lib/nav-config";

export default function OperationsSectionLayout({ children }: { children: React.ReactNode }) {
  return (
    <>
      <SectionTabNav tabs={OPERATIONS_TABS} ariaLabel="設備・生産" />
      {children}
    </>
  );
}
