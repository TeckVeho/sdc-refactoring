import { SectionTabNav } from "@/components/shared/section-tab-nav";
import { ADMIN_TABS } from "@/lib/nav-config";

export default function AdminSectionLayout({ children }: { children: React.ReactNode }) {
  return (
    <>
      <SectionTabNav tabs={ADMIN_TABS} ariaLabel="マスタ管理" />
      {children}
    </>
  );
}
