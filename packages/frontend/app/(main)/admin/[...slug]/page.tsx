import { PlaceholderPage } from "@/components/shared/placeholder-page";

export default async function AdminCatchAllPage({ params }: { params: Promise<{ slug: string[] }> }) {
  const { slug } = await params;
  const path = "/admin/" + slug.join("/");
  return <PlaceholderPage title="管理・マスタ（プレースホルダー）" path={path} />;
}
