import { PlaceholderPage } from "@/components/shared/placeholder-page";

export default async function InventoryCatchAllPage({
  params,
}: {
  params: Promise<{ slug: string[] }>;
}) {
  const { slug } = await params;
  const path = "/inventory/" + slug.join("/");
  return <PlaceholderPage title="入出荷（プレースホルダー）" path={path} />;
}
