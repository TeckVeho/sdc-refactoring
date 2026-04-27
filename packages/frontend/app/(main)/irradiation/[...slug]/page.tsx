import { PlaceholderPage } from "@/components/shared/placeholder-page";

export default async function IrradiationCatchAllPage({
  params,
}: {
  params: Promise<{ slug: string[] }>;
}) {
  const { slug } = await params;
  const path = "/irradiation/" + slug.join("/");
  return <PlaceholderPage title="照射管理（プレースホルダー）" path={path} />;
}
