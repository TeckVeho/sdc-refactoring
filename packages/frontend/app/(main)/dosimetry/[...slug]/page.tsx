import { PlaceholderPage } from "@/components/shared/placeholder-page";

export default async function DosimetryCatchAllPage({
  params,
}: {
  params: Promise<{ slug: string[] }>;
}) {
  const { slug } = await params;
  const path = "/dosimetry/" + slug.join("/");
  return <PlaceholderPage title="線量管理（プレースホルダー）" path={path} />;
}
