import { PlaceholderPage } from "@/components/shared/placeholder-page";

export default async function ReportsCatchAllPage({
  params,
}: {
  params: Promise<{ slug: string[] }>;
}) {
  const { slug } = await params;
  const path = "/reports/" + slug.join("/");
  return <PlaceholderPage title="報告書（プレースホルダー）" path={path} />;
}
