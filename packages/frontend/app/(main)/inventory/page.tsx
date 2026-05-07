import { redirect } from "next/navigation";

export default function InventoryIndexPage() {
  redirect("/inventory/shipment-summary");
}
