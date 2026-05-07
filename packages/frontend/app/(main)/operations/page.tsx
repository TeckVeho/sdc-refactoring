import { redirect } from "next/navigation";

export default function OperationsIndexPage() {
  redirect("/operations/machine-status");
}
