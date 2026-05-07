import { redirect } from "next/navigation";

import { isPublicSampleDataDemo } from "@/lib/demo-mode";
import { apiUrl } from "@/lib/api";

import { LoginForm } from "./login-form";

type HealthPayload = {
  data?: {
    guestApiAccess?: boolean;
    sampleDataMode?: boolean;
    relaxedGuestAuth?: boolean;
  };
};

async function fetchGuestApiAccess(): Promise<{
  guestApiAccess: boolean;
  backendReachable: boolean;
}> {
  try {
    const res = await fetch(apiUrl("/api/health"), {
      cache: "no-store",
    });
    if (!res.ok) {
      return { guestApiAccess: false, backendReachable: true };
    }
    const json = (await res.json()) as HealthPayload;
    return {
      guestApiAccess: Boolean(json?.data?.guestApiAccess),
      backendReachable: true,
    };
  } catch {
    return { guestApiAccess: false, backendReachable: false };
  }
}

export default async function LoginPage() {
  const publicDemo = isPublicSampleDataDemo();
  const { guestApiAccess, backendReachable } = await fetchGuestApiAccess();

  if (publicDemo && guestApiAccess) {
    redirect("/dashboard");
  }

  return (
    <LoginForm
      demoEnvHint={
        publicDemo && guestApiAccess === false && backendReachable
          ? "missing-backend-demo"
          : publicDemo && !backendReachable
            ? "backend-down"
            : undefined
      }
    />
  );
}
