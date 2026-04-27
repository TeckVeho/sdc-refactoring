"use client";

import { useRouter } from "next/navigation";
import { useState } from "react";

import { Button } from "@/components/ui/button";
import { apiUrl } from "@/lib/api";

export function LoginForm() {
  const router = useRouter();
  const [employeeId, setEmployeeId] = useState("dev");
  const [password, setPassword] = useState("devpass");
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError(null);
    setLoading(true);
    try {
      const res = await fetch(apiUrl("/api/auth/login"), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        credentials: "include",
        body: JSON.stringify({ employeeId, password }),
      });
      const json = (await res.json()) as { data?: unknown; error?: string };
      if (!res.ok) {
        setError(json.error ?? "ログインに失敗しました");
        return;
      }
      router.push("/dashboard");
      router.refresh();
    } catch {
      setError("通信エラーが発生しました");
    } finally {
      setLoading(false);
    }
  }

  return (
    <form
      onSubmit={onSubmit}
      className="space-y-5 rounded-2xl border border-cream-300 bg-cream-50 p-8 shadow-sm"
    >
      <div>
        <h1 className="text-xl font-semibold text-ink">ログイン</h1>
        <p className="mt-1 text-sm text-ink-muted">社員番号とパスワードを入力してください。</p>
      </div>
      <div className="space-y-2">
        <label className="block text-sm font-medium text-ink" htmlFor="employeeId">
          社員番号
        </label>
        <input
          id="employeeId"
          name="employeeId"
          autoComplete="username"
          className="w-full rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink outline-none ring-terracotta/0 transition focus:ring-2"
          value={employeeId}
          onChange={(ev) => setEmployeeId(ev.target.value)}
        />
      </div>
      <div className="space-y-2">
        <label className="block text-sm font-medium text-ink" htmlFor="password">
          パスワード
        </label>
        <input
          id="password"
          name="password"
          type="password"
          autoComplete="current-password"
          className="w-full rounded-xl border border-cream-300 bg-white px-3 py-2 text-ink outline-none ring-terracotta/0 transition focus:ring-2"
          value={password}
          onChange={(ev) => setPassword(ev.target.value)}
        />
      </div>
      {error ? <p className="text-sm text-terracotta">{error}</p> : null}
      <Button type="submit" className="w-full" disabled={loading}>
        {loading ? "送信中…" : "ログイン"}
      </Button>
    </form>
  );
}
