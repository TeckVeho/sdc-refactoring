import { Button } from "@/components/ui/button";

type PagePlaceholderProps = {
  title: string;
  description: string;
};

export function PagePlaceholder({ title, description }: PagePlaceholderProps) {
  return (
    <section className="space-y-4">
      <header className="space-y-2">
        <h1 className="text-2xl font-semibold tracking-tight text-foreground">{title}</h1>
        <p className="max-w-3xl text-sm text-muted-foreground">{description}</p>
      </header>
      <div className="rounded-xl border border-border bg-card p-5 shadow-sm">
        <p className="text-sm text-muted-foreground">
          This route is prepared as a Phase 1 scaffold and will be implemented in upcoming phases.
        </p>
        <div className="mt-4">
          <Button variant="secondary">Planned Screen</Button>
        </div>
      </div>
    </section>
  );
}
