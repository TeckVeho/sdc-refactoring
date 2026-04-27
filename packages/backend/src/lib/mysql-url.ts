import { URL } from "node:url";

export type MysqlUrlParts = {
  host: string;
  port: number;
  user: string;
  password: string;
  database: string;
};

export function parseMysqlUrl(databaseUrl: string): MysqlUrlParts {
  const u = new URL(databaseUrl);
  const port = u.port ? Number.parseInt(u.port, 10) : 3306;
  const database = u.pathname.replace(/^\//, "").split("?")[0] ?? "";

  return {
    host: u.hostname,
    port: Number.isFinite(port) ? port : 3306,
    user: decodeURIComponent(u.username),
    password: decodeURIComponent(u.password),
    database,
  };
}
