import session from "express-session";
import MySQLStoreFactory from "express-mysql-session";

import { parseMysqlUrl } from "../lib/mysql-url.js";

const MySQLStore = MySQLStoreFactory(session);

export function createSessionMiddleware() {
  const secret = process.env.SESSION_SECRET;
  if (!secret || secret.length < 8) {
    throw new Error("SESSION_SECRET must be set (min 8 characters)");
  }

  const databaseUrl = process.env.DATABASE_URL;
  if (!databaseUrl) {
    throw new Error("DATABASE_URL is required");
  }

  const { host, port, user, password, database } = parseMysqlUrl(databaseUrl);

  const store = new MySQLStore({
    host,
    port,
    user,
    password,
    database,
    createDatabaseTable: false,
    schema: {
      tableName: "web_sessions",
      columnNames: {
        session_id: "session_id",
        expires: "expires",
        data: "data",
      },
    },
  });

  return session({
    secret,
    store,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: process.env.NODE_ENV === "production",
      maxAge: 1000 * 60 * 60 * 8,
    },
    name: "sdc.sid",
  });
}
