/**
 * アプリより先に必ず評価する（ESM で import が先に束ねられるため、このファイルのみを経由させる）。
 */
import path from "node:path";
import { fileURLToPath } from "node:url";

import dotenv from "dotenv";

const backendEnvPath = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", ".env");
dotenv.config({ path: backendEnvPath });
