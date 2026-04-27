import { createApp } from "./app.js";

const parsedPort = Number.parseInt(process.env.PORT ?? "4100", 10);
const port = Number.isFinite(parsedPort) && parsedPort > 0 ? parsedPort : 4100;

const app = createApp();

app.listen(port, () => {
  console.info(`Backend server listening on http://localhost:${port}`);
});
