import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { defineConfig } from "vite";

// Certificado local do Office (npm run certificado) para testar em https://localhost:3000
const pasta = join(homedir(), ".office-addin-dev-certs");
const https = existsSync(join(pasta, "localhost.crt"))
  ? { cert: readFileSync(join(pasta, "localhost.crt")), key: readFileSync(join(pasta, "localhost.key")) }
  : undefined;

export default defineConfig({
  base: "./",
  server: { port: 3000, https },
  build: { outDir: "dist", target: "es2020" },
});
