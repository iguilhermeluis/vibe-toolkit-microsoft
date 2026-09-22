import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts", "src/auth.ts", "src/sharepoint.ts", "src/copilot.ts", "src/redirect-bridge.ts"],
  format: ["esm"],
  platform: "browser",
  target: "es2022",
  splitting: true,
  dts: true,
  clean: true,
  sourcemap: true,
  outDir: "dist"
});
