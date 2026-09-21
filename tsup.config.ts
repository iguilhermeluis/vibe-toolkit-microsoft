import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm", "iife"],
  globalName: "VibeToolkitMicrosoft",
  dts: true,
  clean: true,
  sourcemap: true,
  treeshake: true,
  outDir: "dist"
});