import { build, context } from "esbuild";
import { mkdir, readFile, writeFile } from "node:fs/promises";

const watch = process.argv.includes("--watch");
await mkdir("dist", { recursive: true });

const common = {
  bundle: true,
  sourcemap: true,
  target: "es2020",
  logLevel: "info"
};

async function buildUiHtml() {
  await build({
    ...common,
    entryPoints: ["src/ui.ts"],
    outfile: "dist/ui.js",
    format: "iife"
  });
  const [template, script] = await Promise.all([
    readFile("src/ui.html", "utf8"),
    readFile("dist/ui.js", "utf8")
  ]);
  await writeFile("dist/ui.html", template.replace("/*__UI_BUNDLE__*/", script));
}

if (watch) {
  const pluginContext = await context({
    ...common,
    entryPoints: ["src/code.ts"],
    outfile: "dist/code.js",
    format: "iife"
  });
  await pluginContext.watch();
  await buildUiHtml();
  console.log("Watching plugin code. Re-run the command after UI source changes.");
} else {
  await Promise.all([
    build({
      ...common,
      entryPoints: ["src/code.ts"],
      outfile: "dist/code.js",
      format: "iife"
    }),
    buildUiHtml()
  ]);
}
