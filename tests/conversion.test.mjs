import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import test from "node:test";
import { build } from "esbuild";

const execFileAsync = promisify(execFile);

function rgbHex(color) {
  const channel = (value) =>
    Math.round(Math.max(0, Math.min(1, value)) * 255)
      .toString(16)
      .padStart(2, "0");
  return `${channel(color.r)}${channel(color.g)}${channel(color.b)}`.toUpperCase();
}

function opacityToTransparency(opacity) {
  return Math.round((1 - Math.max(0, Math.min(1, opacity))) * 100);
}

function normalizeRotation(value) {
  return ((value % 360) + 360) % 360;
}

test("converts normalized RGB to PowerPoint hex", () => {
  assert.equal(rgbHex({ r: 1, g: 0.5, b: 0 }), "FF8000");
  assert.equal(rgbHex({ r: -1, g: 2, b: 0 }), "00FF00");
});

test("converts opacity to clamped transparency", () => {
  assert.equal(opacityToTransparency(1), 0);
  assert.equal(opacityToTransparency(0.25), 75);
  assert.equal(opacityToTransparency(-1), 100);
});

test("normalizes rotations", () => {
  assert.equal(normalizeRotation(-45), 315);
  assert.equal(normalizeRotation(405), 45);
});

test("orders explicitly numbered frames by their numeric prefix", async () => {
  const { sortBySlideOrder } = await importOrderModule();
  const frames = [
    { name: "10-End", x: 0, y: 0, width: 100, height: 100 },
    { name: "2-Content", x: 200, y: 0, width: 100, height: 100 },
    { name: "1-Title", x: 400, y: 0, width: 100, height: 100 }
  ];
  assert.deepEqual(sortBySlideOrder(frames).map((frame) => frame.name), [
    "1-Title",
    "2-Content",
    "10-End"
  ]);
});

test("orders unnumbered frames by canvas rows", async () => {
  const { sortBySlideOrder } = await importOrderModule();
  const frames = [
    { name: "Bottom", x: 0, y: 300, width: 100, height: 100 },
    { name: "Top right", x: 300, y: 5, width: 100, height: 100 },
    { name: "Top left", x: 0, y: 0, width: 100, height: 100 }
  ];
  assert.deepEqual(sortBySlideOrder(frames).map((frame) => frame.name), [
    "Top left",
    "Top right",
    "Bottom"
  ]);
});

test("writes Figma text at the correct PowerPoint point size", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "figma-pptx-test-"));
  const bundlePath = path.join(directory, "pptx.mjs");
  const presentationPath = path.join(directory, "font-size.pptx");

  await build({
    entryPoints: ["src/pptx.ts"],
    outfile: bundlePath,
    bundle: true,
    platform: "node",
    format: "esm"
  });
  const { createPresentation } = await import(`file://${bundlePath}`);
  await createPresentation(
    [{
      id: "slide",
      name: "Font regression",
      width: 1920,
      height: 1080,
      warnings: [],
      layers: [{
        kind: "text",
        id: "title",
        name: "Title",
        x: 100,
        y: 100,
        width: 800,
        height: 120,
        rotation: 0,
        opacity: 1,
        runs: [{
          text: "Editable title",
          fontFamily: "Arial",
          fontStyle: "Regular",
          fontWeight: 400,
          fontSize: 48,
          color: { r: 0, g: 0, b: 0 },
          opacity: 1
        }, {
          text: " Light",
          fontFamily: "Roboto",
          fontStyle: "Light",
          fontWeight: 300,
          fontSize: 48,
          color: { r: 0, g: 0, b: 0 },
          opacity: 1
        }, {
          text: " Medium",
          fontFamily: "Roboto",
          fontStyle: "Medium",
          fontWeight: 500,
          fontSize: 48,
          color: { r: 0, g: 0, b: 0 },
          opacity: 1
        }, {
          text: " Bold",
          fontFamily: "Roboto",
          fontStyle: "Bold",
          fontWeight: 700,
          fontSize: 48,
          color: { r: 0, g: 0, b: 0 },
          opacity: 1
        }],
        align: "LEFT",
        verticalAlign: "TOP"
      }]
    }],
    presentationPath,
    { rasterScale: 2, includeHidden: false, addSpeakerNotes: false }
  );

  await execFileAsync("unzip", ["-q", presentationPath, "ppt/slides/slide1.xml", "-d", directory]);
  const slideXml = await readFile(path.join(directory, "ppt/slides/slide1.xml"), "utf8");
  // 48 Figma px × (13.333 in / 1920 px) × 72 pt/in ≈ 24 pt.
  assert.match(slideXml, /<a:rPr[^>]*sz="2400"/);
  assert.match(slideXml, /<a:latin typeface="Roboto Light"/);
  assert.match(slideXml, /<a:latin typeface="Roboto Medium"/);
  assert.match(slideXml, /<a:rPr[^>]*b="1"[^>]*>[\s\S]*?<a:latin typeface="Roboto"/);
});

async function importOrderModule() {
  const directory = await mkdtemp(path.join(tmpdir(), "figma-order-test-"));
  const bundlePath = path.join(directory, "order.mjs");
  await build({
    entryPoints: ["src/order.ts"],
    outfile: bundlePath,
    bundle: true,
    platform: "node",
    format: "esm"
  });
  return import(`file://${bundlePath}`);
}
