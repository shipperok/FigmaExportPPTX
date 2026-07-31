import assert from "node:assert/strict";
import test from "node:test";

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
