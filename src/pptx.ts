import PptxGenJS from "pptxgenjs";
import { t } from "./i18n";
import type {
  ExportLayer,
  ExportOptions,
  ExportSlide,
  Rgb,
  ShapeLayer,
  TextLayer
} from "./model";

export type ExportResult = {
  warnings: string[];
};

const SLIDE_WIDTH_IN = 13.333;

export async function createPresentation(
  slides: ExportSlide[],
  fileName: string,
  options: ExportOptions
): Promise<ExportResult> {
  if (!slides.length) throw new Error(t(options.locale, "errorNoSlides"));
  const pptx = new PptxGenJS();
  const first = slides[0];
  const slideHeight = SLIDE_WIDTH_IN * (first.height / first.width);
  pptx.defineLayout({
    name: "FIGMA",
    width: SLIDE_WIDTH_IN,
    height: slideHeight
  });
  pptx.layout = "FIGMA";
  pptx.author = "Editable PPTX Export for Figma";
  pptx.subject = "Editable presentation exported from Figma";
  pptx.title = fileName.replace(/\.pptx$/i, "");
  pptx.company = "Figma";
  pptx.theme = {
    headFontFace: "Arial",
    bodyFontFace: "Arial"
  };

  const warnings = new Set<string>();
  for (const source of slides) {
    const slide = pptx.addSlide();
    slide.background = { color: rgbHex(source.background ?? { r: 1, g: 1, b: 1 }) };
    const scale = Math.min(SLIDE_WIDTH_IN / source.width, slideHeight / source.height);
    const offsetX = (SLIDE_WIDTH_IN - source.width * scale) / 2;
    const offsetY = (slideHeight - source.height * scale) / 2;

    if (Math.abs(source.width / source.height - first.width / first.height) > 0.001) {
      warnings.add(t(options.locale, "warningDifferentSizes"));
    }
    source.warnings.forEach((warning) => warnings.add(warning));

    for (const layer of source.layers) {
      addLayer(pptx, slide, layer, scale, offsetX, offsetY);
    }
    if (options.addSpeakerNotes && source.warnings.length) {
      slide.addNotes(`${t(options.locale, "notesHeading")}\n${source.warnings.map((item) => `• ${item}`).join("\n")}`);
    }
  }

  await pptx.writeFile({ fileName });
  return { warnings: [...warnings] };
}

function addLayer(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  layer: ExportLayer,
  scale: number,
  offsetX: number,
  offsetY: number
): void {
  const position = {
    x: offsetX + layer.x * scale,
    y: offsetY + layer.y * scale,
    w: Math.max(layer.width * scale, 0.001),
    h: Math.max(layer.height * scale, 0.001),
    rotate: normalizeRotation(layer.rotation),
    transparency: opacityToTransparency(layer.opacity)
  };
  if (layer.kind === "text") {
    addText(slide, layer, position, scale);
  } else if (layer.kind === "shape") {
    addShape(pptx, slide, layer, position, scale);
  } else {
    slide.addImage({
      data: `data:${layer.mime};base64,${layer.base64}`,
      ...position
    });
  }
}

function addText(
  slide: PptxGenJS.Slide,
  layer: TextLayer,
  position: ReturnType<typeof layerPosition>,
  scale: number
): void {
  const runs = layer.runs.map((run) => ({
    text: run.text,
    options: (() => {
      const font = resolvePowerPointFont(run.fontFamily, run.fontStyle, run.fontWeight, run.italic);
      return {
      fontFace: font.fontFace,
      fontSize: Math.max(pixelsToPoints(run.fontSize, scale), 1),
      color: rgbHex(run.color),
      transparency: opacityToTransparency(run.opacity * (run.color.a ?? 1)),
      bold: font.bold,
      italic: font.italic,
      underline: run.textDecoration === "UNDERLINE" ? { style: "sng" as const } : undefined,
      strike: run.textDecoration === "STRIKETHROUGH" ? "sngStrike" as const : undefined,
      charSpacing: run.letterSpacing ? pixelsToPoints(run.letterSpacing, scale) : undefined,
      breakLine: false
      };
    })()
  }));
  slide.addText(runs, {
    ...position,
    margin: 0,
    breakLine: false,
    align: textAlign(layer.align),
    valign: verticalAlign(layer.verticalAlign),
    lineSpacingMultiple: layer.lineHeight && layer.runs[0]
      ? (layer.lineHeight / layer.runs[0].fontSize) * 1.0
      : undefined,
    paraSpaceAfter: layer.paragraphSpacing
      ? pixelsToPoints(layer.paragraphSpacing, scale)
      : undefined,
    fit: "shrink",
    isTextBox: true
  });
}

function addShape(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  layer: ShapeLayer,
  position: ReturnType<typeof layerPosition>,
  scale: number
): void {
  const fill = layer.fill
    ? { color: rgbHex(layer.fill), transparency: opacityToTransparency(layer.fill.a ?? 1) }
    : { color: "FFFFFF", transparency: 100 };
  const line = layer.stroke
    ? {
        color: rgbHex(layer.stroke),
        transparency: opacityToTransparency(layer.stroke.a ?? 1),
        width: Math.max(pixelsToPoints(layer.strokeWidth ?? 1, scale), 0.1),
        dash: dashType(layer.dash)
      }
    : { color: "FFFFFF", transparency: 100, width: 0 };

  if (layer.shape === "line") {
    slide.addShape(pptx.ShapeType.line, { ...position, fill, line });
    return;
  }
  const rounded = layer.shape === "rect" && (layer.radius ?? 0) > 0;
  slide.addShape(
    layer.shape === "ellipse"
      ? pptx.ShapeType.ellipse
      : rounded
        ? pptx.ShapeType.roundRect
        : pptx.ShapeType.rect,
    { ...position, fill, line }
  );
}

function layerPosition() {
  return { x: 0, y: 0, w: 0, h: 0, rotate: 0, transparency: 0 };
}

function rgbHex(color: Rgb): string {
  const channel = (value: number) =>
    Math.round(Math.max(0, Math.min(1, value)) * 255)
      .toString(16)
      .padStart(2, "0");
  return `${channel(color.r)}${channel(color.g)}${channel(color.b)}`.toUpperCase();
}

function opacityToTransparency(opacity: number): number {
  return Math.round((1 - Math.max(0, Math.min(1, opacity))) * 100);
}

function pixelsToPoints(pixels: number, inchesPerPixel: number): number {
  return pixels * inchesPerPixel * 72;
}

function resolvePowerPointFont(
  family: string,
  figmaStyle: string,
  fontWeight?: number,
  explicitItalic?: boolean
): { fontFace: string; bold: boolean; italic: boolean } {
  const baseFamily = family.trim() || "Arial";
  const italic = explicitItalic ?? /italic|oblique/i.test(figmaStyle);
  const styleWithoutPosture = figmaStyle
    .replace(/\b(?:italic|oblique)\b/gi, "")
    .replace(/\s+/g, " ")
    .trim();
  const plainStyle = /^(?:regular|normal|book|roman)?$/i.test(styleWithoutPosture);
  const standardBold = /^(?:bold)$/i.test(styleWithoutPosture);

  if (plainStyle) {
    return {
      fontFace: baseFamily,
      bold: (fontWeight ?? 400) >= 700,
      italic
    };
  }
  if (standardBold) {
    return { fontFace: baseFamily, bold: true, italic };
  }

  // OOXML exposes only boolean bold/italic switches. Intermediate weights such
  // as Light, Medium and SemiBold must therefore be addressed by their installed
  // face name (for example, "Roboto Medium").
  const familyAlreadyIncludesStyle = baseFamily
    .toLocaleLowerCase()
    .endsWith(styleWithoutPosture.toLocaleLowerCase());
  return {
    fontFace: familyAlreadyIncludesStyle ? baseFamily : `${baseFamily} ${styleWithoutPosture}`,
    bold: false,
    italic
  };
}

function normalizeRotation(value: number): number {
  return ((value % 360) + 360) % 360;
}

function textAlign(value: TextLayer["align"]): "left" | "center" | "right" | "justify" {
  return value === "LEFT"
    ? "left"
    : value === "CENTER"
      ? "center"
      : value === "RIGHT"
        ? "right"
        : "justify";
}

function verticalAlign(value: TextLayer["verticalAlign"]): "top" | "middle" | "bottom" {
  return value === "TOP" ? "top" : value === "CENTER" ? "middle" : "bottom";
}

function dashType(
  dash: number[] | undefined
): "solid" | "dash" {
  return dash?.length ? "dash" : "solid";
}

export const conversionInternals = {
  rgbHex,
  opacityToTransparency,
  normalizeRotation,
  pixelsToPoints,
  resolvePowerPointFont
};
