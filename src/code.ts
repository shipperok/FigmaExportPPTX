/// <reference types="@figma/plugin-typings" />

import type {
  ExportLayer,
  ExportOptions,
  ExportSlide,
  MediaLayer,
  PluginToUiMessage,
  Rgb,
  SelectionSummary,
  ShapeLayer,
  TextLayer,
  TextRun,
  UiToPluginMessage
} from "./model";

figma.showUI(__html__, { width: 400, height: 590, themeColors: true });

const post = (message: PluginToUiMessage) => figma.ui.postMessage(message);

function selectedFrames(): SceneNode[] {
  return figma.currentPage.selection.filter(
    (node) =>
      node.type === "FRAME" ||
      node.type === "COMPONENT" ||
      node.type === "INSTANCE" ||
      node.type === "SECTION"
  );
}

function sendSelection(): void {
  const frames: SelectionSummary = selectedFrames().map((node) => ({
    id: node.id,
    name: node.name,
    width: "width" in node ? node.width : 0,
    height: "height" in node ? node.height : 0,
    type: node.type
  }));
  post({ type: "selection", frames });
}

figma.on("selectionchange", sendSelection);

figma.ui.onmessage = async (message: UiToPluginMessage) => {
  if (message.type === "ready" || message.type === "refresh-selection") {
    sendSelection();
    return;
  }
  if (message.type === "resize") {
    figma.ui.resize(400, Math.max(480, Math.min(720, Math.round(message.height))));
    return;
  }
  if (message.type === "export") {
    try {
      await exportSelection(message.options);
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      post({ type: "error", message: `Не удалось подготовить презентацию: ${detail}` });
    }
  }
};

async function exportSelection(options: ExportOptions): Promise<void> {
  const frames = selectedFrames();
  if (!frames.length) {
    post({ type: "error", message: "Выберите хотя бы один фрейм, компонент или секцию." });
    return;
  }

  post({ type: "export-start", total: frames.length });
  const slides: ExportSlide[] = [];

  for (let index = 0; index < frames.length; index += 1) {
    const frame = frames[index];
    post({
      type: "export-progress",
      current: index + 1,
      total: frames.length,
      name: frame.name
    });
    slides.push(await serializeSlide(frame, options));
  }

  const pageName = safeFileName(figma.currentPage.name || figma.root.name || "Figma export");
  post({
    type: "export-data",
    slides,
    fileName: `${pageName}.pptx`,
    options
  });
};

async function serializeSlide(root: SceneNode, options: ExportOptions): Promise<ExportSlide> {
  if (!("absoluteBoundingBox" in root) || !root.absoluteBoundingBox) {
    throw new Error(`У фрейма «${root.name}» нет размеров.`);
  }
  const box = root.absoluteBoundingBox;
  const layers: ExportLayer[] = [];
  const warnings = new Set<string>();
  const background = solidFill(root);

  if ("children" in root) {
    for (const child of root.children) {
      await serializeNode(child, root, box.x, box.y, options, layers, warnings);
    }
  } else {
    layers.push(await exportMedia(root, box.x, box.y, options.rasterScale, "image/png"));
  }

  return {
    id: root.id,
    name: root.name,
    width: box.width,
    height: box.height,
    background,
    layers,
    warnings: [...warnings]
  };
}

async function serializeNode(
  node: SceneNode,
  slideRoot: SceneNode,
  originX: number,
  originY: number,
  options: ExportOptions,
  output: ExportLayer[],
  warnings: Set<string>
): Promise<void> {
  if (!options.includeHidden && !isVisibleThroughTree(node, slideRoot)) return;
  const box = node.absoluteBoundingBox;
  if (!box || box.width <= 0 || box.height <= 0) return;

  if (node.type === "TEXT") {
    const text = serializeText(node, originX, originY);
    if (text) {
      output.push(text);
    } else {
      output.push(await exportMedia(node, originX, originY, options.rasterScale, "image/svg+xml"));
      warnings.add("Текст с градиентом или изображением сохранён как SVG.");
    }
    return;
  }

  if (canExportNativeShape(node)) {
    output.push(serializeShape(node, originX, originY));
    return;
  }

  if (isContainer(node)) {
    if (requiresFlattening(node)) {
      output.push(await exportMedia(node, originX, originY, options.rasterScale, "image/png"));
      warnings.add("Маски и обрезанные группы сохранены как отдельные PNG-объекты.");
      return;
    }

    if (hasVisiblePaint(node)) {
      if (canExportNativeContainerBackground(node)) {
        output.push(serializeShape(node, originX, originY));
      } else {
        output.push(await exportOwnBackground(node, originX, originY, options.rasterScale));
        warnings.add("Сложные заливки сохранены как отдельные SVG-объекты.");
      }
    }

    for (const child of node.children) {
      await serializeNode(child, slideRoot, originX, originY, options, output, warnings);
    }
    return;
  }

  const vectorLike =
    node.type === "VECTOR" ||
    node.type === "BOOLEAN_OPERATION" ||
    node.type === "STAR" ||
    node.type === "POLYGON";
  output.push(
    await exportMedia(
      node,
      originX,
      originY,
      options.rasterScale,
      vectorLike ? "image/svg+xml" : "image/png"
    )
  );
  warnings.add(
    vectorLike
      ? "Сложные контуры сохранены как отдельные SVG-объекты."
      : "Неподдерживаемые слои сохранены как отдельные PNG-объекты."
  );
}

function serializeText(node: TextNode, originX: number, originY: number): TextLayer | null {
  const box = node.absoluteBoundingBox;
  if (!box) return null;
  const paints = node.fills === figma.mixed ? [] : visiblePaints(node.fills);
  if (paints.some((paint) => paint.type !== "SOLID")) return null;

  const segments = node.getStyledTextSegments([
    "fontName",
    "fontSize",
    "fills",
    "letterSpacing",
    "textDecoration"
  ]);
  const runs: TextRun[] = segments.map((segment) => {
    const fillPaints = visiblePaints(segment.fills);
    const fill = fillPaints.find((paint): paint is SolidPaint => paint.type === "SOLID");
    const fontName = segment.fontName;
    const fontSize = segment.fontSize;
    return {
      text: segment.characters,
      fontFamily: fontName.family,
      fontStyle: fontName.style,
      fontSize,
      color: fill ? rgba(fill.color, fill.opacity) : { r: 0, g: 0, b: 0, a: 1 },
      opacity: fill?.opacity ?? 1,
      letterSpacing: toPixels(segment.letterSpacing, fontSize),
      textDecoration: segment.textDecoration
    };
  });

  const lineHeight =
    node.lineHeight === figma.mixed || node.lineHeight.unit === "AUTO"
      ? undefined
      : node.lineHeight.unit === "PIXELS"
        ? node.lineHeight.value
        : undefined;

  return {
    ...baseLayer(node, originX, originY),
    kind: "text",
    runs,
    align: node.textAlignHorizontal,
    verticalAlign: node.textAlignVertical,
    lineHeight,
    paragraphSpacing: node.paragraphSpacing === figma.mixed ? undefined : node.paragraphSpacing
  };
}

function serializeShape(
  node: RectangleNode | EllipseNode | LineNode | FrameNode | ComponentNode | InstanceNode,
  originX: number,
  originY: number
): ShapeLayer {
  const fill = solidFill(node);
  const stroke = solidStroke(node);
  const radius =
    node.type === "RECTANGLE" || node.type === "FRAME" || node.type === "COMPONENT" || node.type === "INSTANCE"
      ? node.cornerRadius === figma.mixed
        ? 0
        : node.cornerRadius
      : undefined;
  return {
    ...baseLayer(node, originX, originY),
    kind: "shape",
    shape: node.type === "ELLIPSE" ? "ellipse" : node.type === "LINE" ? "line" : "rect",
    fill,
    stroke,
    strokeWidth: stroke && "strokeWeight" in node && node.strokeWeight !== figma.mixed ? node.strokeWeight : 0,
    radius,
    dash: "dashPattern" in node ? [...node.dashPattern] : undefined
  };
}

async function exportOwnBackground(
  node: FrameNode | ComponentNode | InstanceNode | GroupNode,
  originX: number,
  originY: number,
  rasterScale: number
): Promise<MediaLayer> {
  if (node.type === "GROUP") {
    return exportMedia(node, originX, originY, rasterScale, "image/svg+xml");
  }
  const clone = node.clone();
  try {
    for (const child of [...clone.children]) child.remove();
    return await exportMedia(clone, originX, originY, rasterScale, "image/svg+xml", node);
  } finally {
    clone.remove();
  }
}

async function exportMedia(
  node: SceneNode,
  originX: number,
  originY: number,
  rasterScale: number,
  mime: "image/png" | "image/svg+xml",
  geometrySource?: SceneNode
): Promise<MediaLayer> {
  const source = geometrySource ?? node;
  const box = source.absoluteBoundingBox;
  if (!box) throw new Error(`Не удалось определить границы слоя «${source.name}».`);
  const settings: ExportSettings =
    mime === "image/svg+xml"
      ? { format: "SVG", svgOutlineText: false, svgIdAttribute: false }
      : { format: "PNG", constraint: { type: "SCALE", value: rasterScale } };
  const bytes = await node.exportAsync(settings);
  return {
    ...baseLayerFromBox(source, box, originX, originY),
    kind: "media",
    mime,
    base64: bytesToBase64(bytes)
  };
}

function baseLayer(node: SceneNode, originX: number, originY: number) {
  const box = node.absoluteBoundingBox;
  if (!box) throw new Error(`Не удалось определить границы слоя «${node.name}».`);
  return baseLayerFromBox(node, box, originX, originY);
}

function baseLayerFromBox(
  node: SceneNode,
  box: Rect,
  originX: number,
  originY: number
) {
  return {
    id: node.id,
    name: node.name,
    x: box.x - originX,
    y: box.y - originY,
    width: box.width,
    height: box.height,
    rotation: "rotation" in node ? -node.rotation : 0,
    opacity: "opacity" in node ? node.opacity : 1
  };
}

function canExportNativeShape(
  node: SceneNode
): node is RectangleNode | EllipseNode | LineNode {
  if (node.type !== "RECTANGLE" && node.type !== "ELLIPSE" && node.type !== "LINE") return false;
  return paintsAreNative(node.fills) && paintsAreNative(node.strokes) && !hasUnsupportedEffects(node);
}

function canExportNativeContainerBackground(
  node: FrameNode | ComponentNode | InstanceNode | GroupNode
): node is FrameNode | ComponentNode | InstanceNode {
  return (
    node.type !== "GROUP" &&
    paintsAreNative(node.fills) &&
    paintsAreNative(node.strokes) &&
    !hasUnsupportedEffects(node)
  );
}

function isContainer(
  node: SceneNode
): node is FrameNode | ComponentNode | InstanceNode | GroupNode {
  return (
    node.type === "FRAME" ||
    node.type === "COMPONENT" ||
    node.type === "INSTANCE" ||
    node.type === "GROUP"
  );
}

function requiresFlattening(
  node: FrameNode | ComponentNode | InstanceNode | GroupNode
): boolean {
  if ("clipsContent" in node && node.clipsContent) return true;
  return node.children.some((child) => "isMask" in child && child.isMask);
}

function hasUnsupportedEffects(node: SceneNode): boolean {
  return "effects" in node && node.effects.some((effect) => effect.visible !== false);
}

function hasVisiblePaint(node: FrameNode | ComponentNode | InstanceNode | GroupNode): boolean {
  return (
    node.type !== "GROUP" &&
    ((node.fills !== figma.mixed && visiblePaints(node.fills).length > 0) ||
      visiblePaints(node.strokes).length > 0)
  );
}

function paintsAreNative(paints: readonly Paint[] | PluginAPI["mixed"]): boolean {
  if (paints === figma.mixed) return false;
  const visible = visiblePaints(paints);
  return visible.length <= 1 && visible.every((paint) => paint.type === "SOLID");
}

function visiblePaints(paints: readonly Paint[]): Paint[] {
  return paints.filter((paint) => paint.visible !== false);
}

function solidFill(node: SceneNode): Rgb | undefined {
  if (!("fills" in node) || node.fills === figma.mixed) return undefined;
  const visible = visiblePaints(node.fills);
  if (visible.length !== 1 || visible[0].type !== "SOLID") return undefined;
  return rgba(visible[0].color, visible[0].opacity);
}

function solidStroke(node: SceneNode): Rgb | undefined {
  if (!("strokes" in node)) return undefined;
  const visible = visiblePaints(node.strokes);
  if (visible.length !== 1 || visible[0].type !== "SOLID") return undefined;
  return rgba(visible[0].color, visible[0].opacity);
}

function rgba(color: RGB, opacity = 1): Rgb {
  return { r: color.r, g: color.g, b: color.b, a: opacity };
}

function toPixels(spacing: LetterSpacing | PluginAPI["mixed"], fontSize: number): number | undefined {
  if (spacing === figma.mixed) return undefined;
  return spacing.unit === "PIXELS" ? spacing.value : (spacing.value / 100) * fontSize;
}

function isVisibleThroughTree(node: SceneNode, stopAt: SceneNode): boolean {
  let current: BaseNode | null = node;
  while (current && current !== stopAt) {
    if ("visible" in current && !current.visible) return false;
    current = current.parent;
  }
  return true;
}

function bytesToBase64(bytes: Uint8Array): string {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  let result = "";
  for (let index = 0; index < bytes.length; index += 3) {
    const a = bytes[index];
    const b = index + 1 < bytes.length ? bytes[index + 1] : 0;
    const c = index + 2 < bytes.length ? bytes[index + 2] : 0;
    const chunk = (a << 16) | (b << 8) | c;
    result += alphabet[(chunk >> 18) & 63];
    result += alphabet[(chunk >> 12) & 63];
    result += index + 1 < bytes.length ? alphabet[(chunk >> 6) & 63] : "=";
    result += index + 2 < bytes.length ? alphabet[chunk & 63] : "=";
  }
  return result;
}

function safeFileName(value: string): string {
  const cleaned = value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, " ").replace(/\s+/g, " ").trim();
  return cleaned || "Figma export";
}
