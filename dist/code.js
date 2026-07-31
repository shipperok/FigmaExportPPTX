"use strict";
(() => {
  // src/order.ts
  function sortBySlideOrder(items) {
    const numbered = items.map((item) => ({ item, number: numericPrefix(item.name) }));
    const numbers = numbered.map(({ number }) => number);
    if (items.length > 1 && numbers.every((number) => number !== null) && new Set(numbers).size === numbers.length) {
      return numbered.sort((a, b) => a.number - b.number || naturalName(a.item, b.item)).map(({ item }) => item);
    }
    const remaining = [...items].sort(
      (a, b) => a.y - b.y || a.x - b.x || naturalName(a, b)
    );
    const result = [];
    while (remaining.length) {
      const first = remaining.shift();
      const row = [first];
      const tolerance = Math.max(8, first.height * 0.25);
      for (let index = 0; index < remaining.length; ) {
        const candidate = remaining[index];
        const candidateTolerance = Math.max(8, candidate.height * 0.25);
        if (Math.abs(candidate.y - first.y) <= Math.min(tolerance, candidateTolerance)) {
          row.push(candidate);
          remaining.splice(index, 1);
        } else {
          index += 1;
        }
      }
      row.sort((a, b) => a.x - b.x || a.y - b.y || naturalName(a, b));
      result.push(...row);
    }
    return result;
  }
  function numericPrefix(name) {
    const match = name.match(/^\s*(\d+)(?:\D|$)/);
    return match ? Number(match[1]) : null;
  }
  function naturalName(a, b) {
    return a.name.localeCompare(b.name, void 0, { numeric: true, sensitivity: "base" });
  }

  // src/code.ts
  figma.showUI(__html__, { width: 400, height: 590, themeColors: true });
  var post = (message) => figma.ui.postMessage(message);
  function selectedFrames() {
    const selected = figma.currentPage.selection.filter(
      (node) => node.type === "FRAME" || node.type === "COMPONENT" || node.type === "INSTANCE" || node.type === "SECTION"
    );
    return sortBySlideOrder(
      selected.map((node) => {
        const box = node.absoluteBoundingBox;
        return {
          node,
          name: node.name,
          x: box?.x ?? ("x" in node ? node.x : 0),
          y: box?.y ?? ("y" in node ? node.y : 0),
          width: box?.width ?? ("width" in node ? node.width : 0),
          height: box?.height ?? ("height" in node ? node.height : 0)
        };
      })
    ).map(({ node }) => node);
  }
  function sendSelection() {
    const frames = selectedFrames().map((node) => ({
      id: node.id,
      name: node.name,
      width: "width" in node ? node.width : 0,
      height: "height" in node ? node.height : 0,
      type: node.type
    }));
    post({ type: "selection", frames });
  }
  figma.on("selectionchange", sendSelection);
  figma.ui.onmessage = async (message) => {
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
        post({ type: "error", message: `\u041D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u043F\u043E\u0434\u0433\u043E\u0442\u043E\u0432\u0438\u0442\u044C \u043F\u0440\u0435\u0437\u0435\u043D\u0442\u0430\u0446\u0438\u044E: ${detail}` });
      }
    }
  };
  async function exportSelection(options) {
    const frames = selectedFrames();
    if (!frames.length) {
      post({ type: "error", message: "\u0412\u044B\u0431\u0435\u0440\u0438\u0442\u0435 \u0445\u043E\u0442\u044F \u0431\u044B \u043E\u0434\u0438\u043D \u0444\u0440\u0435\u0439\u043C, \u043A\u043E\u043C\u043F\u043E\u043D\u0435\u043D\u0442 \u0438\u043B\u0438 \u0441\u0435\u043A\u0446\u0438\u044E." });
      return;
    }
    post({ type: "export-start", total: frames.length });
    const slides = [];
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
  }
  async function serializeSlide(root, options) {
    if (!("absoluteBoundingBox" in root) || !root.absoluteBoundingBox) {
      throw new Error(`\u0423 \u0444\u0440\u0435\u0439\u043C\u0430 \xAB${root.name}\xBB \u043D\u0435\u0442 \u0440\u0430\u0437\u043C\u0435\u0440\u043E\u0432.`);
    }
    const box = root.absoluteBoundingBox;
    const layers = [];
    const warnings = /* @__PURE__ */ new Set();
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
  async function serializeNode(node, slideRoot, originX, originY, options, output, warnings) {
    if (!options.includeHidden && !isVisibleThroughTree(node, slideRoot)) return;
    const box = node.absoluteBoundingBox;
    if (!box || box.width <= 0 || box.height <= 0) return;
    if (node.type === "TEXT") {
      const text = serializeText(node, originX, originY);
      if (text) {
        output.push(text);
      } else {
        output.push(await exportMedia(node, originX, originY, options.rasterScale, "image/svg+xml"));
        warnings.add("\u0422\u0435\u043A\u0441\u0442 \u0441 \u0433\u0440\u0430\u0434\u0438\u0435\u043D\u0442\u043E\u043C \u0438\u043B\u0438 \u0438\u0437\u043E\u0431\u0440\u0430\u0436\u0435\u043D\u0438\u0435\u043C \u0441\u043E\u0445\u0440\u0430\u043D\u0451\u043D \u043A\u0430\u043A SVG.");
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
        warnings.add("\u041C\u0430\u0441\u043A\u0438 \u0438 \u043E\u0431\u0440\u0435\u0437\u0430\u043D\u043D\u044B\u0435 \u0433\u0440\u0443\u043F\u043F\u044B \u0441\u043E\u0445\u0440\u0430\u043D\u0435\u043D\u044B \u043A\u0430\u043A \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u044B\u0435 PNG-\u043E\u0431\u044A\u0435\u043A\u0442\u044B.");
        return;
      }
      if (hasVisiblePaint(node)) {
        if (canExportNativeContainerBackground(node)) {
          output.push(serializeShape(node, originX, originY));
        } else {
          output.push(await exportOwnBackground(node, originX, originY, options.rasterScale));
          warnings.add("\u0421\u043B\u043E\u0436\u043D\u044B\u0435 \u0437\u0430\u043B\u0438\u0432\u043A\u0438 \u0441\u043E\u0445\u0440\u0430\u043D\u0435\u043D\u044B \u043A\u0430\u043A \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u044B\u0435 SVG-\u043E\u0431\u044A\u0435\u043A\u0442\u044B.");
        }
      }
      for (const child of node.children) {
        await serializeNode(child, slideRoot, originX, originY, options, output, warnings);
      }
      return;
    }
    const vectorLike = node.type === "VECTOR" || node.type === "BOOLEAN_OPERATION" || node.type === "STAR" || node.type === "POLYGON";
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
      vectorLike ? "\u0421\u043B\u043E\u0436\u043D\u044B\u0435 \u043A\u043E\u043D\u0442\u0443\u0440\u044B \u0441\u043E\u0445\u0440\u0430\u043D\u0435\u043D\u044B \u043A\u0430\u043A \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u044B\u0435 SVG-\u043E\u0431\u044A\u0435\u043A\u0442\u044B." : "\u041D\u0435\u043F\u043E\u0434\u0434\u0435\u0440\u0436\u0438\u0432\u0430\u0435\u043C\u044B\u0435 \u0441\u043B\u043E\u0438 \u0441\u043E\u0445\u0440\u0430\u043D\u0435\u043D\u044B \u043A\u0430\u043A \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u044B\u0435 PNG-\u043E\u0431\u044A\u0435\u043A\u0442\u044B."
    );
  }
  function serializeText(node, originX, originY) {
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
    const runs = segments.map((segment) => {
      const fillPaints = visiblePaints(segment.fills);
      const fill = fillPaints.find((paint) => paint.type === "SOLID");
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
    const lineHeight = node.lineHeight === figma.mixed || node.lineHeight.unit === "AUTO" ? void 0 : node.lineHeight.unit === "PIXELS" ? node.lineHeight.value : void 0;
    return {
      ...baseLayer(node, originX, originY),
      kind: "text",
      runs,
      align: node.textAlignHorizontal,
      verticalAlign: node.textAlignVertical,
      lineHeight,
      paragraphSpacing: node.paragraphSpacing === figma.mixed ? void 0 : node.paragraphSpacing
    };
  }
  function serializeShape(node, originX, originY) {
    const fill = solidFill(node);
    const stroke = solidStroke(node);
    const radius = node.type === "RECTANGLE" || node.type === "FRAME" || node.type === "COMPONENT" || node.type === "INSTANCE" ? node.cornerRadius === figma.mixed ? 0 : node.cornerRadius : void 0;
    return {
      ...baseLayer(node, originX, originY),
      kind: "shape",
      shape: node.type === "ELLIPSE" ? "ellipse" : node.type === "LINE" ? "line" : "rect",
      fill,
      stroke,
      strokeWidth: stroke && "strokeWeight" in node && node.strokeWeight !== figma.mixed ? node.strokeWeight : 0,
      radius,
      dash: "dashPattern" in node ? [...node.dashPattern] : void 0
    };
  }
  async function exportOwnBackground(node, originX, originY, rasterScale) {
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
  async function exportMedia(node, originX, originY, rasterScale, mime, geometrySource) {
    const source = geometrySource ?? node;
    const box = ("absoluteRenderBounds" in source ? source.absoluteRenderBounds : null) ?? source.absoluteBoundingBox;
    if (!box) throw new Error(`\u041D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u043E\u043F\u0440\u0435\u0434\u0435\u043B\u0438\u0442\u044C \u0433\u0440\u0430\u043D\u0438\u0446\u044B \u0441\u043B\u043E\u044F \xAB${source.name}\xBB.`);
    const settings = mime === "image/svg+xml" ? { format: "SVG", svgOutlineText: false, svgIdAttribute: false } : { format: "PNG", constraint: { type: "SCALE", value: rasterScale } };
    const bytes = await node.exportAsync(settings);
    const layer = {
      ...baseLayerFromBox(source, box, originX, originY),
      kind: "media",
      mime,
      base64: bytesToBase64(bytes)
    };
    layer.rotation = 0;
    return layer;
  }
  function baseLayer(node, originX, originY) {
    const box = node.absoluteBoundingBox;
    if (!box) throw new Error(`\u041D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u043E\u043F\u0440\u0435\u0434\u0435\u043B\u0438\u0442\u044C \u0433\u0440\u0430\u043D\u0438\u0446\u044B \u0441\u043B\u043E\u044F \xAB${node.name}\xBB.`);
    return baseLayerFromBox(node, box, originX, originY);
  }
  function baseLayerFromBox(node, box, originX, originY) {
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
  function canExportNativeShape(node) {
    if (node.type !== "RECTANGLE" && node.type !== "ELLIPSE" && node.type !== "LINE") return false;
    return paintsAreNative(node.fills) && paintsAreNative(node.strokes) && !hasUnsupportedEffects(node);
  }
  function canExportNativeContainerBackground(node) {
    return node.type !== "GROUP" && paintsAreNative(node.fills) && paintsAreNative(node.strokes) && !hasUnsupportedEffects(node);
  }
  function isContainer(node) {
    return node.type === "FRAME" || node.type === "COMPONENT" || node.type === "INSTANCE" || node.type === "GROUP";
  }
  function requiresFlattening(node) {
    if ("clipsContent" in node && node.clipsContent) return true;
    return node.children.some((child) => "isMask" in child && child.isMask);
  }
  function hasUnsupportedEffects(node) {
    return "effects" in node && node.effects.some((effect) => effect.visible !== false);
  }
  function hasVisiblePaint(node) {
    return node.type !== "GROUP" && (node.fills !== figma.mixed && visiblePaints(node.fills).length > 0 || visiblePaints(node.strokes).length > 0);
  }
  function paintsAreNative(paints) {
    if (paints === figma.mixed) return false;
    const visible = visiblePaints(paints);
    return visible.length <= 1 && visible.every((paint) => paint.type === "SOLID");
  }
  function visiblePaints(paints) {
    return paints.filter((paint) => paint.visible !== false);
  }
  function solidFill(node) {
    if (!("fills" in node) || node.fills === figma.mixed) return void 0;
    const visible = visiblePaints(node.fills);
    if (visible.length !== 1 || visible[0].type !== "SOLID") return void 0;
    return rgba(visible[0].color, visible[0].opacity);
  }
  function solidStroke(node) {
    if (!("strokes" in node)) return void 0;
    const visible = visiblePaints(node.strokes);
    if (visible.length !== 1 || visible[0].type !== "SOLID") return void 0;
    return rgba(visible[0].color, visible[0].opacity);
  }
  function rgba(color, opacity = 1) {
    return { r: color.r, g: color.g, b: color.b, a: opacity };
  }
  function toPixels(spacing, fontSize) {
    if (spacing === figma.mixed) return void 0;
    return spacing.unit === "PIXELS" ? spacing.value : spacing.value / 100 * fontSize;
  }
  function isVisibleThroughTree(node, stopAt) {
    let current = node;
    while (current && current !== stopAt) {
      if ("visible" in current && !current.visible) return false;
      current = current.parent;
    }
    return true;
  }
  function bytesToBase64(bytes) {
    const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let result = "";
    for (let index = 0; index < bytes.length; index += 3) {
      const a = bytes[index];
      const b = index + 1 < bytes.length ? bytes[index + 1] : 0;
      const c = index + 2 < bytes.length ? bytes[index + 2] : 0;
      const chunk = a << 16 | b << 8 | c;
      result += alphabet[chunk >> 18 & 63];
      result += alphabet[chunk >> 12 & 63];
      result += index + 1 < bytes.length ? alphabet[chunk >> 6 & 63] : "=";
      result += index + 2 < bytes.length ? alphabet[chunk & 63] : "=";
    }
    return result;
  }
  function safeFileName(value) {
    const cleaned = value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, " ").replace(/\s+/g, " ").trim();
    return cleaned || "Figma export";
  }
})();
//# sourceMappingURL=code.js.map
