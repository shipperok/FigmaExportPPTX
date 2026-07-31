export type Rgb = { r: number; g: number; b: number; a?: number };
import type { Locale } from "./i18n";

export type ExportOptions = {
  rasterScale: 1 | 2 | 3;
  includeHidden: boolean;
  addSpeakerNotes: boolean;
  locale: Locale;
};

export type SelectionSummary = {
  id: string;
  name: string;
  width: number;
  height: number;
  type: string;
}[];

export type TextRun = {
  text: string;
  fontFamily: string;
  fontStyle: string;
  fontWeight?: number;
  italic?: boolean;
  fontSize: number;
  color: Rgb;
  opacity: number;
  letterSpacing?: number;
  textDecoration?: "NONE" | "UNDERLINE" | "STRIKETHROUGH";
};

export type BaseLayer = {
  id: string;
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  opacity: number;
};

export type ShapeLayer = BaseLayer & {
  kind: "shape";
  shape: "rect" | "ellipse" | "line";
  fill?: Rgb;
  stroke?: Rgb;
  strokeWidth?: number;
  radius?: number;
  dash?: number[];
};

export type TextLayer = BaseLayer & {
  kind: "text";
  runs: TextRun[];
  align: "LEFT" | "CENTER" | "RIGHT" | "JUSTIFIED";
  verticalAlign: "TOP" | "CENTER" | "BOTTOM";
  lineHeight?: number;
  paragraphSpacing?: number;
};

export type MediaLayer = BaseLayer & {
  kind: "media";
  mime: "image/png" | "image/svg+xml";
  base64: string;
};

export type ExportLayer = ShapeLayer | TextLayer | MediaLayer;

export type ExportSlide = {
  id: string;
  name: string;
  width: number;
  height: number;
  background?: Rgb;
  layers: ExportLayer[];
  warnings: string[];
};

export type PluginToUiMessage =
  | { type: "language"; locale: Locale }
  | { type: "selection"; frames: SelectionSummary }
  | { type: "export-start"; total: number }
  | { type: "export-progress"; current: number; total: number; name: string }
  | { type: "export-data"; slides: ExportSlide[]; fileName: string; options: ExportOptions }
  | { type: "error"; message: string };

export type UiToPluginMessage =
  | { type: "ready" }
  | { type: "set-language"; locale: Locale }
  | { type: "refresh-selection" }
  | { type: "export"; options: ExportOptions }
  | { type: "resize"; height: number };
