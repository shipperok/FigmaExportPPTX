import { createPresentation } from "./pptx";
import { t, type Locale, type TranslationKey } from "./i18n";
import type {
  ExportOptions,
  PluginToUiMessage,
  SelectionSummary,
  UiToPluginMessage
} from "./model";

const frameList = requiredElement<HTMLDivElement>("frame-list");
const emptyState = requiredElement<HTMLDivElement>("empty-state");
const exportButton = requiredElement<HTMLButtonElement>("export-button");
const buttonLabel = requiredElement<HTMLSpanElement>("button-label");
const status = requiredElement<HTMLDivElement>("status");
const progressBar = requiredElement<HTMLDivElement>("progress-bar");
const scaleSelect = requiredElement<HTMLSelectElement>("raster-scale");
const includeHidden = requiredElement<HTMLInputElement>("include-hidden");
const addNotes = requiredElement<HTMLInputElement>("add-notes");
const warningBox = requiredElement<HTMLDivElement>("warning-box");

let frames: SelectionSummary = [];
let exporting = false;
let locale: Locale = "ru";
let statusState: { key: TranslationKey; values?: Record<string, string | number> } = { key: "ready" };

window.onmessage = async (event: MessageEvent<{ pluginMessage?: PluginToUiMessage }>) => {
  const message = event.data.pluginMessage;
  if (!message) return;

  if (message.type === "language") {
    setLocale(message.locale);
  } else if (message.type === "selection") {
    frames = message.frames;
    renderFrames();
  } else if (message.type === "export-start") {
    setExporting(true);
    setStatus("preparing");
    setProgress(0);
  } else if (message.type === "export-progress") {
    setStatus("progress", { current: message.current, total: message.total, name: message.name });
    setProgress((message.current / message.total) * 70);
  } else if (message.type === "export-data") {
    try {
      setStatus("building");
      setProgress(82);
      const result = await createPresentation(
        message.slides,
        message.fileName,
        message.options
      );
      setProgress(100);
      setStatus("done");
      showWarnings(result.warnings);
    } catch (error) {
      showError(error instanceof Error ? error.message : String(error));
    } finally {
      setExporting(false);
    }
  } else if (message.type === "error") {
    showError(message.message);
    setExporting(false);
  }
};

exportButton.addEventListener("click", () => {
  if (!frames.length || exporting) return;
  warningBox.hidden = true;
  const options: ExportOptions = {
    rasterScale: Number(scaleSelect.value) as 1 | 2 | 3,
    includeHidden: includeHidden.checked,
    addSpeakerNotes: addNotes.checked,
    locale
  };
  send({ type: "export", options });
});

document.querySelectorAll<HTMLButtonElement>("[data-refresh]").forEach((button) => {
  button.addEventListener("click", () => send({ type: "refresh-selection" }));
});

document.querySelectorAll<HTMLButtonElement>("[data-language]").forEach((button) => {
  button.addEventListener("click", () => {
    const nextLocale = button.dataset.language;
    if (nextLocale !== "ru" && nextLocale !== "en") return;
    setLocale(nextLocale);
    send({ type: "set-language", locale: nextLocale });
  });
});

function renderFrames(): void {
  frameList.replaceChildren();
  emptyState.hidden = frames.length > 0;
  frameList.hidden = frames.length === 0;
  exportButton.disabled = frames.length === 0 || exporting;

  for (const [index, frame] of frames.entries()) {
    const row = document.createElement("div");
    row.className = "frame-row";

    const number = document.createElement("span");
    number.className = "frame-number";
    number.textContent = String(index + 1);

    const copy = document.createElement("div");
    copy.className = "frame-copy";
    const name = document.createElement("strong");
    name.textContent = frame.name;
    const size = document.createElement("span");
    size.textContent = `${Math.round(frame.width)} × ${Math.round(frame.height)} px`;
    copy.append(name, size);
    row.append(number, copy);
    frameList.append(row);
  }

  updateButtonLabel();
}

function setExporting(value: boolean): void {
  exporting = value;
  exportButton.disabled = value || !frames.length;
  exportButton.classList.toggle("is-loading", value);
  scaleSelect.disabled = value;
  includeHidden.disabled = value;
  addNotes.disabled = value;
  updateButtonLabel();
}

function setProgress(value: number): void {
  progressBar.style.width = `${Math.max(0, Math.min(100, value))}%`;
}

function showWarnings(warnings: string[]): void {
  if (!warnings.length) {
    warningBox.hidden = true;
    return;
  }
  warningBox.hidden = false;
  warningBox.textContent = warnings.join(" ");
}

function showError(message: string): void {
  status.textContent = message;
  status.classList.add("is-error");
  setTimeout(() => status.classList.remove("is-error"), 4000);
}

function setLocale(nextLocale: Locale): void {
  locale = nextLocale;
  document.documentElement.lang = locale;
  document.querySelectorAll<HTMLElement>("[data-i18n]").forEach((element) => {
    const key = element.dataset.i18n as TranslationKey | undefined;
    if (key) element.textContent = t(locale, key);
  });
  document.querySelectorAll<HTMLButtonElement>("[data-language]").forEach((button) => {
    const active = button.dataset.language === locale;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-pressed", String(active));
  });
  setStatus(statusState.key, statusState.values);
  renderFrames();
}

function setStatus(key: TranslationKey, values?: Record<string, string | number>): void {
  statusState = { key, values };
  status.textContent = t(locale, key, values);
}

function updateButtonLabel(): void {
  if (exporting) {
    buttonLabel.textContent = t(locale, "exporting");
  } else if (!frames.length) {
    buttonLabel.textContent = t(locale, "selectFrames");
  } else if (locale === "en") {
    buttonLabel.textContent = `Export ${frames.length} slide${frames.length === 1 ? "" : "s"}`;
  } else {
    buttonLabel.textContent = `Экспортировать ${pluralize(frames.length, ["слайд", "слайда", "слайдов"])}`;
  }
}

function send(message: UiToPluginMessage): void {
  parent.postMessage({ pluginMessage: message }, "*");
}

function requiredElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing #${id}`);
  return element as T;
}

function pluralize(value: number, forms: [string, string, string]): string {
  const mod100 = value % 100;
  const mod10 = value % 10;
  const form =
    mod100 >= 11 && mod100 <= 19 ? forms[2] : mod10 === 1 ? forms[0] : mod10 >= 2 && mod10 <= 4 ? forms[1] : forms[2];
  return `${value} ${form}`;
}

send({ type: "ready" });
