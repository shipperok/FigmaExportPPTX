import { createPresentation } from "./pptx";
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

window.onmessage = async (event: MessageEvent<{ pluginMessage?: PluginToUiMessage }>) => {
  const message = event.data.pluginMessage;
  if (!message) return;

  if (message.type === "selection") {
    frames = message.frames;
    renderFrames();
  } else if (message.type === "export-start") {
    setExporting(true);
    status.textContent = "Подготавливаю слои…";
    setProgress(0);
  } else if (message.type === "export-progress") {
    status.textContent = `${message.current} из ${message.total}: ${message.name}`;
    setProgress((message.current / message.total) * 70);
  } else if (message.type === "export-data") {
    try {
      status.textContent = "Собираю файл PowerPoint…";
      setProgress(82);
      const result = await createPresentation(
        message.slides,
        message.fileName,
        message.options
      );
      setProgress(100);
      status.textContent = "Готово — файл сохранён в загрузки.";
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
    addSpeakerNotes: addNotes.checked
  };
  send({ type: "export", options });
});

document.querySelectorAll<HTMLButtonElement>("[data-refresh]").forEach((button) => {
  button.addEventListener("click", () => send({ type: "refresh-selection" }));
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

  buttonLabel.textContent = frames.length
    ? `Экспортировать ${pluralize(frames.length, ["слайд", "слайда", "слайдов"])}`
    : "Выберите фреймы";
}

function setExporting(value: boolean): void {
  exporting = value;
  exportButton.disabled = value || !frames.length;
  exportButton.classList.toggle("is-loading", value);
  scaleSelect.disabled = value;
  includeHidden.disabled = value;
  addNotes.disabled = value;
  if (!value) buttonLabel.textContent = frames.length
    ? `Экспортировать ${pluralize(frames.length, ["слайд", "слайда", "слайдов"])}`
    : "Выберите фреймы";
  else buttonLabel.textContent = "Экспортирую…";
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
