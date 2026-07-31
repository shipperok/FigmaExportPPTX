export type Locale = "ru" | "en";

const translations = {
  ru: {
    eyebrow: "Редактируемый экспорт",
    title: "Фреймы → PowerPoint",
    subtitle: "Текст и простые фигуры останутся редактируемыми. Сложная графика сохранится отдельными объектами.",
    selectedSlides: "Выбранные слайды",
    autoOrder: "Автопорядок",
    refresh: "Обновить",
    emptyTitle: "Выберите фреймы на холсте",
    emptyText: "Номера в названиях задают порядок. Без номеров — сверху вниз и слева направо.",
    settings: "Настройки",
    pngQuality: "Качество PNG",
    pngHint: "Для теней, масок и эффектов",
    hiddenLayers: "Скрытые слои",
    hiddenHint: "Добавить в презентацию",
    notes: "Примечания",
    notesHint: "Записать ограничения экспорта",
    ready: "Готов к экспорту.",
    selectFrames: "Выберите фреймы",
    exporting: "Экспортирую…",
    preparing: "Подготавливаю слои…",
    progress: "{{current}} из {{total}}: {{name}}",
    building: "Собираю файл PowerPoint…",
    done: "Готово — файл сохранён в загрузки.",
    errorPrepare: "Не удалось подготовить презентацию: {{detail}}",
    errorNoSelection: "Выберите хотя бы один фрейм, компонент или секцию.",
    errorNoBounds: "У фрейма «{{name}}» нет размеров.",
    errorNoLayerBounds: "Не удалось определить границы слоя «{{name}}».",
    errorNoSlides: "Нет слайдов для экспорта.",
    warningTextSvg: "Текст с градиентом или изображением сохранён как SVG.",
    warningMasksPng: "Маски и обрезанные группы сохранены как отдельные PNG-объекты.",
    warningFillsSvg: "Сложные заливки сохранены как отдельные SVG-объекты.",
    warningVectorsSvg: "Сложные контуры сохранены как отдельные SVG-объекты.",
    warningUnsupportedPng: "Неподдерживаемые слои сохранены как отдельные PNG-объекты.",
    warningDifferentSizes: "Фреймы разного формата вписаны в размер первого слайда.",
    notesHeading: "Особенности экспорта:"
  },
  en: {
    eyebrow: "Editable export",
    title: "Frames → PowerPoint",
    subtitle: "Text and simple shapes remain editable. Complex artwork is preserved as separate objects.",
    selectedSlides: "Selected slides",
    autoOrder: "Auto order",
    refresh: "Refresh",
    emptyTitle: "Select frames on the canvas",
    emptyText: "Numbered names define the order. Otherwise: top to bottom, then left to right.",
    settings: "Settings",
    pngQuality: "PNG quality",
    pngHint: "For shadows, masks and effects",
    hiddenLayers: "Hidden layers",
    hiddenHint: "Include in the presentation",
    notes: "Speaker notes",
    notesHint: "Record export limitations",
    ready: "Ready to export.",
    selectFrames: "Select frames",
    exporting: "Exporting…",
    preparing: "Preparing layers…",
    progress: "{{current}} of {{total}}: {{name}}",
    building: "Building the PowerPoint file…",
    done: "Done — the file was saved to Downloads.",
    errorPrepare: "Could not prepare the presentation: {{detail}}",
    errorNoSelection: "Select at least one frame, component, or section.",
    errorNoBounds: "Frame “{{name}}” has no dimensions.",
    errorNoLayerBounds: "Could not determine the bounds of layer “{{name}}”.",
    errorNoSlides: "There are no slides to export.",
    warningTextSvg: "Text with a gradient or image fill was saved as SVG.",
    warningMasksPng: "Masks and clipped groups were saved as separate PNG objects.",
    warningFillsSvg: "Complex fills were saved as separate SVG objects.",
    warningVectorsSvg: "Complex paths were saved as separate SVG objects.",
    warningUnsupportedPng: "Unsupported layers were saved as separate PNG objects.",
    warningDifferentSizes: "Frames with different aspect ratios were fitted to the first slide size.",
    notesHeading: "Export notes:"
  }
} as const;

export type TranslationKey = keyof typeof translations.ru;

export function isLocale(value: unknown): value is Locale {
  return value === "ru" || value === "en";
}

export function t(
  locale: Locale,
  key: TranslationKey,
  values: Record<string, string | number> = {}
): string {
  let text: string = translations[locale][key];
  for (const [name, value] of Object.entries(values)) {
    text = text.split(`{{${name}}}`).join(String(value));
  }
  return text;
}
