export type PositionedItem = {
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
};

/**
 * Presentation order:
 * 1. If every frame has a unique numeric prefix, use that explicit numbering.
 * 2. Otherwise read the canvas by rows: top-to-bottom, left-to-right.
 */
export function sortBySlideOrder<T extends PositionedItem>(items: readonly T[]): T[] {
  const numbered = items.map((item) => ({ item, number: numericPrefix(item.name) }));
  const numbers = numbered.map(({ number }) => number);
  if (
    items.length > 1 &&
    numbers.every((number): number is number => number !== null) &&
    new Set(numbers).size === numbers.length
  ) {
    return numbered
      .sort((a, b) => (a.number as number) - (b.number as number) || naturalName(a.item, b.item))
      .map(({ item }) => item);
  }

  const remaining = [...items].sort(
    (a, b) => a.y - b.y || a.x - b.x || naturalName(a, b)
  );
  const result: T[] = [];

  while (remaining.length) {
    const first = remaining.shift()!;
    const row = [first];
    const tolerance = Math.max(8, first.height * 0.25);

    for (let index = 0; index < remaining.length;) {
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

function numericPrefix(name: string): number | null {
  const match = name.match(/^\s*(\d+)(?:\D|$)/);
  return match ? Number(match[1]) : null;
}

function naturalName(a: PositionedItem, b: PositionedItem): number {
  return a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: "base" });
}
