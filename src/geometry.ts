export type AffineTransform = readonly [
  readonly [number, number, number],
  readonly [number, number, number]
];

export type LineGeometry = {
  x: number;
  y: number;
  width: number;
  height: number;
  flipV: boolean;
};

/**
 * Converts a horizontal Figma LineNode in local coordinates into the
 * axis-aligned box and direction expected by a PowerPoint line shape.
 * absoluteTransform includes rotations/scales from all ancestor nodes.
 */
export function lineGeometryFromTransform(
  transform: AffineTransform,
  localWidth: number,
  originX: number,
  originY: number
): LineGeometry {
  const [[a, , translateX], [b, , translateY]] = transform;
  const startX = translateX - originX;
  const startY = translateY - originY;
  const endX = translateX + a * localWidth - originX;
  const endY = translateY + b * localWidth - originY;

  return {
    x: Math.min(startX, endX),
    y: Math.min(startY, endY),
    width: Math.abs(endX - startX),
    height: Math.abs(endY - startY),
    flipV: (endX - startX) * (endY - startY) < 0
  };
}
