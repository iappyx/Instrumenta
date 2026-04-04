/**
 * ColorSwatch.tsx
 * Zeigt eine Farbvorschau (kleines Quadrat) mit optionalem Hex-Label.
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";

interface ColorSwatchProps {
  color: string;        // Hex-Farbe, z.B. "#FF0000"
  size?: number;        // Breite/Höhe in px (default: 20)
  showLabel?: boolean;  // Hex-Code anzeigen (default: false)
  onClick?: (color: string) => void;
  title?: string;
}

const ColorSwatch: React.FC<ColorSwatchProps> = ({
  color,
  size = 20,
  showLabel = false,
  onClick,
  title,
}) => {
  const isTransparent = color === "transparent" || color === "none" || color === "";

  const swatchStyle: React.CSSProperties = {
    width:  size,
    height: size,
    backgroundColor: isTransparent ? "transparent" : color,
    border: "1px solid #ccc",
    borderRadius: 3,
    cursor: onClick ? "pointer" : "default",
    flexShrink: 0,
    backgroundImage: isTransparent
      ? "linear-gradient(45deg, #ccc 25%, transparent 25%), linear-gradient(-45deg, #ccc 25%, transparent 25%), linear-gradient(45deg, transparent 75%, #ccc 75%), linear-gradient(-45deg, transparent 75%, #ccc 75%)"
      : undefined,
    backgroundSize: isTransparent ? "8px 8px" : undefined,
    backgroundPosition: isTransparent ? "0 0, 0 4px, 4px -4px, -4px 0px" : undefined,
  };

  const swatch = (
    <div
      style={swatchStyle}
      onClick={() => onClick?.(color)}
      role={onClick ? "button" : undefined}
      aria-label={title ?? color}
    />
  );

  return (
    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
      {title ? (
        <TooltipHost content={title}>{swatch}</TooltipHost>
      ) : (
        swatch
      )}
      {showLabel && (
        <Text variant="small" style={{ fontFamily: "monospace" }}>
          {isTransparent ? "transparent" : color}
        </Text>
      )}
    </Stack>
  );
};

export default ColorSwatch;
