/**
 * CornerRadiusPanel.tsx
 * Feature-Panel: Eckenradius setzen (vollständige Implementierung).
 *
 * Unterstützte Shape-Typen: roundedRectangle
 * Nicht unterstützt: alle anderen Typen (werden mit Hinweis übersprungen)
 *
 * Pixel → Punkte: 1 px = 0,75 pt (96 DPI)
 * Normierung: adjustment[0] = ptValue / (min(width, height) / 2), geklemmt [0, 1]
 *
 * Undo-Strategie: Session-Snapshot vor Änderung (kein natives Office.js-Undo).
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Separator } from "@fluentui/react/lib/Separator";

import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import {
  applyCornerRadius,
  readCurrentRadiusPx,
} from "../../../features/cornerRadius/CornerRadiusService";

interface Notification {
  message: string;
  type:    NotificationType;
}

// Gängige Radius-Presets
const PRESETS = [
  { label: "Kein Radius",  value: 0 },
  { label: "Klein (4 px)", value: 4 },
  { label: "Mittel (8 px)", value: 8 },
  { label: "Groß (16 px)", value: 16 },
  { label: "Rund (50 px)", value: 50 },
];

const CornerRadiusPanel: React.FC = () => {
  const [radiusInput, setRadiusInput]     = React.useState<string>("8");
  const [notification, setNotification]   = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]         = React.useState(false);
  const [isReading, setIsReading]         = React.useState(false);
  const [currentPx, setCurrentPx]         = React.useState<number | null>(null);

  /** Liest den Radius des ersten selektierten Shapes und zeigt ihn an. */
  const handleReadCurrent = async () => {
    setIsReading(true);
    setCurrentPx(null);
    const px = await readCurrentRadiusPx();
    setIsReading(false);
    if (px !== null) {
      setCurrentPx(px);
      setRadiusInput(String(px));
      setNotification({ message: `Aktueller Radius: ${px} px`, type: "info" });
    } else {
      setNotification({
        message: "Kein Rounded-Rectangle selektiert oder Radius nicht lesbar.",
        type: "warning",
      });
    }
  };

  /** Validiert und wendet den Eckenradius an. */
  const handleApply = async () => {
    const px = parseFloat(radiusInput);

    if (isNaN(px) || px < 0) {
      setNotification({ message: "Bitte einen gültigen Pixelwert eingeben (≥ 0).", type: "error" });
      return;
    }

    setIsRunning(true);
    setNotification(null);

    try {
      const result = await applyCornerRadius(px);

      if (result.applied === 0 && result.skipped === 0 && result.errors.length === 0) {
        setNotification({ message: "Keine Shapes selektiert.", type: "warning" });
      } else if (result.applied === 0 && result.errors.length === 0) {
        setNotification({
          message: `Keine unterstützten Shapes gefunden. ${result.skipped} Shape(s) übersprungen (nur Rounded Rectangle wird unterstützt).`,
          type: "warning",
        });
      } else {
        const skipNote = result.skipped > 0
          ? ` ${result.skipped} übersprungen (nicht Rounded Rectangle).`
          : "";
        const errNote = result.errors.length > 0
          ? ` Fehler bei: ${result.errors.join(", ")}.`
          : "";
        setNotification({
          message: `${result.applied} Shape(s) angepasst.${skipNote}${errNote}`,
          type: result.errors.length > 0 ? "warning" : "success",
        });
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setNotification({ message: `Fehler: ${message}`, type: "error" });
    } finally {
      setIsRunning(false);
    }
  };

  const parsedPx = parseFloat(radiusInput);
  const ptPreview = !isNaN(parsedPx) && parsedPx >= 0
    ? (parsedPx * 0.75).toFixed(2)
    : "—";

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Eckenradius setzen</Text>

      <MessageBar messageBarType={MessageBarType.info} isMultiline>
        Unterstützt: <strong>Rounded Rectangle</strong>. Alle anderen Shape-Typen
        (Rechteck, Ellipse, Gruppe, Tabelle etc.) werden übersprungen.
      </MessageBar>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {/* Eingabefeld */}
      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <TextField
          label="Eckenradius in Pixel:"
          value={radiusInput}
          onChange={(_e, val) => setRadiusInput(val ?? "")}
          type="number"
          min={0}
          step={1}
          suffix="px"
          styles={{ root: { flex: 1 } }}
          description={`≈ ${ptPreview} pt (bei 96 DPI, 1 px = 0,75 pt)`}
        />
        <DefaultButton
          text={isReading ? "…" : "Auslesen"}
          onClick={handleReadCurrent}
          disabled={isReading || isRunning}
          title="Aktuellen Radius des selektierten Shapes auslesen"
          styles={{ root: { marginBottom: 22 } }}
        />
      </Stack>

      {currentPx !== null && (
        <Text variant="small" style={{ color: "#555" }}>
          Gelesener Wert: {currentPx} px
        </Text>
      )}

      {/* Presets */}
      <Separator>Schnell-Presets</Separator>
      <Stack horizontal wrap tokens={{ childrenGap: 6 }}>
        {PRESETS.map((p) => (
          <DefaultButton
            key={p.value}
            text={p.label}
            onClick={() => setRadiusInput(String(p.value))}
            styles={{ root: { minWidth: "auto", padding: "0 10px", height: 28, fontSize: 12 } }}
          />
        ))}
      </Stack>

      {/* Anwenden */}
      {isRunning ? (
        <Spinner size={SpinnerSize.small} label="Wird angewendet…" />
      ) : (
        <PrimaryButton
          text="Auf selektierte Shapes anwenden"
          onClick={handleApply}
          disabled={isNaN(parsedPx) || parsedPx < 0}
          styles={{ root: { width: "100%", marginTop: 8 } }}
        />
      )}

      {/* Hinweise */}
      <Separator />
      <Stack tokens={{ childrenGap: 4 }}>
        <Text variant="small" style={{ color: "#666" }}>
          <strong>Undo-Hinweis:</strong> Office.js bietet kein natives Undo für
          Add-in-Änderungen. Ein Snapshot wird im Session-State gespeichert.
          Das Dokument-Undo (⌘+Z) funktioniert nach Add-in-Operationen möglicherweise
          nicht zuverlässig.
        </Text>
        <Text variant="small" style={{ color: "#666" }}>
          <strong>Skalierung:</strong> Der Radius wird relativ zur Shape-Größe
          normiert. Sehr kleine Shapes können einen kleineren Radius zeigen als
          eingegeben wenn der Wert die maximale Shape-Ausdehnung überschreitet.
        </Text>
      </Stack>
    </Stack>
  );
};

export default CornerRadiusPanel;
