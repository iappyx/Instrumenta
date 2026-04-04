/**
 * GapEqualizerPanel.tsx
 * Feature-Panel: Gap-Equalizer (vollständige Implementierung).
 *
 * Drei Modi:
 * - equal: gleichmäßig verteilen (äußere Shapes bleiben fixiert)
 * - fixed: fester Abstand (erstes Shape bleibt, alle weiteren werden verschoben)
 * - pack:  dicht packen (Abstand = 0, wie fixed mit gap=0)
 *
 * Richtungen: horizontal, vertikal, beide.
 *
 * Zusatz: Live-Vorschau des berechneten Abstands vor dem Anwenden.
 * Schnell-Buttons auch über Ribbon-ExecuteFunction (gapHorizontal/gapVertical).
 */

import * as React from "react";
import { Stack }              from "@fluentui/react/lib/Stack";
import { Text }               from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton, ActionButton } from "@fluentui/react/lib/Button";
import { TextField }          from "@fluentui/react/lib/TextField";
import { ChoiceGroup, IChoiceGroupOption } from "@fluentui/react/lib/ChoiceGroup";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Separator }          from "@fluentui/react/lib/Separator";
import { Icon }               from "@fluentui/react/lib/Icon";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";

import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import {
  equalizeGaps,
  previewGap,
  GapDirection,
  GapMode,
  GapOptions,
  GapPreview,
} from "../../../features/gapEqualizer/GapEqualizerService";

interface Notification {
  message: string;
  type:    NotificationType;
}

// ─── Optionen ─────────────────────────────────────────────────────────────────

const DIRECTION_OPTIONS: IChoiceGroupOption[] = [
  { key: "horizontal", text: "Horizontal" },
  { key: "vertical",   text: "Vertikal"   },
  { key: "both",       text: "Beide"      },
];

const MODE_OPTIONS: IChoiceGroupOption[] = [
  { key: "equal", text: "Gleichmäßig verteilen" },
  { key: "fixed", text: "Fester Abstand"         },
  { key: "pack",  text: "Dicht packen (0 pt)"    },
];

// ─── Hauptkomponente ──────────────────────────────────────────────────────────

const GapEqualizerPanel: React.FC = () => {
  const [direction, setDirection]       = React.useState<GapDirection>("horizontal");
  const [mode, setMode]                 = React.useState<GapMode>("equal");
  const [fixedGapStr, setFixedGapStr]   = React.useState("8");
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]       = React.useState(false);
  const [preview, setPreview]           = React.useState<GapPreview | null>(null);

  const buildOptions = (): GapOptions => ({
    direction,
    mode,
    fixedGap: mode === "fixed" ? parseFloat(fixedGapStr) || 0 : 0,
  });

  // ── Vorschau berechnen ──────────────────────────────────────────────────────
  const handlePreview = async () => {
    setIsRunning(true);
    setNotification(null);
    setPreview(null);
    try {
      const p = await previewGap(buildOptions());
      setPreview(p);
      if (!p.valid) {
        setNotification({ message: p.message, type: "warning" });
      }
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ── Anwenden ────────────────────────────────────────────────────────────────
  const handleApply = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      const result = await equalizeGaps(buildOptions());

      if (result.errors.length > 0) {
        setNotification({ message: result.errors[0], type: "warning" });
        return;
      }

      if (result.adjusted === 0) {
        setNotification({ message: "Abstände bereits gleichmäßig – keine Änderung nötig.", type: "info" });
        return;
      }

      const parts: string[] = [];
      if (!isNaN(result.computedGapH)) parts.push(`H: ${result.computedGapH.toFixed(1)} pt`);
      if (!isNaN(result.computedGapV)) parts.push(`V: ${result.computedGapV.toFixed(1)} pt`);
      setNotification({
        message: `${result.adjusted} Shape(s) verschoben. Abstand → ${parts.join(", ")}.`,
        type: "success",
      });
      setPreview(null);
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ─── Render ─────────────────────────────────────────────────────────────────
  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Gap-Equalizer</Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {/* Info-Hinweis */}
      <MessageBar messageBarType={MessageBarType.info} isMultiline>
        Shapes auf der aktiven Folie selektieren, dann Modus und Richtung wählen.
        Schnellzugriff auch über den Ribbon (Gap H / Gap V).
      </MessageBar>

      {/* Richtung */}
      <ChoiceGroup
        label="Richtung:"
        options={DIRECTION_OPTIONS}
        selectedKey={direction}
        onChange={(_e, opt) => { opt && setDirection(opt.key as GapDirection); setPreview(null); }}
        styles={{ flexContainer: { display: "flex", gap: 12 } }}
      />

      <Separator />

      {/* Modus */}
      <ChoiceGroup
        label="Modus:"
        options={MODE_OPTIONS}
        selectedKey={mode}
        onChange={(_e, opt) => { opt && setMode(opt.key as GapMode); setPreview(null); }}
      />

      {/* Fester Abstand Input */}
      {mode === "fixed" && (
        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
          <TextField
            label="Gewünschter Abstand:"
            value={fixedGapStr}
            onChange={(_e, v) => { setFixedGapStr(v ?? ""); setPreview(null); }}
            type="number"
            min={0}
            suffix="pt"
            styles={{ root: { width: 110 } }}
          />
          <Stack tokens={{ childrenGap: 2 }} styles={{ root: { marginBottom: 4 } }}>
            {[4, 8, 12, 16, 24].map((v) => (
              <ActionButton
                key={v}
                text={`${v}`}
                onClick={() => { setFixedGapStr(String(v)); setPreview(null); }}
                styles={{ root: { height: 20, padding: "0 6px", fontSize: 11, minWidth: "auto" } }}
              />
            ))}
          </Stack>
        </Stack>
      )}

      {mode === "equal" && (
        <Text variant="xSmall" style={{ color: "#666" }}>
          Äußerste Shapes bleiben an ihrer Position. Innere Shapes werden verschoben.
          Mindestens 3 Shapes selektieren.
        </Text>
      )}

      {mode === "fixed" && (
        <Text variant="xSmall" style={{ color: "#666" }}>
          Das erste Shape (ganz links / oben) bleibt fixiert. Alle weiteren werden
          mit dem angegebenen Abstand dahinter positioniert. Mindestens 2 Shapes.
        </Text>
      )}

      {mode === "pack" && (
        <Text variant="xSmall" style={{ color: "#666" }}>
          Shapes werden dicht nebeneinander / untereinander angeordnet (Abstand = 0 pt).
          Das erste Shape bleibt fixiert. Mindestens 2 Shapes.
        </Text>
      )}

      <Separator />

      {/* Vorschau-Ergebnis */}
      {preview?.valid && (
        <PreviewBadge preview={preview} direction={direction} />
      )}

      {/* Buttons */}
      {isRunning ? (
        <Spinner size={SpinnerSize.small} label="Wird berechnet…" />
      ) : (
        <Stack tokens={{ childrenGap: 8 }}>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="Vorschau"
              iconProps={{ iconName: "Calculator" }}
              onClick={handlePreview}
              styles={{ root: { flex: 1 } }}
            />
            <PrimaryButton
              text="Anwenden"
              iconProps={{ iconName: "AlignCenter" }}
              onClick={handleApply}
              styles={{ root: { flex: 1 } }}
            />
          </Stack>

          {/* Schnell-Buttons */}
          <Separator>Schnellzugriff</Separator>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <QuickButton
              label="Gap H"
              iconName="AlignHorizontalCenter"
              onClick={() => {
                setDirection("horizontal");
                setMode("equal");
                setPreview(null);
              }}
            />
            <QuickButton
              label="Gap V"
              iconName="AlignVerticalCenter"
              onClick={() => {
                setDirection("vertical");
                setMode("equal");
                setPreview(null);
              }}
            />
            <QuickButton
              label="Pack H"
              iconName="AlignLeft"
              onClick={() => {
                setDirection("horizontal");
                setMode("pack");
                setPreview(null);
              }}
            />
            <QuickButton
              label="Pack V"
              iconName="AlignTop"
              onClick={() => {
                setDirection("vertical");
                setMode("pack");
                setPreview(null);
              }}
            />
          </Stack>
        </Stack>
      )}

      {/* Undo-Hinweis */}
      <Text variant="xSmall" style={{ color: "#888" }}>
        <strong>Undo:</strong> Session-State-Snapshot vor jeder Operation gespeichert.
        Natives ⌘+Z funktioniert nach Add-in-Operationen möglicherweise nicht.
      </Text>
    </Stack>
  );
};

// ─── Vorschau-Badge ───────────────────────────────────────────────────────────

interface PreviewBadgeProps {
  preview:   GapPreview;
  direction: GapDirection;
}

const PreviewBadge: React.FC<PreviewBadgeProps> = ({ preview, direction }) => (
  <Stack
    horizontal
    verticalAlign="center"
    tokens={{ childrenGap: 8, padding: "8px 10px" }}
    styles={{ root: { background: "#EFF6FC", borderRadius: 4 } }}
  >
    <Icon iconName="Info" style={{ fontSize: 14, color: "#0078D4" }} />
    <Stack tokens={{ childrenGap: 2 }}>
      <Text variant="small" style={{ color: "#0078D4", fontWeight: 600 }}>
        Vorschau ({preview.shapeCount} Shapes)
      </Text>
      {direction !== "vertical" && !isNaN(preview.computedGapH) && (
        <Text variant="xSmall" style={{ color: "#0078D4" }}>
          Horizontal: {preview.computedGapH.toFixed(2)} pt
          {preview.computedGapH < 0 ? " ⚠ Shapes überlappen sich" : ""}
        </Text>
      )}
      {direction !== "horizontal" && !isNaN(preview.computedGapV) && (
        <Text variant="xSmall" style={{ color: "#0078D4" }}>
          Vertikal: {preview.computedGapV.toFixed(2)} pt
          {preview.computedGapV < 0 ? " ⚠ Shapes überlappen sich" : ""}
        </Text>
      )}
    </Stack>
  </Stack>
);

// ─── Schnell-Button ───────────────────────────────────────────────────────────

interface QuickButtonProps {
  label:     string;
  iconName:  string;
  onClick:   () => void;
}

const QuickButton: React.FC<QuickButtonProps> = ({ label, iconName, onClick }) => (
  <DefaultButton
    text={label}
    iconProps={{ iconName }}
    onClick={onClick}
    styles={{ root: { flex: 1, fontSize: 11, padding: "0 4px" } }}
  />
);

export default GapEqualizerPanel;
