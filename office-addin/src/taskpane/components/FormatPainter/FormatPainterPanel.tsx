/**
 * FormatPainterPanel.tsx
 * Feature-Panel: Format Painter+ (vollständige Implementierung).
 *
 * Zwei-Schritt-Workflow:
 * 1. Quelle erfassen: erstes selektiertes Shape → Format speichern
 * 2. Ziele wählen + Eigenschaften auswählen + Anwenden
 *
 * Scope:
 * - Aktuelle Selektion (ohne Quell-Shape)
 * - Gleicher Shape-Typ auf aktiver Slide
 * - Gleicher Shape-Typ auf allen Slides
 *
 * Presets: Eigenschafts-Kombinationen dauerhaft in Document Settings speichern.
 */

import * as React from "react";
import { Stack }              from "@fluentui/react/lib/Stack";
import { Text }               from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton, ActionButton } from "@fluentui/react/lib/Button";
import { Checkbox }           from "@fluentui/react/lib/Checkbox";
import { Separator }          from "@fluentui/react/lib/Separator";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { ChoiceGroup, IChoiceGroupOption } from "@fluentui/react/lib/ChoiceGroup";
import { TextField }          from "@fluentui/react/lib/TextField";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Icon }               from "@fluentui/react/lib/Icon";
import { TooltipHost }        from "@fluentui/react/lib/Tooltip";

import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import ColorSwatch            from "../shared/ColorSwatch";
import {
  captureFormat,
  applyFormat,
  loadPresets,
  savePreset,
  deletePreset,
  CapturedFormat,
  FormatOptions,
  ApplyScope,
  FormatPreset,
} from "../../../features/formatPainter/FormatPainterService";

interface Notification {
  message: string;
  type:    NotificationType;
}

const DEFAULT_OPTIONS: FormatOptions = {
  fill: true, line: true, text: false, geometry: false, shadow: false,
};

const SCOPE_OPTIONS: IChoiceGroupOption[] = [
  { key: "selection",   text: "Aktuelle Selektion" },
  { key: "slideByType", text: "Gleicher Typ auf dieser Slide" },
  { key: "deckByType",  text: "Gleicher Typ auf allen Slides" },
];

const OPTION_LABELS: Record<keyof FormatOptions, string> = {
  fill:     "Füllung",
  line:     "Linie",
  text:     "Schrift",
  geometry: "Geometrie (Adjustments)",
  shadow:   "Schatten",
};

const OPTION_NOTES: Partial<Record<keyof FormatOptions, string>> = {
  fill:     "Nur Solid Fill. Gradient/Pattern wird übersprungen.",
  geometry: "Nur bei identischem Shape-Typ. Eckenradius etc.",
  shadow:   "Eingeschränkt auf Mac (Office.js-Limit).",
};

// ─── Komponente ───────────────────────────────────────────────────────────────

const FormatPainterPanel: React.FC = () => {
  const [captured, setCaptured]       = React.useState<CapturedFormat | null>(null);
  const [options, setOptions]         = React.useState<FormatOptions>(DEFAULT_OPTIONS);
  const [scope, setScope]             = React.useState<ApplyScope>("selection");
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isCapturing, setIsCapturing] = React.useState(false);
  const [isApplying, setIsApplying]   = React.useState(false);
  const [presets, setPresets]         = React.useState<FormatPreset[]>(() => loadPresets());
  const [newPresetName, setNewPresetName] = React.useState("");
  const [showSavePreset, setShowSavePreset] = React.useState(false);

  // ── Quelle erfassen ─────────────────────────────────────────────────────────
  const handleCapture = async () => {
    setIsCapturing(true);
    setNotification(null);
    const fmt = await captureFormat();
    setIsCapturing(false);

    if (fmt) {
      setCaptured(fmt);
      setNotification({ message: `Format erfasst von: „${fmt.shapeName}"`, type: "success" });
    } else {
      setNotification({ message: "Kein Shape selektiert oder Format nicht lesbar.", type: "warning" });
    }
  };

  // ── Format anwenden ─────────────────────────────────────────────────────────
  const handleApply = async () => {
    if (!captured) {
      setNotification({ message: "Bitte zuerst ein Quell-Shape erfassen.", type: "warning" });
      return;
    }
    const anySelected = Object.values(options).some(Boolean);
    if (!anySelected) {
      setNotification({ message: "Bitte mindestens eine Eigenschaft auswählen.", type: "warning" });
      return;
    }

    setIsApplying(true);
    setNotification(null);

    try {
      const result = await applyFormat(captured, options, scope);

      if (result.applied === 0 && result.skipped === 0) {
        setNotification({ message: "Keine Ziel-Shapes gefunden.", type: "warning" });
      } else {
        const skipNote = result.skipped > 0 ? `, ${result.skipped} übersprungen` : "";
        const errNote  = result.errors.length > 0 ? ` Fehler: ${result.errors.join(", ")}` : "";
        setNotification({
          message: `Format auf ${result.applied} Shape(s) übertragen${skipNote}.${errNote}`,
          type: result.errors.length > 0 ? "warning" : "success",
        });
      }
    } catch (err) {
      setNotification({
        message: `Fehler: ${err instanceof Error ? err.message : String(err)}`,
        type: "error",
      });
    } finally {
      setIsApplying(false);
    }
  };

  // ── Optionen ─────────────────────────────────────────────────────────────────
  const toggleOption = (key: keyof FormatOptions) =>
    setOptions((prev) => ({ ...prev, [key]: !prev[key] }));
  const selectAll  = () => setOptions({ fill: true, line: true, text: true, geometry: true, shadow: true });
  const selectNone = () => setOptions({ fill: false, line: false, text: false, geometry: false, shadow: false });

  // ── Presets ──────────────────────────────────────────────────────────────────
  const handleSavePreset = async () => {
    if (!newPresetName.trim()) return;
    const preset: FormatPreset = { name: newPresetName.trim(), options: { ...options } };
    await savePreset(preset);
    setPresets(loadPresets());
    setNewPresetName("");
    setShowSavePreset(false);
    setNotification({ message: `Preset „${preset.name}" gespeichert.`, type: "success" });
  };

  const handleLoadPreset = (p: FormatPreset) => {
    setOptions({ ...p.options });
    setNotification({ message: `Preset „${p.name}" geladen.`, type: "info" });
  };

  const handleDeletePreset = async (name: string) => {
    await deletePreset(name);
    setPresets(loadPresets());
  };

  // ─── Render ─────────────────────────────────────────────────────────────────
  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Format Painter+</Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {/* ── Schritt 1: Quelle erfassen ── */}
      <Separator>Schritt 1: Quell-Shape erfassen</Separator>
      <Text variant="small" style={{ color: "#555" }}>
        Selektiere das Shape, dessen Format übernommen werden soll,
        dann klicke „Format erfassen".
      </Text>

      {isCapturing ? (
        <Spinner size={SpinnerSize.small} label="Liest Format…" />
      ) : (
        <DefaultButton
          text="Format erfassen"
          iconProps={{ iconName: "Eyedropper" }}
          onClick={handleCapture}
          styles={{ root: { width: "100%" } }}
        />
      )}

      {/* Format-Vorschau */}
      {captured && <CapturedFormatPreview fmt={captured} />}

      {/* ── Schritt 2: Eigenschaften ── */}
      <Separator>Schritt 2: Eigenschaften auswählen</Separator>
      <Stack tokens={{ childrenGap: 6 }}>
        {(Object.keys(DEFAULT_OPTIONS) as (keyof FormatOptions)[]).map((key) => (
          <Stack key={key} tokens={{ childrenGap: 2 }}>
            <Checkbox
              label={OPTION_LABELS[key]}
              checked={options[key]}
              onChange={() => toggleOption(key)}
            />
            {OPTION_NOTES[key] && (
              <Text variant="xSmall" style={{ color: "#888", paddingLeft: 24 }}>
                {OPTION_NOTES[key]}
              </Text>
            )}
          </Stack>
        ))}
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <DefaultButton text="Alle"  onClick={selectAll}  styles={{ root: { flex: 1 } }} />
        <DefaultButton text="Keine" onClick={selectNone} styles={{ root: { flex: 1 } }} />
      </Stack>

      {/* ── Presets ── */}
      {(presets.length > 0 || !showSavePreset) && (
        <Stack tokens={{ childrenGap: 4 }}>
          {presets.length > 0 && (
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="xSmall" style={{ color: "#666" }}>Presets:</Text>
              <Stack horizontal wrap tokens={{ childrenGap: 4 }}>
                {presets.map((p) => (
                  <Stack key={p.name} horizontal verticalAlign="center" tokens={{ childrenGap: 2 }}>
                    <ActionButton
                      text={p.name}
                      onClick={() => handleLoadPreset(p)}
                      styles={{ root: { height: 24, fontSize: 12, padding: "0 6px" } }}
                    />
                    <TooltipHost content="Löschen">
                      <ActionButton
                        iconProps={{ iconName: "Delete" }}
                        onClick={() => handleDeletePreset(p.name)}
                        styles={{ root: { height: 24, minWidth: "auto", padding: "0 4px", color: "#A19F9D" } }}
                      />
                    </TooltipHost>
                  </Stack>
                ))}
              </Stack>
            </Stack>
          )}
          {!showSavePreset && (
            <ActionButton
              text="Aktuelle Auswahl als Preset speichern"
              iconProps={{ iconName: "Save" }}
              onClick={() => setShowSavePreset(true)}
              styles={{ root: { fontSize: 12, height: 24, padding: 0 } }}
            />
          )}
        </Stack>
      )}

      {showSavePreset && (
        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
          <TextField
            label="Preset-Name:"
            value={newPresetName}
            onChange={(_e, val) => setNewPresetName(val ?? "")}
            placeholder="z.B. Infront Standard"
            styles={{ root: { flex: 1 } }}
          />
          <PrimaryButton
            text="Speichern"
            onClick={handleSavePreset}
            disabled={!newPresetName.trim()}
            styles={{ root: { marginBottom: 4 } }}
          />
          <DefaultButton
            text="Abbrechen"
            onClick={() => { setShowSavePreset(false); setNewPresetName(""); }}
            styles={{ root: { marginBottom: 4 } }}
          />
        </Stack>
      )}

      {/* ── Schritt 3: Scope + Anwenden ── */}
      <Separator>Schritt 3: Anwenden</Separator>
      <ChoiceGroup
        options={SCOPE_OPTIONS}
        selectedKey={scope}
        onChange={(_e, opt) => opt && setScope(opt.key as ApplyScope)}
      />

      {scope === "deckByType" && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          „Alle Slides" kann bei großen Decks länger dauern.
        </MessageBar>
      )}

      {isApplying ? (
        <Spinner size={SpinnerSize.small} label="Wird übertragen…" />
      ) : (
        <PrimaryButton
          text="Format übertragen"
          iconProps={{ iconName: "PaintBucket" }}
          onClick={handleApply}
          disabled={!captured || isCapturing}
          styles={{ root: { width: "100%", marginTop: 4 } }}
        />
      )}

      {/* ── Hinweis ── */}
      <Separator />
      <Text variant="small" style={{ color: "#666" }}>
        <strong>Undo-Hinweis:</strong> Snapshot vor Übertragung im Session-State.
        Natives ⌘+Z funktioniert nach Add-in-Operationen möglicherweise nicht.
      </Text>
    </Stack>
  );
};

// ─── Format-Vorschau ──────────────────────────────────────────────────────────

const CapturedFormatPreview: React.FC<{ fmt: CapturedFormat }> = ({ fmt }) => (
  <Stack
    tokens={{ padding: "8px 10px", childrenGap: 6 }}
    styles={{ root: { background: "#F3F2F1", borderRadius: 4 } }}
  >
    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
      <Icon iconName="Copy" style={{ fontSize: 14, color: "#003366" }} />
      <Text variant="smallPlus" style={{ fontWeight: 600, color: "#003366" }}>
        Quelle: {fmt.shapeName || "(unbenannt)"}
      </Text>
    </Stack>
    <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
      {fmt.fill && !fmt.fill.unsupported && fmt.fill.color && (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
          <Text variant="xSmall" style={{ color: "#666" }}>Fill:</Text>
          <ColorSwatch color={fmt.fill.color} size={14} showLabel title={fmt.fill.color} />
        </Stack>
      )}
      {fmt.fill?.unsupported && (
        <Text variant="xSmall" style={{ color: "#888" }}>Fill: Gradient/Pattern</Text>
      )}
      {fmt.line?.color && (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
          <Text variant="xSmall" style={{ color: "#666" }}>Linie:</Text>
          <ColorSwatch color={fmt.line.color} size={14} showLabel title={fmt.line.color} />
        </Stack>
      )}
      {fmt.text?.fontName && (
        <Text variant="xSmall" style={{ color: "#555", fontFamily: "monospace" }}>
          {fmt.text.fontName} {fmt.text.fontSize ? `${fmt.text.fontSize}pt` : ""}
          {fmt.text.bold ? " B" : ""}{fmt.text.italic ? " I" : ""}
        </Text>
      )}
      {fmt.adjustments && fmt.adjustments.length > 0 && (
        <Text variant="xSmall" style={{ color: "#555" }}>
          Adj: [{fmt.adjustments.map((a) => a.toFixed(2)).join(", ")}]
        </Text>
      )}
    </Stack>
  </Stack>
);

export default FormatPainterPanel;
