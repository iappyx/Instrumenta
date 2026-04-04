/**
 * ColorPickerPanel.tsx
 * Feature-Panel: Farbwähler (vollständige Implementierung).
 *
 * WICHTIG – Mac/Office.js-Einschränkung (Kategorie C):
 * Ein systemweiter Screen-Pixel-Picker ist im Office Add-in auf Mac NICHT möglich.
 * WKWebView (PowerPoint für Mac) unterstützt die EyeDropper API nicht.
 * Diese Einschränkung ist in TESTING.md dokumentiert.
 *
 * Verfügbare Funktionen:
 * - Hex-Eingabe mit RGB-Anzeige und Live-Vorschau
 * - Farbe aus selektiertem Shape auslesen (Fill / Linie / Schrift)
 * - Markenfarben-Palette (aus BrandConfig)
 * - Zuletzt-verwendet-Palette (Session)
 * - Anwenden auf: Füllung, Linie oder Schriftfarbe aller selektierten Shapes
 */

import * as React from "react";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton, ActionButton } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react/lib/TextField";
import { ChoiceGroup, IChoiceGroupOption } from "@fluentui/react/lib/ChoiceGroup";
import { Separator } from "@fluentui/react/lib/Separator";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";

import ColorSwatch from "../shared/ColorSwatch";
import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import {
  applyColorToShapes,
  pickColorFromShape,
  getRecentColors,
  ColorTarget,
} from "../../../features/colorPicker/ColorPickerService";
import { normalizeHex, hexToRgb } from "../../../utils/colorUtils";
import { DEFAULT_BRAND_CONFIG } from "../../../services/config/BrandConfig";

interface Notification {
  message: string;
  type:    NotificationType;
}

const TARGET_OPTIONS: IChoiceGroupOption[] = [
  { key: "fill", text: "Füllfarbe" },
  { key: "line", text: "Linienfarbe" },
  { key: "font", text: "Schriftfarbe" },
];

// Markenfarben aus der Default-Konfiguration
const BRAND_COLORS = DEFAULT_BRAND_CONFIG.profiles[0].brandColors;

const ColorPickerPanel: React.FC = () => {
  const [hexInput, setHexInput]         = React.useState<string>("#003366");
  const [target, setTarget]             = React.useState<ColorTarget>("fill");
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isApplying, setIsApplying]     = React.useState(false);
  const [isPicking, setIsPicking]       = React.useState(false);
  const [recentColors, setRecentColors] = React.useState<string[]>([]);

  // Zuletzt-verwendet-Farben bei Panel-Mount laden
  React.useEffect(() => {
    setRecentColors(getRecentColors());
  }, []);

  // ─── Abgeleitete Werte ────────────────────────────────────────────────────

  const normHex = normalizeHex(hexInput);
  const rgb     = normHex ? hexToRgb(normHex) : null;
  const isValid = normHex !== null;

  // ─── Handler ──────────────────────────────────────────────────────────────

  /** Hex-Eingabe normalisieren und setzen. */
  const handleHexChange = (_e: React.FormEvent, val?: string) => {
    setHexInput(val ?? "");
  };

  /** Farbe aus Shape auslesen. */
  const handlePickFromShape = async () => {
    setIsPicking(true);
    setNotification(null);

    const picked = await pickColorFromShape(target);

    setIsPicking(false);

    if (picked) {
      setHexInput(picked.hex);
      setNotification({
        message: `Farbe übernommen: ${picked.hex} (${
          target === "fill" ? "Füllung" : target === "line" ? "Linie" : "Schrift"
        })`,
        type: "success",
      });
    } else {
      const targetLabel = target === "fill" ? "Füllfarbe" : target === "line" ? "Linienfarbe" : "Schriftfarbe";
      setNotification({
        message: `${targetLabel} konnte nicht gelesen werden. Shape selektiert? Gradient/kein TextFrame?`,
        type: "warning",
      });
    }
  };

  /** Farbe auf selektierte Shapes anwenden. */
  const handleApply = async () => {
    if (!isValid) {
      setNotification({ message: "Ungültiger Hex-Wert. Format: #RRGGBB", type: "error" });
      return;
    }

    setIsApplying(true);
    setNotification(null);

    try {
      const result = await applyColorToShapes(normHex!, target);

      // Recent Colors aktualisieren
      setRecentColors(getRecentColors());

      if (result.applied === 0 && result.skipped === 0 && result.errors.length === 0) {
        setNotification({ message: "Keine Shapes selektiert.", type: "warning" });
      } else if (result.applied === 0) {
        setNotification({
          message: `Keine Shapes eingefärbt. ${result.skipped} übersprungen (kein TextFrame?).`,
          type: "warning",
        });
      } else {
        const skipNote = result.skipped > 0 ? `, ${result.skipped} übersprungen` : "";
        const errNote  = result.errors.length > 0 ? ` Fehler: ${result.errors.join(", ")}` : "";
        setNotification({
          message: `${result.applied} Shape(s) eingefärbt${skipNote}.${errNote}`,
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

  /** Farbe aus Palette übernehmen. */
  const selectColor = (hex: string) => {
    setHexInput(hex);
    setNotification(null);
  };

  // ─── Render ───────────────────────────────────────────────────────────────

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Farbwähler</Text>

      {/* Hinweis auf fehlenden Screen-Picker */}
      <MessageBar messageBarType={MessageBarType.info} isMultiline>
        Screen-Pixel-Picker nicht verfügbar (Mac/WebKit-Einschränkung).
        Verwende Hex-Eingabe, Markenfarben oder &bdquo;Aus Shape&ldquo;.
      </MessageBar>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {/* ── Farbeingabe + Vorschau ── */}
      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <TextField
          label="Hex-Farbwert:"
          value={hexInput}
          onChange={handleHexChange}
          placeholder="#003366"
          errorMessage={hexInput && !isValid ? "Format: #RRGGBB oder #RGB" : undefined}
          styles={{ root: { flex: 1 } }}
        />
        {/* Farbvorschau-Quadrat */}
        <div style={{ marginBottom: 4 }}>
          <ColorSwatch
            color={normHex ?? "transparent"}
            size={36}
            showLabel={false}
            title={normHex ?? "ungültig"}
          />
        </div>
      </Stack>

      {/* RGB-Anzeige */}
      {rgb ? (
        <Text variant="small" style={{ fontFamily: "monospace", color: "#555" }}>
          RGB({rgb.r}, {rgb.g}, {rgb.b})
        </Text>
      ) : (
        <Text variant="small" style={{ color: "#999" }}>RGB — —</Text>
      )}

      {/* ── Ziel-Auswahl ── */}
      <ChoiceGroup
        label="Anwenden auf:"
        options={TARGET_OPTIONS}
        selectedKey={target}
        onChange={(_e, opt) => opt && setTarget(opt.key as ColorTarget)}
        styles={{ flexContainer: { display: "flex", gap: 12 } }}
      />

      {/* ── Aktions-Buttons ── */}
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        {isApplying ? (
          <Spinner size={SpinnerSize.small} label="Wird angewendet…" styles={{ root: { flex: 1 } }} />
        ) : (
          <PrimaryButton
            text="Auf selektierte Shapes anwenden"
            onClick={handleApply}
            disabled={!isValid || isPicking}
            styles={{ root: { flex: 1 } }}
          />
        )}
        <TooltipHost
          content={`${
            target === "fill" ? "Füllfarbe" : target === "line" ? "Linienfarbe" : "Schriftfarbe"
          } des ersten selektierten Shapes übernehmen`}
        >
          <DefaultButton
            text={isPicking ? "…" : "Aus Shape"}
            onClick={handlePickFromShape}
            disabled={isPicking || isApplying}
            styles={{ root: { minWidth: 90 } }}
          />
        </TooltipHost>
      </Stack>

      {/* ── Markenfarben ── */}
      <Separator>Markenfarben</Separator>
      <Stack horizontal wrap tokens={{ childrenGap: 6 }}>
        {BRAND_COLORS.map((bc) => (
          <TooltipHost key={bc.value} content={bc.name}>
            <ColorSwatch
              color={bc.value}
              size={26}
              onClick={selectColor}
              title={bc.name}
            />
          </TooltipHost>
        ))}
      </Stack>
      <ActionButton
        text="Aus Konfiguration laden…"
        iconProps={{ iconName: "Settings" }}
        styles={{ root: { fontSize: 12, height: 24, padding: 0 } }}
        disabled
        title="Konfigurierbare Markenfarben – vollständig in Schritt 6 (Brand Check)"
      />

      {/* ── Zuletzt verwendet ── */}
      {recentColors.length > 0 && (
        <>
          <Separator>Zuletzt verwendet</Separator>
          <Stack horizontal wrap tokens={{ childrenGap: 6 }}>
            {recentColors.map((c) => (
              <ColorSwatch
                key={c}
                color={c}
                size={26}
                onClick={selectColor}
                title={c}
                showLabel={false}
              />
            ))}
          </Stack>
        </>
      )}

      {/* ── Hinweis ── */}
      <Separator />
      <Text variant="small" style={{ color: "#666" }}>
        <strong>Undo-Hinweis:</strong> Snapshot vor Farbänderung wird im Session-State gespeichert.
        Natives Undo (⌘+Z) funktioniert nach Add-in-Operationen möglicherweise nicht.
      </Text>
    </Stack>
  );
};

export default ColorPickerPanel;
