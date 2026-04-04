/**
 * MasterImportPanel.tsx
 * Feature-Panel: Master / Theme importieren (vollständige Implementierung).
 *
 * API-Einschränkung:
 * Vollständiger SlideMaster-Ersatz ist über Office.js auf Mac nicht möglich.
 * Fallback: Farb- und Font-Mapping über das gesamte Deck.
 *
 * Workflow:
 * 1. Deck scannen → Übersicht aller verwendeten Farben + Fonts
 * 2. Farb-Mappings definieren (Von → Nach) oder Preset wählen
 * 3. Font-Mappings definieren
 * 4. Optionen festlegen + Anwenden
 */

import * as React from "react";
import { Stack }              from "@fluentui/react/lib/Stack";
import { Text }               from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton, ActionButton } from "@fluentui/react/lib/Button";
import { TextField }          from "@fluentui/react/lib/TextField";
import { Checkbox }           from "@fluentui/react/lib/Checkbox";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Separator }          from "@fluentui/react/lib/Separator";
import { Pivot, PivotItem }   from "@fluentui/react/lib/Pivot";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { Icon }               from "@fluentui/react/lib/Icon";
import { TooltipHost }        from "@fluentui/react/lib/Tooltip";

import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import ColorSwatch            from "../shared/ColorSwatch";
import { normalizeHex }       from "../../../utils/colorUtils";
import {
  applyThemeMapping,
  scanDeckTheme,
  ThemeColorMapping,
  ThemeFontMapping,
  ThemeApplyOptions,
  DeckThemeSnapshot,
  INFRONT_PRESETS,
  DEFAULT_THEME_OPTIONS,
  ThemePreset,
} from "../../../features/masterImport/MasterImportService";

interface Notification {
  message: string;
  type:    NotificationType;
}

const EMPTY_COLOR_MAPPING = (): ThemeColorMapping => ({ from: "", to: "", label: "" });
const EMPTY_FONT_MAPPING  = (): ThemeFontMapping  => ({ from: "", to: "" });

// ─── Hauptkomponente ──────────────────────────────────────────────────────────

const MasterImportPanel: React.FC = () => {
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isScanning, setIsScanning]     = React.useState(false);
  const [isApplying, setIsApplying]     = React.useState(false);
  const [snapshot, setSnapshot]         = React.useState<DeckThemeSnapshot | null>(null);

  const [colorMappings, setColorMappings] = React.useState<ThemeColorMapping[]>([EMPTY_COLOR_MAPPING()]);
  const [fontMappings,  setFontMappings]  = React.useState<ThemeFontMapping[]>([EMPTY_FONT_MAPPING()]);
  const [options, setOptions]             = React.useState<ThemeApplyOptions>(DEFAULT_THEME_OPTIONS);

  // ── Deck scannen ────────────────────────────────────────────────────────────
  const handleScan = async () => {
    setIsScanning(true);
    setNotification(null);
    try {
      const result = await scanDeckTheme();
      setSnapshot(result);
      setNotification({
        message: `Scan abgeschlossen: ${result.colors.length} Farbe(n), ${result.fonts.length} Schriftart(en) gefunden.`,
        type: "success",
      });
    } catch (err) {
      setNotification({ message: `Scan-Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsScanning(false); }
  };

  // ── Preset laden ────────────────────────────────────────────────────────────
  const handleLoadPreset = (preset: ThemePreset) => {
    setColorMappings(preset.colorMappings.map((m) => ({ ...m })));
    setFontMappings(preset.fontMappings.map((m) => ({ ...m })));
    setNotification({ message: `Preset „${preset.name}" geladen.`, type: "info" });
  };

  // ── Theme anwenden ──────────────────────────────────────────────────────────
  const handleApply = async () => {
    const validColors = colorMappings.filter((m) => normalizeHex(m.from) && normalizeHex(m.to));
    const validFonts  = fontMappings.filter((m) => m.from.trim() && m.to.trim());

    if (validColors.length === 0 && validFonts.length === 0) {
      setNotification({ message: "Bitte mindestens ein Farb- oder Font-Mapping definieren.", type: "warning" });
      return;
    }

    setIsApplying(true);
    setNotification(null);
    try {
      const result = await applyThemeMapping(validColors, validFonts, options);
      const parts: string[] = [];
      if (result.colorsReplaced > 0) parts.push(`${result.colorsReplaced} Farb-Shape(s)`);
      if (result.fontsReplaced  > 0) parts.push(`${result.fontsReplaced} Font-Shape(s)`);
      const summary = parts.length > 0 ? parts.join(", ") + " angepasst." : "Keine passenden Shapes gefunden.";
      const errNote = result.errors.length > 0 ? ` Fehler: ${result.errors.slice(0, 3).join("; ")}` : "";
      setNotification({
        message: summary + errNote,
        type: result.errors.length > 0 ? "warning" : result.colorsReplaced + result.fontsReplaced > 0 ? "success" : "info",
      });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsApplying(false); }
  };

  // ── Farb-Mapping-Zeilen verwalten ────────────────────────────────────────────
  const updateColorMapping = (i: number, field: keyof ThemeColorMapping, value: string) => {
    setColorMappings((prev) => {
      const next = [...prev];
      next[i] = { ...next[i], [field]: value };
      return next;
    });
  };
  const addColorRow    = () => setColorMappings((prev) => [...prev, EMPTY_COLOR_MAPPING()]);
  const removeColorRow = (i: number) => setColorMappings((prev) => prev.filter((_, idx) => idx !== i));

  // ── Font-Mapping-Zeilen verwalten ────────────────────────────────────────────
  const updateFontMapping = (i: number, field: keyof ThemeFontMapping, value: string) => {
    setFontMappings((prev) => {
      const next = [...prev];
      next[i] = { ...next[i], [field]: value };
      return next;
    });
  };
  const addFontRow    = () => setFontMappings((prev) => [...prev, EMPTY_FONT_MAPPING()]);
  const removeFontRow = (i: number) => setFontMappings((prev) => prev.filter((_, idx) => idx !== i));

  // ── Farbe aus Scan in Mapping übernehmen ─────────────────────────────────────
  const insertColorFromScan = (hex: string) => {
    // In die erste leere "Von"-Zelle eintragen
    const emptyIdx = colorMappings.findIndex((m) => !m.from);
    if (emptyIdx >= 0) {
      updateColorMapping(emptyIdx, "from", hex);
    } else {
      setColorMappings((prev) => [...prev, { from: hex, to: "", label: "" }]);
    }
  };

  // ── Preset-Dropdown ──────────────────────────────────────────────────────────
  const presetOptions: IDropdownOption[] = [
    { key: "__none", text: "— Preset wählen —" },
    ...INFRONT_PRESETS.map((p, i) => ({ key: String(i), text: p.name })),
  ];

  // ─── Render ─────────────────────────────────────────────────────────────────
  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Master / Theme importieren</Text>

      <MessageBar messageBarType={MessageBarType.warning} isMultiline>
        <strong>API-Einschränkung:</strong> Vollständiger SlideMaster-Ersatz ist über Office.js auf Mac nicht möglich.
        Verfügbar: Farben und Schriftarten deck-weit ersetzen.
      </MessageBar>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      <Pivot>

        {/* ── Tab 1: Deck-Scan ── */}
        <PivotItem headerText="Deck-Scan" itemIcon="Search">
          <Stack tokens={{ childrenGap: 10, padding: "12px 0 0 0" }}>
            <Text variant="small" style={{ color: "#555" }}>
              Scannt alle Folien nach verwendeten Farben und Schriftarten.
              Klicke auf eine Farbe, um sie als Quell-Farbe in ein Mapping zu übernehmen.
            </Text>

            {isScanning ? (
              <Spinner size={SpinnerSize.small} label="Scannt Deck…" />
            ) : (
              <DefaultButton
                text="Deck scannen"
                iconProps={{ iconName: "Search" }}
                onClick={handleScan}
                styles={{ root: { width: "100%" } }}
              />
            )}

            {snapshot && (
              <Stack tokens={{ childrenGap: 10 }}>
                {/* Farben */}
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="smallPlus" style={{ fontWeight: 600 }}>
                    Farben ({snapshot.colors.length}):
                  </Text>
                  {snapshot.colors.length === 0 ? (
                    <Text variant="small" style={{ color: "#888" }}>Keine Farben gefunden.</Text>
                  ) : (
                    <Stack horizontal wrap tokens={{ childrenGap: 6 }}>
                      {snapshot.colors.map((hex) => (
                        <TooltipHost key={hex} content={`${hex} – als Quelle übernehmen`}>
                          <div
                            onClick={() => insertColorFromScan(hex)}
                            style={{ cursor: "pointer" }}
                            title={hex}
                          >
                            <ColorSwatch color={hex} size={22} showLabel />
                          </div>
                        </TooltipHost>
                      ))}
                    </Stack>
                  )}
                </Stack>

                <Separator />

                {/* Fonts */}
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="smallPlus" style={{ fontWeight: 600 }}>
                    Schriftarten ({snapshot.fonts.length}):
                  </Text>
                  {snapshot.fonts.length === 0 ? (
                    <Text variant="small" style={{ color: "#888" }}>Keine Schriftarten gefunden.</Text>
                  ) : (
                    <Stack horizontal wrap tokens={{ childrenGap: 6 }}>
                      {snapshot.fonts.map((f) => (
                        <Text
                          key={f}
                          variant="xSmall"
                          style={{
                            background: "#F3F2F1",
                            borderRadius: 3,
                            padding: "2px 6px",
                            fontFamily: f,
                            cursor: "pointer",
                          }}
                          onClick={() => {
                            const emptyIdx = fontMappings.findIndex((m) => !m.from);
                            if (emptyIdx >= 0) {
                              updateFontMapping(emptyIdx, "from", f);
                            } else {
                              setFontMappings((prev) => [...prev, { from: f, to: "" }]);
                            }
                          }}
                          title={`${f} – als Quelle übernehmen`}
                        >
                          {f}
                        </Text>
                      ))}
                    </Stack>
                  )}
                </Stack>
              </Stack>
            )}
          </Stack>
        </PivotItem>

        {/* ── Tab 2: Mappings ── */}
        <PivotItem headerText="Mappings" itemIcon="Switch">
          <Stack tokens={{ childrenGap: 12, padding: "12px 0 0 0" }}>

            {/* Preset */}
            <Dropdown
              label="Preset laden:"
              options={presetOptions}
              defaultSelectedKey="__none"
              onChange={(_e, opt) => {
                if (!opt || opt.key === "__none") return;
                const preset = INFRONT_PRESETS[Number(opt.key)];
                if (preset) handleLoadPreset(preset);
              }}
            />

            <Separator>Farb-Mappings</Separator>
            <Text variant="xSmall" style={{ color: "#666" }}>
              Von-Farbe (Hex) → Nach-Farbe (Hex). Toleranz in Optionen einstellbar.
            </Text>

            {colorMappings.map((mapping, i) => (
              <Stack key={i} horizontal verticalAlign="end" tokens={{ childrenGap: 6 }}>
                <Stack tokens={{ childrenGap: 2 }} styles={{ root: { flex: 1 } }}>
                  <TextField
                    placeholder="#RRGGBB (Von)"
                    value={mapping.from}
                    onChange={(_e, v) => updateColorMapping(i, "from", v ?? "")}
                    prefix="#"
                    styles={{ root: { flex: 1 } }}
                  />
                </Stack>
                {normalizeHex(mapping.from) && (
                  <ColorSwatch color={normalizeHex(mapping.from)!} size={22} />
                )}
                <Icon iconName="Forward" style={{ fontSize: 14, color: "#888", marginBottom: 6 }} />
                <Stack tokens={{ childrenGap: 2 }} styles={{ root: { flex: 1 } }}>
                  <TextField
                    placeholder="#RRGGBB (Nach)"
                    value={mapping.to}
                    onChange={(_e, v) => updateColorMapping(i, "to", v ?? "")}
                    prefix="#"
                    styles={{ root: { flex: 1 } }}
                  />
                </Stack>
                {normalizeHex(mapping.to) && (
                  <ColorSwatch color={normalizeHex(mapping.to)!} size={22} />
                )}
                <ActionButton
                  iconProps={{ iconName: "Delete" }}
                  onClick={() => removeColorRow(i)}
                  disabled={colorMappings.length <= 1}
                  styles={{ root: { minWidth: "auto", padding: "0 4px", height: 28, marginBottom: 2, color: "#A19F9D" } }}
                />
              </Stack>
            ))}

            <ActionButton
              text="Farb-Mapping hinzufügen"
              iconProps={{ iconName: "Add" }}
              onClick={addColorRow}
              disabled={colorMappings.length >= 16}
              styles={{ root: { height: 28, padding: 0 } }}
            />

            <Separator>Font-Mappings</Separator>
            <Text variant="xSmall" style={{ color: "#666" }}>
              Schriftart-Name (Von) → Schriftart-Name (Nach). Exakte Schreibweise erforderlich.
            </Text>

            {fontMappings.map((mapping, i) => (
              <Stack key={i} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                <TextField
                  placeholder="Schriftart (Von)"
                  value={mapping.from}
                  onChange={(_e, v) => updateFontMapping(i, "from", v ?? "")}
                  styles={{ root: { flex: 1 } }}
                />
                <Icon iconName="Forward" style={{ fontSize: 14, color: "#888" }} />
                <TextField
                  placeholder="Schriftart (Nach)"
                  value={mapping.to}
                  onChange={(_e, v) => updateFontMapping(i, "to", v ?? "")}
                  styles={{ root: { flex: 1 } }}
                />
                <ActionButton
                  iconProps={{ iconName: "Delete" }}
                  onClick={() => removeFontRow(i)}
                  disabled={fontMappings.length <= 1}
                  styles={{ root: { minWidth: "auto", padding: "0 4px", height: 28, color: "#A19F9D" } }}
                />
              </Stack>
            ))}

            <ActionButton
              text="Font-Mapping hinzufügen"
              iconProps={{ iconName: "Add" }}
              onClick={addFontRow}
              disabled={fontMappings.length >= 8}
              styles={{ root: { height: 28, padding: 0 } }}
            />
          </Stack>
        </PivotItem>

        {/* ── Tab 3: Optionen + Anwenden ── */}
        <PivotItem headerText="Anwenden" itemIcon="Play">
          <Stack tokens={{ childrenGap: 12, padding: "12px 0 0 0" }}>

            <Stack tokens={{ childrenGap: 6 }}>
              <Text variant="smallPlus" style={{ fontWeight: 600 }}>Farb-Ersatz anwenden auf:</Text>
              <Checkbox
                label="Füllfarben"
                checked={options.targetFill}
                onChange={(_e, c) => setOptions((prev) => ({ ...prev, targetFill: !!c }))}
              />
              <Checkbox
                label="Linienfarben"
                checked={options.targetLine}
                onChange={(_e, c) => setOptions((prev) => ({ ...prev, targetLine: !!c }))}
              />
              <Checkbox
                label="Schriftfarben"
                checked={options.targetFont}
                onChange={(_e, c) => setOptions((prev) => ({ ...prev, targetFont: !!c }))}
              />
            </Stack>

            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="smallPlus" style={{ fontWeight: 600 }}>Farb-Toleranz:</Text>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <TextField
                  value={String(options.colorTolerance)}
                  onChange={(_e, v) => {
                    const n = parseInt(v ?? "0", 10);
                    if (!isNaN(n) && n >= 0 && n <= 30) {
                      setOptions((prev) => ({ ...prev, colorTolerance: n }));
                    }
                  }}
                  type="number"
                  min={0}
                  max={30}
                  styles={{ root: { width: 64 } }}
                />
                <Text variant="xSmall" style={{ color: "#666" }}>
                  0 = exakt, 30 = sehr tolerant
                </Text>
              </Stack>
            </Stack>

            <Separator />

            <Stack tokens={{ childrenGap: 6 }}>
              <Text variant="smallPlus" style={{ fontWeight: 600 }}>Font-Optionen:</Text>
              <Checkbox
                label="Schriftgröße beibehalten"
                checked={options.keepFontSize}
                onChange={(_e, c) => setOptions((prev) => ({ ...prev, keepFontSize: !!c }))}
              />
              <Checkbox
                label="Formatierung beibehalten (Fett/Kursiv)"
                checked={options.keepFontFormatting}
                onChange={(_e, c) => setOptions((prev) => ({ ...prev, keepFontFormatting: !!c }))}
              />
            </Stack>

            <Separator />

            <MessageBar messageBarType={MessageBarType.info} isMultiline>
              <strong>Hinweis:</strong> Diese Operation ersetzt Farben und Fonts im gesamten Deck.
              Ein natives Undo ist nicht verfügbar. Bitte vorher eine Kopie der Datei anlegen.
            </MessageBar>

            {/* Zusammenfassung der aktiven Mappings */}
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="xSmall" style={{ color: "#666" }}>
                Aktive Farb-Mappings: {colorMappings.filter((m) => normalizeHex(m.from) && normalizeHex(m.to)).length}
                {" | "}
                Aktive Font-Mappings: {fontMappings.filter((m) => m.from.trim() && m.to.trim()).length}
              </Text>
            </Stack>

            {isApplying ? (
              <Spinner size={SpinnerSize.small} label="Wird angewendet…" />
            ) : (
              <PrimaryButton
                text="Theme anwenden (gesamtes Deck)"
                iconProps={{ iconName: "Sync" }}
                onClick={handleApply}
                styles={{ root: { width: "100%", marginTop: 4 } }}
              />
            )}

            <Text variant="small" style={{ color: "#666" }}>
              <strong>Undo:</strong> Session-State-Snapshot wird vor der Anwendung gespeichert
              (über FindReplace-Service). Natives ⌘+Z funktioniert nach Add-in-Operationen nicht.
            </Text>
          </Stack>
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

export default MasterImportPanel;
