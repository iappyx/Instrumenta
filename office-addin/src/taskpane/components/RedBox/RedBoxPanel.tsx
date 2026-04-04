/**
 * RedBoxPanel.tsx
 * Feature-Panel: Red Box / Safe-Area-Begrenzung (vollständige Implementierung).
 *
 * Features:
 * - Konfiguration der Abstände (Margins) von allen 4 Seiten
 * - Verknüpfte Margins (alle gleichzeitig ändern)
 * - Farb-Auswahl mit Presets + Hex-Eingabe
 * - Linienstärke + Linientyp (solid/dash/dot)
 * - CSS-basierte Vorschau der Box auf einer Folien-Skizze
 * - Toggle (Ein/Aus) auf aktueller Folie
 * - Auf alle Folien anwenden / Von allen entfernen
 * - Status-Badge (aktuelle Folie + Gesamtanzahl)
 * - Konfiguration persistent in Document.Settings
 */

import * as React from "react";
import { Stack }              from "@fluentui/react/lib/Stack";
import { Text }               from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton, ActionButton } from "@fluentui/react/lib/Button";
import { TextField }          from "@fluentui/react/lib/TextField";
import { Toggle }             from "@fluentui/react/lib/Toggle";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Separator }          from "@fluentui/react/lib/Separator";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { Icon }               from "@fluentui/react/lib/Icon";
import { TooltipHost }        from "@fluentui/react/lib/Tooltip";

import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import ColorSwatch            from "../shared/ColorSwatch";
import { normalizeHex }       from "../../../utils/colorUtils";
import {
  RedBoxConfig,
  RedBoxStatus,
  DEFAULT_REDBOX_CONFIG,
  loadRedBoxConfig,
  saveRedBoxConfig,
  getRedBoxStatus,
  toggleRedBoxOnCurrentSlide,
  addRedBoxToAllSlides,
  removeRedBoxFromAllSlides,
  updateRedBoxOnCurrentSlide,
} from "../../../features/redBox/RedBoxService";

interface Notification {
  message: string;
  type:    NotificationType;
}

// ─── Farb-Presets ─────────────────────────────────────────────────────────────

const COLOR_PRESETS = [
  { hex: "#FF0000", label: "Rot (Standard)" },
  { hex: "#003366", label: "Infront Navy"   },
  { hex: "#000000", label: "Schwarz"        },
  { hex: "#FF8C00", label: "Orange"         },
  { hex: "#0078D4", label: "Blau"           },
];

const DASH_OPTIONS: IDropdownOption[] = [
  { key: "solid", text: "Durchgezogen" },
  { key: "dash",  text: "Gestrichelt"  },
  { key: "dot",   text: "Gepunktet"    },
];

const WEIGHT_PRESETS = [0.5, 1, 1.5, 2, 3];

// ─── Hauptkomponente ──────────────────────────────────────────────────────────

const RedBoxPanel: React.FC = () => {
  const [config, setConfig]             = React.useState<RedBoxConfig>(loadRedBoxConfig);
  const [linked, setLinked]             = React.useState(true);   // Margins verknüpft
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]       = React.useState(false);
  const [status, setStatus]             = React.useState<RedBoxStatus | null>(null);

  // Status laden
  const refreshStatus = () =>
    getRedBoxStatus().then(setStatus).catch(() => {});

  React.useEffect(() => { refreshStatus(); }, []);

  // ── Margin-Update ──────────────────────────────────────────────────────────
  const updateMargin = (side: "marginTop" | "marginRight" | "marginBottom" | "marginLeft", val: string) => {
    const n = parseFloat(val);
    const num = isNaN(n) ? 0 : Math.max(0, n);
    if (linked) {
      setConfig((prev) => ({
        ...prev,
        marginTop:    num,
        marginRight:  num,
        marginBottom: num,
        marginLeft:   num,
      }));
    } else {
      setConfig((prev) => ({ ...prev, [side]: num }));
    }
  };

  // ── Konfiguration speichern ────────────────────────────────────────────────
  const handleSave = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      await saveRedBoxConfig(config);
      setNotification({ message: "Einstellungen gespeichert.", type: "success" });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ── Toggle auf aktueller Folie ─────────────────────────────────────────────
  const handleToggle = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      await saveRedBoxConfig(config);
      const action = await toggleRedBoxOnCurrentSlide(config);
      await refreshStatus();
      setNotification({
        message: action === "added" ? "Red Box eingefügt." : "Red Box entfernt.",
        type: "success",
      });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ── Auf alle Folien anwenden ───────────────────────────────────────────────
  const handleAddAll = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      await saveRedBoxConfig(config);
      const result = await addRedBoxToAllSlides(config);
      await refreshStatus();
      const msg = result.added > 0
        ? `Red Box auf ${result.added} Folie(n) eingefügt.`
        : "Alle Folien haben bereits eine Red Box.";
      const errNote = result.errors.length > 0 ? ` Fehler: ${result.errors[0]}` : "";
      setNotification({ message: msg + errNote, type: result.errors.length > 0 ? "warning" : "success" });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ── Von allen Folien entfernen ─────────────────────────────────────────────
  const handleRemoveAll = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      const result = await removeRedBoxFromAllSlides();
      await refreshStatus();
      setNotification({
        message: result.removed > 0
          ? `Red Box von ${result.removed} Folie(n) entfernt.`
          : "Keine Red Box gefunden.",
        type: result.removed > 0 ? "success" : "info",
      });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ── Aktualisieren (Update vorhandene Box) ──────────────────────────────────
  const handleUpdate = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      await saveRedBoxConfig(config);
      const updated = await updateRedBoxOnCurrentSlide(config);
      await refreshStatus();
      setNotification({
        message: updated ? "Red Box auf aktueller Folie aktualisiert." : "Keine Red Box auf aktueller Folie.",
        type: updated ? "success" : "info",
      });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  const validColor = normalizeHex(config.color) ?? "#FF0000";

  // ─── Render ─────────────────────────────────────────────────────────────────
  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Red Box / Safe Area</Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {/* Status-Badge */}
      <StatusBadge status={status} onRefresh={refreshStatus} />

      {/* Vorschau */}
      <RedBoxPreview config={config} slideW={status?.slideWidth ?? 720} slideH={status?.slideHeight ?? 540} />

      <Separator>Konfiguration</Separator>

      {/* Margins */}
      <Stack tokens={{ childrenGap: 6 }}>
        <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
          <Text variant="smallPlus" style={{ fontWeight: 600 }}>Abstände (pt):</Text>
          <Toggle
            label="Alle verknüpfen"
            inlineLabel
            checked={linked}
            onChange={(_e, c) => setLinked(!!c)}
            styles={{ root: { marginBottom: 0 } }}
          />
        </Stack>

        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          {([
            { key: "marginTop"    as const, label: "Oben"   },
            { key: "marginRight"  as const, label: "Rechts" },
            { key: "marginBottom" as const, label: "Unten"  },
            { key: "marginLeft"   as const, label: "Links"  },
          ]).map(({ key, label }) => (
            <TextField
              key={key}
              label={label}
              value={String(config[key])}
              onChange={(_e, v) => updateMargin(key, v ?? "0")}
              type="number"
              min={0}
              suffix="pt"
              styles={{ root: { width: 80 } }}
            />
          ))}
        </Stack>

        {/* Schnell-Presets für Abstände */}
        <Stack horizontal tokens={{ childrenGap: 6 }}>
          <Text variant="xSmall" style={{ color: "#666", lineHeight: "24px" }}>Schnell:</Text>
          {[10, 15, 20, 25, 30].map((v) => (
            <ActionButton
              key={v}
              text={`${v}`}
              onClick={() => setConfig((prev) => ({
                ...prev, marginTop: v, marginRight: v, marginBottom: v, marginLeft: v,
              }))}
              styles={{ root: { height: 22, padding: "0 6px", fontSize: 11, minWidth: "auto" } }}
            />
          ))}
        </Stack>
      </Stack>

      <Separator />

      {/* Farbe */}
      <Stack tokens={{ childrenGap: 6 }}>
        <Text variant="smallPlus" style={{ fontWeight: 600 }}>Linienfarbe:</Text>
        <Stack horizontal tokens={{ childrenGap: 6 }}>
          {COLOR_PRESETS.map((c) => (
            <TooltipHost key={c.hex} content={c.label}>
              <div
                onClick={() => setConfig((prev) => ({ ...prev, color: c.hex }))}
                style={{
                  cursor: "pointer",
                  border: config.color.toUpperCase() === c.hex ? "2px solid #003366" : "2px solid transparent",
                  borderRadius: 3,
                  padding: 1,
                }}
              >
                <ColorSwatch color={c.hex} size={22} />
              </div>
            </TooltipHost>
          ))}
        </Stack>
        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
          <TextField
            label="Hex:"
            value={config.color}
            onChange={(_e, v) => setConfig((prev) => ({ ...prev, color: v ?? prev.color }))}
            prefix="#"
            styles={{ root: { width: 110 } }}
          />
          <ColorSwatch color={validColor} size={28} />
        </Stack>
      </Stack>

      {/* Linienstärke */}
      <Stack tokens={{ childrenGap: 6 }}>
        <Text variant="smallPlus" style={{ fontWeight: 600 }}>Linienstärke:</Text>
        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
          <TextField
            value={String(config.weight)}
            onChange={(_e, v) => {
              const n = parseFloat(v ?? "1");
              if (!isNaN(n) && n > 0) setConfig((prev) => ({ ...prev, weight: n }));
            }}
            type="number"
            min={0.25}
            step={0.25}
            suffix="pt"
            styles={{ root: { width: 80 } }}
          />
          <Stack horizontal tokens={{ childrenGap: 4 }}>
            {WEIGHT_PRESETS.map((w) => (
              <ActionButton
                key={w}
                text={`${w}`}
                onClick={() => setConfig((prev) => ({ ...prev, weight: w }))}
                styles={{ root: { height: 22, padding: "0 4px", fontSize: 11, minWidth: "auto" } }}
              />
            ))}
          </Stack>
        </Stack>
      </Stack>

      {/* Linientyp */}
      <Dropdown
        label="Linientyp:"
        options={DASH_OPTIONS}
        selectedKey={config.lineDash}
        onChange={(_e, opt) => opt && setConfig((prev) => ({
          ...prev, lineDash: opt.key as RedBoxConfig["lineDash"],
        }))}
        styles={{ root: { maxWidth: 180 } }}
      />

      <Separator>Aktionen</Separator>

      {isRunning ? (
        <Spinner size={SpinnerSize.small} label="Wird ausgeführt…" />
      ) : (
        <Stack tokens={{ childrenGap: 8 }}>
          {/* Primäre Aktionen */}
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text={status?.currentSlideHasBox ? "Entfernen" : "Einfügen"}
              iconProps={{ iconName: status?.currentSlideHasBox ? "Remove" : "Add" }}
              onClick={handleToggle}
              styles={{ root: { flex: 1 } }}
            />
            <DefaultButton
              text="Aktualisieren"
              iconProps={{ iconName: "Refresh" }}
              onClick={handleUpdate}
              disabled={!status?.currentSlideHasBox}
              styles={{ root: { flex: 1 } }}
            />
          </Stack>

          {/* Deck-Aktionen */}
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="Alle Folien"
              iconProps={{ iconName: "AllApps" }}
              onClick={handleAddAll}
              styles={{ root: { flex: 1, fontSize: 12 } }}
            />
            <DefaultButton
              text="Alle entfernen"
              iconProps={{ iconName: "Delete" }}
              onClick={handleRemoveAll}
              disabled={status ? status.totalBoxes === 0 : false}
              styles={{ root: { flex: 1, fontSize: 12 } }}
            />
          </Stack>

          {/* Einstellungen speichern */}
          <ActionButton
            text="Einstellungen speichern (ohne Anwenden)"
            iconProps={{ iconName: "Save" }}
            onClick={handleSave}
            styles={{ root: { height: 28, padding: 0, fontSize: 12 } }}
          />
        </Stack>
      )}

      <Text variant="xSmall" style={{ color: "#888" }}>
        Shape-Name: <code>INFRONT_REDBOX</code> · Auch über Ribbon-Buttons steuerbar.
      </Text>
    </Stack>
  );
};

// ─── Status-Badge ─────────────────────────────────────────────────────────────

interface StatusBadgeProps {
  status:     RedBoxStatus | null;
  onRefresh:  () => void;
}

const StatusBadge: React.FC<StatusBadgeProps> = ({ status, onRefresh }) => {
  if (!status) return null;

  const hasBox  = status.currentSlideHasBox;
  const bg      = hasBox ? "#DFF6DD" : "#F3F2F1";
  const iconCol = hasBox ? "#107C10" : "#A19F9D";
  const icon    = hasBox ? "CheckMark" : "Remove";

  return (
    <Stack
      horizontal
      verticalAlign="center"
      horizontalAlign="space-between"
      tokens={{ childrenGap: 8, padding: "6px 10px" }}
      styles={{ root: { background: bg, borderRadius: 4 } }}
    >
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        <Icon iconName={icon} style={{ fontSize: 14, color: iconCol }} />
        <Stack tokens={{ childrenGap: 0 }}>
          <Text variant="small" style={{ fontWeight: 600 }}>
            Aktuelle Folie: {hasBox ? "Red Box vorhanden" : "Keine Red Box"}
          </Text>
          <Text variant="xSmall" style={{ color: "#666" }}>
            Gesamt im Deck: {status.totalBoxes} ·
            Folie: {status.slideWidth}×{status.slideHeight} pt
          </Text>
        </Stack>
      </Stack>
      <TooltipHost content="Status aktualisieren">
        <ActionButton
          iconProps={{ iconName: "Refresh" }}
          onClick={onRefresh}
          styles={{ root: { height: 22, minWidth: "auto", padding: "0 4px" } }}
        />
      </TooltipHost>
    </Stack>
  );
};

// ─── CSS-Vorschau ─────────────────────────────────────────────────────────────

interface RedBoxPreviewProps {
  config:  RedBoxConfig;
  slideW:  number;
  slideH:  number;
}

const RedBoxPreview: React.FC<RedBoxPreviewProps> = ({ config, slideW, slideH }) => {
  const PREVIEW_W = 200;
  const scale     = PREVIEW_W / slideW;
  const PREVIEW_H = slideH * scale;

  const validColor = normalizeHex(config.color) ?? "#FF0000";

  const boxStyle: React.CSSProperties = {
    position:    "absolute",
    left:        config.marginLeft  * scale,
    top:         config.marginTop   * scale,
    width:       Math.max(0, (slideW - config.marginLeft - config.marginRight)  * scale),
    height:      Math.max(0, (slideH - config.marginTop  - config.marginBottom) * scale),
    border:      `${Math.max(1, config.weight * scale)}px ${config.lineDash === "solid" ? "solid" : config.lineDash === "dash" ? "dashed" : "dotted"} ${validColor}`,
    boxSizing:   "border-box",
    pointerEvents: "none",
  };

  const slideStyle: React.CSSProperties = {
    position:    "relative",
    width:       PREVIEW_W,
    height:      PREVIEW_H,
    background:  "#FFFFFF",
    border:      "1px solid #C8C6C4",
    borderRadius: 2,
    overflow:    "hidden",
    marginBottom: 4,
  };

  return (
    <Stack tokens={{ childrenGap: 4 }} horizontalAlign="center">
      <Text variant="xSmall" style={{ color: "#666" }}>Vorschau (nicht maßstabsgetreu):</Text>
      <div style={slideStyle}>
        {/* Hintergrund-Raster-Andeutung */}
        <div style={{ position: "absolute", inset: 0, background: "repeating-linear-gradient(45deg,#F8F8F8,#F8F8F8 2px,#fff 2px,#fff 10px)" }} />
        <div style={boxStyle} />
      </div>
      <Text variant="xSmall" style={{ color: "#888" }}>
        T:{config.marginTop} R:{config.marginRight} B:{config.marginBottom} L:{config.marginLeft} pt ·
        Linie: {validColor} {config.weight}pt
      </Text>
    </Stack>
  );
};

export default RedBoxPanel;
