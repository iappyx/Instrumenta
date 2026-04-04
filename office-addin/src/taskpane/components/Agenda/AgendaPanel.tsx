/**
 * AgendaPanel.tsx
 * Feature-Panel: Agenda-Wizard (vollständige Implementierung).
 *
 * Workflow:
 * 1. Sektionen konfigurieren (Name + Foliennummer, bis zu 8)
 * 2. Formatierung festlegen (aktiv / inaktiv)
 * 3. Aktive Sektion wählen
 * 4. Auf aktueller Slide einfügen ODER alle Agenda-Shapes aktualisieren
 *
 * Auto-Update: NICHT implementiert – Office.js auf Mac unterstützt
 * document.selectionChanged-Events nicht zuverlässig. Manueller Update-Button.
 * Dokumentiert in TESTING.md.
 *
 * Shape-Erkennung: INFRONT_AGENDA_ITEM_NN-Namenspräfix (1.1+, kein Tag-API nötig).
 * Shape Tags als optionales Supplement (1.5+, try/catch).
 */

import * as React from "react";
import { Stack }            from "@fluentui/react/lib/Stack";
import { Text }             from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton, ActionButton } from "@fluentui/react/lib/Button";
import { TextField }        from "@fluentui/react/lib/TextField";
import { Checkbox }         from "@fluentui/react/lib/Checkbox";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Separator }        from "@fluentui/react/lib/Separator";
import { Pivot, PivotItem } from "@fluentui/react/lib/Pivot";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Icon }             from "@fluentui/react/lib/Icon";

import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import ColorSwatch          from "../shared/ColorSwatch";
import { normalizeHex }     from "../../../utils/colorUtils";
import {
  loadAgendaConfig, saveAgendaConfig,
  insertAgendaOnCurrentSlide, updateAllAgendaShapes,
  removeAllAgendaShapes, findAgendaSlides,
  AgendaConfig, AgendaSection, AgendaFormat,
  DEFAULT_AGENDA_CONFIG,
} from "../../../features/agenda/AgendaService";

interface Notification {
  message: string;
  type:    NotificationType;
}

const MAX_SECTIONS = 8;

// ─── Hauptkomponente ──────────────────────────────────────────────────────────

const AgendaPanel: React.FC = () => {
  const [config, setConfig]             = React.useState<AgendaConfig>(() => loadAgendaConfig());
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]       = React.useState(false);
  const [agendaSlides, setAgendaSlides] = React.useState<Array<{slideIndex: number; count: number}>>([]);

  // Beim Mounten: vorhandene Agenda-Slides suchen
  React.useEffect(() => {
    findAgendaSlides().then(setAgendaSlides).catch(() => {});
  }, []);

  const activeSections = config.sections.filter((s) => s.name.trim().length > 0);

  // ── Konfiguration speichern ─────────────────────────────────────────────────
  const handleSaveConfig = async () => {
    try {
      await saveAgendaConfig(config);
      setNotification({ message: "Konfiguration gespeichert.", type: "success" });
    } catch (err) {
      setNotification({ message: `Fehler beim Speichern: ${err instanceof Error ? err.message : err}`, type: "error" });
    }
  };

  // ── Auf aktueller Slide einfügen ────────────────────────────────────────────
  const handleInsert = async () => {
    if (activeSections.length === 0) {
      setNotification({ message: "Bitte mindestens eine Sektion konfigurieren.", type: "warning" });
      return;
    }
    setIsRunning(true);
    setNotification(null);
    try {
      await saveAgendaConfig(config);
      const result = await insertAgendaOnCurrentSlide(config);
      setAgendaSlides(await findAgendaSlides());
      setNotification({
        message: `${result.updated} Agenda-Shapes eingefügt.${result.errors.length ? ` Fehler: ${result.errors.join(", ")}` : ""}`,
        type: result.errors.length > 0 ? "warning" : "success",
      });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ── Alle Agenda-Shapes aktualisieren ────────────────────────────────────────
  const handleUpdateAll = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      await saveAgendaConfig(config);
      const result = await updateAllAgendaShapes(config);
      if (result.updated === 0) {
        setNotification({ message: "Keine Agenda-Shapes im Deck gefunden. Zuerst einfügen.", type: "info" });
      } else {
        setNotification({
          message: `${result.updated} Agenda-Shape(s) aktualisiert.`,
          type: "success",
        });
      }
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ── Alle Agenda-Shapes entfernen ────────────────────────────────────────────
  const handleRemoveAll = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      const removed = await removeAllAgendaShapes();
      setAgendaSlides([]);
      setNotification({
        message: removed > 0 ? `${removed} Agenda-Shape(s) entfernt.` : "Keine Agenda-Shapes gefunden.",
        type: removed > 0 ? "success" : "info",
      });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ── Sektion aktualisieren ────────────────────────────────────────────────────
  const updateSection = (i: number, field: keyof AgendaSection, value: string | number) => {
    setConfig((prev) => {
      const sections = [...prev.sections];
      sections[i] = { ...sections[i], [field]: value };
      return { ...prev, sections };
    });
  };

  const updateActiveFormat = (field: keyof AgendaFormat, value: string | number | boolean) => {
    setConfig((prev) => ({ ...prev, activeFormat: { ...prev.activeFormat, [field]: value } }));
  };

  const updateInactiveFormat = (field: keyof AgendaFormat, value: string | number | boolean) => {
    setConfig((prev) => ({ ...prev, inactiveFormat: { ...prev.inactiveFormat, [field]: value } }));
  };

  // Aktive-Sektion-Optionen für Dropdown
  const activeIndexOptions: IDropdownOption[] = [
    { key: -1, text: "(Keine aktive Sektion)" },
    ...activeSections.map((s, i) => ({ key: i, text: `${i + 1}. ${s.name}` })),
  ];

  // ─── Render ─────────────────────────────────────────────────────────────────

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Agenda-Assistent</Text>

      <MessageBar messageBarType={MessageBarType.info} isMultiline>
        Auto-Update bei Folien-Wechsel nicht verfügbar (Mac/Office.js-Einschränkung).
        Aktive Sektion manuell wählen und „Alle aktualisieren" klicken.
      </MessageBar>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {/* Status: vorhandene Agenda-Slides */}
      {agendaSlides.length > 0 && (
        <Stack
          horizontal
          verticalAlign="center"
          tokens={{ childrenGap: 6 }}
          styles={{ root: { background: "#EFF6FC", borderRadius: 4, padding: "6px 10px" } }}
        >
          <Icon iconName="TableGroup" style={{ fontSize: 14, color: "#0078D4" }} />
          <Text variant="small" style={{ color: "#0078D4" }}>
            Agenda auf {agendaSlides.length} Folie(n): {agendaSlides.map((s) => s.slideIndex + 1).join(", ")}
          </Text>
        </Stack>
      )}

      <Pivot>
        {/* ── Tab 1: Sektionen ── */}
        <PivotItem headerText="Sektionen" itemIcon="BulletedList">
          <Stack tokens={{ childrenGap: 8, padding: "12px 0 0 0" }}>
            <Text variant="small" style={{ color: "#555" }}>
              Bis zu {MAX_SECTIONS} Sektionen. Leere Einträge werden übersprungen.
            </Text>

            {Array.from({ length: MAX_SECTIONS }, (_, i) => (
              <Stack key={i} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                <Text
                  variant="xSmall"
                  style={{ minWidth: 16, color: "#888", fontWeight: 600, textAlign: "right" }}
                >
                  {i + 1}.
                </Text>
                <TextField
                  placeholder={`Sektion ${i + 1}`}
                  value={config.sections[i]?.name ?? ""}
                  onChange={(_e, v) => updateSection(i, "name", v ?? "")}
                  styles={{ root: { flex: 1 } }}
                />
                <TextField
                  placeholder="Folie"
                  value={config.sections[i]?.slideNumber > 0 ? String(config.sections[i].slideNumber) : ""}
                  onChange={(_e, v) => {
                    const n = parseInt(v ?? "0", 10);
                    updateSection(i, "slideNumber", isNaN(n) ? 0 : n);
                  }}
                  type="number"
                  min={0}
                  styles={{ root: { width: 56 } }}
                  suffix="#"
                />
              </Stack>
            ))}

            <Checkbox
              label="Foliennummer anzeigen"
              checked={config.showPageNumbers}
              onChange={(_e, c) => setConfig((prev) => ({ ...prev, showPageNumbers: !!c }))}
            />

            <DefaultButton
              text="Konfiguration speichern"
              iconProps={{ iconName: "Save" }}
              onClick={handleSaveConfig}
              styles={{ root: { width: "100%", marginTop: 4 } }}
            />
          </Stack>
        </PivotItem>

        {/* ── Tab 2: Formatierung ── */}
        <PivotItem headerText="Format" itemIcon="FontColor">
          <Stack tokens={{ childrenGap: 12, padding: "12px 0 0 0" }}>
            <FormatEditor
              label="Aktive Sektion"
              format={config.activeFormat}
              onUpdate={updateActiveFormat}
            />
            <Separator />
            <FormatEditor
              label="Inaktive Sektionen"
              format={config.inactiveFormat}
              onUpdate={updateInactiveFormat}
            />
            <Separator />
            <Stack tokens={{ childrenGap: 6 }}>
              <Text variant="smallPlus" style={{ fontWeight: 600 }}>Layout (pt):</Text>
              <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
                {(["left", "top", "width", "itemHeight", "gap"] as const).map((field) => (
                  <TextField
                    key={field}
                    label={field === "itemHeight" ? "Höhe" : field === "gap" ? "Abstand" : field === "left" ? "Links" : field === "top" ? "Oben" : "Breite"}
                    value={String(config[field])}
                    onChange={(_e, v) => {
                      const n = parseFloat(v ?? "0");
                      if (!isNaN(n)) setConfig((prev) => ({ ...prev, [field]: n }));
                    }}
                    type="number"
                    suffix="pt"
                    styles={{ root: { width: 72 } }}
                  />
                ))}
              </Stack>
            </Stack>
            <DefaultButton
              text="Formatierung speichern"
              iconProps={{ iconName: "Save" }}
              onClick={handleSaveConfig}
              styles={{ root: { width: "100%" } }}
            />
          </Stack>
        </PivotItem>

        {/* ── Tab 3: Aktionen ── */}
        <PivotItem headerText="Aktionen" itemIcon="Play">
          <Stack tokens={{ childrenGap: 12, padding: "12px 0 0 0" }}>
            {/* Aktive Sektion */}
            <Dropdown
              label="Aktive Sektion:"
              options={activeIndexOptions}
              selectedKey={config.activeIndex}
              onChange={(_e, opt) => opt && setConfig((prev) => ({ ...prev, activeIndex: Number(opt.key) }))}
              disabled={activeSections.length === 0}
            />

            {activeSections.length === 0 && (
              <MessageBar messageBarType={MessageBarType.warning}>
                Noch keine Sektionen konfiguriert. Bitte im Tab „Sektionen" ausfüllen.
              </MessageBar>
            )}

            <Separator />

            {isRunning ? (
              <Spinner size={SpinnerSize.small} label="Arbeitet…" />
            ) : (
              <Stack tokens={{ childrenGap: 8 }}>
                <PrimaryButton
                  text="Auf aktueller Folie einfügen"
                  iconProps={{ iconName: "Add" }}
                  onClick={handleInsert}
                  disabled={activeSections.length === 0}
                  styles={{ root: { width: "100%" } }}
                />
                <DefaultButton
                  text="Alle Agenda-Shapes aktualisieren"
                  iconProps={{ iconName: "Refresh" }}
                  onClick={handleUpdateAll}
                  disabled={agendaSlides.length === 0}
                  styles={{ root: { width: "100%" } }}
                  title="Aktualisiert Text und Formatierung aller INFRONT_AGENDA_ITEM_* Shapes im Deck"
                />
                <DefaultButton
                  text={`Alle Agenda-Shapes entfernen${agendaSlides.length > 0 ? ` (${agendaSlides.length} Folie(n))` : ""}`}
                  iconProps={{ iconName: "Delete" }}
                  onClick={handleRemoveAll}
                  disabled={agendaSlides.length === 0}
                  styles={{ root: { width: "100%" } }}
                />
              </Stack>
            )}

            <Separator />
            <Text variant="small" style={{ color: "#666" }}>
              <strong>Shapes im Deck:</strong> INFRONT_AGENDA_ITEM_01–08
              <br />
              <strong>Undo:</strong> Kein natives Office.js-Undo. Shapes manuell entfernen und neu einfügen.
              <br />
              <strong>Auto-Update:</strong> Nicht verfügbar auf Mac (kein zuverlässiges
              Folien-Wechsel-Event in Office.js).
            </Text>
          </Stack>
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

// ─── Format-Editor ────────────────────────────────────────────────────────────

interface FormatEditorProps {
  label:    string;
  format:   AgendaFormat;
  onUpdate: (field: keyof AgendaFormat, value: string | number | boolean) => void;
}

const FormatEditor: React.FC<FormatEditorProps> = ({ label, format, onUpdate }) => {
  const normColor = normalizeHex(format.color);

  return (
    <Stack tokens={{ childrenGap: 6 }}>
      <Text variant="smallPlus" style={{ fontWeight: 600 }}>{label}:</Text>
      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <TextField
          label="Farbe (Hex):"
          value={format.color}
          onChange={(_e, v) => onUpdate("color", v ?? "")}
          styles={{ root: { flex: 1 } }}
        />
        <div style={{ marginBottom: 4 }}>
          <ColorSwatch color={normColor ?? "transparent"} size={28} />
        </div>
      </Stack>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <TextField
          label="Größe (pt):"
          value={String(format.fontSize)}
          onChange={(_e, v) => {
            const n = parseFloat(v ?? "0");
            if (!isNaN(n) && n > 0) onUpdate("fontSize", n);
          }}
          type="number"
          min={6}
          suffix="pt"
          styles={{ root: { width: 80 } }}
        />
        <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: 26 } }}>
          <Checkbox label="Fett"    checked={format.bold}   onChange={(_e, c) => onUpdate("bold",   !!c)} />
          <Checkbox label="Kursiv"  checked={format.italic} onChange={(_e, c) => onUpdate("italic", !!c)} />
        </Stack>
      </Stack>
    </Stack>
  );
};

export default AgendaPanel;
