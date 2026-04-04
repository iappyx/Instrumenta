/**
 * FindReplacePanel.tsx
 * Feature-Panel: Suchen & Ersetzen (vollständige Implementierung).
 *
 * 3 Tabs: Text · Farbe · Font
 *
 * Hinweis Text-Tab: textRange.text-Zuweisung → intra-Shape-Formatierung
 * (gemischte Schriftgrößen, partielles Bold) geht bei Ersatz verloren.
 * Dokumentiert in TESTING.md (Kategorie B).
 */

import * as React from "react";
import { Stack }              from "@fluentui/react/lib/Stack";
import { Text }               from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { TextField }          from "@fluentui/react/lib/TextField";
import { Checkbox }           from "@fluentui/react/lib/Checkbox";
import { Pivot, PivotItem }   from "@fluentui/react/lib/Pivot";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { Slider }             from "@fluentui/react/lib/Slider";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Separator }          from "@fluentui/react/lib/Separator";
import { Icon }               from "@fluentui/react/lib/Icon";

import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import ColorSwatch            from "../shared/ColorSwatch";
import { normalizeHex }       from "../../../utils/colorUtils";
import {
  SearchScope, TextQuery, ColorQuery, FontReplaceOptions,
  TextMatch, ColorMatch,
  previewText, replaceText,
  previewColor, replaceColor,
  collectFonts, replaceFont,
} from "../../../features/findReplace/FindReplaceService";

interface Notification {
  message: string;
  type:    NotificationType;
}

const SCOPE_OPTIONS: IDropdownOption[] = [
  { key: "allSlides",      text: "Alle Slides" },
  { key: "currentSlide",   text: "Aktuelle Slide" },
  { key: "selectedSlides", text: "Selektierte Slides" },
];

const TARGET_LABEL: Record<"fill" | "line" | "font", string> = {
  fill: "Füllung", line: "Linie", font: "Schrift",
};

// ─── Hauptkomponente ──────────────────────────────────────────────────────────

const FindReplacePanel: React.FC = () => {
  const [notification, setNotification] = React.useState<Notification | null>(null);

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Suchen &amp; Ersetzen</Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      <Pivot>
        <PivotItem headerText="Text" itemIcon="Search">
          <TextTab setParentNotification={setNotification} />
        </PivotItem>
        <PivotItem headerText="Farbe" itemIcon="BucketColor">
          <ColorTab setParentNotification={setNotification} />
        </PivotItem>
        <PivotItem headerText="Font" itemIcon="Font">
          <FontTab setParentNotification={setNotification} />
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

// ─── Tab 1: Text ──────────────────────────────────────────────────────────────

const TextTab: React.FC<{ setParentNotification: (n: Notification | null) => void }> = ({
  setParentNotification,
}) => {
  const [search, setSearch]           = React.useState("");
  const [replacement, setReplacement] = React.useState("");
  const [caseSensitive, setCaseSensitive] = React.useState(false);
  const [wholeWord, setWholeWord]     = React.useState(false);
  const [scope, setScope]             = React.useState<SearchScope>("allSlides");
  const [isRunning, setIsRunning]     = React.useState(false);
  const [matches, setMatches]         = React.useState<TextMatch[] | null>(null);

  const query: TextQuery = { search, replacement, caseSensitive, wholeWord };

  const handlePreview = async () => {
    if (!search) { setParentNotification({ message: "Suchtext eingeben.", type: "warning" }); return; }
    setIsRunning(true);
    setMatches(null);
    try {
      const found = await previewText(query, scope);
      setMatches(found);
      const total = found.reduce((s, m) => s + m.count, 0);
      setParentNotification(
        found.length === 0
          ? { message: "Kein Treffer gefunden.", type: "info" }
          : { message: `${total} Treffer in ${found.length} Shape(s) gefunden.`, type: "info" }
      );
    } catch (err) {
      setParentNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally {
      setIsRunning(false);
    }
  };

  const handleReplaceAll = async () => {
    if (!search) { setParentNotification({ message: "Suchtext eingeben.", type: "warning" }); return; }
    setIsRunning(true);
    setMatches(null);
    try {
      const result = await replaceText(query, scope);
      setParentNotification(
        result.replaced === 0
          ? { message: "Keine Treffer ersetzt.", type: "warning" }
          : { message: `${result.replaced} Shape(s) geändert.${result.errors.length ? ` Fehler: ${result.errors.join(", ")}` : ""}`, type: "success" }
      );
    } catch (err) {
      setParentNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally {
      setIsRunning(false);
    }
  };

  return (
    <Stack tokens={{ childrenGap: 10, padding: "12px 0 0 0" }}>
      <MessageBar messageBarType={MessageBarType.warning} isMultiline>
        Achtung: Beim Ersetzen geht intra-Shape-Formatierung (gemischte Schriftgrößen,
        partielles Bold) verloren. Office.js-Einschränkung – TESTING.md.
      </MessageBar>

      <TextField
        label="Suchen:"
        value={search}
        onChange={(_e, v) => { setSearch(v ?? ""); setMatches(null); }}
        placeholder="Suchtext…"
      />
      <TextField
        label="Ersetzen durch:"
        value={replacement}
        onChange={(_e, v) => setReplacement(v ?? "")}
        placeholder="Ersatztext (leer = löschen)…"
      />

      <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
        <Checkbox label="Groß-/Kleinschreibung" checked={caseSensitive} onChange={(_e, c) => setCaseSensitive(!!c)} />
        <Checkbox label="Ganzes Wort" checked={wholeWord} onChange={(_e, c) => setWholeWord(!!c)} />
      </Stack>

      <Dropdown
        label="Scope:"
        options={SCOPE_OPTIONS}
        selectedKey={scope}
        onChange={(_e, o) => o && setScope(o.key as SearchScope)}
        styles={{ root: { maxWidth: 220 } }}
      />

      {/* Vorschau-Ergebnisse */}
      {matches && matches.length > 0 && (
        <Stack
          tokens={{ padding: "6px 8px", childrenGap: 3 }}
          styles={{ root: { background: "#FFF4CE", borderRadius: 4, maxHeight: 150, overflowY: "auto" } }}
        >
          {matches.map((m, i) => (
            <Stack key={i} horizontal tokens={{ childrenGap: 6 }}>
              <Text variant="xSmall" style={{ color: "#605E5C", minWidth: 50 }}>
                Folie {m.slideIndex + 1}
              </Text>
              <Text variant="xSmall" style={{ fontWeight: 600, minWidth: 80, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                {m.shapeName}
              </Text>
              <Text variant="xSmall" style={{ color: "#A19F9D" }}>({m.count}×)</Text>
            </Stack>
          ))}
        </Stack>
      )}

      {isRunning ? (
        <Spinner size={SpinnerSize.small} label="Arbeitet…" />
      ) : (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Vorschau" onClick={handlePreview} disabled={!search} styles={{ root: { flex: 1 } }} />
          <PrimaryButton text="Alle ersetzen" onClick={handleReplaceAll} disabled={!search} styles={{ root: { flex: 1 } }} />
        </Stack>
      )}
    </Stack>
  );
};

// ─── Tab 2: Farbe ─────────────────────────────────────────────────────────────

const ColorTab: React.FC<{ setParentNotification: (n: Notification | null) => void }> = ({
  setParentNotification,
}) => {
  const [searchHex, setSearchHex]   = React.useState("#003366");
  const [replaceHex, setReplaceHex] = React.useState("#1A5276");
  const [tolerance, setTolerance]   = React.useState(10);
  const [targetFill, setTargetFill] = React.useState(true);
  const [targetLine, setTargetLine] = React.useState(false);
  const [targetFont, setTargetFont] = React.useState(false);
  const [scope, setScope]           = React.useState<SearchScope>("allSlides");
  const [isRunning, setIsRunning]   = React.useState(false);
  const [matches, setMatches]       = React.useState<ColorMatch[] | null>(null);

  const normSearch  = normalizeHex(searchHex);
  const normReplace = normalizeHex(replaceHex);

  const query: ColorQuery = {
    searchHex:  searchHex,
    replaceHex: replaceHex,
    tolerance,
    targetFill, targetLine, targetFont,
  };

  const handlePreview = async () => {
    if (!normSearch) { setParentNotification({ message: "Gültige Suchfarbe eingeben.", type: "warning" }); return; }
    if (!targetFill && !targetLine && !targetFont) {
      setParentNotification({ message: "Bitte mindestens ein Ziel wählen.", type: "warning" }); return;
    }
    setIsRunning(true); setMatches(null);
    try {
      const found = await previewColor(query, scope);
      setMatches(found);
      setParentNotification(
        found.length === 0
          ? { message: "Keine passenden Farben gefunden.", type: "info" }
          : { message: `${found.length} Shape(s) mit passender Farbe gefunden.`, type: "info" }
      );
    } catch (err) {
      setParentNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  const handleReplaceAll = async () => {
    if (!normSearch || !normReplace) {
      setParentNotification({ message: "Gültige Hex-Farben eingeben.", type: "error" }); return;
    }
    if (!targetFill && !targetLine && !targetFont) {
      setParentNotification({ message: "Bitte mindestens ein Ziel wählen.", type: "warning" }); return;
    }
    setIsRunning(true); setMatches(null);
    try {
      const result = await replaceColor(query, scope);
      setParentNotification(
        result.replaced === 0
          ? { message: "Keine Farben ersetzt.", type: "warning" }
          : { message: `${result.replaced} Shape(s) umgefärbt.`, type: "success" }
      );
    } catch (err) {
      setParentNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  return (
    <Stack tokens={{ childrenGap: 10, padding: "12px 0 0 0" }}>
      {/* Suchfarbe */}
      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <TextField
          label="Suchfarbe (Hex):"
          value={searchHex}
          onChange={(_e, v) => { setSearchHex(v ?? ""); setMatches(null); }}
          styles={{ root: { flex: 1 } }}
          errorMessage={searchHex && !normSearch ? "Format: #RRGGBB" : undefined}
        />
        <div style={{ marginBottom: 4 }}>
          <ColorSwatch color={normSearch ?? "transparent"} size={32} />
        </div>
      </Stack>

      {/* Ersatzfarbe */}
      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <TextField
          label="Ersatzfarbe (Hex):"
          value={replaceHex}
          onChange={(_e, v) => setReplaceHex(v ?? "")}
          styles={{ root: { flex: 1 } }}
          errorMessage={replaceHex && !normReplace ? "Format: #RRGGBB" : undefined}
        />
        <div style={{ marginBottom: 4 }}>
          <ColorSwatch color={normReplace ?? "transparent"} size={32} />
        </div>
      </Stack>

      <Slider
        label={`Toleranz: ${tolerance}`}
        min={0} max={30} step={1}
        value={tolerance}
        onChange={(v) => setTolerance(v)}
        showValue={false}
        styles={{ root: { maxWidth: 280 } }}
      />

      <Stack tokens={{ childrenGap: 4 }}>
        <Text variant="smallPlus" style={{ fontWeight: 600 }}>Anwenden auf:</Text>
        <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
          <Checkbox label="Füllfarbe"  checked={targetFill} onChange={(_e, c) => setTargetFill(!!c)} />
          <Checkbox label="Linienfarbe" checked={targetLine} onChange={(_e, c) => setTargetLine(!!c)} />
          <Checkbox label="Schriftfarbe" checked={targetFont} onChange={(_e, c) => setTargetFont(!!c)} />
        </Stack>
      </Stack>

      <Dropdown
        label="Scope:"
        options={SCOPE_OPTIONS}
        selectedKey={scope}
        onChange={(_e, o) => o && setScope(o.key as SearchScope)}
        styles={{ root: { maxWidth: 220 } }}
      />

      {/* Vorschau */}
      {matches && matches.length > 0 && (
        <Stack
          tokens={{ padding: "6px 8px", childrenGap: 3 }}
          styles={{ root: { background: "#FFF4CE", borderRadius: 4, maxHeight: 130, overflowY: "auto" } }}
        >
          {matches.slice(0, 20).map((m, i) => (
            <Stack key={i} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
              <Text variant="xSmall" style={{ color: "#605E5C", minWidth: 50 }}>Folie {m.slideIndex + 1}</Text>
              <Text variant="xSmall" style={{ fontWeight: 600, flex: 1, overflow: "hidden", textOverflow: "ellipsis" }}>{m.shapeName}</Text>
              <Text variant="xSmall" style={{ color: "#666" }}>{TARGET_LABEL[m.target]}</Text>
              <ColorSwatch color={m.foundColor} size={12} />
            </Stack>
          ))}
          {matches.length > 20 && (
            <Text variant="xSmall" style={{ color: "#666" }}>… und {matches.length - 20} weitere</Text>
          )}
        </Stack>
      )}

      {isRunning ? (
        <Spinner size={SpinnerSize.small} label="Arbeitet…" />
      ) : (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Vorschau" onClick={handlePreview} styles={{ root: { flex: 1 } }} />
          <PrimaryButton text="Alle ersetzen" onClick={handleReplaceAll} styles={{ root: { flex: 1 } }} />
        </Stack>
      )}
    </Stack>
  );
};

// ─── Tab 3: Font ──────────────────────────────────────────────────────────────

const FontTab: React.FC<{ setParentNotification: (n: Notification | null) => void }> = ({
  setParentNotification,
}) => {
  const [fonts, setFonts]           = React.useState<string[]>([]);
  const [fromFont, setFromFont]     = React.useState<string>("");
  const [toFont, setToFont]         = React.useState<string>("");
  const [keepSize, setKeepSize]     = React.useState(true);
  const [keepFormat, setKeepFormat] = React.useState(true);
  const [scope, setScope]           = React.useState<SearchScope>("allSlides");
  const [isLoading, setIsLoading]   = React.useState(false);
  const [isRunning, setIsRunning]   = React.useState(false);

  const options: FontReplaceOptions = { keepSize, keepFormatting: keepFormat };

  const handleCollect = async () => {
    setIsLoading(true);
    try {
      const found = await collectFonts();
      setFonts(found);
      setParentNotification(
        found.length === 0
          ? { message: "Keine Schriftarten im Deck gefunden.", type: "info" }
          : { message: `${found.length} Schriftarten gefunden.`, type: "info" }
      );
    } catch (err) {
      setParentNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsLoading(false); }
  };

  const handleReplaceAll = async () => {
    if (!fromFont || !toFont) {
      setParentNotification({ message: "Quell- und Ziel-Schriftart wählen.", type: "warning" }); return;
    }
    if (fromFont === toFont) {
      setParentNotification({ message: "Quell- und Ziel-Schriftart sind identisch.", type: "warning" }); return;
    }
    setIsRunning(true);
    try {
      const result = await replaceFont(fromFont, toFont, options, scope);
      setParentNotification(
        result.replaced === 0
          ? { message: `Keine Shapes mit Schriftart „${fromFont}" gefunden.`, type: "warning" }
          : { message: `${result.replaced} Shape(s) auf „${toFont}" umgestellt.`, type: "success" }
      );
      // Font-Liste neu laden
      await handleCollect();
    } catch (err) {
      setParentNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  const fontOptions: IDropdownOption[] = fonts.map((f) => ({ key: f, text: f }));

  return (
    <Stack tokens={{ childrenGap: 10, padding: "12px 0 0 0" }}>
      {/* Schriftarten laden */}
      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
        <Text variant="small" style={{ flex: 1, color: "#555" }}>
          {fonts.length > 0
            ? `${fonts.length} Schriftarten im Deck:`
            : "Schriftarten im Deck laden:"}
        </Text>
        {isLoading ? (
          <Spinner size={SpinnerSize.small} />
        ) : (
          <DefaultButton
            text="Laden"
            iconProps={{ iconName: "Refresh" }}
            onClick={handleCollect}
            styles={{ root: { minWidth: 80 } }}
          />
        )}
      </Stack>

      {fonts.length > 0 && (
        <Stack horizontal wrap tokens={{ childrenGap: 4 }}>
          {fonts.map((f) => (
            <span
              key={f}
              onClick={() => setFromFont(f)}
              style={{
                padding: "2px 8px",
                background: f === fromFont ? "#003366" : "#EDEBE9",
                color: f === fromFont ? "#fff" : "#323130",
                borderRadius: 12,
                fontSize: 12,
                cursor: "pointer",
              }}
            >
              {f}
            </span>
          ))}
        </Stack>
      )}

      <Separator />

      {/* Quelle */}
      <Dropdown
        label="Ersetze Schriftart:"
        placeholder="Schriftart wählen…"
        options={fontOptions}
        selectedKey={fromFont || undefined}
        onChange={(_e, o) => o && setFromFont(String(o.key))}
        disabled={fonts.length === 0}
      />

      {/* Ziel */}
      <TextField
        label="Durch Schriftart:"
        value={toFont}
        onChange={(_e, v) => setToFont(v ?? "")}
        placeholder="Ziel-Schriftname eingeben…"
        description="Genauer Name (z.B. Calibri, Arial)"
      />

      <Stack tokens={{ childrenGap: 4 }}>
        <Checkbox label="Schriftgröße beibehalten" checked={keepSize} onChange={(_e, c) => setKeepSize(!!c)} />
        <Checkbox label="Formatierung beibehalten (Bold, Italic)" checked={keepFormat} onChange={(_e, c) => setKeepFormat(!!c)} />
      </Stack>

      <Dropdown
        label="Scope:"
        options={SCOPE_OPTIONS}
        selectedKey={scope}
        onChange={(_e, o) => o && setScope(o.key as SearchScope)}
        styles={{ root: { maxWidth: 220 } }}
      />

      {isRunning ? (
        <Spinner size={SpinnerSize.small} label="Schriftarten werden ersetzt…" />
      ) : (
        <PrimaryButton
          text="Alle ersetzen"
          iconProps={{ iconName: "Font" }}
          onClick={handleReplaceAll}
          disabled={!fromFont || !toFont}
          styles={{ root: { width: "100%", marginTop: 4 } }}
        />
      )}
    </Stack>
  );
};

export default FindReplacePanel;
