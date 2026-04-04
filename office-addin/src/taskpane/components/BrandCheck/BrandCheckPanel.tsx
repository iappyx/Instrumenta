/**
 * BrandCheckPanel.tsx
 * Feature-Panel: Brand Compliance Check (vollständige Implementierung).
 *
 * Prüft alle Shapes aller Slides auf Brand-Konformität.
 * Zeigt Verstöße gruppiert nach Slide mit Navigation und Fix-Aktionen.
 * CSV-Export der vollständigen Verstoßliste.
 *
 * API-Hinweis: Font-Details in Tabellenzellen (pro Run) sind über
 * Office.js nicht zugänglich – nur Zell-Text prüfbar. Dokumentiert in TESTING.md.
 */

import * as React from "react";
import { Stack }              from "@fluentui/react/lib/Stack";
import { Text }               from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton, ActionButton } from "@fluentui/react/lib/Button";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Separator }          from "@fluentui/react/lib/Separator";
import { ProgressIndicator }  from "@fluentui/react/lib/ProgressIndicator";
import { Icon }               from "@fluentui/react/lib/Icon";

import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import ColorSwatch            from "../shared/ColorSwatch";
import {
  runBrandCheck,
  fixViolations,
  goToSlide,
  exportViolationsAsCsv,
  Violation,
  ViolationType,
  BrandCheckResult,
} from "../../../features/brandCheck/BrandCheckService";
import { loadBrandConfig, saveBrandConfig } from "../../../services/config/ConfigService";
import { DEFAULT_BRAND_CONFIG }             from "../../../services/config/BrandConfig";

interface Notification {
  message: string;
  type:    NotificationType;
}

interface ProgressState {
  current: number;
  total:   number;
}

// ─── Hilfsfunktionen ──────────────────────────────────────────────────────────

const TYPE_LABELS: Record<ViolationType, string> = {
  "font-name":  "Schriftart",
  "font-size":  "Schriftgröße",
  "font-color": "Schriftfarbe",
  "fill-color": "Füllfarbe",
  "line-color": "Linienfarbe",
};

const TYPE_ICONS: Record<ViolationType, string> = {
  "font-name":  "Font",
  "font-size":  "FontSize",
  "font-color": "FontColor",
  "fill-color": "BucketColor",
  "line-color": "LineStyle",
};

function groupBySlide(violations: Violation[]): Map<number, Violation[]> {
  const map = new Map<number, Violation[]>();
  for (const v of violations) {
    const list = map.get(v.slideIndex) ?? [];
    list.push(v);
    map.set(v.slideIndex, list);
  }
  return map;
}

function summarize(violations: Violation[]): Record<ViolationType, number> {
  const counts: Record<ViolationType, number> = {
    "font-name":  0, "font-size":  0,
    "font-color": 0, "fill-color": 0, "line-color": 0,
  };
  for (const v of violations) counts[v.type]++;
  return counts;
}

// ─── Komponente ───────────────────────────────────────────────────────────────

const BrandCheckPanel: React.FC = () => {
  const [isRunning, setIsRunning]         = React.useState(false);
  const [isFixing, setIsFixing]           = React.useState(false);
  const [result, setResult]               = React.useState<BrandCheckResult | null>(null);
  const [progress, setProgress]           = React.useState<ProgressState | null>(null);
  const [notification, setNotification]   = React.useState<Notification | null>(null);
  const [selectedProfile, setSelectedProfile] = React.useState<string>(() =>
    loadBrandConfig().activeProfile
  );
  const [expandedSlides, setExpandedSlides]   = React.useState<Set<number>>(new Set());
  const [selectedViolations, setSelectedViolations] = React.useState<Set<number>>(new Set());

  // Profile-Optionen aus Konfiguration
  const config         = React.useMemo(() => loadBrandConfig(), []);
  const profileOptions = React.useMemo<IDropdownOption[]>(
    () => config.profiles.map((p) => ({ key: p.name, text: p.name })),
    [config]
  );

  // ── Scan starten ────────────────────────────────────────────────────────────
  const handleRun = async () => {
    setIsRunning(true);
    setResult(null);
    setNotification(null);
    setExpandedSlides(new Set());
    setSelectedViolations(new Set());
    setProgress({ current: 0, total: 1 });

    try {
      // Aktives Profil in Config speichern
      if (config.activeProfile !== selectedProfile) {
        config.activeProfile = selectedProfile;
        await saveBrandConfig(config).catch(() => { /* nicht kritisch */ });
      }

      const checkResult = await runBrandCheck((done, total) => {
        setProgress({ current: done, total });
      });

      setResult(checkResult);

      const n = checkResult.violations.length;
      if (n === 0) {
        setNotification({
          message: `Keine Verstöße gefunden! ${checkResult.shapeCount} Shapes geprüft (${checkResult.durationMs} ms).`,
          type: "success",
        });
      } else {
        setNotification({
          message: `${n} Verstoß${n !== 1 ? "e" : ""} gefunden in ${checkResult.shapeCount} Shapes (${checkResult.durationMs} ms).`,
          type: "warning",
        });
        // Alle Slides mit Verstößen aufklappen
        const slides = new Set(checkResult.violations.map((v) => v.slideIndex));
        setExpandedSlides(slides);
      }
    } catch (err) {
      setNotification({
        message: `Fehler: ${err instanceof Error ? err.message : String(err)}`,
        type: "error",
      });
    } finally {
      setIsRunning(false);
      setProgress(null);
    }
  };

  // ── Auto-Fix ausführen ──────────────────────────────────────────────────────
  const handleFixAll = async () => {
    if (!result) return;
    const fixable = result.violations.filter((v) => v.fixable);
    if (fixable.length === 0) {
      setNotification({ message: "Keine behebbare Verstöße.", type: "info" });
      return;
    }

    setIsFixing(true);
    setNotification(null);
    try {
      const fixResult = await fixViolations(fixable);
      setNotification({
        message: `${fixResult.fixed} Verstoß${fixResult.fixed !== 1 ? "e" : ""} behoben.${
          fixResult.errors.length > 0 ? ` Fehler bei: ${fixResult.errors.join(", ")}` : ""
        }`,
        type: fixResult.errors.length > 0 ? "warning" : "success",
      });
      // Erneut prüfen um aktualisierten Status zu zeigen
      await handleRun();
    } catch (err) {
      setNotification({
        message: `Fix-Fehler: ${err instanceof Error ? err.message : String(err)}`,
        type: "error",
      });
    } finally {
      setIsFixing(false);
    }
  };

  // ── Zur Slide navigieren ────────────────────────────────────────────────────
  const handleNavigate = async (v: Violation) => {
    try {
      await goToSlide(v.slideId);
    } catch (err) {
      setNotification({
        message: `Navigation fehlgeschlagen: ${err instanceof Error ? err.message : String(err)}`,
        type: "error",
      });
    }
  };

  // ── Einzelnen Fix anwenden ──────────────────────────────────────────────────
  const handleFixOne = async (v: Violation) => {
    try {
      await fixViolations([v]);
      setNotification({ message: `"${v.shapeName}": ${TYPE_LABELS[v.type]} behoben.`, type: "success" });
      await handleRun();
    } catch (err) {
      setNotification({
        message: `Fehler: ${err instanceof Error ? err.message : String(err)}`,
        type: "error",
      });
    }
  };

  // ── CSV-Export ──────────────────────────────────────────────────────────────
  const handleExportCsv = () => {
    if (!result || result.violations.length === 0) return;
    exportViolationsAsCsv(result.violations, selectedProfile);
  };

  // ── Slide ein-/ausklappen ───────────────────────────────────────────────────
  const toggleSlide = (idx: number) => {
    setExpandedSlides((prev) => {
      const next = new Set(prev);
      next.has(idx) ? next.delete(idx) : next.add(idx);
      return next;
    });
  };

  // ─── Render ─────────────────────────────────────────────────────────────────
  const violations    = result?.violations ?? [];
  const bySlide       = groupBySlide(violations);
  const summary       = summarize(violations);
  const fixableCount  = violations.filter((v) => v.fixable).length;

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Brand Compliance Check</Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {/* ── Profil-Auswahl ── */}
      <Dropdown
        label="Konfigurationsprofil:"
        options={profileOptions}
        selectedKey={selectedProfile}
        onChange={(_e, opt) => opt && setSelectedProfile(String(opt.key))}
        styles={{ root: { maxWidth: 220 } }}
      />

      {/* ── Fortschritts-Indikator ── */}
      {isRunning && progress && (
        <ProgressIndicator
          label={`Prüfe Folie ${progress.current} von ${progress.total}…`}
          percentComplete={progress.total > 0 ? progress.current / progress.total : 0}
        />
      )}

      {/* ── Aktions-Buttons ── */}
      <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
        {isRunning ? (
          <Spinner size={SpinnerSize.small} label="Prüft…" />
        ) : (
          <PrimaryButton
            text="Jetzt prüfen"
            iconProps={{ iconName: "CheckList" }}
            onClick={handleRun}
            disabled={isFixing}
            styles={{ root: { flex: 1 } }}
          />
        )}
        {violations.length > 0 && (
          <>
            <DefaultButton
              text={isFixing ? "Behebt…" : `Auto-Fix (${fixableCount})`}
              iconProps={{ iconName: "Repair" }}
              onClick={handleFixAll}
              disabled={isRunning || isFixing || fixableCount === 0}
            />
            <DefaultButton
              text="CSV"
              iconProps={{ iconName: "Download" }}
              onClick={handleExportCsv}
              disabled={isRunning}
              title="Verstoßliste als CSV exportieren"
            />
          </>
        )}
      </Stack>

      {/* ── API-Hinweis Tabellen ── */}
      <MessageBar messageBarType={MessageBarType.info} isMultiline>
        Font-Details in Tabellenzellen (pro Text-Run) sind über Office.js nicht zugänglich.
        Nur Zelltext wird erkannt. Dokumentiert in TESTING.md.
      </MessageBar>

      {/* ── Zusammenfassung ── */}
      {violations.length > 0 && (
        <>
          <Separator>Zusammenfassung</Separator>
          <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
            {(Object.keys(summary) as ViolationType[])
              .filter((t) => summary[t] > 0)
              .map((t) => (
                <Stack
                  key={t}
                  horizontal
                  verticalAlign="center"
                  tokens={{ childrenGap: 4 }}
                  styles={{
                    root: {
                      background: "#FFF4CE",
                      border: "1px solid #F7C948",
                      borderRadius: 4,
                      padding: "2px 8px",
                    },
                  }}
                >
                  <Icon iconName={TYPE_ICONS[t]} style={{ fontSize: 12 }} />
                  <Text variant="small">
                    {TYPE_LABELS[t]}: <strong>{summary[t]}</strong>
                  </Text>
                </Stack>
              ))}
          </Stack>
        </>
      )}

      {/* ── Verstöße nach Slide ── */}
      {bySlide.size > 0 && (
        <>
          <Separator>Verstöße ({violations.length})</Separator>
          {Array.from(bySlide.entries())
            .sort(([a], [b]) => a - b)
            .map(([slideIdx, slViolations]) => (
              <SlideViolationGroup
                key={slideIdx}
                slideIndex={slideIdx}
                violations={slViolations}
                expanded={expandedSlides.has(slideIdx)}
                onToggle={() => toggleSlide(slideIdx)}
                onNavigate={handleNavigate}
                onFixOne={handleFixOne}
              />
            ))}
        </>
      )}

      {/* ── Kein Ergebnis ── */}
      {result && violations.length === 0 && (
        <Stack
          horizontalAlign="center"
          tokens={{ padding: 16 }}
          style={{ background: "#DFF6DD", borderRadius: 4 }}
        >
          <Icon iconName="CheckMark" style={{ fontSize: 24, color: "#107C10" }} />
          <Text variant="mediumPlus" style={{ color: "#107C10", marginTop: 8 }}>
            Alle Shapes sind brand-konform!
          </Text>
        </Stack>
      )}
    </Stack>
  );
};

// ─── Slide-Gruppe ─────────────────────────────────────────────────────────────

interface SlideGroupProps {
  slideIndex: number;
  violations: Violation[];
  expanded:   boolean;
  onToggle:   () => void;
  onNavigate: (v: Violation) => void;
  onFixOne:   (v: Violation) => void;
}

const SlideViolationGroup: React.FC<SlideGroupProps> = ({
  slideIndex, violations, expanded, onToggle, onNavigate, onFixOne,
}) => (
  <Stack
    styles={{
      root: {
        border: "1px solid #EDEBE9",
        borderRadius: 4,
        overflow: "hidden",
      },
    }}
  >
    {/* Header */}
    <Stack
      horizontal
      verticalAlign="center"
      horizontalAlign="space-between"
      styles={{
        root: {
          padding: "6px 10px",
          background: "#F3F2F1",
          cursor: "pointer",
        },
      }}
      onClick={onToggle}
    >
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
        <Icon
          iconName={expanded ? "ChevronDown" : "ChevronRight"}
          style={{ fontSize: 12, color: "#605E5C" }}
        />
        <Text variant="smallPlus" style={{ fontWeight: 600 }}>
          Folie {slideIndex + 1}
        </Text>
      </Stack>
      <Stack horizontal tokens={{ childrenGap: 6 }}>
        <Text variant="small" style={{ color: "#A19F9D" }}>
          {violations.length} Verstoß{violations.length !== 1 ? "e" : ""}
        </Text>
        <ActionButton
          text="Zur Folie"
          iconProps={{ iconName: "NavigateForward" }}
          styles={{ root: { height: 22, fontSize: 11, padding: "0 4px" } }}
          onClick={(e) => { e.stopPropagation(); onNavigate(violations[0]); }}
        />
      </Stack>
    </Stack>

    {/* Violations */}
    {expanded && (
      <Stack tokens={{ padding: "4px 0" }}>
        {violations.map((v, i) => (
          <ViolationRow
            key={`${v.shapeId}-${v.type}-${i}`}
            violation={v}
            onNavigate={onNavigate}
            onFix={onFixOne}
          />
        ))}
      </Stack>
    )}
  </Stack>
);

// ─── Einzelner Verstoß ────────────────────────────────────────────────────────

interface ViolationRowProps {
  violation:  Violation;
  onNavigate: (v: Violation) => void;
  onFix:      (v: Violation) => void;
}

const ViolationRow: React.FC<ViolationRowProps> = ({ violation: v, onNavigate, onFix }) => {
  const isColor = v.type === "fill-color" || v.type === "line-color" || v.type === "font-color";

  return (
    <Stack
      horizontal
      verticalAlign="center"
      horizontalAlign="space-between"
      styles={{
        root: {
          padding: "5px 10px",
          borderBottom: "1px solid #F3F2F1",
          ":hover": { background: "#FAFAFA" },
        },
      }}
    >
      {/* Icon + Details */}
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }} styles={{ root: { flex: 1, minWidth: 0 } }}>
        <Icon
          iconName={TYPE_ICONS[v.type]}
          style={{ fontSize: 14, color: "#A80000", flexShrink: 0 }}
        />
        <Stack styles={{ root: { minWidth: 0 } }}>
          <Text
            variant="small"
            style={{ fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}
            title={v.shapeName}
          >
            {v.shapeName || "(unbenannt)"}
          </Text>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
            <Text variant="xSmall" style={{ color: "#605E5C" }}>
              {TYPE_LABELS[v.type]}:
            </Text>
            {isColor ? (
              <ColorSwatch color={v.found} size={14} showLabel title={v.found} />
            ) : (
              <Text variant="xSmall" style={{ fontFamily: "monospace", color: "#A80000" }}>
                {v.found}
              </Text>
            )}
          </Stack>
        </Stack>
      </Stack>

      {/* Aktionen */}
      <Stack horizontal tokens={{ childrenGap: 4 }} styles={{ root: { flexShrink: 0 } }}>
        <ActionButton
          iconProps={{ iconName: "NavigateForward" }}
          styles={{ root: { minWidth: "auto", height: 24, padding: "0 4px" } }}
          onClick={() => onNavigate(v)}
          title="Zur Folie navigieren"
        />
        {v.fixable && (
          <ActionButton
            iconProps={{ iconName: "Repair" }}
            styles={{ root: { minWidth: "auto", height: 24, padding: "0 4px", color: "#107C10" } }}
            onClick={() => onFix(v)}
            title="Automatisch beheben"
          />
        )}
      </Stack>
    </Stack>
  );
};

export default BrandCheckPanel;
