/**
 * ReviewPanel.tsx
 * Feature-Panel: Review / Kommentare / Annotationen (vollständige Implementierung).
 *
 * Zwei Ansichten (über URL-Parameter ?view=):
 * - "review"      → Kommentar/Highlight einfügen + Statistik
 * - "my-comments" → Übersicht aller Kommentare im Deck mit Navigation + Löschen
 *
 * Shape-Namenkonvention:
 * INFRONT_COMMENT_{timestamp}   – Kommentare (Post-It)
 * INFRONT_HIGHLIGHT_{timestamp} – Markierungen (gefülltes Rechteck)
 *
 * Annotation-Entfernung auch über Ribbon-Button "Kommentare entfernen"
 * (commands.ts → removeComments).
 */

import * as React from "react";
import { Stack }              from "@fluentui/react/lib/Stack";
import { Text }               from "@fluentui/react/lib/Text";
import { PrimaryButton, DefaultButton, ActionButton } from "@fluentui/react/lib/Button";
import { TextField }          from "@fluentui/react/lib/TextField";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Separator }          from "@fluentui/react/lib/Separator";
import { Pivot, PivotItem }   from "@fluentui/react/lib/Pivot";
import { Checkbox }           from "@fluentui/react/lib/Checkbox";
import { Icon }               from "@fluentui/react/lib/Icon";
import { TooltipHost }        from "@fluentui/react/lib/Tooltip";

import NotificationBar, { NotificationType } from "../shared/NotificationBar";
import ColorSwatch            from "../shared/ColorSwatch";
import {
  addCommentToCurrentSlide,
  addHighlightToCurrentSlide,
  findAllComments,
  removeCommentByName,
  removeAllAnnotations,
  goToSlideById,
  getReviewStats,
  CommentInfo,
  ReviewStats,
} from "../../../features/review/ReviewService";

interface Notification {
  message: string;
  type:    NotificationType;
}

// ─── Farb-Optionen für Kommentare ─────────────────────────────────────────────

const COMMENT_COLORS = [
  { hex: "#FFFACD", label: "Gelb (Standard)" },
  { hex: "#FFE4E1", label: "Rosa" },
  { hex: "#E1F5FE", label: "Hellblau" },
  { hex: "#E8F5E9", label: "Hellgrün" },
  { hex: "#F3E5F5", label: "Lila" },
];

const HIGHLIGHT_COLORS = [
  { hex: "#FFFF00", label: "Gelb (Standard)" },
  { hex: "#FF9800", label: "Orange" },
  { hex: "#4FC3F7", label: "Blau" },
  { hex: "#81C784", label: "Grün" },
  { hex: "#FF8A80", label: "Rot" },
];

// ─── Hilfsfunktionen ──────────────────────────────────────────────────────────

function getUserName(): string {
  try {
    return Office.context.userProfile?.displayName ?? "Unbekannt";
  } catch {
    return "Unbekannt";
  }
}

function formatTimestampPreview(): string {
  return new Date().toLocaleString("de-DE", {
    day: "2-digit", month: "2-digit", year: "numeric",
    hour: "2-digit", minute: "2-digit",
  });
}

// ─── Hauptkomponente ──────────────────────────────────────────────────────────

const ReviewPanel: React.FC = () => {
  const view = new URLSearchParams(window.location.search).get("view");
  const isMyComments = view === "my-comments";

  if (isMyComments) {
    return <MyCommentsView />;
  }
  return <AddAnnotationView />;
};

// ─── Ansicht: Kommentar / Highlight hinzufügen ────────────────────────────────

const AddAnnotationView: React.FC = () => {
  const [notification, setNotification] = React.useState<Notification | null>(null);
  const [isRunning, setIsRunning]       = React.useState(false);
  const [stats, setStats]               = React.useState<ReviewStats | null>(null);

  // Kommentar-State
  const [commentText, setCommentText]     = React.useState("");
  const [commentColor, setCommentColor]   = React.useState(COMMENT_COLORS[0].hex);

  // Highlight-State
  const [highlightColor, setHighlightColor] = React.useState(HIGHLIGHT_COLORS[0].hex);
  const [highlightWidth,  setHighlightWidth]  = React.useState(200);
  const [highlightHeight, setHighlightHeight] = React.useState(80);

  const author    = getUserName();
  const stampPreview = `[${author} – ${formatTimestampPreview()}]`;

  // Stats beim Mounten laden
  React.useEffect(() => {
    getReviewStats().then(setStats).catch(() => {});
  }, []);

  const refreshStats = () => getReviewStats().then(setStats).catch(() => {});

  // ── Kommentar einfügen ────────────────────────────────────────────────────
  const handleAddComment = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      const name = await addCommentToCurrentSlide({
        text:   commentText,
        author,
        color:  commentColor,
      });
      await refreshStats();
      setCommentText("");
      setNotification({ message: `Kommentar eingefügt: „${name}".`, type: "success" });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ── Highlight einfügen ────────────────────────────────────────────────────
  const handleAddHighlight = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      const name = await addHighlightToCurrentSlide({
        color:  highlightColor,
        width:  highlightWidth,
        height: highlightHeight,
      });
      await refreshStats();
      setNotification({ message: `Markierung eingefügt: „${name}".`, type: "success" });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  // ── Alle entfernen ────────────────────────────────────────────────────────
  const handleRemoveAll = async () => {
    setIsRunning(true);
    setNotification(null);
    try {
      const result = await removeAllAnnotations();
      setStats({ comments: 0, highlights: 0 });
      const msg = result.comments + result.highlights > 0
        ? `${result.comments} Kommentar(e) und ${result.highlights} Markierung(en) entfernt.`
        : "Keine Annotationen gefunden.";
      setNotification({ message: msg, type: result.comments + result.highlights > 0 ? "success" : "info" });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsRunning(false); }
  };

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Review &amp; Annotationen</Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {/* Statistik-Banner */}
      {stats && (stats.comments > 0 || stats.highlights > 0) && (
        <Stack
          horizontal
          verticalAlign="center"
          tokens={{ childrenGap: 8 }}
          styles={{ root: { background: "#EFF6FC", borderRadius: 4, padding: "6px 10px" } }}
        >
          <Icon iconName="Comment" style={{ fontSize: 14, color: "#0078D4" }} />
          <Text variant="small" style={{ color: "#0078D4" }}>
            {stats.comments} Kommentar(e) · {stats.highlights} Markierung(en) im Deck
          </Text>
        </Stack>
      )}

      <Pivot>
        {/* ── Tab 1: Kommentar ── */}
        <PivotItem headerText="Kommentar" itemIcon="Comment">
          <Stack tokens={{ childrenGap: 10, padding: "12px 0 0 0" }}>
            <Text variant="xSmall" style={{ color: "#555", fontStyle: "italic" }}>
              {stampPreview}
            </Text>

            <TextField
              label="Kommentartext:"
              value={commentText}
              onChange={(_e, v) => setCommentText(v ?? "")}
              multiline
              rows={4}
              placeholder="Kommentar eingeben… (optional – Zeitstempel wird immer eingefügt)"
            />

            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="xSmall" style={{ color: "#666" }}>Hintergrundfarbe:</Text>
              <Stack horizontal tokens={{ childrenGap: 6 }}>
                {COMMENT_COLORS.map((c) => (
                  <TooltipHost key={c.hex} content={c.label}>
                    <div
                      onClick={() => setCommentColor(c.hex)}
                      style={{
                        cursor: "pointer",
                        border: commentColor === c.hex ? "2px solid #003366" : "2px solid transparent",
                        borderRadius: 3,
                        padding: 1,
                      }}
                    >
                      <ColorSwatch color={c.hex} size={20} />
                    </div>
                  </TooltipHost>
                ))}
              </Stack>
            </Stack>

            {isRunning ? (
              <Spinner size={SpinnerSize.small} label="Wird eingefügt…" />
            ) : (
              <PrimaryButton
                text="Kommentar auf aktiver Folie einfügen"
                iconProps={{ iconName: "CommentAdd" }}
                onClick={handleAddComment}
                styles={{ root: { width: "100%" } }}
              />
            )}
          </Stack>
        </PivotItem>

        {/* ── Tab 2: Markierung ── */}
        <PivotItem headerText="Markierung" itemIcon="RectangleShape">
          <Stack tokens={{ childrenGap: 10, padding: "12px 0 0 0" }}>
            <Text variant="small" style={{ color: "#555" }}>
              Fügt ein farbiges Rechteck als Markierung/Callout-Fläche ein.
              Position und Größe können danach in PowerPoint angepasst werden.
            </Text>

            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="xSmall" style={{ color: "#666" }}>Farbe:</Text>
              <Stack horizontal tokens={{ childrenGap: 6 }}>
                {HIGHLIGHT_COLORS.map((c) => (
                  <TooltipHost key={c.hex} content={c.label}>
                    <div
                      onClick={() => setHighlightColor(c.hex)}
                      style={{
                        cursor: "pointer",
                        border: highlightColor === c.hex ? "2px solid #003366" : "2px solid transparent",
                        borderRadius: 3,
                        padding: 1,
                      }}
                    >
                      <ColorSwatch color={c.hex} size={20} />
                    </div>
                  </TooltipHost>
                ))}
              </Stack>
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <TextField
                label="Breite (pt):"
                value={String(highlightWidth)}
                onChange={(_e, v) => {
                  const n = parseInt(v ?? "0", 10);
                  if (!isNaN(n) && n > 0) setHighlightWidth(n);
                }}
                type="number"
                min={10}
                suffix="pt"
                styles={{ root: { width: 80 } }}
              />
              <TextField
                label="Höhe (pt):"
                value={String(highlightHeight)}
                onChange={(_e, v) => {
                  const n = parseInt(v ?? "0", 10);
                  if (!isNaN(n) && n > 0) setHighlightHeight(n);
                }}
                type="number"
                min={10}
                suffix="pt"
                styles={{ root: { width: 80 } }}
              />
            </Stack>

            {isRunning ? (
              <Spinner size={SpinnerSize.small} label="Wird eingefügt…" />
            ) : (
              <PrimaryButton
                text="Markierung auf aktiver Folie einfügen"
                iconProps={{ iconName: "RectangleShape" }}
                onClick={handleAddHighlight}
                styles={{ root: { width: "100%" } }}
              />
            )}
          </Stack>
        </PivotItem>

        {/* ── Tab 3: Verwalten ── */}
        <PivotItem headerText="Verwalten" itemIcon="Settings">
          <Stack tokens={{ childrenGap: 12, padding: "12px 0 0 0" }}>
            <Text variant="small" style={{ color: "#555" }}>
              Entfernt alle Infront-Annotationen (Kommentare + Markierungen) aus dem gesamten Deck.
              Für Einzellöschen: Tab „Meine Kommentare" (Ribbon-Button).
            </Text>

            <Stack
              tokens={{ childrenGap: 6, padding: "8px 10px" }}
              styles={{ root: { background: "#F3F2F1", borderRadius: 4 } }}
            >
              <Text variant="small">
                <strong>Kommentare im Deck:</strong> {stats?.comments ?? "—"}
              </Text>
              <Text variant="small">
                <strong>Markierungen im Deck:</strong> {stats?.highlights ?? "—"}
              </Text>
              <ActionButton
                text="Statistik aktualisieren"
                iconProps={{ iconName: "Refresh" }}
                onClick={refreshStats}
                styles={{ root: { height: 24, padding: 0, fontSize: 12 } }}
              />
            </Stack>

            <Separator />

            {isRunning ? (
              <Spinner size={SpinnerSize.small} label="Wird entfernt…" />
            ) : (
              <DefaultButton
                text={`Alle Annotationen entfernen${stats && stats.comments + stats.highlights > 0 ? ` (${stats.comments + stats.highlights})` : ""}`}
                iconProps={{ iconName: "Delete" }}
                onClick={handleRemoveAll}
                disabled={stats ? stats.comments + stats.highlights === 0 : false}
                styles={{ root: { width: "100%" } }}
              />
            )}

            <Text variant="xSmall" style={{ color: "#888" }}>
              Alternativ: Ribbon-Button „Kommentare entfernen" in der Review-Gruppe.
            </Text>
          </Stack>
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

// ─── Ansicht: Meine Kommentare (Übersicht) ────────────────────────────────────

const MyCommentsView: React.FC = () => {
  const [notification, setNotification]   = React.useState<Notification | null>(null);
  const [comments, setComments]           = React.useState<CommentInfo[]>([]);
  const [isLoading, setIsLoading]         = React.useState(false);
  const [filterAuthor, setFilterAuthor]   = React.useState("");
  const [onlyMine, setOnlyMine]           = React.useState(false);
  const currentUser = getUserName();

  const loadComments = async () => {
    setIsLoading(true);
    setNotification(null);
    try {
      const result = await findAllComments();
      setComments(result);
      if (result.length === 0) {
        setNotification({ message: "Keine Kommentare im Deck gefunden.", type: "info" });
      }
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    } finally { setIsLoading(false); }
  };

  React.useEffect(() => { loadComments(); }, []);

  const handleNavigate = async (comment: CommentInfo) => {
    try {
      await goToSlideById(comment.slideId);
    } catch (err) {
      setNotification({ message: `Navigation fehlgeschlagen: ${err instanceof Error ? err.message : err}`, type: "error" });
    }
  };

  const handleDelete = async (comment: CommentInfo) => {
    try {
      await removeCommentByName(comment.shapeName);
      setComments((prev) => prev.filter((c) => c.shapeName !== comment.shapeName));
      setNotification({ message: `Kommentar auf Folie ${comment.slideIndex + 1} entfernt.`, type: "success" });
    } catch (err) {
      setNotification({ message: `Fehler: ${err instanceof Error ? err.message : err}`, type: "error" });
    }
  };

  // Filter anwenden
  const filtered = comments.filter((c) => {
    if (onlyMine && c.author !== currentUser) return false;
    if (filterAuthor.trim() && !c.author.toLowerCase().includes(filterAuthor.toLowerCase())) return false;
    return true;
  });

  return (
    <Stack className="panel-container" tokens={{ childrenGap: 12 }}>
      <Text className="panel-title">Meine Kommentare</Text>

      <NotificationBar
        message={notification?.message ?? null}
        type={notification?.type}
        onDismiss={() => setNotification(null)}
      />

      {/* Filter */}
      <Stack tokens={{ childrenGap: 6 }}>
        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
          <TextField
            label="Filter nach Autor:"
            value={filterAuthor}
            onChange={(_e, v) => setFilterAuthor(v ?? "")}
            placeholder="Name…"
            styles={{ root: { flex: 1 } }}
          />
          <DefaultButton
            text="Laden"
            iconProps={{ iconName: "Refresh" }}
            onClick={loadComments}
            styles={{ root: { marginBottom: 1 } }}
          />
        </Stack>
        <Checkbox
          label={`Nur meine Kommentare (${currentUser})`}
          checked={onlyMine}
          onChange={(_e, c) => setOnlyMine(!!c)}
        />
      </Stack>

      <Text variant="xSmall" style={{ color: "#666" }}>
        {filtered.length} von {comments.length} Kommentar(en)
      </Text>

      <Separator />

      {isLoading ? (
        <Spinner size={SpinnerSize.small} label="Lade Kommentare…" />
      ) : filtered.length === 0 ? (
        <Text variant="small" style={{ color: "#888" }}>
          Keine Kommentare gefunden.
        </Text>
      ) : (
        <Stack tokens={{ childrenGap: 8 }}>
          {filtered.map((c) => (
            <CommentCard
              key={c.shapeName}
              comment={c}
              onNavigate={handleNavigate}
              onDelete={handleDelete}
            />
          ))}
        </Stack>
      )}
    </Stack>
  );
};

// ─── Kommentar-Karte ──────────────────────────────────────────────────────────

interface CommentCardProps {
  comment:    CommentInfo;
  onNavigate: (c: CommentInfo) => void;
  onDelete:   (c: CommentInfo) => void;
}

const CommentCard: React.FC<CommentCardProps> = ({ comment, onNavigate, onDelete }) => (
  <Stack
    tokens={{ childrenGap: 4, padding: "8px 10px" }}
    styles={{ root: { background: "#FFFACD", borderRadius: 4, border: "1px solid #FFA500" } }}
  >
    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
        <Icon iconName="Comment" style={{ fontSize: 13, color: "#FFA500" }} />
        <Text variant="smallPlus" style={{ fontWeight: 600 }}>
          Folie {comment.slideIndex + 1}
        </Text>
        <Text variant="xSmall" style={{ color: "#666" }}>
          · {comment.author}
        </Text>
      </Stack>
      <Stack horizontal tokens={{ childrenGap: 2 }}>
        <TooltipHost content="Zur Folie navigieren">
          <ActionButton
            iconProps={{ iconName: "NavigateExternalInline" }}
            onClick={() => onNavigate(comment)}
            styles={{ root: { height: 22, minWidth: "auto", padding: "0 4px" } }}
          />
        </TooltipHost>
        <TooltipHost content="Kommentar löschen">
          <ActionButton
            iconProps={{ iconName: "Delete" }}
            onClick={() => onDelete(comment)}
            styles={{ root: { height: 22, minWidth: "auto", padding: "0 4px", color: "#A19F9D" } }}
          />
        </TooltipHost>
      </Stack>
    </Stack>
    {comment.timestamp && (
      <Text variant="xSmall" style={{ color: "#888" }}>{comment.timestamp}</Text>
    )}
    {comment.commentText && (
      <Text variant="small" style={{ color: "#333" }}>
        {comment.commentText.slice(0, 120)}{comment.commentText.length > 120 ? "…" : ""}
      </Text>
    )}
  </Stack>
);

export default ReviewPanel;
