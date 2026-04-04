/**
 * commands.ts
 *
 * Ribbon-Funktionen für ExecuteFunction-Buttons im Infront Toolkit.
 * Diese Datei wird als unsichtbare commands.html geladen und registriert
 * alle Handler, die direkt aus dem Ribbon ausgeführt werden (ohne Task Pane).
 *
 * Jede Funktion MUSS event.completed() aufrufen, sonst hängt PowerPoint.
 */

/* global Office */

import {
  toggleRedBoxOnCurrentSlide,
  addRedBoxToAllSlides,
  removeRedBoxFromAllSlides,
} from "../features/redBox/RedBoxService";

Office.onReady(() => {
  // Alle ExecuteFunction-Handler werden als Properties auf `window` registriert,
  // damit das Office-Runtime sie per Name aufrufen kann.
  (window as Record<string, unknown>)["gapHorizontal"]  = gapHorizontal;
  (window as Record<string, unknown>)["gapVertical"]    = gapVertical;
  (window as Record<string, unknown>)["toggleRedBox"]   = toggleRedBox;
  (window as Record<string, unknown>)["redBoxAllSlides"] = redBoxAllSlides;
  (window as Record<string, unknown>)["removeRedBoxAll"] = removeRedBoxAll;
  (window as Record<string, unknown>)["removeComments"] = removeComments;
  (window as Record<string, unknown>)["addHighlight"]   = addHighlight;
});

/**
 * Gleicht horizontale Abstände der selektierten Shapes an.
 * Implementierung: Schritt 12 (Gap Equalizer).
 */
function gapHorizontal(event: Office.AddinCommands.Event): void {
  PowerPoint.run(async (context) => {
    const selection = context.presentation.getSelectedShapes();
    selection.load("items");
    await context.sync();

    if (selection.items.length < 3) {
      showNotification("Gap H", "Bitte mindestens 3 Shapes selektieren.");
      event.completed();
      return;
    }

    // Shapes nach left-Position sortieren
    const shapes = selection.items.map((s) => {
      s.load(["left", "width"]);
      return s;
    });
    await context.sync();

    shapes.sort((a, b) => a.left - b.left);

    const first = shapes[0];
    const last  = shapes[shapes.length - 1];
    const totalWidth = (last.left + last.width) - first.left;
    const totalShapeWidth = shapes.reduce((sum, s) => sum + s.width, 0);
    const gap = (totalWidth - totalShapeWidth) / (shapes.length - 1);

    let cursor = first.left + first.width;
    for (let i = 1; i < shapes.length - 1; i++) {
      shapes[i].left = cursor + gap;
      cursor = shapes[i].left + shapes[i].width;
    }

    await context.sync();
    event.completed();
  }).catch((err: Error) => {
    console.error("[InfrontToolkit] gapHorizontal:", err);
    event.completed();
  });
}

/**
 * Gleicht vertikale Abstände der selektierten Shapes an.
 * Implementierung: Schritt 12 (Gap Equalizer).
 */
function gapVertical(event: Office.AddinCommands.Event): void {
  PowerPoint.run(async (context) => {
    const selection = context.presentation.getSelectedShapes();
    selection.load("items");
    await context.sync();

    if (selection.items.length < 3) {
      showNotification("Gap V", "Bitte mindestens 3 Shapes selektieren.");
      event.completed();
      return;
    }

    const shapes = selection.items.map((s) => {
      s.load(["top", "height"]);
      return s;
    });
    await context.sync();

    shapes.sort((a, b) => a.top - b.top);

    const first = shapes[0];
    const last  = shapes[shapes.length - 1];
    const totalHeight = (last.top + last.height) - first.top;
    const totalShapeHeight = shapes.reduce((sum, s) => sum + s.height, 0);
    const gap = (totalHeight - totalShapeHeight) / (shapes.length - 1);

    let cursor = first.top + first.height;
    for (let i = 1; i < shapes.length - 1; i++) {
      shapes[i].top = cursor + gap;
      cursor = shapes[i].top + shapes[i].height;
    }

    await context.sync();
    event.completed();
  }).catch((err: Error) => {
    console.error("[InfrontToolkit] gapVertical:", err);
    event.completed();
  });
}

/**
 * Schaltet die Red Box (INFRONT_REDBOX) auf der aktiven Slide ein/aus.
 * Delegiert an RedBoxService (Schritt 13).
 */
function toggleRedBox(event: Office.AddinCommands.Event): void {
  toggleRedBoxOnCurrentSlide()
    .then(() => event.completed())
    .catch((err: Error) => {
      console.error("[InfrontToolkit] toggleRedBox:", err);
      event.completed();
    });
}

/**
 * Fügt INFRONT_REDBOX auf allen Slides ein.
 * Delegiert an RedBoxService (Schritt 13).
 */
function redBoxAllSlides(event: Office.AddinCommands.Event): void {
  addRedBoxToAllSlides()
    .then(() => event.completed())
    .catch((err: Error) => {
      console.error("[InfrontToolkit] redBoxAllSlides:", err);
      event.completed();
    });
}

/**
 * Entfernt alle INFRONT_REDBOX-Shapes aus dem gesamten Deck.
 * Delegiert an RedBoxService (Schritt 13).
 */
function removeRedBoxAll(event: Office.AddinCommands.Event): void {
  removeRedBoxFromAllSlides()
    .then((r) => {
      console.log(`[InfrontToolkit] removeRedBoxAll: ${r.removed} Box(en) entfernt.`);
      event.completed();
    })
    .catch((err: Error) => {
      console.error("[InfrontToolkit] removeRedBoxAll:", err);
      event.completed();
    });
}

/**
 * Entfernt alle Infront-Kommentar-Shapes (INFRONT_COMMENT_*) aus dem Deck.
 * Vollständige Implementierung: Schritt 11.
 */
function removeComments(event: Office.AddinCommands.Event): void {
  PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    let removed = 0;
    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items/name");
      await context.sync();

      for (const shape of shapes.items) {
        if (shape.name.startsWith("INFRONT_COMMENT_") ||
            shape.name.startsWith("INFRONT_HIGHLIGHT_")) {
          shape.delete();
          removed++;
        }
      }
    }

    await context.sync();
    console.log(`[InfrontToolkit] removeComments: ${removed} Kommentar(e) entfernt.`);
    event.completed();
  }).catch((err: Error) => {
    console.error("[InfrontToolkit] removeComments:", err);
    event.completed();
  });
}

/**
 * Fügt ein Highlight-Rechteck auf der aktiven Slide ein.
 * Vollständige Implementierung: Schritt 11.
 */
function addHighlight(event: Office.AddinCommands.Event): void {
  PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;

    const timestamp = new Date().toLocaleString("de-DE");
    const shapeName = `INFRONT_HIGHLIGHT_${Date.now()}`;

    const box = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    box.name = shapeName;
    box.load(["left", "top", "width", "height"]);
    await context.sync();

    // Standard-Position: mittig auf der Slide
    box.left   = 100;
    box.top    = 100;
    box.width  = 200;
    box.height = 80;
    box.fill.setSolidColor("#FFFF00");
    box.lineFormat.color = "#FFA500";
    box.lineFormat.weight = 1;

    const tf = box.textFrame;
    tf.textRange.text = `[Markierung – ${timestamp}]`;

    await context.sync();
    console.log(`[InfrontToolkit] addHighlight: "${shapeName}" eingefügt.`);
    event.completed();
  }).catch((err: Error) => {
    console.error("[InfrontToolkit] addHighlight:", err);
    event.completed();
  });
}

/**
 * Zeigt eine einfache Benachrichtigung über Office-Notification-API.
 * Fallback auf console.log wenn nicht verfügbar.
 */
function showNotification(title: string, message: string): void {
  console.log(`[InfrontToolkit] ${title}: ${message}`);
  // Notification-API ist im Commands-Kontext nicht verfügbar –
  // Rückmeldung erfolgt über console.log.
  // Für UI-Feedback bitte entsprechenden Task-Pane-Button nutzen.
}
