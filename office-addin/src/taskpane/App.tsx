/**
 * App.tsx – Root-Komponente der Infront Toolkit Task Pane.
 *
 * Routing: liest den `?view=` URL-Parameter und rendert die passende
 * Feature-Komponente. Jeder Ribbon-Button öffnet die Task Pane mit
 * einer unterschiedlichen URL (z.B. taskpane.html?view=corner-radius).
 */

import * as React from "react";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import { ThemeProvider, createTheme } from "@fluentui/react/lib/Theme";
import { Text } from "@fluentui/react/lib/Text";
import { Stack } from "@fluentui/react/lib/Stack";

import CornerRadiusPanel   from "./components/CornerRadius/CornerRadiusPanel";
import ColorPickerPanel    from "./components/ColorPicker/ColorPickerPanel";
import BrandCheckPanel     from "./components/BrandCheck/BrandCheckPanel";
import FormatPainterPanel  from "./components/FormatPainter/FormatPainterPanel";
import FindReplacePanel    from "./components/FindReplace/FindReplacePanel";
import GapEqualizerPanel   from "./components/GapEqualizer/GapEqualizerPanel";
import AgendaPanel         from "./components/Agenda/AgendaPanel";
import MasterImportPanel   from "./components/MasterImport/MasterImportPanel";
import ReviewPanel         from "./components/Review/ReviewPanel";
import RedBoxPanel         from "./components/RedBox/RedBoxPanel";

import "./App.css";

initializeIcons();

/** Infront-Markenfarbe als Fluent-UI-Theme */
const infrontTheme = createTheme({
  palette: {
    themePrimary:   "#003366",
    themeSecondary: "#1A5276",
    themeDark:      "#002244",
    white:          "#FFFFFF",
    neutralLight:   "#F3F4F6",
  },
});

/** Mapping: URL-View-Parameter → Komponente */
const VIEW_MAP: Record<string, React.ComponentType> = {
  "corner-radius":    CornerRadiusPanel,
  "color-picker":     ColorPickerPanel,
  "brand-check":      BrandCheckPanel,
  "format-painter":   FormatPainterPanel,
  "find-replace":     FindReplacePanel,
  "gap-equalizer":    GapEqualizerPanel,
  "agenda":           AgendaPanel,
  "master-import":    MasterImportPanel,
  "review":           ReviewPanel,
  "red-box-settings": RedBoxPanel,
  "my-comments":      ReviewPanel,
};

/** Liest den `view`-Parameter aus der aktuellen URL. */
function getCurrentView(): string {
  const params = new URLSearchParams(window.location.search);
  return params.get("view") ?? "welcome";
}

const App: React.FC = () => {
  const view = getCurrentView();
  const PanelComponent = VIEW_MAP[view];

  return (
    <ThemeProvider theme={infrontTheme}>
      <Stack className="app-container">
        <Stack className="app-header" horizontal verticalAlign="center">
          <Text variant="mediumPlus" className="app-title">
            Infront Toolkit
          </Text>
        </Stack>
        <Stack.Item grow className="app-content">
          {PanelComponent ? (
            <PanelComponent />
          ) : (
            <WelcomePanel />
          )}
        </Stack.Item>
      </Stack>
    </ThemeProvider>
  );
};

/** Startseite wenn kein view-Parameter gesetzt */
const WelcomePanel: React.FC = () => (
  <Stack tokens={{ padding: 16, childrenGap: 8 }}>
    <Text variant="large">Willkommen im Infront Toolkit</Text>
    <Text variant="small" style={{ color: "#666" }}>
      Bitte verwende einen der Ribbon-Buttons, um ein Feature zu öffnen.
    </Text>
  </Stack>
);

export default App;
