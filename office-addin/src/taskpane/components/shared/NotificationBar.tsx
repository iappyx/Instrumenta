/**
 * NotificationBar.tsx
 * Zeigt Erfolgs-, Warn- und Fehlermeldungen inline in der Task Pane an.
 */

import * as React from "react";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";

export type NotificationType = "success" | "warning" | "error" | "info";

interface NotificationBarProps {
  message: string | null;
  type?: NotificationType;
  onDismiss?: () => void;
}

const TYPE_MAP: Record<NotificationType, MessageBarType> = {
  success: MessageBarType.success,
  warning: MessageBarType.warning,
  error:   MessageBarType.error,
  info:    MessageBarType.info,
};

const NotificationBar: React.FC<NotificationBarProps> = ({
  message,
  type = "info",
  onDismiss,
}) => {
  if (!message) return null;

  return (
    <MessageBar
      messageBarType={TYPE_MAP[type]}
      isMultiline={false}
      onDismiss={onDismiss}
      dismissButtonAriaLabel="Schließen"
      styles={{ root: { marginBottom: 8 } }}
    >
      {message}
    </MessageBar>
  );
};

export default NotificationBar;
