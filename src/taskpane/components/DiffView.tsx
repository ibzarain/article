import * as React from "react";
import {
  Card,
  Button,
  tokens,
  makeStyles,
  Divider,
} from "@fluentui/react-components";
import {
  CheckmarkCircleFilled,
  DismissCircleFilled,
  CheckmarkCircleRegular,
  DismissCircleRegular,
} from "@fluentui/react-icons";
import { DocumentChange } from "../types/changes";

interface DiffViewProps {
  changes: DocumentChange[];
  onAccept: (id: string) => Promise<void>;
  onReject: (id: string) => Promise<void>;
  onAcceptAll: () => Promise<void>;
  onRejectAll: () => Promise<void>;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    maxHeight: "400px",
    overflowY: "auto",
    padding: "8px",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "8px 12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: "8px",
  },
  headerTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  headerActions: {
    display: "flex",
    gap: "8px",
  },
  changeCard: {
    padding: "12px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
  },
  changeHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    marginBottom: "8px",
  },
  changeType: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    textTransform: "uppercase",
    padding: "2px 8px",
    borderRadius: tokens.borderRadiusSmall,
  },
  editType: {
    backgroundColor: tokens.colorPaletteBlueBackground2,
    color: tokens.colorPaletteBlueForeground2,
  },
  insertType: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground2,
  },
  deleteType: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorPaletteRedForeground2,
  },
  formatType: {
    backgroundColor: tokens.colorPalettePurpleBackground2,
    color: tokens.colorPalettePurpleForeground2,
  },
  changeDescription: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground2,
    flex: 1,
    marginRight: "8px",
  },
  changeActions: {
    display: "flex",
    gap: "4px",
  },
  diffContent: {
    marginTop: "8px",
    padding: "8px",
    borderRadius: tokens.borderRadiusSmall,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase300,
  },
  oldText: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
    padding: "4px 8px",
    borderRadius: tokens.borderRadiusSmall,
    marginBottom: "4px",
    textDecoration: "line-through",
  },
  newText: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
    padding: "4px 8px",
    borderRadius: tokens.borderRadiusSmall,
  },
  insertedText: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
    padding: "4px 8px",
    borderRadius: tokens.borderRadiusSmall,
  },
  deletedText: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
    padding: "4px 8px",
    borderRadius: tokens.borderRadiusSmall,
    textDecoration: "line-through",
  },
  formatInfo: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
  },
  emptyState: {
    textAlign: "center",
    padding: "24px",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
  },
});

const DiffView: React.FC<DiffViewProps> = ({
  changes,
  onAccept,
  onReject,
  onAcceptAll,
  onRejectAll,
}) => {
  const styles = useStyles();

  const getTypeClass = (type: DocumentChange["type"]) => {
    switch (type) {
      case "edit":
        return styles.editType;
      case "insert":
        return styles.insertType;
      case "delete":
        return styles.deleteType;
      case "format":
        return styles.formatType;
      default:
        return "";
    }
  };

  const formatChangeDescription = (change: DocumentChange): string => {
    if (change.type === "format" && change.formatChanges) {
      const formatParts: string[] = [];
      if (change.formatChanges.bold !== undefined) {
        formatParts.push(change.formatChanges.bold ? "bold" : "not bold");
      }
      if (change.formatChanges.italic !== undefined) {
        formatParts.push(change.formatChanges.italic ? "italic" : "not italic");
      }
      if (change.formatChanges.underline !== undefined) {
        formatParts.push(change.formatChanges.underline ? "underlined" : "not underlined");
      }
      if (change.formatChanges.fontSize !== undefined) {
        formatParts.push(`font size ${change.formatChanges.fontSize}pt`);
      }
      if (change.formatChanges.fontColor !== undefined) {
        formatParts.push(`color ${change.formatChanges.fontColor}`);
      }
      return formatParts.join(", ");
    }
    return change.description;
  };

  if (changes.length === 0) {
    return (
      <div className={styles.emptyState}>
        No changes made yet. Start editing to see changes here!
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <div className={styles.headerTitle}>
          Changes ({changes.length})
        </div>
        <div className={styles.headerActions}>
          <Button
            size="small"
            appearance="subtle"
            icon={<CheckmarkCircleRegular />}
            onClick={onAcceptAll}
          >
            Accept All
          </Button>
          <Button
            size="small"
            appearance="subtle"
            icon={<DismissCircleRegular />}
            onClick={onRejectAll}
          >
            Reject All
          </Button>
        </div>
      </div>

      {changes.map((change) => (
        <Card key={change.id} className={styles.changeCard}>
          <div className={styles.changeHeader}>
            <div className={styles.changeDescription}>
              <span className={`${styles.changeType} ${getTypeClass(change.type)}`}>
                {change.type}
              </span>
              <div style={{ marginTop: "4px" }}>{formatChangeDescription(change)}</div>
            </div>
            <div className={styles.changeActions}>
              <Button
                size="small"
                appearance="subtle"
                icon={<CheckmarkCircleFilled />}
                onClick={() => onAccept(change.id)}
                title="Accept change"
              />
              <Button
                size="small"
                appearance="subtle"
                icon={<DismissCircleFilled />}
                onClick={() => onReject(change.id)}
                title="Reject change"
              />
            </div>
          </div>

          <div className={styles.diffContent}>
            {change.type === "edit" && (
              <>
                {change.oldText && (
                  <div className={styles.oldText}>
                    <strong>Old:</strong> {change.oldText}
                  </div>
                )}
                {change.newText && (
                  <div className={styles.newText}>
                    <strong>New:</strong> {change.newText}
                  </div>
                )}
              </>
            )}
            {change.type === "insert" && change.newText && (
              <div className={styles.insertedText}>
                <strong>Inserted:</strong> {change.newText}
              </div>
            )}
            {change.type === "delete" && change.oldText && (
              <div className={styles.deletedText}>
                <strong>Deleted:</strong> {change.oldText}
              </div>
            )}
            {change.type === "format" && (
              <div className={styles.formatInfo}>
                Applied formatting to: "{change.searchText}"
              </div>
            )}
          </div>
        </Card>
      ))}
    </div>
  );
};

export default DiffView;
