import * as React from "react";
import {
  makeStyles,
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
    height: "100%",
    backgroundColor: "#252526",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "12px 16px",
    backgroundColor: "#2d2d30",
    borderBottom: "1px solid #3e3e42",
  },
  headerTitle: {
    fontSize: "13px",
    fontWeight: "600",
    color: "#cccccc",
  },
  headerActions: {
    display: "flex",
    gap: "4px",
  },
  headerButton: {
    padding: "4px 8px",
    fontSize: "12px",
    backgroundColor: "transparent",
    color: "#cccccc",
    border: "none",
    borderRadius: "4px",
    cursor: "pointer",
    transition: "background 0.2s ease",
    "&:hover": {
      backgroundColor: "#3e3e42",
    },
  },
  changesList: {
    flex: 1,
    overflowY: "auto",
    padding: "12px",
    scrollbarWidth: "thin",
    scrollbarColor: "#424242 #252526",
    "&::-webkit-scrollbar": {
      width: "8px",
    },
    "&::-webkit-scrollbar-track": {
      background: "#252526",
    },
    "&::-webkit-scrollbar-thumb": {
      background: "#424242",
      borderRadius: "4px",
      "&:hover": {
        background: "#4e4e4e",
      },
    },
  },
  changeCard: {
    padding: "12px",
    marginBottom: "8px",
    backgroundColor: "#1e1e1e",
    border: "1px solid #3e3e42",
    borderRadius: "6px",
    transition: "border-color 0.2s ease",
    "&:hover": {
      borderColor: "#007acc",
    },
  },
  changeHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    marginBottom: "10px",
  },
  changeType: {
    fontSize: "11px",
    fontWeight: "600",
    textTransform: "uppercase",
    padding: "3px 8px",
    borderRadius: "4px",
    letterSpacing: "0.5px",
  },
  editType: {
    backgroundColor: "#264f78",
    color: "#75beff",
  },
  insertType: {
    backgroundColor: "#1e4620",
    color: "#89d185",
  },
  deleteType: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
  },
  formatType: {
    backgroundColor: "#4a148c",
    color: "#c586c0",
  },
  changeDescription: {
    fontSize: "13px",
    color: "#cccccc",
    flex: 1,
    marginRight: "8px",
    lineHeight: "1.5",
  },
  changeActions: {
    display: "flex",
    gap: "4px",
  },
  actionButton: {
    padding: "6px",
    backgroundColor: "transparent",
    border: "none",
    borderRadius: "4px",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    transition: "background 0.2s ease",
    "&:hover": {
      backgroundColor: "#3e3e42",
    },
  },
  diffContent: {
    marginTop: "10px",
    padding: "10px",
    borderRadius: "4px",
    fontFamily: "'SF Mono', Monaco, 'Cascadia Code', 'Roboto Mono', Consolas, 'Courier New', monospace",
    fontSize: "12px",
    backgroundColor: "#1a1a1a",
    border: "1px solid #2d2d30",
  },
  oldText: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
    padding: "6px 10px",
    borderRadius: "4px",
    marginBottom: "6px",
    textDecoration: "line-through",
    display: "block",
  },
  newText: {
    backgroundColor: "#1e4620",
    color: "#89d185",
    padding: "6px 10px",
    borderRadius: "4px",
    display: "block",
  },
  insertedText: {
    backgroundColor: "#1e4620",
    color: "#89d185",
    padding: "6px 10px",
    borderRadius: "4px",
    display: "block",
  },
  deletedText: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
    padding: "6px 10px",
    borderRadius: "4px",
    textDecoration: "line-through",
    display: "block",
  },
  formatInfo: {
    fontSize: "12px",
    color: "#858585",
    fontStyle: "italic",
  },
  emptyState: {
    textAlign: "center",
    padding: "40px 24px",
    color: "#858585",
    fontSize: "13px",
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
        {changes.length > 0 && (
          <div className={styles.headerActions}>
            <button
              className={styles.headerButton}
              onClick={onAcceptAll}
              title="Accept all changes"
            >
              <CheckmarkCircleRegular style={{ fontSize: "14px", marginRight: "4px" }} />
              Accept All
            </button>
            <button
              className={styles.headerButton}
              onClick={onRejectAll}
              title="Reject all changes"
            >
              <DismissCircleRegular style={{ fontSize: "14px", marginRight: "4px" }} />
              Reject All
            </button>
          </div>
        )}
      </div>

      <div className={styles.changesList}>
        {changes.length === 0 ? (
          <div className={styles.emptyState}>
            No changes made yet. Start editing to see changes here!
          </div>
        ) : (
          changes.map((change) => (
            <div key={change.id} className={styles.changeCard}>
              <div className={styles.changeHeader}>
                <div className={styles.changeDescription}>
                  <span className={`${styles.changeType} ${getTypeClass(change.type)}`}>
                    {change.type}
                  </span>
                  <div style={{ marginTop: "6px", fontSize: "12px", color: "#858585" }}>
                    {formatChangeDescription(change)}
                  </div>
                </div>
                <div className={styles.changeActions}>
                  <button
                    className={styles.actionButton}
                    onClick={() => onAccept(change.id)}
                    title="Accept change"
                  >
                    <CheckmarkCircleFilled style={{ fontSize: "16px", color: "#89d185" }} />
                  </button>
                  <button
                    className={styles.actionButton}
                    onClick={() => onReject(change.id)}
                    title="Reject change"
                  >
                    <DismissCircleFilled style={{ fontSize: "16px", color: "#f48771" }} />
                  </button>
                </div>
              </div>

              <div className={styles.diffContent}>
                {change.type === "edit" && (
                  <>
                    {change.oldText && (
                      <div className={styles.oldText}>
                        <span style={{ opacity: 0.7 }}>−</span> {change.oldText}
                      </div>
                    )}
                    {change.newText && (
                      <div className={styles.newText}>
                        <span style={{ opacity: 0.7 }}>+</span> {change.newText}
                      </div>
                    )}
                  </>
                )}
                {change.type === "insert" && change.newText && (
                  <div className={styles.insertedText}>
                    <span style={{ opacity: 0.7 }}>+</span> {change.newText}
                  </div>
                )}
                {change.type === "delete" && change.oldText && (
                  <div className={styles.deletedText}>
                    <span style={{ opacity: 0.7 }}>−</span> {change.oldText}
                  </div>
                )}
                {change.type === "format" && (
                  <div className={styles.formatInfo}>
                    Applied formatting to: "{change.searchText}"
                  </div>
                )}
              </div>
            </div>
          ))
        )}
      </div>
    </div>
  );
};

export default DiffView;
