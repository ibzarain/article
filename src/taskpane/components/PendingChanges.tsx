import * as React from "react";
import { useState, useEffect } from "react";
import {
  makeStyles,
  Button,
  Spinner,
} from "@fluentui/react-components";
import {
  CheckmarkCircleFilled,
  DismissCircleFilled,
} from "@fluentui/react-icons";
import { DocumentChange, ChangeTracking } from "../types/changes";

interface PendingChangesProps {
  changeTracker: ChangeTracking;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    backgroundColor: "#252526",
    borderTop: "1px solid #3e3e42",
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
    gap: "8px",
  },
  headerButton: {
    padding: "4px 12px",
    fontSize: "12px",
    backgroundColor: "transparent",
    color: "#cccccc",
    border: "1px solid #3e3e42",
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
    padding: "10px",
    marginBottom: "8px",
    backgroundColor: "#1e1e1e",
    border: "1px solid #3e3e42",
    borderRadius: "6px",
    fontSize: "12px",
  },
  changeDescription: {
    fontSize: "12px",
    color: "#cccccc",
    marginBottom: "8px",
    lineHeight: "1.5",
  },
  changeActions: {
    display: "flex",
    gap: "8px",
    justifyContent: "flex-end",
  },
  actionButton: {
    padding: "6px 12px",
    fontSize: "12px",
    borderRadius: "4px",
    border: "none",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    gap: "4px",
    transition: "all 0.2s ease",
  },
  acceptButton: {
    backgroundColor: "#1e4620",
    color: "#89d185",
    "&:hover": {
      backgroundColor: "#2d5a2f",
    },
  },
  rejectButton: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
    "&:hover": {
      backgroundColor: "#6a2d2d",
    },
  },
  emptyState: {
    textAlign: "center",
    padding: "40px 24px",
    color: "#858585",
    fontSize: "13px",
  },
  loading: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "20px",
    gap: "8px",
    color: "#858585",
  },
});

const PendingChanges: React.FC<PendingChangesProps> = ({ changeTracker }) => {
  const [changes, setChanges] = React.useState<DocumentChange[]>([]);
  const [isAccepting, setIsAccepting] = React.useState<string | null>(null);
  const [isRejecting, setIsRejecting] = React.useState<string | null>(null);
  const styles = useStyles();

  // Subscribe to changes
  useEffect(() => {
    const updateChanges = () => {
      setChanges([...changeTracker.changes.filter(c => !c.applied)]);
    };

    // Initial load
    updateChanges();

    // Poll for changes (since we can't easily subscribe to array changes)
    const interval = setInterval(updateChanges, 500);

    return () => clearInterval(interval);
  }, [changeTracker]);

  const handleAccept = async (changeId: string) => {
    setIsAccepting(changeId);
    try {
      await changeTracker.acceptChange(changeId);
      setChanges(prev => prev.filter(c => c.id !== changeId));
    } catch (error) {
      console.error("Error accepting change:", error);
    } finally {
      setIsAccepting(null);
    }
  };

  const handleReject = async (changeId: string) => {
    setIsRejecting(changeId);
    try {
      await changeTracker.rejectChange(changeId);
      setChanges(prev => prev.filter(c => c.id !== changeId));
    } catch (error) {
      console.error("Error rejecting change:", error);
    } finally {
      setIsRejecting(null);
    }
  };

  const handleAcceptAll = async () => {
    const pendingChanges = changes.filter(c => !c.applied);
    for (const change of pendingChanges) {
      await handleAccept(change.id);
    }
  };

  const handleRejectAll = async () => {
    const pendingChanges = changes.filter(c => !c.applied);
    for (const change of pendingChanges) {
      await handleReject(change.id);
    }
  };

  if (changes.length === 0) {
    return null; // Don't show if no pending changes
  }

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <div className={styles.headerTitle}>
          Pending Changes ({changes.length})
        </div>
        {changes.length > 0 && (
          <div className={styles.headerActions}>
            <button
              className={styles.headerButton}
              onClick={handleAcceptAll}
              title="Accept all changes"
            >
              Accept All
            </button>
            <button
              className={styles.headerButton}
              onClick={handleRejectAll}
              title="Reject all changes"
            >
              Reject All
            </button>
          </div>
        )}
      </div>

      <div className={styles.changesList}>
        {changes.map((change) => (
          <div key={change.id} className={styles.changeCard}>
            <div className={styles.changeDescription}>
              {change.description}
            </div>
            <div className={styles.changeActions}>
              <button
                className={`${styles.actionButton} ${styles.acceptButton}`}
                onClick={() => handleAccept(change.id)}
                disabled={isAccepting === change.id || isRejecting === change.id}
              >
                {isAccepting === change.id ? (
                  <>
                    <Spinner size="tiny" />
                    Accepting...
                  </>
                ) : (
                  <>
                    <CheckmarkCircleFilled style={{ fontSize: "14px" }} />
                    Accept
                  </>
                )}
              </button>
              <button
                className={`${styles.actionButton} ${styles.rejectButton}`}
                onClick={() => handleReject(change.id)}
                disabled={isAccepting === change.id || isRejecting === change.id}
              >
                {isRejecting === change.id ? (
                  <>
                    <Spinner size="tiny" />
                    Rejecting...
                  </>
                ) : (
                  <>
                    <DismissCircleFilled style={{ fontSize: "14px" }} />
                    Reject
                  </>
                )}
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

export default PendingChanges;
