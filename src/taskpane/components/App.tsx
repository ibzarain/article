import * as React from "react";
import { useState, useMemo } from "react";
import TextInsertion from "./TextInsertion";
import AgentChat from "./AgentChat";
import ApiKeyConfig from "./ApiKeyConfig";
import { makeStyles } from "@fluentui/react-components";
import { insertText } from "../taskpane";
import { createWordAgent } from "../agent/wordAgent";
// Agent type will be inferred from createWordAgent return type

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    backgroundColor: "#1e1e1e",
    display: "flex",
    flexDirection: "column",
    color: "#cccccc",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
  },
  container: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    padding: "0",
    maxWidth: "100%",
    height: "100vh",
    overflow: "hidden",
  },
  tabContainer: {
    display: "flex",
    gap: "0",
    backgroundColor: "#252526",
    borderBottom: "1px solid #3e3e42",
    padding: "0 8px",
  },
  tabButton: {
    minWidth: "120px",
    backgroundColor: "transparent",
    color: "#cccccc",
    border: "none",
    borderRadius: "0",
    padding: "8px 16px",
    fontSize: "13px",
    fontWeight: "400",
    cursor: "pointer",
    transition: "all 0.2s ease",
    ":hover": {
      backgroundColor: "#2a2d2e",
    },
  },
  activeTab: {
    backgroundColor: "#1e1e1e",
    color: "#ffffff",
    borderBottom: "2px solid #007acc",
  },
  tabPanel: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
  },
});

const App: React.FC<AppProps> = () => {
  // Check for API key from environment variable first, then localStorage
  const getInitialApiKey = (): string => {
    // Check environment variable (set by webpack)
    const envKey = (process.env as any).OPENAI_API_KEY;
    if (envKey) {
      return envKey;
    }
    // Fall back to localStorage
    const storedKey = localStorage.getItem("openai_api_key");
    return storedKey || "";
  };

  const [apiKey, setApiKey] = useState<string>(getInitialApiKey);
  const [activeTab, setActiveTab] = useState<"agent" | "manual">("agent");
  const styles = useStyles();

  const agent = useMemo(() => {
    if (!apiKey) {
      return null;
    }
    try {
      // Create agent with change tracking callback
      // The callback will be set up in AgentChat component
      return createWordAgent(apiKey);
    } catch (error) {
      console.error("Failed to create agent:", error);
      return null;
    }
  }, [apiKey]);

  const handleApiKeySet = (key: string) => {
    setApiKey(key);
  };

  return (
    <div className={styles.root}>
      <div className={styles.container}>
        <div className={styles.tabContainer}>
          <button
            className={`${styles.tabButton} ${activeTab === "agent" ? styles.activeTab : ""}`}
            onClick={() => setActiveTab("agent")}
          >
            AI Agent
          </button>
          <button
            className={`${styles.tabButton} ${activeTab === "manual" ? styles.activeTab : ""}`}
            onClick={() => setActiveTab("manual")}
          >
            Manual Edit
          </button>
        </div>

        <div className={styles.tabPanel}>
          {activeTab === "agent" ? (
            <>
              {!apiKey ? (
                <ApiKeyConfig onApiKeySet={handleApiKeySet} />
              ) : agent ? (
                <AgentChat agent={agent} />
              ) : (
                <div>Error creating agent. Please check your API key.</div>
              )}
            </>
          ) : (
            <TextInsertion insertText={insertText} />
          )}
        </div>
      </div>
    </div>
  );
};

export default App;
