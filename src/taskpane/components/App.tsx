import * as React from "react";
import { useState, useMemo } from "react";
import TextInsertion from "./TextInsertion";
import AgentChat from "./AgentChat";
import ApiKeyConfig from "./ApiKeyConfig";
import { makeStyles, tokens, Button } from "@fluentui/react-components";
import { insertText } from "../taskpane";
import { createWordAgent } from "../agent/wordAgent";
// Agent type will be inferred from createWordAgent return type

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    flexDirection: "column",
  },
  container: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    padding: "24px",
    maxWidth: "100%",
  },
  tabContainer: {
    display: "flex",
    gap: "8px",
    marginBottom: "16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    paddingBottom: "8px",
  },
  tabButton: {
    minWidth: "120px",
  },
  activeTab: {
    fontWeight: tokens.fontWeightSemibold,
  },
  tabPanel: {
    paddingTop: "16px",
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
          <Button
            appearance={activeTab === "agent" ? "primary" : "subtle"}
            className={styles.tabButton}
            onClick={() => setActiveTab("agent")}
          >
            AI Agent
          </Button>
          <Button
            appearance={activeTab === "manual" ? "primary" : "subtle"}
            className={styles.tabButton}
            onClick={() => setActiveTab("manual")}
          >
            Manual Edit
          </Button>
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
