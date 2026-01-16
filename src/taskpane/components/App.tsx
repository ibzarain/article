import * as React from "react";
import { useState, useMemo } from "react";
import AgentChat from "./AgentChat";
import ApiKeyConfig from "./ApiKeyConfig";
import { makeStyles } from "@fluentui/react-components";
import { createWordAgent } from "../agent/wordAgent";
// Agent type will be inferred from createWordAgent return type

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    width: "100%",
    height: "100vh",
    backgroundColor: "#1e1e1e",
    display: "flex",
    flexDirection: "column",
    color: "#cccccc",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
    overflow: "hidden",
  },
  container: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    width: "100%",
    height: "100%",
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
        {!apiKey ? (
          <ApiKeyConfig onApiKeySet={handleApiKeySet} />
        ) : agent ? (
          <AgentChat agent={agent} />
        ) : (
          <div>Error creating agent. Please check your API key.</div>
        )}
      </div>
    </div>
  );
};

export default App;
