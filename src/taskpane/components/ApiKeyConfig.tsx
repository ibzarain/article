import * as React from "react";
import { useState, useEffect } from "react";
import {
  makeStyles,
} from "@fluentui/react-components";
import { KeyRegular, CheckmarkCircleFilled } from "@fluentui/react-icons";

interface ApiKeyConfigProps {
  onApiKeySet: (apiKey: string) => void;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "20px",
    width: "100%",
    maxWidth: "600px",
    margin: "0 auto",
    padding: "40px 24px",
    backgroundColor: "#1e1e1e",
    color: "#cccccc",
  },
  field: {
    width: "100%",
  },
  label: {
    fontSize: "13px",
    fontWeight: "600",
    color: "#cccccc",
    marginBottom: "8px",
    display: "block",
  },
  input: {
    width: "100%",
    padding: "10px 14px",
    fontSize: "14px",
    backgroundColor: "#252526",
    color: "#cccccc",
    border: "1px solid #3e3e42",
    borderRadius: "6px",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
    "&:focus": {
      outline: "none",
      borderColor: "#007acc",
      boxShadow: "0 0 0 1px #007acc",
    } as any,
    "&::placeholder": {
      color: "#6a6a6a",
    },
    "&:disabled": {
      opacity: 0.5,
      cursor: "not-allowed",
    },
  },
  helperText: {
    fontSize: "12px",
    color: "#858585",
    marginTop: "6px",
    lineHeight: "1.5",
    "& code": {
      backgroundColor: "#252526",
      padding: "2px 6px",
      borderRadius: "3px",
      fontFamily: "monospace",
      fontSize: "11px",
    },
    "& a": {
      color: "#007acc",
      textDecoration: "none",
      "&:hover": {
        textDecoration: "underline",
      },
    },
  },
  button: {
    minWidth: "140px",
    padding: "10px 20px",
    fontSize: "14px",
    fontWeight: "500",
    backgroundColor: "#007acc",
    color: "#ffffff",
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    transition: "all 0.2s ease",
    "&:hover:not(:disabled)": {
      backgroundColor: "#005a9e",
    },
    "&:disabled": {
      opacity: 0.5,
      cursor: "not-allowed",
    },
  },
  messageBar: {
    marginTop: "12px",
    padding: "10px 14px",
    borderRadius: "6px",
    fontSize: "13px",
    display: "flex",
    alignItems: "center",
  },
  successMessage: {
    backgroundColor: "#1e4620",
    color: "#89d185",
    border: "1px solid #2d5a2f",
  },
  successIcon: {
    marginRight: "8px",
    fontSize: "16px",
  },
});

const API_KEY_STORAGE_KEY = "openai_api_key";

const ApiKeyConfig: React.FC<ApiKeyConfigProps> = ({ onApiKeySet }) => {
  const [apiKey, setApiKey] = useState<string>("");
  const [isConfigured, setIsConfigured] = useState<boolean>(false);
  const styles = useStyles();

  useEffect(() => {
    // Check for API key from environment variable first
    const envKey = (process.env as any).OPENAI_API_KEY;
    if (envKey) {
      setApiKey(envKey);
      setIsConfigured(true);
      onApiKeySet(envKey);
      return;
    }
    
    // Check if API key is already stored in localStorage
    const storedKey = localStorage.getItem(API_KEY_STORAGE_KEY);
    if (storedKey) {
      setApiKey(storedKey);
      setIsConfigured(true);
      onApiKeySet(storedKey);
    }
  }, [onApiKeySet]);

  const handleSave = () => {
    if (!apiKey.trim()) {
      return;
    }

    localStorage.setItem(API_KEY_STORAGE_KEY, apiKey.trim());
    setIsConfigured(true);
    onApiKeySet(apiKey.trim());
  };

  const handleChange = (value: string) => {
    setApiKey(value);
    setIsConfigured(false);
  };

  return (
    <div className={styles.container}>
      <div className={styles.field}>
        <label className={styles.label}>OpenAI API Key</label>
        <input
          type="password"
          className={styles.input}
          value={apiKey}
          onChange={(e) => handleChange(e.target.value)}
          placeholder="sk-..."
          disabled={isConfigured}
          onKeyPress={(e) => {
            if (e.key === "Enter" && apiKey.trim()) {
              handleSave();
            }
          }}
        />
        <div className={styles.helperText}>
          {((process.env as any).OPENAI_API_KEY ? (
            <>
              <strong>Using API key from .env file.</strong> To change it, update your .env file and restart the dev server.
            </>
          ) : (
            <>
              Your API key is stored locally and never sent to our servers. Get your key from{" "}
              <a href="https://platform.openai.com/api-keys" target="_blank" rel="noopener noreferrer">
                platform.openai.com
              </a>
              <br />
              <br />
              <strong>Tip:</strong> You can also add your API key to a <code>.env</code> file in the project root:
              <br />
              <code>OPENAI_API_KEY=sk-your-key-here</code>
            </>
          ))}
        </div>
      </div>

      {isConfigured && (
        <div className={`${styles.messageBar} ${styles.successMessage}`}>
          <CheckmarkCircleFilled className={styles.successIcon} />
          API key configured successfully!
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <button
          disabled={!apiKey.trim() || isConfigured}
          onClick={handleSave}
          className={styles.button}
        >
          <KeyRegular style={{ fontSize: "16px", marginRight: "6px", verticalAlign: "middle" }} />
          {isConfigured ? "Configured" : "Save API Key"}
        </button>
      </div>
    </div>
  );
};

export default ApiKeyConfig;
