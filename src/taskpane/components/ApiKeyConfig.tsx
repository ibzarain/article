import * as React from "react";
import { useState, useEffect } from "react";
import {
  Button,
  Field,
  Input,
  tokens,
  makeStyles,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import { KeyRegular, CheckmarkCircleFilled } from "@fluentui/react-icons";

interface ApiKeyConfigProps {
  onApiKeySet: (apiKey: string) => void;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    width: "100%",
    maxWidth: "600px",
    margin: "0 auto",
    padding: "24px",
  },
  field: {
    width: "100%",
  },
  input: {
    width: "100%",
  },
  button: {
    minWidth: "120px",
  },
  messageBar: {
    marginTop: "8px",
  },
  successIcon: {
    color: tokens.colorPaletteGreenForeground1,
    marginRight: "8px",
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
      <Field label="OpenAI API Key" className={styles.field}>
        <Input
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
        <div style={{ marginTop: "8px", fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
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
              <code style={{ fontSize: tokens.fontSizeBase100 }}>OPENAI_API_KEY=sk-your-key-here</code>
            </>
          ))}
        </div>
      </Field>

      {isConfigured && (
        <MessageBar intent="success" className={styles.messageBar}>
          <MessageBarBody>
            <CheckmarkCircleFilled className={styles.successIcon} />
            API key configured successfully!
          </MessageBarBody>
        </MessageBar>
      )}

      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <Button
          appearance="primary"
          disabled={!apiKey.trim() || isConfigured}
          onClick={handleSave}
          className={styles.button}
          icon={<KeyRegular />}
        >
          {isConfigured ? "Configured" : "Save API Key"}
        </Button>
      </div>
    </div>
  );
};

export default ApiKeyConfig;
