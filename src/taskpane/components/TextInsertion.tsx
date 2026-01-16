import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles, Spinner } from "@fluentui/react-components";

/* global HTMLTextAreaElement */

interface TextInsertionProps {
  insertText: (text: string, searchText: string) => Promise<void>;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "20px",
    width: "100%",
    maxWidth: "600px",
    margin: "0 auto",
  },
  field: {
    width: "100%",
  },
  label: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "8px",
    color: tokens.colorNeutralForeground1,
  },
  textarea: {
    width: "100%",
    minHeight: "80px",
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase300,
  },
  buttonContainer: {
    display: "flex",
    justifyContent: "flex-end",
    marginTop: "8px",
  },
  button: {
    minWidth: "120px",
  },
  helperText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginTop: "4px",
  },
  statusMessage: {
    fontSize: tokens.fontSizeBase300,
    padding: "8px 12px",
    borderRadius: tokens.borderRadiusMedium,
    marginTop: "8px",
  },
  success: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground2,
  },
  error: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorPaletteRedForeground2,
  },
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [textToInsert, setTextToInsert] = useState<string>("");
  const [searchText, setSearchText] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [statusMessage, setStatusMessage] = useState<{ text: string; type: "success" | "error" | null }>({ text: "", type: null });

  const handleTextInsertion = async () => {
    if (!textToInsert.trim() || !searchText.trim()) {
      return;
    }
    
    setIsLoading(true);
    setStatusMessage({ text: "", type: null });
    
    try {
      await props.insertText(textToInsert, searchText);
      setStatusMessage({ text: "Text inserted successfully!", type: "success" });
      // Clear inputs after successful insertion
      setTextToInsert("");
      setSearchText("");
    } catch (error) {
      setStatusMessage({ text: `Error: ${error instanceof Error ? error.message : "Failed to insert text"}`, type: "error" });
    } finally {
      setIsLoading(false);
    }
  };

  const handleTextToInsertChange = (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setTextToInsert(event.target.value);
    setStatusMessage({ text: "", type: null });
  };

  const handleSearchTextChange = (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setSearchText(event.target.value);
    setStatusMessage({ text: "", type: null });
  };

  const styles = useStyles();
  const isDisabled = !textToInsert.trim() || !searchText.trim() || isLoading;

  return (
    <div className={styles.container}>
      <div className={styles.field}>
        <div className={styles.label}>Text to insert</div>
        <Textarea
          className={styles.textarea}
          value={textToInsert}
          onChange={handleTextToInsertChange}
          placeholder='e.g., "hello"'
          disabled={isLoading}
          resize="vertical"
        />
        <div className={styles.helperText}>Enter the text you want to insert into the document</div>
      </div>

      <div className={styles.field}>
        <div className={styles.label}>Find text (insert before)</div>
        <Textarea
          className={styles.textarea}
          value={searchText}
          onChange={handleSearchTextChange}
          placeholder='e.g., "This modal is very"'
          disabled={isLoading}
          resize="vertical"
        />
        <div className={styles.helperText}>Enter text to find. The insertion will happen before the first match (case-insensitive)</div>
      </div>

      {statusMessage.text && (
        <div className={`${styles.statusMessage} ${statusMessage.type === "success" ? styles.success : styles.error}`}>
          {statusMessage.text}
        </div>
      )}

      <div className={styles.buttonContainer}>
        <Button
          appearance="primary"
          disabled={isDisabled}
          size="large"
          onClick={handleTextInsertion}
          className={styles.button}
          icon={isLoading ? <Spinner size="tiny" /> : undefined}
        >
          {isLoading ? "Inserting..." : "Insert Text"}
        </Button>
      </div>
    </div>
  );
};

export default TextInsertion;
