import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";

/* global HTMLTextAreaElement */

interface TextInsertionProps {
  insertText: (text: string, searchText: string) => Promise<void>;
}

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [textToInsert, setTextToInsert] = useState<string>("");
  const [searchText, setSearchText] = useState<string>("");

  const handleTextInsertion = async () => {
    if (!textToInsert.trim() || !searchText.trim()) {
      return;
    }
    await props.insertText(textToInsert, searchText);
  };

  const handleTextToInsertChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setTextToInsert(event.target.value);
  };

  const handleSearchTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setSearchText(event.target.value);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Text to insert">
        <Textarea size="large" value={textToInsert} onChange={handleTextToInsertChange} placeholder="Enter text to insert..." />
      </Field>
      <Field className={styles.textAreaField} size="large" label="Find text (insert before)">
        <Textarea size="large" value={searchText} onChange={handleSearchTextChange} placeholder="Enter text to find (case-insensitive)..." />
      </Field>
      <Field className={styles.instructions}>Click the button to insert text before the found text.</Field>
      <Button appearance="primary" disabled={!textToInsert.trim() || !searchText.trim()} size="large" onClick={handleTextInsertion}>
        Insert text
      </Button>
    </div>
  );
};

export default TextInsertion;
