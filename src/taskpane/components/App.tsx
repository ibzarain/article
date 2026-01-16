import * as React from "react";
import TextInsertion from "./TextInsertion";
import { makeStyles, tokens } from "@fluentui/react-components";
import { insertText } from "../taskpane";

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
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.container}>
        <TextInsertion insertText={insertText} />
      </div>
    </div>
  );
};

export default App;
