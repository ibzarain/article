import * as React from "react";
import { useState, useRef, useEffect } from "react";
import {
  Button,
  Textarea,
  tokens,
  makeStyles,
  Spinner,
  Card,
  MessageBar,
  MessageBarBody,
  Divider,
} from "@fluentui/react-components";
import { SendRegular, SparkleFilled } from "@fluentui/react-icons";
import { Agent } from "ai";
import { generateAgentResponse } from "../agent/wordAgent";
import DiffView from "./DiffView";
import { createChangeTracker } from "../utils/changeTracker";
import { DocumentChange, ChangeTracking } from "../types/changes";
import { setChangeTracker } from "../tools/wordEditWithTracking";

interface AgentChatProps {
  agent: Agent;
}

interface Message {
  role: "user" | "assistant";
  content: string;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    width: "100%",
    maxWidth: "800px",
    margin: "0 auto",
    height: "100%",
  },
  messagesContainer: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    overflowY: "auto",
    padding: "16px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    minHeight: "400px",
    maxHeight: "600px",
  },
  message: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  userMessage: {
    alignSelf: "flex-end",
    maxWidth: "80%",
  },
  assistantMessage: {
    alignSelf: "flex-start",
    maxWidth: "80%",
  },
  messageCard: {
    padding: "12px 16px",
  },
  userCard: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
  },
  assistantCard: {
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  inputContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  inputRow: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-end",
  },
  textarea: {
    flex: 1,
    minHeight: "80px",
    fontFamily: tokens.fontFamilyBase,
    fontSize: tokens.fontSizeBase300,
  },
  sendButton: {
    minWidth: "100px",
    height: "fit-content",
  },
  statusBar: {
    padding: "8px 12px",
    borderRadius: tokens.borderRadiusMedium,
  },
  thinking: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  emptyState: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "12px",
    padding: "40px",
    color: tokens.colorNeutralForeground3,
    textAlign: "center",
  },
  emptyStateIcon: {
    fontSize: "48px",
  },
  emptyStateText: {
    fontSize: tokens.fontSizeBase400,
  },
});

const AgentChat: React.FC<AgentChatProps> = ({ agent }) => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const [changeTracker] = useState<ChangeTracking>(() => createChangeTracker());
  const [changes, setChanges] = useState<DocumentChange[]>([]);
  const styles = useStyles();

  // Set up change tracking callback for the tools
  useEffect(() => {
    setChangeTracker((change: DocumentChange) => {
      changeTracker.addChange(change);
      // Update state to trigger re-render
      setChanges([...changeTracker.changes]);
    });
  }, [changeTracker]);

  // Sync changes state with tracker
  useEffect(() => {
    setChanges([...changeTracker.changes]);
  }, [changeTracker]);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  const handleSend = async () => {
    if (!input.trim() || isLoading) {
      return;
    }

    const userMessage = input.trim();
    setInput("");
    setError(null);
    setIsLoading(true);

    // Add user message
    const newMessages: Message[] = [...messages, { role: "user", content: userMessage }];
    setMessages(newMessages);

    try {
      // Get response from agent (changes will be tracked automatically via the agent's onChange callback)
      const response = await generateAgentResponse(agent, userMessage);

      // Add assistant response
      setMessages([
        ...newMessages,
        { role: "assistant", content: response },
      ]);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : "An error occurred";
      setError(errorMessage);
      setMessages([
        ...newMessages,
        {
          role: "assistant",
          content: `Error: ${errorMessage}`,
        },
      ]);
    } finally {
      setIsLoading(false);
    }
  };

  const handleKeyPress = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleAcceptChange = async (id: string) => {
    try {
      await changeTracker.acceptChange(id);
      setChanges([...changeTracker.changes]);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to accept change");
    }
  };

  const handleRejectChange = async (id: string) => {
    try {
      await changeTracker.rejectChange(id);
      setChanges([...changeTracker.changes]);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to reject change");
    }
  };

  const handleAcceptAll = async () => {
    try {
      await changeTracker.acceptAll();
      setChanges([...changeTracker.changes]);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to accept all changes");
    }
  };

  const handleRejectAll = async () => {
    try {
      await changeTracker.rejectAll();
      setChanges([...changeTracker.changes]);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to reject all changes");
    }
  };

  return (
    <div className={styles.container}>
      <div style={{ display: "flex", gap: "16px", height: "100%" }}>
        <div style={{ flex: 1, display: "flex", flexDirection: "column" }}>
          <div className={styles.messagesContainer}>
            {messages.length === 0 ? (
              <div className={styles.emptyState}>
                <SparkleFilled className={styles.emptyStateIcon} />
                <div className={styles.emptyStateText}>
                  <strong>AI Document Editor</strong>
                  <br />
                  Ask me to edit your Word document! I can read, edit, insert, delete, and format text.
                  <br />
                  <br />
                  Try: "Make the first paragraph bold" or "Replace 'hello' with 'hi'"
                </div>
              </div>
            ) : (
              messages.map((message, index) => (
                <div
                  key={index}
                  className={`${styles.message} ${
                    message.role === "user" ? styles.userMessage : styles.assistantMessage
                  }`}
                >
                  <Card
                    className={`${styles.messageCard} ${
                      message.role === "user" ? styles.userCard : styles.assistantCard
                    }`}
                  >
                    {message.content}
                  </Card>
                </div>
              ))
            )}
            {isLoading && (
              <div className={styles.thinking}>
                <Spinner size="tiny" />
                <span>Thinking and editing...</span>
              </div>
            )}
            <div ref={messagesEndRef} />
          </div>

          {error && (
            <MessageBar intent="error">
              <MessageBarBody>{error}</MessageBarBody>
            </MessageBar>
          )}

          <div className={styles.inputContainer}>
            <div className={styles.inputRow}>
              <Textarea
                className={styles.textarea}
                value={input}
                onChange={(e) => setInput(e.target.value)}
                onKeyPress={handleKeyPress}
                placeholder="Ask me to edit your document... (e.g., 'Make the title bold' or 'Replace all instances of X with Y')"
                disabled={isLoading}
                resize="vertical"
              />
              <Button
                appearance="primary"
                disabled={!input.trim() || isLoading}
                onClick={handleSend}
                className={styles.sendButton}
                icon={isLoading ? <Spinner size="tiny" /> : <SendRegular />}
              >
                {isLoading ? "Sending..." : "Send"}
              </Button>
            </div>
          </div>
        </div>

        <Divider vertical />

        <div style={{ width: "350px", display: "flex", flexDirection: "column" }}>
          <DiffView
            changes={changes}
            onAccept={handleAcceptChange}
            onReject={handleRejectChange}
            onAcceptAll={handleAcceptAll}
            onRejectAll={handleRejectAll}
          />
        </div>
      </div>
    </div>
  );
};

export default AgentChat;
