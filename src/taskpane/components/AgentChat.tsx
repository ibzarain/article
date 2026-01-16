import * as React from "react";
import { useState, useRef, useEffect } from "react";
import {
  tokens,
  makeStyles,
  Spinner,
} from "@fluentui/react-components";
import { SendRegular, SparkleFilled } from "@fluentui/react-icons";
import { generateAgentResponse } from "../agent/wordAgent";
import DiffView from "./DiffView";
import { createChangeTracker } from "../utils/changeTracker";
import { DocumentChange, ChangeTracking } from "../types/changes";
import { setChangeTracker } from "../tools/wordEditWithTracking";

interface AgentChatProps {
  agent: ReturnType<typeof import("../agent/wordAgent").createWordAgent>;
}

interface Message {
  role: "user" | "assistant";
  content: string;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    width: "100%",
    height: "100%",
    backgroundColor: "#1e1e1e",
    color: "#cccccc",
    overflow: "hidden",
  },
  chatPanel: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    backgroundColor: "#1e1e1e",
    borderRight: "1px solid #3e3e42",
    height: "100%",
    overflow: "hidden",
  },
  messagesContainer: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    overflowY: "auto",
    overflowX: "hidden",
    padding: "24px",
    scrollbarWidth: "thin",
    scrollbarColor: "#424242 #1e1e1e",
    "&::-webkit-scrollbar": {
      width: "8px",
    },
    "&::-webkit-scrollbar-track": {
      background: "#1e1e1e",
    },
    "&::-webkit-scrollbar-thumb": {
      background: "#424242",
      borderRadius: "4px",
      "&:hover": {
        background: "#4e4e4e",
      },
    },
  },
  message: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    maxWidth: "85%",
  },
  userMessage: {
    alignSelf: "flex-end",
  },
  assistantMessage: {
    alignSelf: "flex-start",
  },
  messageBubble: {
    padding: "12px 16px",
    borderRadius: "8px",
    fontSize: "14px",
    lineHeight: "1.5",
    wordWrap: "break-word",
  },
  userBubble: {
    backgroundColor: "#007acc",
    color: "#ffffff",
    borderBottomRightRadius: "4px",
  },
  assistantBubble: {
    backgroundColor: "#252526",
    color: "#cccccc",
    border: "1px solid #3e3e42",
    borderBottomLeftRadius: "4px",
  },
  inputContainer: {
    padding: "16px 24px",
    borderTop: "1px solid #3e3e42",
    backgroundColor: "#252526",
    flexShrink: 0,
  },
  inputRow: {
    display: "flex",
    gap: "12px",
    alignItems: "flex-end",
  },
  textarea: {
    flex: 1,
    minHeight: "44px",
    maxHeight: "200px",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
    fontSize: "14px",
    backgroundColor: "#1e1e1e",
    color: "#cccccc",
    border: "1px solid #3e3e42",
    borderRadius: "22px",
    padding: "10px 20px",
    resize: "none",
    lineHeight: "1.5",
    "&:focus": {
      outline: "none",
      borderColor: "#007acc",
      boxShadow: "0 0 0 1px #007acc",
    } as any,
    "&::placeholder": {
      color: "#6a6a6a",
    },
  },
  sendButton: {
    width: "44px",
    height: "44px",
    minWidth: "44px",
    backgroundColor: "#007acc",
    color: "#ffffff",
    border: "none",
    borderRadius: "50%",
    fontSize: "14px",
    fontWeight: "500",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    transition: "all 0.2s ease",
    "&:hover:not(:disabled)": {
      backgroundColor: "#005a9e",
      transform: "scale(1.05)",
    },
    "&:active:not(:disabled)": {
      transform: "scale(0.95)",
    },
    "&:disabled": {
      opacity: 0.5,
      cursor: "not-allowed",
    },
  },
  thinking: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    color: "#858585",
    fontSize: "13px",
    fontStyle: "italic",
    padding: "8px 16px",
  },
  emptyState: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "16px",
    padding: "60px 40px",
    color: "#858585",
    textAlign: "center",
  },
  emptyStateIcon: {
    fontSize: "48px",
    color: "#007acc",
    opacity: 0.7,
  },
  emptyStateText: {
    fontSize: "14px",
    lineHeight: "1.6",
    maxWidth: "400px",
  },
  divider: {
    width: "1px",
    backgroundColor: "#3e3e42",
    margin: "0",
  },
  changesPanel: {
    width: "380px",
    display: "flex",
    flexDirection: "column",
    backgroundColor: "#252526",
    borderLeft: "1px solid #3e3e42",
  },
});

const AgentChat: React.FC<AgentChatProps> = ({ agent }) => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
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

  // Auto-resize textarea
  useEffect(() => {
    const textarea = textareaRef.current;
    if (textarea) {
      textarea.style.height = "auto";
      textarea.style.height = `${Math.min(textarea.scrollHeight, 200)}px`;
    }
  }, [input]);

  const handleSend = async () => {
    if (!input.trim() || isLoading) {
      return;
    }

    const userMessage = input.trim();
    setInput("");
    setError(null);
    setIsLoading(true);
    
    // Reset textarea height
    if (textareaRef.current) {
      textareaRef.current.style.height = "auto";
    }

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
      <div className={styles.chatPanel}>
        <div className={styles.messagesContainer}>
          {messages.length === 0 ? (
            <div className={styles.emptyState}>
              <SparkleFilled className={styles.emptyStateIcon} />
              <div className={styles.emptyStateText}>
                <strong style={{ color: "#ffffff", fontSize: "16px", marginBottom: "8px", display: "block" }}>
                  AI Document Editor
                </strong>
                Ask me to edit your Word document! I can read, edit, insert, delete, and format text.
                <br />
                <br />
                <span style={{ color: "#858585", fontSize: "13px" }}>
                  Try: "Make the first paragraph bold" or "Replace 'hello' with 'hi'"
                </span>
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
                <div
                  className={`${styles.messageBubble} ${
                    message.role === "user" ? styles.userBubble : styles.assistantBubble
                  }`}
                >
                  {message.content}
                </div>
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
          <div style={{ 
            padding: "12px 24px", 
            backgroundColor: "#3a1f1f", 
            color: "#f48771", 
            borderTop: "1px solid #5a2f2f",
            fontSize: "13px"
          }}>
            {error}
          </div>
        )}

        <div className={styles.inputContainer}>
          <div className={styles.inputRow}>
            <textarea
              ref={textareaRef}
              className={styles.textarea}
              value={input}
              onChange={(e) => setInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !e.shiftKey) {
                  e.preventDefault();
                  handleSend();
                }
              }}
              placeholder="Ask me to edit your document..."
              disabled={isLoading}
              rows={1}
            />
            <button
              disabled={!input.trim() || isLoading}
              onClick={handleSend}
              className={styles.sendButton}
              title="Send message"
            >
              {isLoading ? (
                <Spinner size="tiny" />
              ) : (
                <SendRegular style={{ fontSize: "18px" }} />
              )}
            </button>
          </div>
        </div>
      </div>

      <div className={styles.divider} />

      <div className={styles.changesPanel}>
        <DiffView
          changes={changes}
          onAccept={handleAcceptChange}
          onReject={handleRejectChange}
          onAcceptAll={handleAcceptAll}
          onRejectAll={handleRejectAll}
        />
      </div>
    </div>
  );
};

export default AgentChat;
