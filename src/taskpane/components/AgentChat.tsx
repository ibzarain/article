import * as React from "react";
import { useState, useRef, useEffect } from "react";
import {
  tokens,
  makeStyles,
  Spinner,
} from "@fluentui/react-components";
import { SendRegular, SparkleFilled, CheckmarkCircleFilled, DismissCircleFilled } from "@fluentui/react-icons";
import { generateAgentResponse } from "../agent/wordAgent";
import { createChangeTracker } from "../utils/changeTracker";
import { DocumentChange, ChangeTracking } from "../types/changes";
import { setChangeTracker } from "../tools/wordEditWithTracking";

interface AgentChatProps {
  agent: ReturnType<typeof import("../agent/wordAgent").createWordAgent>;
}

interface Message {
  role: "user" | "assistant";
  content: string;
  changes?: DocumentChange[];
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
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
  changesSection: {
    marginTop: "12px",
    paddingTop: "12px",
    borderTop: "1px solid #3e3e42",
  },
  changesTitle: {
    fontSize: "12px",
    fontWeight: "600",
    color: "#858585",
    marginBottom: "8px",
    textTransform: "uppercase",
    letterSpacing: "0.5px",
  },
  changeCard: {
    padding: "10px",
    marginBottom: "8px",
    backgroundColor: "#1e1e1e",
    border: "1px solid #3e3e42",
    borderRadius: "6px",
    fontSize: "12px",
  },
  changeHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    marginBottom: "8px",
  },
  changeType: {
    fontSize: "10px",
    fontWeight: "600",
    textTransform: "uppercase",
    padding: "2px 6px",
    borderRadius: "4px",
    letterSpacing: "0.5px",
  },
  editType: {
    backgroundColor: "#264f78",
    color: "#75beff",
  },
  insertType: {
    backgroundColor: "#1e4620",
    color: "#89d185",
  },
  deleteType: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
  },
  formatType: {
    backgroundColor: "#4a148c",
    color: "#c586c0",
  },
  changeDescription: {
    fontSize: "12px",
    color: "#cccccc",
    flex: 1,
    marginLeft: "8px",
  },
  changeActions: {
    display: "flex",
    gap: "4px",
  },
  actionButton: {
    padding: "4px",
    backgroundColor: "transparent",
    border: "none",
    borderRadius: "4px",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    transition: "background 0.2s ease",
    "&:hover": {
      backgroundColor: "#3e3e42",
    },
  },
  diffContent: {
    marginTop: "8px",
    padding: "8px",
    borderRadius: "4px",
    fontFamily: "'SF Mono', Monaco, 'Cascadia Code', 'Roboto Mono', Consolas, 'Courier New', monospace",
    fontSize: "11px",
    backgroundColor: "#1a1a1a",
    border: "1px solid #2d2d30",
  },
  oldText: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
    padding: "4px 8px",
    borderRadius: "4px",
    marginBottom: "4px",
    textDecoration: "line-through",
    display: "block",
  },
  newText: {
    backgroundColor: "#1e4620",
    color: "#89d185",
    padding: "4px 8px",
    borderRadius: "4px",
    display: "block",
  },
  insertedText: {
    backgroundColor: "#1e4620",
    color: "#89d185",
    padding: "4px 8px",
    borderRadius: "4px",
    display: "block",
  },
  deletedText: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
    padding: "4px 8px",
    borderRadius: "4px",
    textDecoration: "line-through",
    display: "block",
  },
  formatInfo: {
    fontSize: "11px",
    color: "#858585",
    fontStyle: "italic",
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
});

const AgentChat: React.FC<AgentChatProps> = ({ agent }) => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const [changeTracker] = useState<ChangeTracking>(() => createChangeTracker());
  const currentMessageChangesRef = useRef<DocumentChange[]>([]);
  const styles = useStyles();

  // Set up change tracking callback for the tools
  useEffect(() => {
    setChangeTracker((change: DocumentChange) => {
      changeTracker.addChange(change);
      // Add to current message's changes
      currentMessageChangesRef.current.push(change);
    });
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

    // Clear changes for this new message
    currentMessageChangesRef.current = [];

    // Add user message
    const newMessages: Message[] = [...messages, { role: "user", content: userMessage }];
    setMessages(newMessages);

    try {
      // Get response from agent (changes will be tracked automatically via the agent's onChange callback)
      const response = await generateAgentResponse(agent, userMessage);

      // Get the changes that were made during this response
      const changesForThisMessage = [...currentMessageChangesRef.current];

      // Add assistant response with changes
      setMessages([
        ...newMessages,
        { 
          role: "assistant", 
          content: response,
          changes: changesForThisMessage.length > 0 ? changesForThisMessage : undefined
        },
      ]);

      // Clear the changes ref for next message
      currentMessageChangesRef.current = [];
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
      currentMessageChangesRef.current = [];
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

  const handleAcceptChange = async (messageIndex: number, changeId: string) => {
    try {
      await changeTracker.acceptChange(changeId);
      // Update the message to remove the accepted change
      setMessages(prev => {
        const updated = [...prev];
        if (updated[messageIndex]?.changes) {
          updated[messageIndex] = {
            ...updated[messageIndex],
            changes: updated[messageIndex].changes!.filter(c => c.id !== changeId)
          };
        }
        return updated;
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to accept change");
    }
  };

  const handleRejectChange = async (messageIndex: number, changeId: string) => {
    try {
      await changeTracker.rejectChange(changeId);
      // Update the message to remove the rejected change
      setMessages(prev => {
        const updated = [...prev];
        if (updated[messageIndex]?.changes) {
          updated[messageIndex] = {
            ...updated[messageIndex],
            changes: updated[messageIndex].changes!.filter(c => c.id !== changeId)
          };
        }
        return updated;
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to reject change");
    }
  };

  const getTypeClass = (type: DocumentChange["type"]) => {
    switch (type) {
      case "edit":
        return styles.editType;
      case "insert":
        return styles.insertType;
      case "delete":
        return styles.deleteType;
      case "format":
        return styles.formatType;
      default:
        return "";
    }
  };

  const formatChangeDescription = (change: DocumentChange): string => {
    if (change.type === "format" && change.formatChanges) {
      const formatParts: string[] = [];
      if (change.formatChanges.bold !== undefined) {
        formatParts.push(change.formatChanges.bold ? "bold" : "not bold");
      }
      if (change.formatChanges.italic !== undefined) {
        formatParts.push(change.formatChanges.italic ? "italic" : "not italic");
      }
      if (change.formatChanges.underline !== undefined) {
        formatParts.push(change.formatChanges.underline ? "underlined" : "not underlined");
      }
      if (change.formatChanges.fontSize !== undefined) {
        formatParts.push(`font size ${change.formatChanges.fontSize}pt`);
      }
      if (change.formatChanges.fontColor !== undefined) {
        formatParts.push(`color ${change.formatChanges.fontColor}`);
      }
      return formatParts.join(", ");
    }
    return change.description;
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
                  {message.changes && message.changes.length > 0 && (
                    <div className={styles.changesSection}>
                      <div className={styles.changesTitle}>
                        Changes ({message.changes.length})
                      </div>
                      {message.changes.map((change) => (
                        <div key={change.id} className={styles.changeCard}>
                          <div className={styles.changeHeader}>
                            <div style={{ display: "flex", alignItems: "center", flex: 1 }}>
                              <span className={`${styles.changeType} ${getTypeClass(change.type)}`}>
                                {change.type}
                              </span>
                              <div className={styles.changeDescription}>
                                {formatChangeDescription(change)}
                              </div>
                            </div>
                            <div className={styles.changeActions}>
                              <button
                                className={styles.actionButton}
                                onClick={() => handleAcceptChange(index, change.id)}
                                title="Accept change"
                              >
                                <CheckmarkCircleFilled style={{ fontSize: "14px", color: "#89d185" }} />
                              </button>
                              <button
                                className={styles.actionButton}
                                onClick={() => handleRejectChange(index, change.id)}
                                title="Reject change"
                              >
                                <DismissCircleFilled style={{ fontSize: "14px", color: "#f48771" }} />
                              </button>
                            </div>
                          </div>

                          <div className={styles.diffContent}>
                            {change.type === "edit" && (
                              <>
                                {change.oldText && (
                                  <div className={styles.oldText}>
                                    <span style={{ opacity: 0.7 }}>−</span> {change.oldText}
                                  </div>
                                )}
                                {change.newText && (
                                  <div className={styles.newText}>
                                    <span style={{ opacity: 0.7 }}>+</span> {change.newText}
                                  </div>
                                )}
                              </>
                            )}
                            {change.type === "insert" && change.newText && (
                              <div className={styles.insertedText}>
                                <span style={{ opacity: 0.7 }}>+</span> {change.newText}
                              </div>
                            )}
                            {change.type === "delete" && change.oldText && (
                              <div className={styles.deletedText}>
                                <span style={{ opacity: 0.7 }}>−</span> {change.oldText}
                              </div>
                            )}
                            {change.type === "format" && (
                              <div className={styles.formatInfo}>
                                Applied formatting to: "{change.searchText}"
                              </div>
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
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
    </div>
  );
};

export default AgentChat;
