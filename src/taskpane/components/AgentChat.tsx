import * as React from "react";
import { useState, useRef, useEffect } from "react";
import {
  tokens,
  makeStyles,
  Spinner,
} from "@fluentui/react-components";
import { SendRegular, SparkleFilled } from "@fluentui/react-icons";
import { generateAgentResponse } from "../agent/wordAgent";
import { createChangeTracker } from "../utils/changeTracker";
import { DocumentChange, ChangeTracking } from "../types/changes";
import { setChangeTracker } from "../tools/wordEditWithTracking";
import { setArticleChangeTracker } from "../tools/articleEditTools";
import { setFastArticleChangeTracker } from "../tools/fastArticleEdit";
import { setHybridArticleChangeTracker, executeArticleInstructionsHybrid } from "../tools/hybridArticleEdit";
import PendingChanges from "./PendingChanges";

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
    flexDirection: "column",
    width: "100%",
    height: "100%",
    backgroundColor: "#0d1117",
    color: "#c9d1d9",
    overflow: "hidden",
  },
  chatPanel: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    backgroundColor: "#0d1117",
    height: "100%",
    overflow: "hidden",
  },
  messagesContainer: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    gap: "20px",
    overflowY: "auto",
    overflowX: "hidden",
    padding: "20px 24px",
    scrollbarWidth: "thin",
    scrollbarColor: "#30363d #0d1117",
    "&::-webkit-scrollbar": {
      width: "10px",
    },
    "&::-webkit-scrollbar-track": {
      background: "#0d1117",
    },
    "&::-webkit-scrollbar-thumb": {
      background: "#30363d",
      borderRadius: "5px",
      "&:hover": {
        background: "#484f58",
      },
    },
  },
  message: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    maxWidth: "85%",
  },
  userMessage: {
    alignSelf: "flex-end",
  },
  assistantMessage: {
    alignSelf: "flex-start",
  },
  messageBubble: {
    padding: "14px 18px",
    borderRadius: "12px",
    fontSize: "14px",
    lineHeight: "1.6",
    wordWrap: "break-word",
    boxShadow: "0 1px 2px rgba(0, 0, 0, 0.1)",
  },
  userBubble: {
    backgroundColor: "#0969da",
    color: "#ffffff",
    borderBottomRightRadius: "4px",
  },
  assistantBubble: {
    backgroundColor: "#161b22",
    color: "#c9d1d9",
    border: "1px solid #30363d",
    borderBottomLeftRadius: "4px",
  },
  inputContainer: {
    padding: "16px 20px",
    borderTop: "1px solid #21262d",
    backgroundColor: "#0d1117",
    flexShrink: 0,
  },
  inputRow: {
    position: "relative",
    display: "flex",
    alignItems: "center",
  },
  textarea: {
    flex: 1,
    minHeight: "44px",
    maxHeight: "200px",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
    fontSize: "14px",
    backgroundColor: "#0d1117",
    color: "#c9d1d9",
    border: "1px solid #30363d",
    borderRadius: "12px",
    padding: "10px 50px 10px 16px",
    resize: "none",
    lineHeight: "1.5",
    whiteSpace: "pre-wrap",
    wordWrap: "break-word",
    overflowWrap: "break-word",
    textDecoration: "none",
    textDecorationLine: "none",
    "&:focus": {
      outline: "none",
      borderColor: "#1f6feb",
      boxShadow: "0 0 0 3px rgba(31, 111, 235, 0.1)",
    } as any,
    "&::placeholder": {
      color: "#6e7681",
    },
  },
  sendButton: {
    position: "absolute",
    right: "8px",
    bottom: "8px",
    width: "28px",
    height: "28px",
    minWidth: "28px",
    backgroundColor: "#1f6feb",
    color: "#ffffff",
    border: "none",
    borderRadius: "6px",
    fontSize: "14px",
    fontWeight: "500",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    transition: "all 0.15s ease",
    "&:hover:not(:disabled)": {
      backgroundColor: "#0969da",
      transform: "scale(1.05)",
    },
    "&:active:not(:disabled)": {
      transform: "scale(0.95)",
      backgroundColor: "#0860ca",
    },
    "&:disabled": {
      opacity: 0.4,
      cursor: "not-allowed",
      backgroundColor: "#30363d",
    },
  },
  thinking: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    color: "#8b949e",
    fontSize: "13px",
    fontStyle: "normal",
    padding: "10px 18px",
  },
  emptyState: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "20px",
    padding: "60px 40px",
    color: "#8b949e",
    textAlign: "center",
  },
  emptyStateIcon: {
    fontSize: "48px",
    color: "#1f6feb",
    opacity: 0.8,
  },
  emptyStateText: {
    fontSize: "14px",
    lineHeight: "1.7",
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
    const trackChange = async (change: DocumentChange) => {
      await changeTracker.addChange(change);
      // Add to current message's changes
      currentMessageChangesRef.current.push(change);
    };

    setChangeTracker(trackChange);
    setArticleChangeTracker(trackChange);
    setFastArticleChangeTracker(trackChange);
    setHybridArticleChangeTracker(trackChange);
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

  // Handle paste events to preserve formatting
  const handlePaste = (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    e.preventDefault();
    const pastedText = e.clipboardData.getData('text/plain');
    
    // Insert the pasted text at the cursor position, preserving formatting
    const textarea = e.currentTarget;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const currentValue = input;
    
    const newValue = currentValue.substring(0, start) + pastedText + currentValue.substring(end);
    setInput(newValue);
    
    // Set cursor position after the pasted text using requestAnimationFrame for better timing
    requestAnimationFrame(() => {
      const newCursorPos = start + pastedText.length;
      textarea.setSelectionRange(newCursorPos, newCursorPos);
      // Trigger resize after paste
      textarea.style.height = "auto";
      textarea.style.height = `${Math.min(textarea.scrollHeight, 200)}px`;
    });
  };

  const handleSend = async () => {
    if (!input.trim() || isLoading) {
      return;
    }

    // Preserve exact formatting - don't trim, keep as-is
    const userMessage = input;
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
      // HYBRID PATH: Algorithm for parsing/finding, minimal AI for insertion
      // Match ARTICLE X-Y where X is any letter and Y is any number (e.g., A-1, X-67)
      // Also match instructions that start with article name followed by edits
      const hasArticleInstructions = /ARTICLE\s+[A-Z]-\d+/i.test(userMessage) && (/\.\d+\s+(Add|Delete|Substitute|Replace)/i.test(userMessage) || /\.\d+\s+[A-Z]/i.test(userMessage));

      let response: string;
      if (hasArticleInstructions) {
        // Hybrid execution: Algorithm parses/finds, AI only for final insertion (fast like Cursor)
        const result = await executeArticleInstructionsHybrid(userMessage, agent.apiKey, agent.model);
        response = result.success
          ? `Applied ${result.results?.length || 0} article operation(s) successfully. ${result.results?.join('; ') || ''}`
          : `Error: ${result.error || 'Unknown error'}`;
      } else {
        // Get response from agent (changes will be tracked automatically via the agent's onChange callback)
        response = await generateAgentResponse(agent, userMessage);
      }

      // Add assistant response (changes are now shown inline in document)
      setMessages([
        ...newMessages,
        {
          role: "assistant",
          content: response,
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


  return (
    <div className={styles.container}>
      <div className={styles.chatPanel}>
        <div className={styles.messagesContainer}>
          {messages.length === 0 ? (
            <div className={styles.emptyState}>
              <SparkleFilled className={styles.emptyStateIcon} />
              <div className={styles.emptyStateText}>
                <strong style={{ color: "#f0f6fc", fontSize: "18px", marginBottom: "12px", display: "block", fontWeight: "600" }}>
                  AI Document Editor
                </strong>
                Ask me to edit your Word document! I can read, edit, insert, delete, and format text.
                <br />
                <br />
                <span style={{ color: "#8b949e", fontSize: "13px" }}>
                  Try: "Make the first paragraph bold" or "Replace 'hello' with 'hi'"
                </span>
              </div>
            </div>
          ) : (
            messages.map((message, index) => (
              <div
                key={index}
                className={`${styles.message} ${message.role === "user" ? styles.userMessage : styles.assistantMessage
                  }`}
              >
                <div
                  className={`${styles.messageBubble} ${message.role === "user" ? styles.userBubble : styles.assistantBubble
                    }`}
                  style={{ whiteSpace: "pre-wrap" }}
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
              onPaste={handlePaste}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !e.shiftKey) {
                  e.preventDefault();
                  handleSend();
                }
              }}
              placeholder="Ask me to edit your document..."
              disabled={isLoading}
              rows={1}
              spellCheck={false}
              autoComplete="off"
              data-gramm="false"
              data-gramm_editor="false"
              data-enable-grammarly="false"
            />
            <button
              disabled={!input.trim() || isLoading}
              onClick={handleSend}
              className={styles.sendButton}
              title="Send message (Enter)"
              type="button"
            >
              {isLoading ? (
                <Spinner size="tiny" />
              ) : (
                <SendRegular style={{ fontSize: "14px" }} />
              )}
            </button>
          </div>
        </div>
      </div>

      {/* Pending Changes Panel */}
      <PendingChanges changeTracker={changeTracker} />
    </div>
  );
};

export default AgentChat;
