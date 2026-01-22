import * as React from "react";
import { useState, useRef, useEffect } from "react";
import {
  tokens,
  makeStyles,
  Spinner,
} from "@fluentui/react-components";
import { SparkleFilled, CheckmarkCircleFilled, DismissCircleFilled, ArrowUpRegular, EditRegular, ChatRegular, ChevronDownRegular } from "@fluentui/react-icons";
import { generateAgentResponse } from "../agent/wordAgent";
import { createChangeTracker } from "../utils/changeTracker";
import { DocumentChange, ChangeTracking } from "../types/changes";
import { setChangeTracker } from "../tools/wordEditWithTracking";
import { setArticleChangeTracker } from "../tools/articleEditTools";
import { setFastArticleChangeTracker } from "../tools/fastArticleEdit";
import { setHybridArticleChangeTracker, executeArticleInstructionsHybrid } from "../tools/hybridArticleEdit";

interface AgentChatProps {
  agent: ReturnType<typeof import("../agent/wordAgent").createWordAgent>;
}

interface Message {
  role: "user" | "assistant";
  content: string;
  changes?: Array<DocumentChange & { decision?: "accepted" | "rejected" }>;
  messageId?: string;
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
    gap: "0",
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
    gap: "4px",
    maxWidth: "85%",
  },
  messageWithGap: {
    marginTop: "12px",
  },
  userMessage: {
    alignSelf: "flex-end",
  },
  assistantMessage: {
    alignSelf: "flex-start",
  },
  messageBubble: {
    padding: "12px 16px",
    borderRadius: "12px",
    fontSize: "13px",
    lineHeight: "1.5",
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
    padding: "6px 10px",
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
    minHeight: "40px",
    maxHeight: "200px",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
    fontSize: "13px",
    backgroundColor: "#0d1117",
    color: "#c9d1d9",
    border: "1px solid #30363d",
    borderRadius: "8px",
    padding: "6px 12px",
    paddingBottom: "36px",
    position: "relative",
    resize: "none",
    overflowY: "auto",
    lineHeight: "1.5",
    whiteSpace: "pre-wrap",
    wordWrap: "break-word",
    overflowWrap: "break-word",
    textDecoration: "none",
    textDecorationLine: "none",
    scrollbarWidth: "thin",
    scrollbarColor: "#30363d #0d1117",
    "&::-webkit-scrollbar": {
      width: "8px",
    },
    "&::-webkit-scrollbar-track": {
      background: "#0d1117",
    },
    "&::-webkit-scrollbar-thumb": {
      background: "#30363d",
      borderRadius: "6px",
      "&:hover": {
        background: "#484f58",
      },
    },
    "&:focus": {
      outline: "none",
      borderColor: "#1f6feb",
      boxShadow: "0 0 0 3px rgba(31, 111, 235, 0.1)",
    } as any,
    "&::placeholder": {
      color: "#6e7681",
    },
  },
  buttonRow: {
    position: "absolute",
    bottom: "2px",
    left: "2px",
    right: "2px",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    pointerEvents: "none",
    zIndex: 10,
  },
  buttonRowLeft: {
    display: "flex",
    gap: "4px",
    alignItems: "center",
    pointerEvents: "auto",
  },
  buttonRowRight: {
    display: "flex",
    gap: "4px",
    alignItems: "center",
    pointerEvents: "auto",
  },
  modeSelector: {
    padding: "4px 8px",
    fontSize: "10px",
    borderRadius: "5px",
    border: "1px solid #30363d",
    backgroundColor: "#21262d",
    color: "#c9d1d9",
    cursor: "pointer",
    fontWeight: "500",
    transition: "background 0.15s ease, border-color 0.15s ease",
    whiteSpace: "nowrap",
    height: "24px",
    display: "flex",
    alignItems: "center",
    gap: "4px",
    position: "relative",
    "&:hover": {
      backgroundColor: "#30363d",
      borderColor: "#484f58",
    } as any,
  },
  modeSelectorDropdown: {
    position: "absolute",
    bottom: "100%",
    left: "0",
    marginBottom: "4px",
    backgroundColor: "#1c2128",
    border: "1px solid #30363d",
    borderRadius: "6px",
    overflow: "hidden",
    boxShadow: "0 4px 12px rgba(0, 0, 0, 0.3)",
    zIndex: 100,
    minWidth: "100px",
  },
  modeSelectorOption: {
    padding: "6px 12px",
    fontSize: "11px",
    color: "#c9d1d9",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    gap: "6px",
    backgroundColor: "transparent",
    border: "none",
    width: "100%",
    textAlign: "left",
    "&:hover": {
      backgroundColor: "#30363d",
    } as any,
  },
  sendButton: {
    width: "24px",
    height: "24px",
    minWidth: "24px",
    backgroundColor: "#1f6feb",
    color: "#ffffff",
    border: "none",
    borderRadius: "5px",
    fontSize: "12px",
    fontWeight: "500",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    transition: "all 0.15s ease",
    pointerEvents: "auto",
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
  bulkButton: {
    padding: "4px 8px",
    fontSize: "10px",
    borderRadius: "5px",
    border: "none",
    backgroundColor: "rgba(137, 209, 133, 0.15)",
    color: "#89d185",
    cursor: "pointer",
    fontWeight: "500",
    transition: "background 0.15s ease",
    whiteSpace: "nowrap",
    height: "24px",
    display: "flex",
    alignItems: "center",
    pointerEvents: "auto",
    "&:hover:not(:disabled)": {
      backgroundColor: "rgba(137, 209, 133, 0.25)",
    } as any,
    "&:active:not(:disabled)": {
      backgroundColor: "rgba(137, 209, 133, 0.35)",
    } as any,
    "&:disabled": {
      opacity: 0.5,
      cursor: "not-allowed",
      backgroundColor: "rgba(48, 54, 61, 0.5)",
    },
  },
  bulkButtonReject: {
    backgroundColor: "rgba(244, 135, 113, 0.15)",
    color: "#f48771",
    "&:hover:not(:disabled)": {
      backgroundColor: "rgba(244, 135, 113, 0.25)",
    } as any,
    "&:active:not(:disabled)": {
      backgroundColor: "rgba(244, 135, 113, 0.35)",
    } as any,
  },
  thinking: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    color: "#8b949e",
    fontSize: "12px",
    fontStyle: "normal",
    padding: "8px 16px",
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
    fontSize: "13px",
    lineHeight: "1.6",
    maxWidth: "400px",
  },
  changeBlock: {
    border: "1px solid #30363d",
    borderRadius: "8px",
    overflow: "hidden",
    backgroundColor: "#161b22",
  },
  changeHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "6px 10px",
    backgroundColor: "#1c2128",
    borderBottom: "1px solid #30363d",
    fontSize: "11px",
  },
  changeHeaderLeft: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    minWidth: 0,
  },
  changeHeaderMeta: {
    display: "flex",
    flexDirection: "column",
    minWidth: 0,
  },
  changeHeaderSecondary: {
    fontSize: "11px",
    color: "#8b949e",
    marginTop: "2px",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    maxWidth: "520px",
  },
  decisionPill: {
    padding: "2px 6px",
    borderRadius: "999px",
    fontSize: "10px",
    fontWeight: 600,
    border: "1px solid #30363d",
  },
  acceptedPill: {
    backgroundColor: "#1e4620",
    color: "#89d185",
    borderColor: "#2d5a2f",
  } as any,
  rejectedPill: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
    borderColor: "#6a2d2d",
  } as any,
  changeType: {
    fontSize: "10px",
    fontWeight: "600",
    textTransform: "uppercase",
    padding: "2px 5px",
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
  changeActions: {
    display: "flex",
    gap: "6px",
  },
  changeActionButton: {
    padding: "3px 6px",
    fontSize: "10px",
    borderRadius: "4px",
    border: "none",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    gap: "3px",
    transition: "all 0.15s ease",
    fontWeight: "500",
  },
  acceptButton: {
    backgroundColor: "#1e4620",
    color: "#89d185",
    "&:hover:not(:disabled)": {
      backgroundColor: "#2d5a2f",
    },
    "&:disabled": {
      opacity: 0.5,
      cursor: "not-allowed",
    },
  },
  rejectButton: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
    "&:hover:not(:disabled)": {
      backgroundColor: "#6a2d2d",
    },
    "&:disabled": {
      opacity: 0.5,
      cursor: "not-allowed",
    },
  },
  changeContent: {
    padding: "10px",
    fontFamily: "'SF Mono', Monaco, 'Cascadia Code', 'Roboto Mono', Consolas, 'Courier New', monospace",
    fontSize: "11px",
    lineHeight: "1.5",
  },
  diffLine: {
    padding: "4px 8px",
    borderRadius: "4px",
    marginBottom: "4px",
    display: "block",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  diffOld: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
    textDecoration: "line-through",
  },
  diffNew: {
    backgroundColor: "#1e4620",
    color: "#89d185",
  },
  diffInsert: {
    backgroundColor: "#1e4620",
    color: "#89d185",
  },
  diffDelete: {
    backgroundColor: "#5a1d1d",
    color: "#f48771",
    textDecoration: "line-through",
  },
  changeDescription: {
    fontSize: "12px",
    color: "#8b949e",
    marginTop: "4px",
  },
  changesContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    marginTop: "8px",
  },
});

const AgentChat: React.FC<AgentChatProps> = ({ agent }) => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [mode, setMode] = useState<"edit" | "ask">("edit");
  const [showModeDropdown, setShowModeDropdown] = useState<boolean>(false);
  const modeDropdownRef = useRef<HTMLDivElement>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const [changeTracker] = useState<ChangeTracking>(() => createChangeTracker());
  const currentMessageChangesRef = useRef<DocumentChange[]>([]);
  const [processingChanges, setProcessingChanges] = useState<Set<string>>(new Set());
  const [bulkIsProcessing, setBulkIsProcessing] = useState<boolean>(false);
  const messageIdCounter = useRef<number>(0);
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

  // Close mode dropdown when clicking outside
  useEffect(() => {
    if (!showModeDropdown) {
      return undefined;
    }

    const handleClickOutside = (event: MouseEvent) => {
      if (modeDropdownRef.current && !modeDropdownRef.current.contains(event.target as Node)) {
        setShowModeDropdown(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [showModeDropdown]);

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
    const userMessageId = `msg-${messageIdCounter.current++}`;
    const newMessages: Message[] = [...messages, { role: "user", content: userMessage, messageId: userMessageId }];
    setMessages(newMessages);

    try {
      // HYBRID PATH: Algorithm for parsing/finding, minimal AI for insertion
      // Match ARTICLE X-Y where X is any letter and Y is any number (e.g., A-1, X-67)
      // Also match instructions that start with article name followed by edits
      const hasArticleInstructions = /ARTICLE\s+[A-Z]-\d+/i.test(userMessage) && (/\.\d+\s+(Add|Delete|Substitute|Replace)/i.test(userMessage) || /\.\d+\s+[A-Z]/i.test(userMessage));

      let response: string;
      let hybridSucceeded: boolean | null = null;
      let hybridError: string | undefined;
      if (hasArticleInstructions) {
        // Hybrid execution: Algorithm parses/finds, AI only for final insertion (fast like Cursor)
        const result = await executeArticleInstructionsHybrid(userMessage, agent.apiKey, agent.model);
        hybridSucceeded = result.success;
        hybridError = result.error;
        // Keep the visible chat response minimal; UI will show the diffs + accept/reject.
        response = result.success ? "" : `Error: ${result.error || 'Unknown error'}`;
      } else {
        // Get response from agent (changes will be tracked automatically via the agent's onChange callback)
        response = await generateAgentResponse(agent, userMessage);
      }

      // Get changes for this message
      const messageChanges = [...currentMessageChangesRef.current];
      const assistantMessageId = `msg-${messageIdCounter.current++}`;

      // For ARTICLE/hybrid instructions: keep assistant text minimal and avoid verbose summaries.
      if (hasArticleInstructions) {
        if (!hybridSucceeded) {
          response = `Error: ${hybridError || "Unknown error"}`;
        } else if (messageChanges.length === 0) {
          response = "No changes were necessary.";
        } else {
          response = `Proposed ${messageChanges.length} change(s). Review and accept/reject below.`;
        }
      }

      // Add assistant response with associated changes
      setMessages([
        ...newMessages,
        {
          role: "assistant",
          content: response,
          changes: messageChanges.length > 0 ? messageChanges : undefined,
          messageId: assistantMessageId,
        },
      ]);

      // Clear the changes ref for next message
      currentMessageChangesRef.current = [];
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : "An error occurred";
      setError(errorMessage);
      const errorMessageId = `msg-${messageIdCounter.current++}`;
      setMessages([
        ...newMessages,
        {
          role: "assistant",
          content: `Error: ${errorMessage}`,
          messageId: errorMessageId,
        },
      ]);
      currentMessageChangesRef.current = [];
    } finally {
      setIsLoading(false);
    }
  };

  const markDecision = (changeId: string, decision: "accepted" | "rejected", messageId?: string) => {
    setMessages(prev =>
      prev.map(msg => {
        if (messageId && msg.messageId !== messageId) return msg;
        if (!msg.changes) return msg;
        const idx = msg.changes.findIndex(c => c.id === changeId);
        if (idx === -1) return msg;
        const nextChanges = msg.changes.map(c => (c.id === changeId ? { ...c, decision } : c));
        return { ...msg, changes: nextChanges };
      })
    );
  };

  const handleAcceptChange = async (changeId: string, messageId?: string) => {
    setProcessingChanges(prev => new Set(prev).add(changeId));
    try {
      await changeTracker.acceptChange(changeId);
      // Keep the change card in chat; just mark as accepted.
      markDecision(changeId, "accepted", messageId);
    } catch (error) {
      console.error("Error accepting change:", error);
    } finally {
      setProcessingChanges(prev => {
        const next = new Set(prev);
        next.delete(changeId);
        return next;
      });
    }
  };

  const handleRejectChange = async (changeId: string, messageId?: string) => {
    setProcessingChanges(prev => new Set(prev).add(changeId));
    try {
      await changeTracker.rejectChange(changeId);
      // Keep the change card in chat; just mark as rejected.
      markDecision(changeId, "rejected", messageId);
    } catch (error) {
      console.error("Error rejecting change:", error);
    } finally {
      setProcessingChanges(prev => {
        const next = new Set(prev);
        next.delete(changeId);
        return next;
      });
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

  const getSecondaryHeaderText = (change: DocumentChange): string | null => {
    // Avoid repeating huge "Inserted ... before ..." strings (diff block already shows content).
    if (change.type === "format") {
      return formatChangeDescription(change);
    }
    if (change.searchText) {
      return `Near "${change.searchText}"`;
    }
    if (change.location) {
      return change.location;
    }
    return null;
  };

  const getAllPendingChangeRefs = () => {
    const refs: Array<{ changeId: string; messageId?: string }> = [];
    for (const msg of messages) {
      if (!msg.changes || msg.changes.length === 0) continue;
      for (const c of msg.changes) {
        if (!c.decision) {
          refs.push({ changeId: c.id, messageId: msg.messageId });
        }
      }
    }
    return refs;
  };

  const pendingChangeCount = getAllPendingChangeRefs().length;

  const handleAcceptAll = async () => {
    const refs = getAllPendingChangeRefs();
    if (refs.length === 0) return;
    setBulkIsProcessing(true);
    try {
      for (const { changeId, messageId } of refs) {
        // eslint-disable-next-line no-await-in-loop
        await handleAcceptChange(changeId, messageId);
      }
    } finally {
      setBulkIsProcessing(false);
    }
  };

  const handleRejectAll = async () => {
    const refs = getAllPendingChangeRefs();
    if (refs.length === 0) return;
    setBulkIsProcessing(true);
    try {
      for (const { changeId, messageId } of refs) {
        // eslint-disable-next-line no-await-in-loop
        await handleRejectChange(changeId, messageId);
      }
    } finally {
      setBulkIsProcessing(false);
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
            messages.map((message, index) => {
              const prevMessage = index > 0 ? messages[index - 1] : null;
              const shouldHaveGap = index === 0 || 
                (prevMessage && prevMessage.role !== message.role);
              
              return (
              <div
                key={message.messageId || index}
                className={`${styles.message} ${message.role === "user" ? styles.userMessage : styles.assistantMessage
                  } ${shouldHaveGap ? styles.messageWithGap : ""}`}
              >
                <div
                  className={`${styles.messageBubble} ${message.role === "user" ? styles.userBubble : styles.assistantBubble
                    }`}
                  style={{ whiteSpace: "pre-wrap" }}
                >
                  {message.content}
                </div>
                
                {/* Show changes inline for assistant messages */}
                {message.role === "assistant" && message.changes && message.changes.length > 0 && (
                  <div className={styles.changesContainer}>
                    {message.changes.map((change) => {
                      const isProcessing = processingChanges.has(change.id);
                      const decision = change.decision;
                      const secondary = getSecondaryHeaderText(change);
                      return (
                        <div key={change.id} className={styles.changeBlock}>
                          <div className={styles.changeHeader}>
                            <div className={styles.changeHeaderLeft}>
                              <span className={`${styles.changeType} ${getTypeClass(change.type)}`}>
                                {change.type}
                              </span>
                              <div className={styles.changeHeaderMeta}>
                                {secondary && (
                                  <div className={styles.changeHeaderSecondary}>
                                    {secondary}
                                  </div>
                                )}
                              </div>
                            </div>
                            <div className={styles.changeActions}>
                              {decision ? (
                                <span
                                  className={`${styles.decisionPill} ${
                                    decision === "accepted" ? styles.acceptedPill : styles.rejectedPill
                                  }`}
                                >
                                  {decision === "accepted" ? "Accepted" : "Rejected"}
                                </span>
                              ) : (
                                <>
                                  <button
                                    className={`${styles.changeActionButton} ${styles.acceptButton}`}
                                    onClick={() => handleAcceptChange(change.id, message.messageId)}
                                    disabled={isProcessing || bulkIsProcessing}
                                    title="Accept change"
                                  >
                                    {isProcessing ? (
                                      <Spinner size="tiny" />
                                    ) : (
                                      <>
                                        <CheckmarkCircleFilled style={{ fontSize: "12px" }} />
                                        Accept
                                      </>
                                    )}
                                  </button>
                                  <button
                                    className={`${styles.changeActionButton} ${styles.rejectButton}`}
                                    onClick={() => handleRejectChange(change.id, message.messageId)}
                                    disabled={isProcessing || bulkIsProcessing}
                                    title="Reject change"
                                  >
                                    {isProcessing ? (
                                      <Spinner size="tiny" />
                                    ) : (
                                      <>
                                        <DismissCircleFilled style={{ fontSize: "12px" }} />
                                        Reject
                                      </>
                                    )}
                                  </button>
                                </>
                              )}
                            </div>
                          </div>
                          <div className={styles.changeContent}>
                            {change.type === "edit" && (
                              <>
                                {change.oldText && (
                                  <div className={`${styles.diffLine} ${styles.diffOld}`}>
                                    <span style={{ opacity: 0.7 }}>−</span> {change.oldText}
                                  </div>
                                )}
                                {change.newText && (
                                  <div className={`${styles.diffLine} ${styles.diffNew}`}>
                                    <span style={{ opacity: 0.7 }}>+</span> {change.newText}
                                  </div>
                                )}
                              </>
                            )}
                            {change.type === "insert" && change.newText && (
                              <div className={`${styles.diffLine} ${styles.diffInsert}`}>
                                <span style={{ opacity: 0.7 }}>+</span> {change.newText}
                              </div>
                            )}
                            {change.type === "delete" && change.oldText && (
                              <div className={`${styles.diffLine} ${styles.diffDelete}`}>
                                <span style={{ opacity: 0.7 }}>−</span> {change.oldText}
                              </div>
                            )}
                            {change.type === "format" && (
                              <div className={styles.changeDescription}>
                                Applied formatting to: "{change.searchText || 'selected text'}"
                              </div>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
            })
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
              placeholder={mode === "edit" ? "Ask me to edit your document..." : "Ask me a question..."}
              disabled={isLoading}
              rows={1}
              spellCheck={false}
              autoComplete="off"
              data-gramm="false"
              data-gramm_editor="false"
              data-enable-grammarly="false"
            />
            <div className={styles.buttonRow}>
              <div className={styles.buttonRowLeft} ref={modeDropdownRef}>
                <button
                  className={styles.modeSelector}
                  type="button"
                  onClick={() => setShowModeDropdown(!showModeDropdown)}
                  title={`${mode === "edit" ? "Edit" : "Ask"} mode`}
                >
                  {mode === "edit" ? (
                    <EditRegular style={{ fontSize: "12px" }} />
                  ) : (
                    <ChatRegular style={{ fontSize: "12px" }} />
                  )}
                  {mode === "edit" ? "Edit" : "Ask"}
                  <ChevronDownRegular style={{ fontSize: "10px", marginLeft: "2px" }} />
                </button>
                {showModeDropdown && (
                  <div className={styles.modeSelectorDropdown}>
                    <button
                      className={styles.modeSelectorOption}
                      type="button"
                      onClick={() => {
                        setMode("edit");
                        setShowModeDropdown(false);
                      }}
                    >
                      <EditRegular style={{ fontSize: "12px" }} />
                      Edit
                    </button>
                    <button
                      className={styles.modeSelectorOption}
                      type="button"
                      onClick={() => {
                        setMode("ask");
                        setShowModeDropdown(false);
                      }}
                    >
                      <ChatRegular style={{ fontSize: "12px" }} />
                      Ask
                    </button>
                  </div>
                )}
              </div>
              <div className={styles.buttonRowRight}>
                {pendingChangeCount > 0 && (
                  <>
                    <button
                      className={styles.bulkButton}
                      type="button"
                      onClick={handleAcceptAll}
                      disabled={bulkIsProcessing || isLoading}
                      title="Accept all pending changes"
                    >
                      Accept all ({pendingChangeCount})
                    </button>
                    <button
                      className={`${styles.bulkButton} ${styles.bulkButtonReject}`}
                      type="button"
                      onClick={handleRejectAll}
                      disabled={bulkIsProcessing || isLoading}
                      title="Reject all pending changes"
                    >
                      Reject all
                    </button>
                  </>
                )}
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
                    <ArrowUpRegular style={{ fontSize: "14px" }} />
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default AgentChat;
