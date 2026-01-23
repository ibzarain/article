import * as React from "react";
import { useState, useRef, useEffect, useMemo } from "react";
import {
  tokens,
  makeStyles,
  Spinner,
} from "@fluentui/react-components";
import { DocumentRegular, CheckmarkCircleFilled, DismissCircleFilled, ArrowUpRegular, EditRegular, ChatRegular, ChevronDownRegular, WeatherSunnyRegular, WeatherMoonRegular } from "@fluentui/react-icons";
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

type ChecklistStatus = "pending" | "in_progress" | "done" | "error";

interface ChecklistStep {
  id: string;
  text: string;
  status: ChecklistStatus;
  error?: string;
}

interface Message {
  role: "user" | "assistant";
  content: string;
  changes?: Array<DocumentChange & { decision?: "accepted" | "rejected" }>;
  checklist?: ChecklistStep[];
  messageId?: string;
}

const createStyles = (isLight: boolean): any => ({
  container: {
    display: "flex",
    flexDirection: "column",
    width: "100%",
    height: "100%",
    backgroundColor: isLight ? "#ffffff" : "#0d1117",
    color: isLight ? "#24292f" : "#c9d1d9",
    overflow: "hidden",
  },
  chatPanel: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    backgroundColor: isLight ? "#ffffff" : "#0d1117",
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
    scrollbarColor: isLight ? "#d0d7de #ffffff" : "#30363d #0d1117",
    "&::-webkit-scrollbar": {
      width: "10px",
    },
    "&::-webkit-scrollbar-track": {
      background: isLight ? "#ffffff" : "#0d1117",
    },
    "&::-webkit-scrollbar-thumb": {
      background: isLight ? "#d0d7de" : "#30363d",
      borderRadius: "5px",
      "&:hover": {
        background: isLight ? "#b1bac4" : "#484f58",
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
    backgroundColor: isLight ? "#f6f8fa" : "#161b22",
    color: isLight ? "#24292f" : "#c9d1d9",
    border: isLight ? "1px solid #d0d7de" : "1px solid #30363d",
    borderBottomLeftRadius: "4px",
  },
  inputContainer: {
    padding: "10px 10px 6px 10px",
    borderTop: isLight ? "1px solid #d0d7de" : "1px solid #21262d",
    backgroundColor: isLight ? "#ffffff" : "#0d1117",
    flexShrink: 0,
  },
  inputRow: {
    display: "flex",
    flexDirection: "column",
    alignItems: "stretch",
    gap: "6px",
  },
  inputBox: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    backgroundColor: isLight ? "#ffffff" : "#0d1117",
    border: isLight ? "1px solid #d0d7de" : "1px solid #30363d",
    borderRadius: "8px",
    padding: "4px 4px 6px 4px",
    "&:focus-within": {
      borderColor: "#1f6feb",
      boxShadow: isLight ? "0 0 0 3px rgba(31, 111, 235, 0.15)" : "0 0 0 3px rgba(31, 111, 235, 0.1)",
    } as any,
  },
  textarea: {
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
    fontSize: "13px",
    backgroundColor: "transparent",
    color: isLight ? "#24292f" : "#c9d1d9",
    border: "none",
    borderRadius: "6px",
    padding: "4px 4px",
    scrollbarGutter: "stable",
    position: "relative",
    resize: "none",
    overflowY: "hidden",
    lineHeight: "1.5",
    whiteSpace: "pre-wrap",
    wordWrap: "break-word",
    overflowWrap: "break-word",
    textDecoration: "none",
    textDecorationLine: "none",
    scrollbarWidth: "thin",
    scrollbarColor: isLight ? "#d0d7de #ffffff" : "#30363d #0d1117",
    minHeight: "24px",
    boxSizing: "border-box",
    "&::-webkit-scrollbar": {
      width: "8px",
    },
    "&::-webkit-scrollbar-track": {
      background: isLight ? "#ffffff" : "#0d1117",
    },
    "&::-webkit-scrollbar-thumb": {
      background: isLight ? "#d0d7de" : "#30363d",
      borderRadius: "6px",
      "&:hover": {
        background: isLight ? "#b1bac4" : "#484f58",
      },
    },
    "&:focus": {
      outline: "none",
    } as any,
    "&::placeholder": {
      color: isLight ? "#656d76" : "#6e7681",
    },
  },
  buttonRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    backgroundColor: "transparent",
    borderRadius: "6px",
    padding: "2px 2px",
    pointerEvents: "auto",
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
  modeSelectWrap: {
    position: "relative",
    display: "flex",
    alignItems: "center",
    height: "24px",
    borderRadius: "6px",
    backgroundColor: isLight ? "#f6f8fa" : "#21262d",
    border: isLight ? "1px solid #d0d7de" : "1px solid #30363d",
    color: isLight ? "#24292f" : "#c9d1d9",
    pointerEvents: "auto",
    userSelect: "none",
    "&:hover": {
      backgroundColor: isLight ? "#e7ecf0" : "#30363d",
    } as any,
    "&:focus-within": {
      borderColor: "#1f6feb",
      boxShadow: isLight ? "0 0 0 2px rgba(31, 111, 235, 0.2)" : "0 0 0 2px rgba(31, 111, 235, 0.15)",
    } as any,
  },
  modeSelectIcon: {
    display: "flex",
    alignItems: "center",
    color: isLight ? "#24292f" : "#c9d1d9",
    opacity: 0.9,
    pointerEvents: "none",
  },
  modeSelectButton: {
    appearance: "none",
    backgroundColor: "transparent",
    color: isLight ? "#24292f" : "#c9d1d9",
    border: "none",
    fontSize: "10px",
    fontWeight: 700,
    letterSpacing: "0.2px",
    outline: "none",
    cursor: "pointer",
    padding: "4px 6px",
    display: "flex",
    alignItems: "center",
    gap: "6px",
    minWidth: 0,
    width: "100%",
    height: "100%",
    textAlign: "left",
  },
  modeSelectChevron: {
    display: "flex",
    alignItems: "center",
    color: isLight ? "#656d76" : "#8b949e",
    opacity: 0.95,
    pointerEvents: "none",
  },
  modeMenu: {
    position: "absolute",
    left: 0,
    bottom: "calc(100% + 6px)",
    minWidth: "100px",
    backgroundColor: isLight ? "#ffffff" : "#161b22",
    border: isLight ? "1px solid #d0d7de" : "1px solid #30363d",
    borderRadius: "6px",
    boxShadow: isLight ? "0 10px 28px rgba(0,0,0,0.15)" : "0 10px 28px rgba(0,0,0,0.45)",
    padding: "4px",
    zIndex: 50,
  },
  modeMenuItem: {
    width: "100%",
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "6px 8px",
    borderRadius: "4px",
    border: "none",
    backgroundColor: "transparent",
    color: isLight ? "#24292f" : "#c9d1d9",
    cursor: "pointer",
    fontSize: "12px",
    fontWeight: 600,
    textAlign: "left",
    "&:hover": {
      backgroundColor: isLight ? "#f6f8fa" : "#21262d",
    } as any,
    "&:focus": {
      outline: "none",
      backgroundColor: isLight ? "#f6f8fa" : "#21262d",
    } as any,
  },
  modeMenuItemActive: {
    backgroundColor: "rgba(31, 111, 235, 0.15)",
    "&:hover": {
      backgroundColor: "rgba(31, 111, 235, 0.22)",
    } as any,
  },
  modeSelector: {
    padding: "4px 8px",
    fontSize: "10px",
    borderRadius: "5px",
    border: "none",
    backgroundColor: "#21262d",
    color: "#c9d1d9",
    cursor: "pointer",
    fontWeight: "500",
    transition: "background 0.15s ease",
    whiteSpace: "nowrap",
    height: "24px",
    display: "flex",
    alignItems: "center",
    gap: "4px",
    "&:hover": {
      backgroundColor: "#30363d",
    } as any,
  },
  modeSelectorActive: {
    backgroundColor: "#1f6feb",
    color: "#ffffff",
    "&:hover": {
      backgroundColor: "#0969da",
    } as any,
  },
  sendButton: {
    width: "24px",
    height: "24px",
    minWidth: "24px",
    backgroundColor: "#1f6feb",
    color: "#ffffff",
    border: "none",
    borderRadius: "4px",
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
      opacity: 1,
      cursor: "not-allowed",
      backgroundColor: isLight ? "#656d76" : "#30363d",
      color: isLight ? "#ffffff" : "#c9d1d9",
    },
  },
  themeToggleButton: {
    width: "24px",
    height: "24px",
    minWidth: "24px",
    backgroundColor: "transparent",
    color: isLight ? "#24292f" : "#c9d1d9",
    border: "none",
    borderRadius: "4px",
    fontSize: "12px",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    transition: "color 0.15s ease",
    pointerEvents: "auto",
    padding: 0,
    "&:hover": {
      color: "#fbbf24",
    },
    "&:active": {
      transform: "scale(0.95)",
    },
  },
  bulkButton: {
    padding: "3px 6px",
    fontSize: "10px",
    borderRadius: "4px",
    border: isLight ? "1px solid #2da44e" : "none",
    backgroundColor: isLight ? "#dafbe1" : "#1e4620",
    color: isLight ? "#1a7f37" : "#89d185",
    cursor: "pointer",
    fontWeight: "500",
    transition: "all 0.15s ease",
    whiteSpace: "nowrap",
    height: "24px",
    display: "flex",
    alignItems: "center",
    gap: "3px",
    pointerEvents: "auto",
    "&:hover:not(:disabled)": {
      backgroundColor: isLight ? "#aceebb" : "#2d5a2f",
    },
    "&:disabled": {
      opacity: 0.5,
      cursor: "not-allowed",
    },
  },
  bulkButtonReject: {
    backgroundColor: isLight ? "#ffebe9" : "#5a1d1d",
    color: isLight ? "#cf222e" : "#f48771",
    border: isLight ? "1px solid #da3633" : "none",
    "&:hover:not(:disabled)": {
      backgroundColor: isLight ? "#ffc1cc" : "#6a2d2d",
    } as any,
  },
  thinking: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    color: isLight ? "#656d76" : "#8b949e",
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
    color: isLight ? "#656d76" : "#8b949e",
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
  checklistContainer: {
    border: isLight ? "1px solid #d0d7de" : "1px solid #30363d",
    borderRadius: "8px",
    padding: "8px 10px",
    backgroundColor: isLight ? "#ffffff" : "#0f141a",
    display: "flex",
    flexDirection: "column",
    gap: "6px",
  },
  checklistHeader: {
    fontSize: "10px",
    fontWeight: 700,
    letterSpacing: "0.6px",
    textTransform: "uppercase",
    color: isLight ? "#57606a" : "#8b949e",
  },
  checklistItem: {
    display: "flex",
    alignItems: "flex-start",
    gap: "8px",
  },
  checklistIcon: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "16px",
    height: "16px",
    marginTop: "2px",
    flexShrink: 0,
  },
  checklistPendingDot: {
    width: "8px",
    height: "8px",
    borderRadius: "50%",
    backgroundColor: isLight ? "#8c959f" : "#6e7681",
  },
  checklistText: {
    fontSize: "12px",
    lineHeight: "1.4",
    whiteSpace: "pre-wrap",
  },
  changeBlock: {
    border: isLight ? "1px solid #d0d7de" : "1px solid #30363d",
    borderRadius: "8px",
    overflow: "hidden",
    backgroundColor: isLight ? "#f6f8fa" : "#161b22",
  },
  changeHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "6px 10px",
    backgroundColor: isLight ? "#ffffff" : "#1c2128",
    borderBottom: isLight ? "1px solid #d0d7de" : "1px solid #30363d",
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
    color: isLight ? "#656d76" : "#8b949e",
    marginTop: "2px",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    maxWidth: "520px",
  },
  decisionPill: {
    padding: "2px 6px",
    borderRadius: "4px",
    fontSize: "10px",
    fontWeight: 600,
    border: isLight ? "1px solid #d0d7de" : "1px solid #30363d",
  },
  acceptedPill: {
    backgroundColor: isLight ? "#dafbe1" : "#1e4620",
    color: isLight ? "#1a7f37" : "#89d185",
    borderColor: isLight ? "#2da44e" : "#2d5a2f",
  } as any,
  rejectedPill: {
    backgroundColor: isLight ? "#ffebe9" : "#5a1d1d",
    color: isLight ? "#cf222e" : "#f48771",
    borderColor: isLight ? "#da3633" : "#6a2d2d",
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
    backgroundColor: isLight ? "#dafbe1" : "#1e4620",
    color: isLight ? "#1a7f37" : "#89d185",
  },
  deleteType: {
    backgroundColor: isLight ? "#ffebe9" : "#5a1d1d",
    color: isLight ? "#cf222e" : "#f48771",
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
    backgroundColor: isLight ? "#dafbe1" : "#1e4620",
    color: isLight ? "#1a7f37" : "#89d185",
    border: isLight ? "1px solid #2da44e" : "none",
    "&:hover:not(:disabled)": {
      backgroundColor: isLight ? "#aceebb" : "#2d5a2f",
    },
    "&:disabled": {
      opacity: 0.5,
      cursor: "not-allowed",
    },
  },
  rejectButton: {
    backgroundColor: isLight ? "#ffebe9" : "#5a1d1d",
    color: isLight ? "#cf222e" : "#f48771",
    border: isLight ? "1px solid #da3633" : "none",
    "&:hover:not(:disabled)": {
      backgroundColor: isLight ? "#ffc1cc" : "#6a2d2d",
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
    backgroundColor: isLight ? "#ffebe9" : "#5a1d1d",
    color: isLight ? "#cf222e" : "#f48771",
    textDecoration: "line-through",
    border: isLight ? "1px solid #ff8182" : "none",
  },
  diffNew: {
    backgroundColor: isLight ? "#dafbe1" : "#1e4620",
    color: isLight ? "#1a7f37" : "#89d185",
    border: isLight ? "1px solid #2da44e" : "none",
  },
  diffInsert: {
    backgroundColor: isLight ? "#dafbe1" : "#1e4620",
    color: isLight ? "#1a7f37" : "#89d185",
    border: isLight ? "1px solid #2da44e" : "none",
  },
  diffDelete: {
    backgroundColor: isLight ? "#ffebe9" : "#5a1d1d",
    color: isLight ? "#cf222e" : "#f48771",
    textDecoration: "line-through",
    border: isLight ? "1px solid #ff8182" : "none",
  },
  changeDescription: {
    fontSize: "12px",
    color: isLight ? "#656d76" : "#8b949e",
    marginTop: "4px",
  },
  changesContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    marginTop: "8px",
  },
});

// IMPORTANT: `makeStyles` returns a hook. Hooks must be called at the top level of the component,
// not inside `useMemo`/callbacks/conditionals. We create stable hooks for each theme and call both.
const useLightStyles = makeStyles(createStyles(true) as any);
const useDarkStyles = makeStyles(createStyles(false) as any);

const AgentChat: React.FC<AgentChatProps> = ({ agent }) => {
  // Detect system theme preference
  const getInitialTheme = (): boolean => {
    try {
      const stored = localStorage.getItem("amico_article_theme");
      if (stored !== null) {
        return stored === "light";
      }
    } catch {
      // In some Office/embedded webviews, localStorage can be unavailable or throw.
    }
    // Check system preference
    if (window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches) {
      return true;
    }
    return false;
  };

  const [isLightTheme, setIsLightTheme] = useState<boolean>(() => getInitialTheme());
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [mode, setMode] = useState<"edit" | "ask">("edit");
  const [modeMenuOpen, setModeMenuOpen] = useState<boolean>(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const modeMenuRef = useRef<HTMLDivElement>(null);
  const [changeTracker] = useState<ChangeTracking>(() => createChangeTracker());
  const currentMessageChangesRef = useRef<DocumentChange[]>([]);
  const [processingChanges, setProcessingChanges] = useState<Set<string>>(new Set());
  const [bulkIsProcessing, setBulkIsProcessing] = useState<boolean>(false);
  const messageIdCounter = useRef<number>(0);
  const lightStyles = useLightStyles();
  const darkStyles = useDarkStyles();
  const styles = isLightTheme ? lightStyles : darkStyles;

  const toggleTheme = () => {
    const newTheme = !isLightTheme;
    setIsLightTheme(newTheme);
    try {
      localStorage.setItem("amico_article_theme", newTheme ? "light" : "dark");
    } catch {
      // Ignore if storage is unavailable.
    }
  };

  useEffect(() => {
    if (!modeMenuOpen) {
      return () => { };
    }

    const onMouseDown = (e: MouseEvent) => {
      const el = modeMenuRef.current;
      if (!el) {
        return;
      }
      if (!el.contains(e.target as Node)) {
        setModeMenuOpen(false);
      }
    };

    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        setModeMenuOpen(false);
      }
    };

    window.addEventListener("mousedown", onMouseDown);
    window.addEventListener("keydown", onKeyDown);
    return () => {
      window.removeEventListener("mousedown", onMouseDown);
      window.removeEventListener("keydown", onKeyDown);
    };
  }, [modeMenuOpen]);

  const resizeTextarea = React.useCallback(() => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    const MAX_LINES = 12;
    const computed = window.getComputedStyle(textarea);
    const fontSize = parseFloat(computed.fontSize || "13");
    const rawLineHeight = computed.lineHeight || "";
    const lineHeight =
      rawLineHeight === "normal"
        ? fontSize * 1.5
        : /^[0-9.]+$/.test(rawLineHeight)
          ? fontSize * parseFloat(rawLineHeight)
          : parseFloat(rawLineHeight || `${fontSize * 1.5}`);

    const paddingTop = parseFloat(computed.paddingTop || "0");
    const paddingBottom = parseFloat(computed.paddingBottom || "0");
    const borderTop = parseFloat(computed.borderTopWidth || "0");
    const borderBottom = parseFloat(computed.borderBottomWidth || "0");

    // Calculate single line height and max height
    const singleLineHeight = lineHeight + paddingTop + paddingBottom + borderTop + borderBottom;
    const maxHeightPx = lineHeight * MAX_LINES + paddingTop + paddingBottom + borderTop + borderBottom;

    // Temporarily set overflow to hidden and height to auto to get accurate scrollHeight
    textarea.style.overflowY = "hidden";
    textarea.style.height = "auto";

    // Force a reflow to ensure scrollHeight is calculated
    void textarea.offsetHeight;

    const scrollHeight = textarea.scrollHeight;

    // Set height to scrollHeight, but ensure minimum is singleLineHeight and cap at maxHeightPx
    const nextHeight = Math.max(singleLineHeight, Math.min(scrollHeight, maxHeightPx));
    textarea.style.height = `${nextHeight}px`;

    // Only show scrollbar when content exceeds max height
    textarea.style.overflowY = scrollHeight > maxHeightPx ? "auto" : "hidden";
  }, []);

  const articleHeaderRegex = /ARTICLE\s+[A-Z]-\d+/i;
  const instructionKeywordRegex = /\b(?:Add|Delete|Substitute|Substitue|Replace|Insert|Remove|Change|Update|Modify|Include)\b/i;
  const instructionLineRegex = /^\s*(?:\.\d+|\d+\.|\d+\))?\s*(?:Add|Delete|Substitute|Substitue|Replace|Insert|Remove|Change|Update|Modify|Include)\b/i;

  const splitInstructions = (text: string): { steps: string[]; headerLine?: string } => {
    const normalized = text.replace(/\r\n/g, "\n");
    const lines = normalized.split("\n");
    const headerIndex = lines.findIndex(line => articleHeaderRegex.test(line));
    const headerLine = headerIndex >= 0 ? lines[headerIndex].trim() : undefined;
    const workingLines = [...lines];
    if (headerIndex >= 0) {
      workingLines[headerIndex] = "";
    }

    const findStartIndices = (regex: RegExp) => {
      const indices: number[] = [];
      for (let i = 0; i < workingLines.length; i += 1) {
        if (regex.test(workingLines[i])) {
          indices.push(i);
        }
      }
      return indices;
    };

    const startIndices = findStartIndices(instructionLineRegex);

    if (startIndices.length < 2) {
      return { steps: [text], headerLine };
    }

    const steps = startIndices
      .map((start, idx) => {
        const end = idx + 1 < startIndices.length ? startIndices[idx + 1] : workingLines.length;
        const slice = workingLines.slice(start, end);
        while (slice.length > 0 && slice[0].trim() === "") slice.shift();
        while (slice.length > 0 && slice[slice.length - 1].trim() === "") slice.pop();
        return slice.join("\n").trimEnd();
      })
      .filter(step => step.trim().length > 0);

    return { steps: steps.length > 0 ? steps : [text], headerLine };
  };

  const isArticleInstruction = (text: string): boolean => articleHeaderRegex.test(text);
  const isArticleEditInstruction = (text: string): boolean =>
    isArticleInstruction(text) && instructionKeywordRegex.test(text);

  const updateChecklistStep = (messageId: string, stepId: string, updates: Partial<ChecklistStep>) => {
    setMessages(prev =>
      prev.map(msg => {
        if (msg.messageId !== messageId || !msg.checklist) return msg;
        const nextChecklist = msg.checklist.map(step => (step.id === stepId ? { ...step, ...updates } : step));
        return { ...msg, checklist: nextChecklist };
      })
    );
  };

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

  // Initialize and auto-resize textarea
  useEffect(() => {
    resizeTextarea();
  }, [input, resizeTextarea]);

  // Initialize textarea height on mount
  useEffect(() => {
    resizeTextarea();
  }, [resizeTextarea]);

  // Handle paste events to preserve formatting
  const handlePaste = (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    e.preventDefault();
    // Some Office/Edge webviews can return undefined here; always coerce to string.
    const pastedText =
      (e.clipboardData?.getData("text/plain") ??
        // Fallback for older clipboard APIs (Office/IE legacy)
        (window as any)?.clipboardData?.getData("Text") ??
        "") + "";

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
      resizeTextarea();
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

    // Reset textarea height (will be handled by resizeTextarea when input changes)

    // Clear changes for this new message
    currentMessageChangesRef.current = [];

    // Add user message
    const userMessageId = `msg-${messageIdCounter.current++}`;
    const newMessages: Message[] = [...messages, { role: "user", content: userMessage, messageId: userMessageId }];
    setMessages(newMessages);

    try {
      const askOnlyAgent =
        mode === "ask"
          ? {
            ...agent,
            tools: {},
            system: `${(agent as any).system || ""}\n\nMODE: ASK. Do not call any tools. Do not edit the document. Answer conversationally and concisely.`,
          }
          : agent;

      const { steps, headerLine } = splitInstructions(userMessage);
      const isMultiStep = steps.length > 1;

      if (isMultiStep) {
        const checklistMessageId = `msg-${messageIdCounter.current++}`;
        const checklistSteps: ChecklistStep[] = steps.map((step, index) => ({
          id: `step-${checklistMessageId}-${index}`,
          text: step.trim(),
          status: "pending",
        }));

        setMessages([
          ...newMessages,
          {
            role: "assistant",
            content: `Plan: ${checklistSteps.length} step(s).`,
            checklist: checklistSteps,
            messageId: checklistMessageId,
          },
        ]);

        for (let i = 0; i < checklistSteps.length; i += 1) {
          const step = checklistSteps[i];
          updateChecklistStep(checklistMessageId, step.id, { status: "in_progress" });
          currentMessageChangesRef.current = [];

          const stepInstruction = headerLine && !isArticleInstruction(steps[i])
            ? `${headerLine}\n${steps[i]}`
            : steps[i];
          const shouldUseHybrid = mode === "edit" && isArticleEditInstruction(stepInstruction);

          try {
            let response: string;
            let hybridSucceeded: boolean | null = null;
            let hybridError: string | undefined;

            if (shouldUseHybrid) {
              const result = await executeArticleInstructionsHybrid(stepInstruction, agent.apiKey, agent.model);
              hybridSucceeded = result.success;
              hybridError = result.error;
              response = result.success ? "" : `Error: ${result.error || "Unknown error"}`;
            } else {
              response = await generateAgentResponse(askOnlyAgent, stepInstruction);
            }

            const stepChanges = [...currentMessageChangesRef.current];
            let stepContent = response;
            if (shouldUseHybrid) {
              if (!hybridSucceeded) {
                stepContent = `Error: ${hybridError || "Unknown error"}`;
              } else if (stepChanges.length === 0) {
                stepContent = "No changes were necessary.";
              } else {
                stepContent = `Proposed ${stepChanges.length} change(s). Review and accept/reject below.`;
              }
            } else if (!stepContent) {
              stepContent = `Step ${i + 1} complete.`;
            }

            const assistantStepMessageId = `msg-${messageIdCounter.current++}`;
            setMessages(prev => [
              ...prev,
              {
                role: "assistant",
                content: stepContent,
                changes: stepChanges.length > 0 ? stepChanges : undefined,
                messageId: assistantStepMessageId,
              },
            ]);

            updateChecklistStep(checklistMessageId, step.id, hybridSucceeded === false
              ? { status: "error", error: hybridError }
              : { status: "done" });

            currentMessageChangesRef.current = [];
          } catch (stepError) {
            const errorMessage = stepError instanceof Error ? stepError.message : "An error occurred";
            updateChecklistStep(checklistMessageId, step.id, { status: "error", error: errorMessage });
            const assistantStepMessageId = `msg-${messageIdCounter.current++}`;
            setMessages(prev => [
              ...prev,
              {
                role: "assistant",
                content: `Error: ${errorMessage}`,
                messageId: assistantStepMessageId,
              },
            ]);
            currentMessageChangesRef.current = [];
            break;
          }
        }
        return;
      }

      const hasArticleInstructions = mode === "edit" && isArticleEditInstruction(userMessage);

      let response: string;
      let hybridSucceeded: boolean | null = null;
      let hybridError: string | undefined;
      if (hasArticleInstructions) {
        // Hybrid execution: Algorithm parses/finds, AI only for final insertion (fast like Cursor)
        const result = await executeArticleInstructionsHybrid(userMessage, agent.apiKey, agent.model);
        hybridSucceeded = result.success;
        hybridError = result.error;
        // Keep the visible chat response minimal; UI will show the diffs + accept/reject.
        response = result.success ? "" : `Error: ${result.error || "Unknown error"}`;
      } else {
        // Get response from agent (changes will be tracked automatically via the agent's onChange callback)
        response = await generateAgentResponse(askOnlyAgent, userMessage);
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
              <DocumentRegular className={styles.emptyStateIcon} />
              <div className={styles.emptyStateText}>
                <strong style={{ color: isLightTheme ? "#24292f" : "#f0f6fc", fontSize: "18px", marginBottom: "12px", display: "block", fontWeight: "600" }}>
                  Amico Article
                </strong>
                <br />
                <br />
                <span style={{ color: isLightTheme ? "#656d76" : "#8b949e", fontSize: "13px" }}>
                  <div style={{ marginTop: "8px" }}>
                    • Ask me anything.<br />
                    • The more detail, the better.<br />
                    • It's faster than typing.<br />
                  </div>
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

                  {message.role === "assistant" && message.checklist && message.checklist.length > 0 && (
                    <div className={styles.checklistContainer}>
                      <div className={styles.checklistHeader}>Checklist</div>
                      {message.checklist.map((step) => (
                        <div key={step.id} className={styles.checklistItem}>
                          <div className={styles.checklistIcon}>
                            {step.status === "done" ? (
                              <CheckmarkCircleFilled style={{ fontSize: "14px", color: "#1a7f37" }} />
                            ) : step.status === "error" ? (
                              <DismissCircleFilled style={{ fontSize: "14px", color: "#cf222e" }} />
                            ) : step.status === "in_progress" ? (
                              <Spinner size="tiny" />
                            ) : (
                              <span className={styles.checklistPendingDot} />
                            )}
                          </div>
                          <div className={styles.checklistText}>{step.text}</div>
                        </div>
                      ))}
                    </div>
                  )}

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
                                    className={`${styles.decisionPill} ${decision === "accepted" ? styles.acceptedPill : styles.rejectedPill
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
              <span>Processing...</span>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>

        {error && (
          <div style={{
            padding: "12px 24px",
            backgroundColor: isLightTheme ? "#fff1f2" : "#3a1f1f",
            color: isLightTheme ? "#cf222e" : "#f48771",
            borderTop: isLightTheme ? "1px solid #ff8182" : "1px solid #5a2f2f",
            fontSize: "13px"
          }}>
            {error}
          </div>
        )}

        <div className={styles.inputContainer}>
          <div className={styles.inputRow}>
            <div className={styles.inputBox}>
              <textarea
                ref={textareaRef}
                className={styles.textarea}
                value={input}
                onChange={(e) => setInput(e.target.value)}
                onInput={() => {
                  // Resize immediately as user types
                  requestAnimationFrame(() => {
                    resizeTextarea();
                  });
                }}
                onPaste={handlePaste}
                onKeyDown={(e) => {
                  if (e.key === "Enter" && !e.shiftKey) {
                    e.preventDefault();
                    handleSend();
                  }
                }}
                placeholder={mode === "edit" ? "Describe the changes you want to make..." : "Ask a question..."}
                disabled={isLoading}
                rows={1}
                spellCheck={false}
                autoComplete="off"
                data-gramm="false"
                data-gramm_editor="false"
                data-enable-grammarly="false"
              />
              <div className={styles.buttonRow}>
                <div className={styles.buttonRowLeft}>
                  <div className={styles.modeSelectWrap} ref={modeMenuRef} title="Mode">
                    <button
                      type="button"
                      className={styles.modeSelectButton}
                      onClick={() => setModeMenuOpen((v) => !v)}
                      disabled={isLoading}
                      aria-haspopup="listbox"
                      aria-expanded={modeMenuOpen}
                      aria-label="Mode"
                    >
                      <span className={styles.modeSelectIcon}>
                        {mode === "edit" ? <EditRegular style={{ fontSize: "12px" }} /> : <ChatRegular style={{ fontSize: "12px" }} />}
                      </span>
                      <span style={{ minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                        {mode === "edit" ? "Edit" : "Ask"}
                      </span>
                      <span className={styles.modeSelectChevron}>
                        <ChevronDownRegular style={{ fontSize: "14px" }} />
                      </span>
                    </button>

                    {modeMenuOpen && (
                      <div className={styles.modeMenu} role="listbox" aria-label="Mode options">
                        <button
                          type="button"
                          className={`${styles.modeMenuItem} ${mode === "edit" ? styles.modeMenuItemActive : ""}`}
                          onClick={() => {
                            setMode("edit");
                            setModeMenuOpen(false);
                          }}
                        >
                          <EditRegular style={{ fontSize: "12px", color: isLightTheme ? "#24292f" : "#c9d1d9" }} />
                          <span>Edit</span>
                        </button>

                        <button
                          type="button"
                          className={`${styles.modeMenuItem} ${mode === "ask" ? styles.modeMenuItemActive : ""}`}
                          onClick={() => {
                            setMode("ask");
                            setModeMenuOpen(false);
                          }}
                        >
                          <ChatRegular style={{ fontSize: "12px", color: isLightTheme ? "#24292f" : "#c9d1d9" }} />
                          <span>Ask</span>
                        </button>
                      </div>
                    )}
                  </div>
                  <button
                    type="button"
                    className={styles.themeToggleButton}
                    onClick={toggleTheme}
                    title={isLightTheme ? "Switch to dark theme" : "Switch to light theme"}
                    aria-label={isLightTheme ? "Switch to dark theme" : "Switch to light theme"}
                  >
                    {isLightTheme ? (
                      <WeatherMoonRegular style={{ fontSize: "16px" }} />
                    ) : (
                      <WeatherSunnyRegular style={{ fontSize: "16px" }} />
                    )}
                  </button>
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
                        {bulkIsProcessing ? (
                          <Spinner size="tiny" />
                        ) : (
                          <>
                            <CheckmarkCircleFilled style={{ fontSize: "12px" }} />
                            Accept all ({pendingChangeCount})
                          </>
                        )}
                      </button>
                      <button
                        className={`${styles.bulkButton} ${styles.bulkButtonReject}`}
                        type="button"
                        onClick={handleRejectAll}
                        disabled={bulkIsProcessing || isLoading}
                        title="Reject all pending changes"
                      >
                        {bulkIsProcessing ? (
                          <Spinner size="tiny" />
                        ) : (
                          <>
                            <DismissCircleFilled style={{ fontSize: "12px" }} />
                            Reject all
                          </>
                        )}
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
                      <ArrowUpRegular style={{ fontSize: "12px" }} />
                    )}
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default AgentChat;
