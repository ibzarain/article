/**
 * Types for tracking document changes
 */

export interface DocumentChange {
  id: string;
  type: 'edit' | 'insert' | 'delete' | 'format';
  timestamp: Date;
  description: string;
  oldText?: string;
  newText?: string;
  searchText?: string;
  location?: string;
  /**
   * Optional scoping hints to prevent diff rendering from matching the wrong place.
   * When present, inline diff rendering + accept/reject should prefer these boundaries.
   */
  articleName?: string;
  articleStartParagraphIndex?: number;
  articleEndParagraphIndex?: number;
  /**
   * Absolute paragraph index in the document body for the intended target.
   * When present, inline diff rendering should operate on that paragraph.
   */
  targetParagraphIndex?: number;
  /**
   * End paragraph index when the change spans multiple paragraphs.
   * Used with targetParagraphIndex to define a range of paragraphs.
   */
  targetEndParagraphIndex?: number;
  formatChanges?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontSize?: number;
    fontColor?: string;
    highlightColor?: string;
  };
  applied: boolean;
  canUndo: boolean;
}

export interface ChangeTracking {
  changes: DocumentChange[];
  addChange: (change: DocumentChange) => Promise<void>;
  removeChange: (id: string) => void;
  acceptChange: (id: string) => Promise<void>;
  rejectChange: (id: string) => Promise<void>;
  acceptAll: () => Promise<void>;
  rejectAll: () => Promise<void>;
  clear: () => void;
}
