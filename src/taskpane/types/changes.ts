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
  addChange: (change: DocumentChange) => void;
  removeChange: (id: string) => void;
  acceptChange: (id: string) => Promise<void>;
  rejectChange: (id: string) => Promise<void>;
  acceptAll: () => Promise<void>;
  rejectAll: () => Promise<void>;
  clear: () => void;
}
