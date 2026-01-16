import { DocumentChange, ChangeTracking } from '../types/changes';
import { acceptInlineChange, rejectInlineChange } from './inlineDiffRenderer';

/**
 * Creates a change tracking system for document edits
 */
export function createChangeTracker(): ChangeTracking {
  const changes: DocumentChange[] = [];

  const addChange = async (change: DocumentChange) => {
    changes.push(change);
  };

  const removeChange = (id: string) => {
    const index = changes.findIndex(c => c.id === id);
    if (index !== -1) {
      changes.splice(index, 1);
    }
  };

  const acceptChange = async (id: string) => {
    const change = changes.find(c => c.id === id);
    if (!change) return;

    try {
      // Use inline diff renderer to accept the change
      await acceptInlineChange(change);
      
      // Mark as applied
      change.applied = true;
      change.canUndo = true;
      
      // Optionally remove from tracking after acceptance
      // removeChange(id);
    } catch (error) {
      console.error('Error accepting change:', error);
      throw error;
    }
  };

  const rejectChange = async (id: string) => {
    const change = changes.find(c => c.id === id);
    if (!change) return;

    try {
      // Use inline diff renderer to reject the change
      await rejectInlineChange(change);
      
      // Remove the change from tracking
      removeChange(id);
    } catch (error) {
      console.error('Error rejecting change:', error);
      throw error;
    }
  };

  const acceptAll = async () => {
    for (const change of changes) {
      if (!change.applied) {
        await acceptChange(change.id);
      }
    }
  };

  const rejectAll = async () => {
    const changeIds = [...changes.map(c => c.id)];
    for (const id of changeIds) {
      await rejectChange(id);
    }
  };

  const clear = () => {
    changes.length = 0;
  };

  return {
    changes,
    addChange,
    removeChange,
    acceptChange,
    rejectChange,
    acceptAll,
    rejectAll,
    clear,
  };
}
