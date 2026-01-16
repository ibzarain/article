import { DocumentChange, ChangeTracking } from '../types/changes';

/**
 * Creates a change tracking system for document edits
 */
export function createChangeTracker(): ChangeTracking {
  const changes: DocumentChange[] = [];

  const addChange = (change: DocumentChange) => {
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
    if (change && !change.applied) {
      // Mark as applied - the change is already in the document
      change.applied = true;
      change.canUndo = true;
    }
  };

  const rejectChange = async (id: string) => {
    const change = changes.find(c => c.id === id);
    if (!change) return;

    try {
      // Undo the change based on its type
      await Word.run(async (context) => {
        if (change.type === 'edit' && change.searchText && change.oldText) {
          // Revert edit: replace newText back with oldText
          const searchResults = context.document.body.search(change.newText || '', {
            matchCase: false,
            matchWholeWord: false,
          });
          context.load(searchResults, 'items');
          await context.sync();
          
          if (searchResults.items.length > 0) {
            // Find the exact match (simplified - in production you'd want better matching)
            searchResults.items[0].insertText(change.oldText, Word.InsertLocation.replace);
            await context.sync();
          }
        } else if (change.type === 'insert' && change.searchText) {
          // Remove inserted text
          const searchResults = context.document.body.search(change.newText || '', {
            matchCase: false,
            matchWholeWord: false,
          });
          context.load(searchResults, 'items');
          await context.sync();
          
          if (searchResults.items.length > 0) {
            searchResults.items[0].delete();
            await context.sync();
          }
        } else if (change.type === 'delete' && change.oldText) {
          // Restore deleted text
          const searchResults = context.document.body.search(change.searchText || '', {
            matchCase: false,
            matchWholeWord: false,
          });
          context.load(searchResults, 'items');
          await context.sync();
          
          if (searchResults.items.length > 0) {
            searchResults.items[0].insertText(change.oldText, Word.InsertLocation.after);
            await context.sync();
          }
        } else if (change.type === 'format' && change.searchText) {
          // Revert formatting (simplified - would need to track original format)
          const searchResults = context.document.body.search(change.searchText, {
            matchCase: false,
            matchWholeWord: false,
          });
          context.load(searchResults, 'items');
          context.load(searchResults, 'font');
          await context.sync();
          
          if (searchResults.items.length > 0 && change.formatChanges) {
            const font = searchResults.items[0].font;
            if (change.formatChanges.bold !== undefined) {
              font.bold = !change.formatChanges.bold;
            }
            if (change.formatChanges.italic !== undefined) {
              font.italic = !change.formatChanges.italic;
            }
            if (change.formatChanges.underline !== undefined) {
              font.underline = change.formatChanges.underline ? 'none' : 'single';
            }
            await context.sync();
          }
        }
      });

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
