/* global Word */
import { DocumentChange } from '../types/changes';

/**
 * Renders an inline diff in the Word document showing old text (strikethrough/red) 
 * and new text (green) with accept/undo buttons
 */
export async function renderInlineDiff(change: DocumentChange): Promise<void> {
  try {
    await Word.run(async (context) => {
      if (change.type === 'edit' && change.searchText && change.oldText && change.newText) {
        // Find the text to replace
        const searchResults = context.document.body.search(change.searchText, {
          matchCase: false,
          matchWholeWord: false,
        });
        
        context.load(searchResults, 'items');
        await context.sync();
        
        if (searchResults.items.length === 0) {
          console.warn(`Text "${change.searchText}" not found for inline diff`);
          return;
        }
        
        const range = searchResults.items[0];
        
        // Store original formatting
        context.load(range.font, ['bold', 'italic', 'underline', 'size', 'color', 'highlightColor']);
        context.load(range, 'text');
        await context.sync();
        
        // Create the diff display: old text (strikethrough, red) + new text (green)
        const oldTextDisplay = change.oldText;
        const newTextDisplay = change.newText;
        
        // Replace with old text first (strikethrough, red)
        range.insertText(oldTextDisplay, Word.InsertLocation.replace);
        await context.sync();
        
        // Apply strikethrough and red color to old text
        range.font.strikeThrough = true;
        range.font.color = '#f48771'; // Red color
        await context.sync();
        
        // Insert new text after old text with green highlighting
        const newRange = range.insertText(` ${newTextDisplay}`, Word.InsertLocation.after);
        await context.sync();
        
        // Apply green color to new text
        newRange.font.color = '#89d185'; // Green color
        await context.sync();
        
      } else if (change.type === 'insert' && change.newText) {
        // For insertions, find the already-inserted text and apply green highlighting
        // The text was already inserted by the tool, we just need to highlight it
        const searchResults = context.document.body.search(change.newText, {
          matchCase: false,
          matchWholeWord: false,
        });
        
        context.load(searchResults, 'items');
        await context.sync();
        
        if (searchResults.items.length === 0) {
          console.warn(`Inserted text "${change.newText}" not found for highlighting`);
          return;
        }
        
        // Find the most recently inserted text (should be the last match or one near searchText)
        let insertRange: Word.Range | null = null;
        
        if (change.searchText && change.searchText !== change.location) {
          // Try to find text near the searchText location
          const locationResults = context.document.body.search(change.searchText, {
            matchCase: false,
            matchWholeWord: false,
          });
          context.load(locationResults, 'items');
          await context.sync();
          
          if (locationResults.items.length > 0) {
            const locationRange = locationResults.items[0];
            // Find the inserted text closest to this location
            for (const range of searchResults.items) {
              context.load(range, 'text');
              await context.sync();
              if (range.text.trim() === change.newText.trim()) {
                insertRange = range;
                break;
              }
            }
          }
        }
        
        // If we didn't find it near searchText, use the first match
        if (!insertRange && searchResults.items.length > 0) {
          insertRange = searchResults.items[0];
        }
        
        if (insertRange) {
          // Check if already green, if not apply green color
          context.load(insertRange.font, 'color');
          await context.sync();
          if (insertRange.font.color !== '#89d185') {
            insertRange.font.color = '#89d185';
            await context.sync();
          }
        }
        
      } else if (change.type === 'delete' && change.oldText) {
        // For deletions, show old text with strikethrough and red
        const searchResults = context.document.body.search(change.searchText || change.oldText, {
          matchCase: false,
          matchWholeWord: false,
        });
        
        context.load(searchResults, 'items');
        await context.sync();
        
        if (searchResults.items.length === 0) {
          console.warn(`Text not found for deletion diff`);
          return;
        }
        
        const range = searchResults.items[0];
        context.load(range, 'text');
        await context.sync();
        
        // Apply strikethrough and red color
        range.font.strikeThrough = true;
        range.font.color = '#f48771';
        await context.sync();
      }
    });
  } catch (error) {
    // Best-effort: inline diff rendering should never break the primary operation
    // (insert/edit/delete). Office.js can throw opaque errors (e.g. "We couldn't find
    // the item you requested") depending on range invalidation / timing.
    console.error('Error rendering inline diff (best-effort):', {
      changeId: change.id,
      changeType: change.type,
      error,
    });
    return;
  }
}

/**
 * Accepts a change by removing the diff markers and keeping only the new text
 */
export async function acceptInlineChange(change: DocumentChange): Promise<void> {
  try {
    await Word.run(async (context) => {
      if (change.type === 'edit' && change.newText && change.oldText) {
        // Find the new text (green highlighted) and old text (strikethrough red)
        // Search for new text with green highlighting
        const newTextResults = context.document.body.search(change.newText, {
          matchCase: false,
          matchWholeWord: false,
        });
        context.load(newTextResults, 'items');
        await context.sync();
        
        // Find and accept the new text (remove green highlighting)
        for (const range of newTextResults.items) {
          context.load(range, ['text', 'font']);
          await context.sync();
          
          // Check if this is our highlighted new text (green)
          if (range.text.trim() === change.newText.trim() && range.font.color === '#89d185') {
            // Remove green color, keep text with default color
            range.font.color = null;
            await context.sync();
            break; // Only accept the first match
          }
        }
        
        // Find and remove old text (strikethrough red)
        const oldTextResults = context.document.body.search(change.oldText, {
          matchCase: false,
          matchWholeWord: false,
        });
        context.load(oldTextResults, 'items');
        await context.sync();
        
        for (const range of oldTextResults.items) {
          context.load(range, ['text', 'font']);
          await context.sync();
          
          // Check if this is our strikethrough old text (red)
          if (range.text.trim() === change.oldText.trim() && range.font.strikeThrough && range.font.color === '#f48771') {
            // Delete the old text (removes red strikethrough)
            range.delete();
            await context.sync();
            break; // Only remove the first match
          }
        }
        
      } else if (change.type === 'insert' && change.newText) {
        // Find the inserted text (green highlighted) and remove highlighting
        const searchResults = context.document.body.search(change.newText, {
          matchCase: false,
          matchWholeWord: false,
        });
        context.load(searchResults, 'items');
        await context.sync();
        
        for (const range of searchResults.items) {
          context.load(range, ['text', 'font']);
          await context.sync();
          
          // Check if this is our highlighted inserted text (green)
          if (range.text.trim() === change.newText.trim() && range.font.color === '#89d185') {
            // Remove green color, keep text with default color
            range.font.color = null;
            await context.sync();
            break;
          }
        }
        
      } else if (change.type === 'delete' && change.oldText) {
        // Find and remove the deleted text (strikethrough red)
        const searchResults = context.document.body.search(change.oldText, {
          matchCase: false,
          matchWholeWord: false,
        });
        context.load(searchResults, 'items');
        await context.sync();
        
        for (const range of searchResults.items) {
          context.load(range, ['text', 'font']);
          await context.sync();
          
          // Check if this is our strikethrough deleted text (red)
          if (range.text.trim() === change.oldText.trim() && range.font.strikeThrough && range.font.color === '#f48771') {
            // Delete the text (removes red strikethrough)
            range.delete();
            await context.sync();
            break;
          }
        }
      }
    });
  } catch (error) {
    console.error('Error accepting inline change:', error);
    throw error;
  }
}

/**
 * Rejects a change by removing the diff and reverting to original state
 */
export async function rejectInlineChange(change: DocumentChange): Promise<void> {
  try {
    await Word.run(async (context) => {
      if (change.type === 'edit' && change.oldText && change.newText) {
        // Remove new text (green highlighted), restore old text (remove strikethrough)
        // First, find and remove the new text
        const newTextResults = context.document.body.search(change.newText, {
          matchCase: false,
          matchWholeWord: false,
        });
        context.load(newTextResults, 'items');
        await context.sync();
        
        for (const range of newTextResults.items) {
          context.load(range, ['text', 'font']);
          await context.sync();
          
          // Check if this is our highlighted new text (green)
          if (range.text.trim() === change.newText.trim() && range.font.color === '#89d185') {
            // Remove green text
            range.delete();
            await context.sync();
            break;
          }
        }
        
        // Then, restore the old text (remove strikethrough and red)
        const oldTextResults = context.document.body.search(change.oldText, {
          matchCase: false,
          matchWholeWord: false,
        });
        context.load(oldTextResults, 'items');
        await context.sync();
        
        for (const range of oldTextResults.items) {
          context.load(range, ['text', 'font']);
          await context.sync();
          
          // Check if this is our strikethrough old text (red)
          if (range.text.trim() === change.oldText.trim() && range.font.strikeThrough && range.font.color === '#f48771') {
            // Remove strikethrough and red color, revert to default text color
            range.font.strikeThrough = false;
            range.font.color = null;
            await context.sync();
            break;
          }
        }
        
      } else if (change.type === 'insert' && change.newText) {
        // Remove the inserted text
        const searchResults = context.document.body.search(change.newText, {
          matchCase: false,
          matchWholeWord: false,
        });
        context.load(searchResults, 'items');
        await context.sync();
        
        for (const range of searchResults.items) {
          context.load(range, ['text', 'font']);
          await context.sync();
          
          // Check if this is our highlighted inserted text (green)
          if (range.text.trim() === change.newText.trim() && range.font.color === '#89d185') {
            // Remove green text
            range.delete();
            await context.sync();
            break;
          }
        }
        
      } else if (change.type === 'delete' && change.oldText) {
        // Restore the deleted text (remove strikethrough and red)
        const searchResults = context.document.body.search(change.oldText, {
          matchCase: false,
          matchWholeWord: false,
        });
        context.load(searchResults, 'items');
        await context.sync();
        
        for (const range of searchResults.items) {
          context.load(range, ['text', 'font']);
          await context.sync();
          
          // Check if this is our strikethrough deleted text (red)
          if (range.text.trim() === change.oldText.trim() && range.font.strikeThrough && range.font.color === '#f48771') {
            // Remove strikethrough and red color, revert to default text color
            range.font.strikeThrough = false;
            range.font.color = null;
            await context.sync();
            break;
          }
        }
      }
    });
  } catch (error) {
    console.error('Error rejecting inline change:', error);
    throw error;
  }
}
