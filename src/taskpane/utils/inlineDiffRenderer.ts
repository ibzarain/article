/* global Word */
import { DocumentChange } from '../types/changes';

async function resolveScopedSearchRoot(context: Word.RequestContext, change: DocumentChange): Promise<Word.Range | Word.Paragraph> {
  const paragraphs = context.document.body.paragraphs;
  context.load(paragraphs, 'items');
  await context.sync();

  // Strongest: explicit paragraph index
  if (typeof change.targetParagraphIndex === 'number' && change.targetParagraphIndex >= 0) {
    const p = paragraphs.items[change.targetParagraphIndex];
    if (p) {
      return p;
    }
  }

  // Next: article boundary indices
  if (
    typeof change.articleStartParagraphIndex === 'number' &&
    typeof change.articleEndParagraphIndex === 'number' &&
    change.articleStartParagraphIndex >= 0 &&
    change.articleEndParagraphIndex >= change.articleStartParagraphIndex
  ) {
    const start = paragraphs.items[change.articleStartParagraphIndex];
    const end = paragraphs.items[change.articleEndParagraphIndex];
    if (start && end) {
      const startRange = start.getRange('Start');
      const endRange = end.getRange('End');
      return startRange.expandTo(endRange);
    }
  }

  // Fallback: whole body
  return context.document.body.getRange('Whole');
}

async function scopedSearch(
  context: Word.RequestContext,
  root: Word.Range | Word.Paragraph,
  text: string
): Promise<Word.RangeCollection> {
  const range = (root as any).getRange ? (root as Word.Paragraph).getRange('Whole') : (root as Word.Range);
  const results = range.search(text, { matchCase: false, matchWholeWord: false });
  context.load(results, 'items');
  await context.sync();
  return results;
}

/**
 * Renders an inline diff in the Word document showing old text (strikethrough/red) 
 * and new text (green) with accept/undo buttons
 */
export async function renderInlineDiff(change: DocumentChange): Promise<void> {
  try {
    await Word.run(async (context) => {
      if (change.type === 'edit' && change.searchText && change.oldText && change.newText) {
        // Check if this is a multi-paragraph edit
        if (
          typeof change.targetParagraphIndex === 'number' &&
          typeof change.targetEndParagraphIndex === 'number' &&
          change.targetEndParagraphIndex >= change.targetParagraphIndex
        ) {
          // In-place replacement: keep the same bullet/number, replace content within the existing paragraph(s).
          // Do NOT insert a new paragraph — replace the content of the target range so green new + red old
          // appear in the same list item(s).
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'items');
          await context.sync();

          const firstParagraph = paragraphs.items[change.targetParagraphIndex];
          const lastParagraph = paragraphs.items[change.targetEndParagraphIndex];
          if (!firstParagraph || !lastParagraph) {
            console.warn(`Target paragraph range ${change.targetParagraphIndex}..${change.targetEndParagraphIndex} not found`);
            return;
          }

          const startRange = firstParagraph.getRange('Start');
          const endRange = lastParagraph.getRange('End');
          const targetRange = startRange.expandTo(endRange);

          // Replace content in place: new text (green) then old text (red strikethrough)
          targetRange.insertText(change.newText, Word.InsertLocation.replace);
          await context.sync();

          targetRange.font.color = '#89d185'; // Green color
          await context.sync();

          const oldRange = targetRange.insertText(`\n${change.oldText}`, Word.InsertLocation.after);
          await context.sync();

          oldRange.font.strikeThrough = true;
          oldRange.font.color = '#f48771'; // Red color
          await context.sync();
        } else {
          // Single paragraph edit or text-based search
          const root = await resolveScopedSearchRoot(context, change);
          const searchResults = await scopedSearch(context, root, change.searchText);

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

          // Replace with NEW text first (green)
          range.insertText(newTextDisplay, Word.InsertLocation.replace);
          await context.sync();

          // Apply green color to new text
          range.font.color = '#89d185'; // Green color
          await context.sync();

          // Insert old text after new text (strikethrough, red), so it starts at the same spot.
          const oldRange = range.insertText(`\n${oldTextDisplay}`, Word.InsertLocation.after);
          await context.sync();

          // Apply strikethrough and red color to old text
          oldRange.font.strikeThrough = true;
          oldRange.font.color = '#f48771'; // Red color
          await context.sync();
        }
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
        // Check if this is a multi-paragraph deletion
        if (
          typeof change.targetParagraphIndex === 'number' &&
          typeof change.targetEndParagraphIndex === 'number' &&
          change.targetEndParagraphIndex > change.targetParagraphIndex
        ) {
          // Multi-paragraph deletion: apply formatting to all paragraphs in range
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'items');
          await context.sync();

          for (let i = change.targetParagraphIndex; i <= change.targetEndParagraphIndex; i++) {
            const p = paragraphs.items[i];
            if (p) {
              const pRange = p.getRange('Whole');
              pRange.font.strikeThrough = true;
              pRange.font.color = '#f48771';
            }
          }
          await context.sync();
        } else {
          // Single paragraph or text-based search
          const root = await resolveScopedSearchRoot(context, change);
          const searchResults = await scopedSearch(context, root, change.searchText || change.oldText);

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
        // Check if this is a multi-paragraph edit (in-place: same bullet, content replaced)
        if (
          typeof change.targetParagraphIndex === 'number' &&
          typeof change.targetEndParagraphIndex === 'number' &&
          change.targetEndParagraphIndex >= change.targetParagraphIndex
        ) {
          // In-place edit: target range holds newText (green) + "\n" + oldText (red). Accept = keep new, remove red.
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'items');
          await context.sync();

          const firstP = paragraphs.items[change.targetParagraphIndex];
          const lastP = paragraphs.items[change.targetEndParagraphIndex];
          if (!firstP || !lastP) return;

          const root = firstP.getRange('Start').expandTo(lastP.getRange('End'));
          const newTextResults = await scopedSearch(context, root, change.newText);
          const oldTextResults = await scopedSearch(context, root, change.oldText);

          for (const range of newTextResults.items) {
            context.load(range, ['text', 'font']);
            await context.sync();
            if (range.text.trim() === change.newText!.trim() && range.font.color === '#89d185') {
              range.font.color = null;
              await context.sync();
              break;
            }
          }

          for (const range of oldTextResults.items) {
            context.load(range, ['text', 'font']);
            await context.sync();
            if (range.text.trim() === change.oldText!.trim() && range.font.strikeThrough && range.font.color === '#f48771') {
              range.delete();
              await context.sync();
              break;
            }
          }
        } else {
          const root = await resolveScopedSearchRoot(context, change);
          // Find the new text (green highlighted) and old text (strikethrough red)
          // Search for new text with green highlighting
          const newTextResults = await scopedSearch(context, root, change.newText);

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
          const oldTextResults = await scopedSearch(context, root, change.oldText);

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
        }
      } else if (change.type === 'insert' && change.newText) {
        const root = await resolveScopedSearchRoot(context, change);
        // Find the inserted text (green highlighted) and remove highlighting
        const searchResults = await scopedSearch(context, root, change.newText);

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
        // Check if this is a multi-paragraph deletion
        if (
          typeof change.targetParagraphIndex === 'number' &&
          typeof change.targetEndParagraphIndex === 'number' &&
          change.targetEndParagraphIndex > change.targetParagraphIndex
        ) {
          // Multi-paragraph deletion: delete all marked paragraphs
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'items');
          await context.sync();

          // Delete from end to start to avoid index shifting issues
          for (let i = change.targetEndParagraphIndex; i >= change.targetParagraphIndex; i--) {
            const p = paragraphs.items[i];
            if (p) {
              context.load(p, 'font');
              await context.sync();

              // Only delete if it's marked as deletion (red strikethrough)
              if (p.font.strikeThrough && p.font.color === '#f48771') {
                p.delete();
              }
            }
          }
          await context.sync();
        } else {
          const root = await resolveScopedSearchRoot(context, change);
          // Find and remove the deleted text (strikethrough red)
          const searchResults = await scopedSearch(context, root, change.oldText);

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
        // Check if this is a multi-paragraph edit (in-place)
        if (
          typeof change.targetParagraphIndex === 'number' &&
          typeof change.targetEndParagraphIndex === 'number' &&
          change.targetEndParagraphIndex >= change.targetParagraphIndex
        ) {
          // In-place edit rejection: delete green new text, restore old (remove strikethrough/red)
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'items');
          await context.sync();

          const firstP = paragraphs.items[change.targetParagraphIndex];
          const lastP = paragraphs.items[change.targetEndParagraphIndex];
          if (!firstP || !lastP) return;

          const root = firstP.getRange('Start').expandTo(lastP.getRange('End'));
          const newTextResults = await scopedSearch(context, root, change.newText);
          const oldTextResults = await scopedSearch(context, root, change.oldText);

          for (const range of newTextResults.items) {
            context.load(range, ['text', 'font']);
            await context.sync();
            if (range.text.trim() === change.newText!.trim() && range.font.color === '#89d185') {
              range.delete();
              await context.sync();
              break;
            }
          }

          for (const range of oldTextResults.items) {
            context.load(range, ['text', 'font']);
            await context.sync();
            if (range.text.trim() === change.oldText!.trim() && range.font.strikeThrough && range.font.color === '#f48771') {
              range.font.strikeThrough = false;
              range.font.color = null;
              await context.sync();
              break;
            }
          }
        } else {
          const root = await resolveScopedSearchRoot(context, change);
          // Remove new text (green highlighted), restore old text (remove strikethrough)
          // First, find and remove the new text
          const newTextResults = await scopedSearch(context, root, change.newText);

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
          const oldTextResults = await scopedSearch(context, root, change.oldText);

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
        }
      } else if (change.type === 'insert' && change.newText) {
        const root = await resolveScopedSearchRoot(context, change);
        // Remove the inserted text
        const searchResults = await scopedSearch(context, root, change.newText);

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
        // Check if this is a multi-paragraph deletion
        if (
          typeof change.targetParagraphIndex === 'number' &&
          typeof change.targetEndParagraphIndex === 'number' &&
          change.targetEndParagraphIndex > change.targetParagraphIndex
        ) {
          // Multi-paragraph deletion: restore all marked paragraphs
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'items');
          await context.sync();

          for (let i = change.targetParagraphIndex; i <= change.targetEndParagraphIndex; i++) {
            const p = paragraphs.items[i];
            if (p) {
              context.load(p, 'font');
              await context.sync();

              // Restore if it's marked as deletion (red strikethrough)
              if (p.font.strikeThrough && p.font.color === '#f48771') {
                p.font.strikeThrough = false;
                p.font.color = null;
              }
            }
          }
          await context.sync();
        } else {
          const root = await resolveScopedSearchRoot(context, change);
          // Restore the deleted text (remove strikethrough and red)
          const searchResults = await scopedSearch(context, root, change.oldText);

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
      }
    });
  } catch (error) {
    console.error('Error rejecting inline change:', error);
    throw error;
  }
}
