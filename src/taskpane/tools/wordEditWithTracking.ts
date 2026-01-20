/* global Word */
import { tool } from 'ai';
import { z } from 'zod';
import { DocumentChange } from '../types/changes';
import { renderInlineDiff } from '../utils/inlineDiffRenderer';

// Global change tracker - will be set by the agent wrapper
let changeTracker: ((change: DocumentChange) => Promise<void>) | null = null;

export function setChangeTracker(tracker: (change: DocumentChange) => Promise<void>) {
  changeTracker = tracker;
}

function generateChangeId(): string {
  return `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

async function trackChange(change: Omit<DocumentChange, 'id' | 'timestamp' | 'applied' | 'canUndo'>): Promise<void> {
  if (changeTracker) {
    const changeObj: DocumentChange = {
      ...change,
      id: generateChangeId(),
      timestamp: new Date(),
      applied: false, // Changes are pending until accepted
      canUndo: true,
    };
    
    // Render inline diff in the document
    await renderInlineDiff(changeObj);
    
    // Track the change
    await changeTracker(changeObj);
  }
}

/**
 * Tool to edit/replace text in the Word document (with change tracking)
 * Enhanced to preserve all formatting including font, style, and paragraph formatting
 */
export const editDocumentTool = tool({
  description: 'Edit or replace text in the Word document. Finds the specified text and replaces it with new text. Automatically preserves all formatting including font properties (bold, italic, size, color), paragraph styles, and list formatting. Assesses context to maintain formatting consistency.',
  parameters: z.object({
    searchText: z.string().describe('The text to find and replace in the document'),
    newText: z.string().describe('The new text to replace the found text with'),
    replaceAll: z.boolean().optional().default(false).describe('If true, replaces all occurrences. If false, replaces only the first occurrence.'),
    matchCase: z.boolean().optional().default(false).describe('Whether the search should be case-sensitive'),
    matchWholeWord: z.boolean().optional().default(false).describe('Whether to match whole words only'),
  }),
  execute: async ({ searchText, newText, replaceAll, matchCase, matchWholeWord }) => {
    try {
      const result = await Word.run(async (context) => {
        const searchResults = context.document.body.search(searchText, {
          matchCase: matchCase || false,
          matchWholeWord: matchWholeWord || false,
        });
        
        context.load(searchResults, 'items');
        await context.sync();
        
        if (searchResults.items.length === 0) {
          throw new Error(`Text "${searchText}" not found in document`);
        }
        
        const itemsToReplace = replaceAll ? searchResults.items : [searchResults.items[0]];
        let replacementCount = 0;
        
        for (const item of itemsToReplace) {
          // Load current text to get the actual old text (may differ from searchText)
          context.load(item, 'text');
          await context.sync();
          const actualOldText = item.text;
          
          replacementCount++;
          
          // Track the change (will render inline diff)
          await trackChange({
            type: 'edit',
            description: `Replaced "${actualOldText}" with "${newText}"`,
            oldText: actualOldText,
            newText: newText,
            searchText: searchText,
          });
        }
        
        await context.sync();
        
        return {
          replaced: replacementCount,
          totalFound: searchResults.items.length,
        };
      });
      
      return {
        success: true,
        replaced: result.replaced,
        totalFound: result.totalFound,
        message: `Replaced ${result.replaced} occurrence(s) of "${searchText}" with "${newText}"`,
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error editing document',
      };
    }
  },
});

/**
 * Tool to insert text into the Word document (with change tracking)
 * Enhanced to assess format context and insert appropriately (paragraph vs inline)
 */
export const insertTextTool = tool({
  description: 'Insert text into the Word document at a specific location. Automatically assesses format context: if inserting after a list item/bullet point, creates a new bullet point; if inserting in a paragraph, creates a new paragraph; if inserting inline in a sentence, inserts inline. Preserves all formatting including lists, styles, and paragraph formatting.',
  parameters: z.object({
    text: z.string().describe('The text to insert'),
    location: z.enum(['before', 'after', 'beginning', 'end', 'inline']).describe('Where to insert: "before" or "after" a search text, "beginning" or "end" of document, or "inline" to insert within the found text (sentence context)'),
    searchText: z.string().optional().describe('Required if location is "before", "after", or "inline". The text to search for to determine insertion point.'),
  }),
  execute: async ({ text, location, searchText }) => {
    try {
      const result = await Word.run(async (context) => {
        let insertLocation: Word.InsertLocation;
        let range: Word.Range;
        let targetParagraph: Word.Paragraph | null = null;
        let foundRange: Word.Range | null = null;
        
        if (location === 'beginning') {
          const firstParagraph = context.document.body.paragraphs.getFirst();
          range = firstParagraph.getRange('Start');
          insertLocation = Word.InsertLocation.before;
          targetParagraph = firstParagraph;
        } else if (location === 'end') {
          const lastParagraph = context.document.body.paragraphs.getLast();
          range = lastParagraph.getRange('End');
          insertLocation = Word.InsertLocation.after;
          targetParagraph = lastParagraph;
        } else if (location === 'before' || location === 'after' || location === 'inline') {
          if (!searchText) {
            throw new Error('searchText is required when location is "before", "after", or "inline"');
          }
          
          const searchResults = context.document.body.search(searchText, {
            matchCase: false,
            matchWholeWord: false,
          });
          
          context.load(searchResults, 'items');
          await context.sync();
          
          if (searchResults.items.length === 0) {
            throw new Error(`Search text "${searchText}" not found in document`);
          }
          
          foundRange = searchResults.items[0];
          // Get the paragraph containing the found text
          targetParagraph = foundRange.paragraphs.getFirst();
          context.load(targetParagraph, ['listItem', 'list', 'text', 'style']);
          
          if (location === 'inline') {
            // For inline insertion, insert at the end of the found range
            range = foundRange;
            insertLocation = Word.InsertLocation.after;
          } else if (location === 'before') {
            range = foundRange;
            insertLocation = Word.InsertLocation.before;
          } else {
            // For "after", we'll insert after the paragraph end
            range = targetParagraph.getRange('End');
            insertLocation = Word.InsertLocation.after;
          }
        } else {
          throw new Error(`Invalid location: ${location}`);
        }
        
        await context.sync();
        
        // Assess format context
        const isListItem = targetParagraph ? targetParagraph.listItem : false;
        const listObject = targetParagraph && targetParagraph.list ? targetParagraph.list : null;
        const paragraphText = targetParagraph ? targetParagraph.text : '';
        const isEndOfParagraph = foundRange && foundRange.text.trim().length > 0 
          ? paragraphText.trim().endsWith(foundRange.text.trim()) 
          : false;
        
        let insertedRange: Word.Range;
        
        if (location === 'inline') {
          // Inline insertion: insert text directly after found text, preserving formatting
          insertedRange = foundRange!.insertText(` ${text}`, Word.InsertLocation.after);
          await context.sync();
        } else if (location === 'after' && isListItem && targetParagraph) {
          // List item context: insert as new paragraph with list formatting
          const newParagraph = targetParagraph.insertParagraph(text, Word.InsertLocation.after);
          context.load(newParagraph, ['listItem']);
          await context.sync();
          
          // List formatting should be preserved automatically when inserting after a list item
          // Note: listItem is read-only, so we can't set it directly
          
          insertedRange = newParagraph.getRange();
        } else if (location === 'after' && targetParagraph) {
          // Paragraph context: insert as new paragraph, preserving paragraph style
          const newParagraph = targetParagraph.insertParagraph(text, Word.InsertLocation.after);
          context.load(newParagraph, ['style']);
          await context.sync();
          
          // Preserve paragraph style if it exists
          if (targetParagraph.style && targetParagraph.style !== 'Normal') {
            newParagraph.style = targetParagraph.style;
            await context.sync();
          }
          
          insertedRange = newParagraph.getRange();
        } else {
          // Standard text insertion for other cases
          insertedRange = range.insertText(text, insertLocation);
          await context.sync();
        }
        
        // Apply green color to inserted text immediately
        insertedRange.font.color = '#89d185';
        await context.sync();
        
        // Track the change (will render inline diff)
        await trackChange({
          type: 'insert',
          description: `Inserted "${text}" ${location}${searchText ? ` "${searchText}"` : ''}`,
          newText: text,
          searchText: searchText || location,
          location: location,
        });
        
        return {
          inserted: true,
        };
      });
      
      return {
        success: true,
        message: `Text inserted successfully at ${location}`,
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error inserting text',
      };
    }
  },
});

/**
 * Tool to delete text from the Word document (with change tracking)
 * Enhanced to assess context and handle paragraph vs inline deletion appropriately
 */
export const deleteTextTool = tool({
  description: 'Delete text from the Word document. Finds the specified text and removes it. Automatically assesses context: if deleting an entire paragraph, removes the paragraph; if deleting inline text, removes only the text while preserving paragraph structure and formatting.',
  parameters: z.object({
    searchText: z.string().describe('The text to find and delete from the document'),
    deleteAll: z.boolean().optional().default(false).describe('If true, deletes all occurrences. If false, deletes only the first occurrence.'),
    matchCase: z.boolean().optional().default(false).describe('Whether the search should be case-sensitive'),
    matchWholeWord: z.boolean().optional().default(false).describe('Whether to match whole words only'),
  }),
  execute: async ({ searchText, deleteAll, matchCase, matchWholeWord }) => {
    try {
      const result = await Word.run(async (context) => {
        const searchResults = context.document.body.search(searchText, {
          matchCase: matchCase || false,
          matchWholeWord: matchWholeWord || false,
        });
        
        context.load(searchResults, 'items');
        await context.sync();
        
        if (searchResults.items.length === 0) {
          throw new Error(`Text "${searchText}" not found in document`);
        }
        
        const itemsToDelete = deleteAll ? searchResults.items : [searchResults.items[0]];
        let deletionCount = 0;
        
        for (const item of itemsToDelete) {
          // Load the text to be deleted
          context.load(item, 'text');
          await context.sync();
          
          const deletedText = item.text;
          
          deletionCount++;
          
          // Track the change (will render inline diff - show deleted text with strikethrough)
          // Don't actually delete yet, the inline diff renderer will show it with strikethrough
          await trackChange({
            type: 'delete',
            description: `Deleted "${deletedText}"`,
            oldText: deletedText,
            searchText: searchText,
          });
        }
        
        await context.sync();
        
        return {
          deleted: deletionCount,
          totalFound: searchResults.items.length,
        };
      });
      
      return {
        success: true,
        deleted: result.deleted,
        totalFound: result.totalFound,
        message: `Deleted ${result.deleted} occurrence(s) of "${searchText}"`,
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error deleting text',
      };
    }
  },
});

/**
 * Tool to format text in the Word document (with change tracking)
 */
export const formatTextTool = tool({
  description: 'Format text in the Word document. Finds text and applies formatting like bold, italic, font size, color, etc. Use searchText: "*" or "all" to format the entire document. Use searchText: "beginning" or "end" to format text at those locations.',
  parameters: z.object({
    searchText: z.string().describe('The text to find and format. Use "*" or "all" to format the entire document. Use "beginning" or "end" for those locations.'),
    bold: z.boolean().optional().describe('Make the text bold'),
    italic: z.boolean().optional().describe('Make the text italic'),
    underline: z.boolean().optional().describe('Underline the text'),
    fontSize: z.number().optional().describe('Set the font size in points'),
    fontColor: z.string().optional().describe('Set the font color (e.g., "red", "blue", "#FF0000")'),
    highlightColor: z.string().optional().describe('Set the highlight color (e.g., "yellow", "green", "#FFFF00")'),
    formatAll: z.boolean().optional().default(false).describe('If true, formats all occurrences. If false, formats only the first occurrence.'),
  }),
  execute: async ({ searchText, bold, italic, underline, fontSize, fontColor, highlightColor, formatAll }) => {
    try {
      const result = await Word.run(async (context) => {
        let itemsToFormat: Word.Range[] = [];
        
        // Handle special cases: "*", "all", "beginning", "end"
        if (searchText === '*' || searchText.toLowerCase() === 'all') {
          // Format entire document
          const bodyRange = context.document.body.getRange('Whole');
          itemsToFormat = [bodyRange];
        } else if (searchText.toLowerCase() === 'beginning') {
          // Format from beginning
          const startRange = context.document.body.getRange('Start');
          itemsToFormat = [startRange];
        } else if (searchText.toLowerCase() === 'end') {
          // Format from end
          const endRange = context.document.body.getRange('End');
          itemsToFormat = [endRange];
        } else {
          // Normal search
          const searchResults = context.document.body.search(searchText, {
            matchCase: false,
            matchWholeWord: false,
          });
          
          context.load(searchResults, 'items');
          await context.sync();
          
          if (searchResults.items.length === 0) {
            throw new Error(`Text "${searchText}" not found in document`);
          }
          
          itemsToFormat = formatAll ? searchResults.items : [searchResults.items[0]];
        }
        
        let formattedCount = 0;
        
        const formatChanges: DocumentChange['formatChanges'] = {};
        if (bold !== undefined) formatChanges.bold = bold;
        if (italic !== undefined) formatChanges.italic = italic;
        if (underline !== undefined) formatChanges.underline = underline;
        if (fontSize !== undefined) formatChanges.fontSize = fontSize;
        if (fontColor !== undefined) formatChanges.fontColor = fontColor;
        if (highlightColor !== undefined) formatChanges.highlightColor = highlightColor;
        
        for (const item of itemsToFormat) {
          const font = item.font;
          
          if (bold !== undefined) {
            font.bold = bold;
          }
          if (italic !== undefined) {
            font.italic = italic;
          }
          if (underline !== undefined) {
            font.underline = underline ? Word.UnderlineType.single : Word.UnderlineType.none;
          }
          if (fontSize !== undefined) {
            font.size = fontSize;
          }
          if (fontColor !== undefined) {
            font.color = fontColor;
          }
          if (highlightColor !== undefined) {
            font.highlightColor = highlightColor;
          }
          
          formattedCount++;
        }
        
        await context.sync();
        
        // Track the change (formatting is applied immediately but we still track it)
        if (formattedCount > 0) {
          const description = searchText === '*' || searchText.toLowerCase() === 'all'
            ? 'Formatted entire document'
            : `Formatted "${searchText}"`;
            
          await trackChange({
            type: 'format',
            description,
            searchText: searchText,
            formatChanges: formatChanges,
          });
        }
        
        return {
          formatted: formattedCount,
          totalFound: itemsToFormat.length,
        };
      });
      
      return {
        success: true,
        formatted: result.formatted,
        totalFound: result.totalFound,
        message: `Formatted ${result.formatted} occurrence(s)`,
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error formatting text',
      };
    }
  },
});

// Re-export readDocumentTool from the original file (no tracking needed for reads)
export { readDocumentTool } from './wordEdit';
