/* global Word */
import { tool } from 'ai';
import { z } from 'zod';
import { DocumentChange } from '../types/changes';

// Global change tracker - will be set by the agent wrapper
let changeTracker: ((change: DocumentChange) => void) | null = null;

export function setChangeTracker(tracker: (change: DocumentChange) => void) {
  changeTracker = tracker;
}

function generateChangeId(): string {
  return `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

function trackChange(change: Omit<DocumentChange, 'id' | 'timestamp' | 'applied' | 'canUndo'>): void {
  if (changeTracker) {
    changeTracker({
      ...change,
      id: generateChangeId(),
      timestamp: new Date(),
      applied: true, // Changes are applied immediately
      canUndo: true,
    });
  }
}

/**
 * Tool to edit/replace text in the Word document (with change tracking)
 */
export const editDocumentTool = tool({
  description: 'Edit or replace text in the Word document. Finds the specified text and replaces it with new text. Preserves formatting of the surrounding text.',
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
          item.insertText(newText, Word.InsertLocation.replace);
          replacementCount++;
          
          // Track the change
          trackChange({
            type: 'edit',
            description: `Replaced "${searchText}" with "${newText}"`,
            oldText: searchText,
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
 */
export const insertTextTool = tool({
  description: 'Insert text into the Word document at a specific location. Can insert before or after found text, or at the beginning/end of the document.',
  parameters: z.object({
    text: z.string().describe('The text to insert'),
    location: z.enum(['before', 'after', 'beginning', 'end']).describe('Where to insert the text: "before" or "after" a search text, or at "beginning" or "end" of document'),
    searchText: z.string().optional().describe('Required if location is "before" or "after". The text to search for to determine insertion point.'),
  }),
  execute: async ({ text, location, searchText }) => {
    try {
      const result = await Word.run(async (context) => {
        let insertLocation: Word.InsertLocation;
        let range: Word.Range;
        
        if (location === 'beginning') {
          range = context.document.body.getRange('Start');
          insertLocation = Word.InsertLocation.before;
        } else if (location === 'end') {
          range = context.document.body.getRange('End');
          insertLocation = Word.InsertLocation.after;
        } else if (location === 'before' || location === 'after') {
          if (!searchText) {
            throw new Error('searchText is required when location is "before" or "after"');
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
          
          range = searchResults.items[0];
          insertLocation = location === 'before' ? Word.InsertLocation.before : Word.InsertLocation.after;
        } else {
          throw new Error(`Invalid location: ${location}`);
        }
        
        range.insertText(text, insertLocation);
        await context.sync();
        
        // Track the change
        trackChange({
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
 */
export const deleteTextTool = tool({
  description: 'Delete text from the Word document. Finds the specified text and removes it.',
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
          // Get the text before deleting
          context.load(item, 'text');
          await context.sync();
          const deletedText = item.text;
          
          item.delete();
          deletionCount++;
          
          // Track the change
          trackChange({
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
  description: 'Format text in the Word document. Finds text and applies formatting like bold, italic, font size, color, etc.',
  parameters: z.object({
    searchText: z.string().describe('The text to find and format'),
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
        const searchResults = context.document.body.search(searchText, {
          matchCase: false,
          matchWholeWord: false,
        });
        
        context.load(searchResults, 'items');
        await context.sync();
        
        if (searchResults.items.length === 0) {
          throw new Error(`Text "${searchText}" not found in document`);
        }
        
        const itemsToFormat = formatAll ? searchResults.items : [searchResults.items[0]];
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
            font.underline = underline ? 'single' : 'none';
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
        
        // Track the change
        if (formattedCount > 0) {
          trackChange({
            type: 'format',
            description: `Formatted "${searchText}"`,
            searchText: searchText,
            formatChanges: formatChanges,
          });
        }
        
        return {
          formatted: formattedCount,
          totalFound: searchResults.items.length,
        };
      });
      
      return {
        success: true,
        formatted: result.formatted,
        totalFound: result.totalFound,
        message: `Formatted ${result.formatted} occurrence(s) of "${searchText}"`,
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
