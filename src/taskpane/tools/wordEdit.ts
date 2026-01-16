/* global Word */
import { tool } from 'ai';
import { z } from 'zod';

/**
 * Tool to read text from the Word document
 */
export const readDocumentTool = tool({
  description: 'Read text content from the Word document. Use this to see what is currently in the document before making edits.',
  parameters: z.object({
    startLocation: z.string().optional().describe('Optional: Start location to read from (e.g., "beginning", "end", or search text)'),
    length: z.number().optional().describe('Optional: Number of characters to read. If not specified, reads the entire document.'),
  }),
  execute: async ({ startLocation, length }) => {
    try {
      const result = await Word.run(async (context) => {
        let range: Word.Range;
        
        if (startLocation === 'beginning') {
          range = context.document.body.getRange('Start');
        } else if (startLocation === 'end') {
          range = context.document.body.getRange('End');
        } else if (startLocation) {
          // Search for the location text
          const searchResults = context.document.body.search(startLocation, {
            matchCase: false,
            matchWholeWord: false,
          });
          context.load(searchResults, 'items');
          await context.sync();
          
          if (searchResults.items.length === 0) {
            throw new Error(`Location "${startLocation}" not found in document`);
          }
          range = searchResults.items[0].getRange('Start');
        } else {
          range = context.document.body;
        }
        
        context.load(range, 'text');
        await context.sync();
        
        let text = range.text;
        if (length && length > 0) {
          text = text.substring(0, length);
        }
        
        return {
          text,
          length: text.length,
        };
      });
      
      return {
        success: true,
        content: result.text,
        length: result.length,
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error reading document',
      };
    }
  },
});

/**
 * Tool to edit/replace text in the Word document
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
 * Tool to insert text into the Word document
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
 * Tool to delete text from the Word document
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
          item.delete();
          deletionCount++;
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
 * Tool to format text in the Word document
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
