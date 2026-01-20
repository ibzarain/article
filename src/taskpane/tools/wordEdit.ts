/* global Word */
import { tool } from 'ai';
import { z } from 'zod';

/**
 * Tool to read text from the Word document
 */
export const readDocumentTool = tool({
  description: 'Search the Word document for a query and return contextual snippets around each match.',
  parameters: z.object({
    query: z.string().min(1).describe('Text to search for in the document.'),
    contextChars: z.number().optional().default(800).describe('Number of characters of context to include before and after each match.'),
    maxMatches: z.number().optional().describe('Optional cap on number of snippets returned (does not change totalFound).'),
    matchCase: z.boolean().optional().default(false).describe('Whether the search should be case-sensitive.'),
    matchWholeWord: z.boolean().optional().default(false).describe('Whether to match whole words only.'),
  }),
  execute: async ({ query, contextChars = 800, maxMatches, matchCase, matchWholeWord }) => {
    try {
      const result = await Word.run(async (context) => {
        const range = context.document.body.getRange('Whole');
        context.load(range, 'text');
        await context.sync();

        const text = range.text || '';
        const safeContextChars = Math.max(0, Math.floor(contextChars || 0));
        const safeMaxMatches = typeof maxMatches === 'number' && maxMatches > 0 ? Math.floor(maxMatches) : undefined;

        const escapeRegExp = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const escapedQuery = escapeRegExp(query);
        const pattern = matchWholeWord ? `\\b${escapedQuery}\\b` : escapedQuery;
        const flags = matchCase ? 'g' : 'gi';
        const regex = new RegExp(pattern, flags);

        const matches: Array<{
          matchText: string;
          snippet: string;
          matchStart: number;
          matchEnd: number;
          snippetStart: number;
          snippetEnd: number;
        }> = [];

        let totalFound = 0;
        let match: RegExpExecArray | null;

        while ((match = regex.exec(text)) !== null) {
          totalFound++;

          const matchStart = match.index;
          const matchEnd = matchStart + match[0].length;
          const snippetStart = Math.max(0, matchStart - safeContextChars);
          const snippetEnd = Math.min(text.length, matchEnd + safeContextChars);

          if (!safeMaxMatches || matches.length < safeMaxMatches) {
            matches.push({
              matchText: match[0],
              snippet: text.slice(snippetStart, snippetEnd),
              matchStart,
              matchEnd,
              snippetStart,
              snippetEnd,
            });
          }
        }

        return {
          matches,
          totalFound,
          documentLength: text.length,
        };
      });
      
      return {
        success: true,
        query,
        content: result.matches,
        totalFound: result.totalFound,
        documentLength: result.documentLength,
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
