/* global Word */
import { extractArticle, parseArticleName, ArticleBoundaries } from '../utils/articleExtractor';
import { DocumentChange } from '../types/changes';
import { generateAgentResponse } from '../agent/wordAgent';
import { createWordAgent } from '../agent/wordAgent';

// Global change tracker - will be set by the agent wrapper
let changeTracker: ((change: DocumentChange) => Promise<void>) | null = null;

export function setHybridArticleChangeTracker(tracker: (change: DocumentChange) => Promise<void>) {
  changeTracker = tracker;
}

/**
 * Creates a scoped readDocument tool that only reads content within article boundaries
 */
function createScopedReadDocumentTool(articleBoundaries: ArticleBoundaries) {
  return {
    description: 'Read text content from the Word document. This tool is scoped to only read content within the current article section.',
    parameters: {
      type: 'object',
      properties: {
        startLocation: {
          type: 'string',
          description: 'Optional: Start location to read from (e.g., "beginning", "end", or search text)',
        },
        length: {
          type: 'number',
          description: 'Optional: Number of characters to read. If not specified, reads the entire article.',
        },
      },
      required: [],
    },
    execute: async ({ startLocation, length }: { startLocation?: string; length?: number }) => {
      try {
        const result = await Word.run(async (context) => {
          // Get paragraphs for the article
          const paragraphs = context.document.body.paragraphs;
          const startParagraph = paragraphs.items[articleBoundaries.startParagraphIndex];
          const endParagraph = paragraphs.items[articleBoundaries.endParagraphIndex];
          
          const startRange = startParagraph.getRange('Start');
          const endRange = endParagraph.getRange('End');
          const articleRange = startRange.expandTo(endRange);
          
          let range: Word.Range;
          
          if (startLocation === 'beginning') {
            range = startRange;
          } else if (startLocation === 'end') {
            range = endRange;
          } else if (startLocation) {
            // Search within the article range only
            const searchResults = articleRange.search(startLocation, {
              matchCase: false,
              matchWholeWord: false,
            });
            context.load(searchResults, 'items');
            await context.sync();
            
            if (searchResults.items.length === 0) {
              throw new Error(`Location "${startLocation}" not found in article`);
            }
            range = searchResults.items[0].getRange('Start');
          } else {
            range = articleRange;
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
          error: error instanceof Error ? error.message : 'Unknown error reading article',
        };
      }
    },
  };
}

/**
 * Creates scoped editing tools that only work within article boundaries
 */
function createScopedEditTools(articleBoundaries: ArticleBoundaries) {
  return {
    editDocument: {
      description: 'Edit or replace text in the article. Automatically preserves all formatting. Only edits within the current article section.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to find and replace in the article' },
          newText: { type: 'string', description: 'The new text to replace the found text with' },
          replaceAll: { type: 'boolean', description: 'If true, replaces all occurrences. If false, replaces only the first occurrence.' },
          matchCase: { type: 'boolean', description: 'Whether the search should be case-sensitive' },
          matchWholeWord: { type: 'boolean', description: 'Whether to match whole words only' },
        },
        required: ['searchText', 'newText'],
      },
      execute: async ({ searchText, newText, replaceAll, matchCase, matchWholeWord }: any) => {
        try {
          const result = await Word.run(async (context) => {
            // Get article range
            const paragraphs = context.document.body.paragraphs;
            const startParagraph = paragraphs.items[articleBoundaries.startParagraphIndex];
            const endParagraph = paragraphs.items[articleBoundaries.endParagraphIndex];
            const startRange = startParagraph.getRange('Start');
            const endRange = endParagraph.getRange('End');
            const articleRange = startRange.expandTo(endRange);
            
            // Search only within the article range
            const searchResults = articleRange.search(searchText, {
              matchCase: matchCase || false,
              matchWholeWord: matchWholeWord || false,
            });
            
            context.load(searchResults, 'items');
            await context.sync();
            
            if (searchResults.items.length === 0) {
              throw new Error(`Text "${searchText}" not found in article`);
            }
            
            const itemsToReplace = replaceAll ? searchResults.items : [searchResults.items[0]];
            let replacementCount = 0;
            
            for (const item of itemsToReplace) {
              context.load(item, 'text');
              await context.sync();
              const actualOldText = item.text;
              
              replacementCount++;
              
              // Track the change
              if (changeTracker) {
                await changeTracker({
                  type: 'edit',
                  description: `Replaced "${actualOldText}" with "${newText}"`,
                  oldText: actualOldText,
                  newText: newText,
                  searchText: searchText,
                  id: `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                  timestamp: new Date(),
                  applied: false,
                  canUndo: true,
                });
              }
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
            error: error instanceof Error ? error.message : 'Unknown error editing article',
          };
        }
      },
    },
    insertText: {
      description: 'Insert text into the article at a specific location. Only works within the current article section.',
      parameters: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'The text to insert' },
          location: { 
            type: 'string', 
            enum: ['before', 'after', 'beginning', 'end', 'inline'],
            description: 'Where to insert: "before" or "after" a search text, "beginning" or "end" of article, or "inline" to insert within the found text',
          },
          searchText: { type: 'string', description: 'Required if location is "before", "after", or "inline". The text to search for to determine insertion point.' },
        },
        required: ['text', 'location'],
      },
      execute: async ({ text, location, searchText }: any) => {
        try {
          const result = await Word.run(async (context) => {
            // Get article range
            const paragraphs = context.document.body.paragraphs;
            const startParagraph = paragraphs.items[articleBoundaries.startParagraphIndex];
            const endParagraph = paragraphs.items[articleBoundaries.endParagraphIndex];
            const startRange = startParagraph.getRange('Start');
            const endRange = endParagraph.getRange('End');
            const articleRange = startRange.expandTo(endRange);
            
            let insertLocation: Word.InsertLocation;
            let range: Word.Range;
            
            if (location === 'beginning') {
              range = startRange;
              insertLocation = Word.InsertLocation.after;
            } else if (location === 'end') {
              range = endRange;
              insertLocation = Word.InsertLocation.before;
            } else if (location === 'before' || location === 'after' || location === 'inline') {
              if (!searchText) {
                throw new Error('searchText is required when location is "before", "after", or "inline"');
              }
              
              // Search only within article range
              const searchResults = articleRange.search(searchText, {
                matchCase: false,
                matchWholeWord: false,
              });
              
              context.load(searchResults, 'items');
              await context.sync();
              
              if (searchResults.items.length === 0) {
                throw new Error(`Search text "${searchText}" not found in article`);
              }
              
              const foundRange = searchResults.items[0];
              const targetParagraph = foundRange.paragraphs.getFirst();
              context.load(targetParagraph, ['listItem', 'list', 'text', 'style']);
              
              if (location === 'inline') {
                range = foundRange;
                insertLocation = Word.InsertLocation.after;
              } else if (location === 'before') {
                range = foundRange;
                insertLocation = Word.InsertLocation.before;
              } else {
                range = targetParagraph.getRange('End');
                insertLocation = Word.InsertLocation.after;
              }
            } else {
              throw new Error(`Invalid location: ${location}`);
            }
            
            await context.sync();
            
            // Insert the text
            const insertedRange = range.insertText(text, insertLocation);
            await context.sync();
            
            // Track the change
            if (changeTracker) {
              await changeTracker({
                type: 'insert',
                description: `Inserted "${text}" ${location}${searchText ? ` "${searchText}"` : ''}`,
                newText: text,
                searchText: searchText || location,
                location: location,
                id: `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                timestamp: new Date(),
                applied: false,
                canUndo: true,
              });
            }
            
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
    },
    deleteText: {
      description: 'Delete text from the article. Only works within the current article section.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to find and delete from the article' },
          deleteAll: { type: 'boolean', description: 'If true, deletes all occurrences. If false, deletes only the first occurrence.' },
          matchCase: { type: 'boolean', description: 'Whether the search should be case-sensitive' },
          matchWholeWord: { type: 'boolean', description: 'Whether to match whole words only' },
        },
        required: ['searchText'],
      },
      execute: async ({ searchText, deleteAll, matchCase, matchWholeWord }: any) => {
        try {
          const result = await Word.run(async (context) => {
            // Get article range
            const paragraphs = context.document.body.paragraphs;
            const startParagraph = paragraphs.items[articleBoundaries.startParagraphIndex];
            const endParagraph = paragraphs.items[articleBoundaries.endParagraphIndex];
            const startRange = startParagraph.getRange('Start');
            const endRange = endParagraph.getRange('End');
            const articleRange = startRange.expandTo(endRange);
            
            // Search only within article range
            const searchResults = articleRange.search(searchText, {
              matchCase: matchCase || false,
              matchWholeWord: matchWholeWord || false,
            });
            
            context.load(searchResults, 'items');
            await context.sync();
            
            if (searchResults.items.length === 0) {
              throw new Error(`Text "${searchText}" not found in article`);
            }
            
            const itemsToDelete = deleteAll ? searchResults.items : [searchResults.items[0]];
            let deletionCount = 0;
            
            for (const item of itemsToDelete) {
              context.load(item, 'text');
              await context.sync();
              const deletedText = item.text;
              
              deletionCount++;
              
              // Track the change
              if (changeTracker) {
                await changeTracker({
                  type: 'delete',
                  description: `Deleted "${deletedText}"`,
                  oldText: deletedText,
                  searchText: searchText,
                  id: `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                  timestamp: new Date(),
                  applied: false,
                  canUndo: true,
                });
              }
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
    },
  };
}

/**
 * Executes article instructions using hybrid approach:
 * 1. Extracts only the relevant article content
 * 2. Passes only that article to AI for processing
 * 3. AI makes edits only within that article
 */
export async function executeArticleInstructionsHybrid(
  instruction: string,
  apiKey: string,
  model: string
): Promise<{ success: boolean; error?: string; results?: string[] }> {
  try {
    // Parse article name from instruction
    const articleName = parseArticleName(instruction);
    if (!articleName) {
      return {
        success: false,
        error: 'Could not parse article name from instruction. Expected format: "ARTICLE A-1" or "A-1"',
      };
    }
    
    // Extract article boundaries
    const articleBoundaries = await extractArticle(`ARTICLE ${articleName}`);
    if (!articleBoundaries) {
      return {
        success: false,
        error: `Article "${articleName}" not found in document`,
      };
    }
    
    // Create scoped tools that only work within the article
    const scopedReadDocument = createScopedReadDocumentTool(articleBoundaries);
    const scopedEditTools = createScopedEditTools(articleBoundaries);
    
    // Create a scoped agent with only article content
    // Include the article content directly in the prompt so AI knows what it's working with
    const articleContentPreview = articleBoundaries.articleContent.length > 2000 
      ? articleBoundaries.articleContent.substring(0, 2000) + '...'
      : articleBoundaries.articleContent;
    
    const scopedAgent = {
      apiKey,
      model,
      tools: {
        readDocument: scopedReadDocument,
        ...scopedEditTools,
      },
      system: `You are a helpful AI assistant that can edit Word documents. You are currently working ONLY within ARTICLE ${articleName}.

IMPORTANT: You can ONLY read and edit content within ARTICLE ${articleName}. All your tools are scoped to this article only.

CURRENT ARTICLE CONTENT (for reference):
${articleContentPreview}

AVAILABLE TOOLS:
- readDocument: Read text content from ARTICLE ${articleName} only
- editDocument: Find and replace text within ARTICLE ${articleName} only
- insertText: Insert new text within ARTICLE ${articleName} only
- deleteText: Delete text from ARTICLE ${articleName} only

CRITICAL RULES:
1. You MUST only work within ARTICLE ${articleName} - do not attempt to edit content outside this article
2. ALWAYS use the tools to make changes - don't just describe what you would do
3. Be PRECISE with searchText - use unique, specific text that appears exactly where you want to make changes
4. When inserting text, assess format context to preserve formatting (lists, paragraphs, etc.)
5. When editing text, automatically preserves all formatting
6. Use readDocument tool if you need to see the full article content before making edits

The user has provided the following instructions for ARTICLE ${articleName}:
${instruction}

Please execute these instructions by using the available tools.`,
    };
    
    // Generate response using the scoped agent
    // The agent will only see and work with the article content
    const response = await generateAgentResponse(scopedAgent, instruction);
    
    return {
      success: true,
      results: [response],
    };
  } catch (error) {
    console.error('Error executing article instructions:', error);
    return {
      success: false,
      error: error instanceof Error ? error.message : 'Unknown error executing article instructions',
    };
  }
}
