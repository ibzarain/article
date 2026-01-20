/* global Word */
import { extractArticle, parseArticleName, ArticleBoundaries } from '../utils/articleExtractor';
import { DocumentChange } from '../types/changes';
import { generateAgentResponse } from '../agent/wordAgent';
import { createWordAgent } from '../agent/wordAgent';
import { renderInlineDiff } from '../utils/inlineDiffRenderer';

// Global change tracker - will be set by the agent wrapper
let changeTracker: ((change: DocumentChange) => Promise<void>) | null = null;

export function setHybridArticleChangeTracker(tracker: (change: DocumentChange) => Promise<void>) {
  changeTracker = tracker;
}

/**
 * Creates a scoped readDocument tool that only reads content within article boundaries
 * This is a proper search tool that returns matches with context snippets
 */
function createScopedReadDocumentTool(articleBoundaries: ArticleBoundaries) {
  return {
    description: 'Search ARTICLE content for a query and return contextual snippets around each match. This is a SEARCH tool - use it to find exact text before editing. Returns matches with snippets showing context.',
    parameters: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Text to search for in the article. This is a search query - use it to find exact text before making edits.',
        },
        contextChars: {
          type: 'number',
          description: 'Number of characters of context to include before and after each match. Default: 800',
        },
        maxMatches: {
          type: 'number',
          description: 'Optional cap on number of snippets returned',
        },
        matchCase: {
          type: 'boolean',
          description: 'Whether the search should be case-sensitive. Default: false',
        },
        matchWholeWord: {
          type: 'boolean',
          description: 'Whether to match whole words only. Default: false',
        },
      },
      required: ['query'],
    },
    execute: async ({ query, contextChars = 800, maxMatches, matchCase = false, matchWholeWord = false }: { 
      query: string; 
      contextChars?: number; 
      maxMatches?: number; 
      matchCase?: boolean; 
      matchWholeWord?: boolean;
    }) => {
      try {
        const result = await Word.run(async (context) => {
          // Get article range
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'items');
          await context.sync();
          const startParagraph = paragraphs.items[articleBoundaries.startParagraphIndex];
          const endParagraph = paragraphs.items[articleBoundaries.endParagraphIndex];
          
          const startRange = startParagraph.getRange('Start');
          const endRange = endParagraph.getRange('End');
          const articleRange = startRange.expandTo(endRange);
          
          // Get article text for regex search
          context.load(articleRange, 'text');
          await context.sync();
          
          const text = articleRange.text || '';
          const safeContextChars = Math.max(0, Math.floor(contextChars || 0));
          const safeMaxMatches = typeof maxMatches === 'number' && maxMatches > 0 ? Math.floor(maxMatches) : undefined;

          // Escape regex special characters
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
            articleLength: text.length,
          };
        });
        
        return {
          success: true,
          query,
          content: result.matches,
          totalFound: result.totalFound,
          articleLength: result.articleLength,
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
function createScopedEditTools(articleBoundaries: ArticleBoundaries, articleName: string) {
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
            context.load(paragraphs, 'items');
            await context.sync();
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
              
              // Track the change (renderInlineDiff will handle the visual diff)
              if (changeTracker) {
                const changeObj: DocumentChange = {
                  type: 'edit',
                  description: `Replaced "${actualOldText}" with "${newText}"`,
                  oldText: actualOldText,
                  newText: newText,
                  searchText: searchText,
                  id: `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                  timestamp: new Date(),
                  applied: false,
                  canUndo: true,
                };
                
                // Render inline diff (this will show old text with strikethrough/red and new text in green)
                await renderInlineDiff(changeObj);
                
                // Track the change
                await changeTracker(changeObj);
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
      description: 'Insert text into the article at a specific location. Only works within the current article section. Use searchText from readDocument results to identify where to insert. The searchText should be unique text near the insertion point - it can be the exact matchText from readDocument or nearby unique text from the snippets.',
      parameters: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'The text to insert' },
          location: { 
            type: 'string', 
            enum: ['before', 'after', 'beginning', 'end', 'inline'],
            description: 'Where to insert: "before" or "after" a search text, "beginning" or "end" of article, or "inline" to insert within the found text',
          },
          searchText: { type: 'string', description: 'Required if location is "before", "after", or "inline". Use text from readDocument results that uniquely identifies the insertion point. Can be the matchText or nearby unique text from the snippets.' },
        },
        required: ['text', 'location'],
      },
      execute: async ({ text, location, searchText }: any) => {
        try {
          const result = await Word.run(async (context) => {
            // Get article range
            const paragraphs = context.document.body.paragraphs;
            context.load(paragraphs, 'items');
            await context.sync();
            const startParagraph = paragraphs.items[articleBoundaries.startParagraphIndex];
            const endParagraph = paragraphs.items[articleBoundaries.endParagraphIndex];
            const startRange = startParagraph.getRange('Start');
            const endRange = endParagraph.getRange('End');
            const articleRange = startRange.expandTo(endRange);
            
            let insertLocation: Word.InsertLocation;
            let range: Word.Range;
            let targetParagraph: Word.Paragraph | null = null;
            
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
              // Try multiple search strategies to find the text
              let searchResults = articleRange.search(searchText, {
                matchCase: false,
                matchWholeWord: false,
              });
              
              context.load(searchResults, 'items');
              await context.sync();
              
              // If not found, try with different whitespace handling
              if (searchResults.items.length === 0) {
                // Try trimming the search text
                const trimmedSearch = searchText.trim();
                if (trimmedSearch !== searchText) {
                  searchResults = articleRange.search(trimmedSearch, {
                    matchCase: false,
                    matchWholeWord: false,
                  });
                  context.load(searchResults, 'items');
                  await context.sync();
                }
              }
              
              // If still not found, try without punctuation at the end
              if (searchResults.items.length === 0 && searchText && /[.:;]/.test(searchText)) {
                const withoutPunct = searchText.replace(/[.:;]+$/, '').trim();
                if (withoutPunct && withoutPunct !== searchText) {
                  searchResults = articleRange.search(withoutPunct, {
                    matchCase: false,
                    matchWholeWord: false,
                  });
                  context.load(searchResults, 'items');
                  await context.sync();
                }
              }
              
              // Also try with punctuation added if original didn't have it
              if (searchResults.items.length === 0 && searchText && !/[.:;]$/.test(searchText)) {
                const withPunct = searchText + ':';
                searchResults = articleRange.search(withPunct, {
                  matchCase: false,
                  matchWholeWord: false,
                });
                context.load(searchResults, 'items');
                await context.sync();
              }
              
              if (searchResults.items.length === 0) {
                // Get article content snippet for better error message
                context.load(articleRange, 'text');
                await context.sync();
                const articlePreview = articleRange.text.substring(0, 500);
                throw new Error(`Search text "${searchText}" not found in ARTICLE ${articleName}. Searched within article content. Please use readDocument first to find the exact text. Article preview: ${articlePreview}...`);
              }
              
              // Log the search for debugging
              console.log(`[insertText] Searching for: "${searchText}", found ${searchResults.items.length} match(es) in ARTICLE ${articleName}`);
              
              // Use the first match (most relevant)
              const foundRange = searchResults.items[0];
              targetParagraph = foundRange.paragraphs.getFirst();
              context.load(targetParagraph, ['listItem', 'list', 'text', 'style']);
              
              // Get paragraph text to check context
              context.load(targetParagraph, 'text');
              await context.sync();
              const paragraphText = targetParagraph.text || '';
              
              // Check if the found text is at the very beginning of the paragraph
              // (allowing for minimal whitespace)
              // Use the actual found text from the range, not searchText (which might have different punctuation)
              context.load(foundRange, 'text');
              await context.sync();
              const actualFoundText = foundRange.text || '';
              const foundTextStart = paragraphText.toLowerCase().indexOf(actualFoundText.toLowerCase());
              const textBeforeMatch = foundTextStart >= 0 ? paragraphText.substring(0, foundTextStart).trim() : '';
              
              if (location === 'inline') {
                // Inline: insert right after the found text (within the sentence)
                range = foundRange;
                insertLocation = Word.InsertLocation.after;
              } else if (location === 'before') {
                // "Before" means: insert right before the found text
                // Check if the text is at the very start of the paragraph
                if (textBeforeMatch.length === 0) {
                  // Text is at paragraph start - insert as new paragraph before this paragraph
                  range = targetParagraph.getRange('Start');
                  insertLocation = Word.InsertLocation.before;
                } else {
                  // Text is in the middle or end of paragraph - insert right before the found text
                  range = foundRange;
                  insertLocation = Word.InsertLocation.before;
                }
              } else {
                // For "after", check if we should insert after paragraph or after text
                // If text is at end of paragraph, insert after paragraph; otherwise after text
                const textAfterMatch = paragraphText.substring(foundTextStart + foundRange.text.length).trim();
                if (textAfterMatch.length === 0 || textAfterMatch.length < 5) {
                  // Text is at or near end of paragraph - insert after paragraph
                  range = targetParagraph.getRange('End');
                  insertLocation = Word.InsertLocation.after;
                } else {
                  // Text is in middle - insert right after the found text
                  range = foundRange;
                  insertLocation = Word.InsertLocation.after;
                }
              }
            } else {
              throw new Error(`Invalid location: ${location}`);
            }
            
            await context.sync();
            
            // Insert the text intelligently based on location and context
            let insertedRange: Word.Range;
            
            // Check if we're inserting at paragraph boundaries
            const paragraphStart = targetParagraph ? targetParagraph.getRange('Start') : null;
            const paragraphEnd = targetParagraph ? targetParagraph.getRange('End') : null;
            
            if (location === 'before' && targetParagraph && paragraphStart) {
              // Check if range is at paragraph start by comparing positions
              context.load(range, 'start');
              context.load(paragraphStart, 'start');
              await context.sync();
              
              if (range.start === paragraphStart.start) {
                // Inserting before paragraph start - create new paragraph
                const newParagraph = targetParagraph.insertParagraph(text, Word.InsertLocation.before);
                context.load(newParagraph, ['style']);
                await context.sync();
                
                // Preserve paragraph style if it exists
                if (targetParagraph.style && targetParagraph.style !== 'Normal') {
                  newParagraph.style = targetParagraph.style;
                  await context.sync();
                }
                
                insertedRange = newParagraph.getRange();
              } else {
                // Inserting right before found text - add space if needed
                const textToInsert = text.endsWith(' ') || text.endsWith('\n') ? text : text + ' ';
                insertedRange = range.insertText(textToInsert, Word.InsertLocation.before);
              }
            } else if (location === 'after' && targetParagraph && paragraphEnd) {
              // Check if range is at paragraph end
              context.load(range, 'start');
              context.load(paragraphEnd, 'start');
              await context.sync();
              
              if (range.start === paragraphEnd.start) {
                // Inserting after paragraph end - create new paragraph
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
                // Inserting right after found text - add space if needed
                const textToInsert = text.startsWith(' ') || text.startsWith('\n') ? text : ' ' + text;
                insertedRange = range.insertText(textToInsert, Word.InsertLocation.after);
              }
            } else if (location === 'after' || location === 'inline') {
              // Inserting after found text - add space if needed
              const textToInsert = text.startsWith(' ') || text.startsWith('\n') ? text : ' ' + text;
              insertedRange = range.insertText(textToInsert, Word.InsertLocation.after);
            } else {
              // Regular text insertion
              insertedRange = range.insertText(text, insertLocation);
            }
            
            await context.sync();
            
            // Apply green color to inserted text immediately
            insertedRange.font.color = '#89d185';
            await context.sync();
            
            // Track the change (text is already green from above)
            if (changeTracker) {
              const changeObj: DocumentChange = {
                type: 'insert',
                description: `Inserted "${text}" ${location}${searchText ? ` "${searchText}"` : ''}`,
                newText: text,
                searchText: searchText || location,
                location: location,
                id: `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                timestamp: new Date(),
                applied: false,
                canUndo: true,
              };
              
              // Render inline diff (will ensure text is green if not already)
              await renderInlineDiff(changeObj);
              
              // Track the change
              await changeTracker(changeObj);
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
            context.load(paragraphs, 'items');
            await context.sync();
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
              
              // Track the change (renderInlineDiff will handle the visual diff)
              if (changeTracker) {
                const changeObj: DocumentChange = {
                  type: 'delete',
                  description: `Deleted "${deletedText}"`,
                  oldText: deletedText,
                  searchText: searchText,
                  id: `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                  timestamp: new Date(),
                  applied: false,
                  canUndo: true,
                };
                
                // Render inline diff (this will show deleted text with strikethrough/red)
                await renderInlineDiff(changeObj);
                
                // Track the change
                await changeTracker(changeObj);
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
    const scopedEditTools = createScopedEditTools(articleBoundaries, articleName);
    
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
- readDocument: SEARCH tool - Search ARTICLE ${articleName} for a query and return contextual snippets around each match. MANDATORY: Call this FIRST before any insert/edit/delete. Returns matches with snippets showing context. Use the matchText from results as searchText.
- editDocument: Find and replace text within ARTICLE ${articleName} only. Requires searchText from readDocument results.
- insertText: Insert new text within ARTICLE ${articleName} only. MANDATORY: Requires searchText from readDocument results. If user says "before X", you MUST find X via readDocument first, then use that matchText as searchText with location: "before".
- deleteText: Delete text from ARTICLE ${articleName} only. Requires searchText from readDocument results.

MANDATORY WORKFLOW - FOLLOW THIS EXACTLY:
1. UNDERSTAND the user's instruction. If they say "before 'The Construction Manager shall:'", you MUST find that exact text or a close variation.

2. ALWAYS call readDocument FIRST with the text the user specified:
   - If user says "before 'The Construction Manager shall:'", search for "The Construction Manager shall" (with or without colon)
   - Review the readDocument results - you MUST see matches before proceeding
   - If no matches found, try variations: "Construction Manager shall", "The Construction Manager", etc.
   - DO NOT proceed to insert/edit until you've found the location via readDocument

3. Once readDocument returns matches:
   - Use the EXACT matchText from the results as searchText for insertText
   - If multiple matches, use the first one (or the one that makes sense in context)
   - Call insertText with location: "before" and the matchText as searchText

4. CRITICAL: If readDocument doesn't find the text after multiple searches:
   - Report what you searched for and that it wasn't found
   - DO NOT default to "beginning" or "end" - that's wrong!
   - Ask the user for clarification or an alternative search term

5. NEVER insert at "beginning" or "end" unless the user explicitly asks for that. If user says "before X", you MUST find X first.

6. Use ONE tool call at a time. Wait for the tool result before deciding the next action.

7. For insertText: location "before"/"after"/"inline" requires searchText from readDocument results.

The user has provided the following instructions for ARTICLE ${articleName}:
${instruction}

Use your AI intelligence to understand where to make the changes, then use the tools to execute them.`,
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
