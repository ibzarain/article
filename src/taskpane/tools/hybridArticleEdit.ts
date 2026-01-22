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
    description: 'Search ARTICLE content for a query and return contextual snippets around each match. This is a SEARCH tool - use it to find exact text before editing. Returns matches with snippets showing context. If query is "*" or "all", returns the full article content.',
    parameters: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Text to search for in the article. This is a search query - use it to find exact text before making edits. Use "*" or "all" to get the full article content.',
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

          // If query is "*" or "all", return full article content
          if (query === '*' || query.toLowerCase() === 'all') {
            console.log(`[readDocument] Returning full article content (${text.length} characters)`);
            return {
              matches: [{
                matchText: 'FULL ARTICLE CONTENT',
                snippet: text,
                matchStart: 0,
                matchEnd: text.length,
                snippetStart: 0,
                snippetEnd: text.length,
              }],
              totalFound: 1,
              articleLength: text.length,
              fullContent: text, // Add full content flag
            };
          }
          const safeContextChars = Math.max(0, Math.floor(contextChars || 0));
          const safeMaxMatches = typeof maxMatches === 'number' && maxMatches > 0 ? Math.floor(maxMatches) : undefined;

          // Escape regex special characters
          const escapeRegExp = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

          // Try multiple search patterns to handle punctuation variations
          const searchPatterns: string[] = [];

          // Original query
          searchPatterns.push(query);

          // Without trailing punctuation
          if (/[.:;]$/.test(query)) {
            searchPatterns.push(query.replace(/[.:;]+$/, ''));
          }

          // With punctuation added
          if (!/[.:;]$/.test(query)) {
            searchPatterns.push(query + ':');
            searchPatterns.push(query + '.');
          }

          // Trimmed version
          const trimmed = query.trim();
          if (trimmed !== query) {
            searchPatterns.push(trimmed);
          }

          // Remove duplicates
          const uniquePatterns = Array.from(new Set(searchPatterns));

          const matches: Array<{
            matchText: string;
            snippet: string;
            matchStart: number;
            matchEnd: number;
            snippetStart: number;
            snippetEnd: number;
          }> = [];

          let totalFound = 0;
          const foundPositions = new Set<number>(); // Track positions to avoid duplicates

          // Try each pattern
          for (const patternQuery of uniquePatterns) {
            const escapedQuery = escapeRegExp(patternQuery);
            const pattern = matchWholeWord ? `\\b${escapedQuery}\\b` : escapedQuery;
            const flags = matchCase ? 'g' : 'gi';
            const regex = new RegExp(pattern, flags);

            let match: RegExpExecArray | null;
            while ((match = regex.exec(text)) !== null) {
              // Avoid counting the same match twice
              if (!foundPositions.has(match.index)) {
                foundPositions.add(match.index);
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
            }
          }

          return {
            matches,
            totalFound,
            articleLength: text.length,
            fullContent: undefined as string | undefined, // No full content for search queries
          };
        }) as {
          matches: Array<{
            matchText: string;
            snippet: string;
            matchStart: number;
            matchEnd: number;
            snippetStart: number;
            snippetEnd: number;
          }>;
          totalFound: number;
          articleLength: number;
          fullContent?: string;
        };

        // Log the result
        if (query === '*' || query.toLowerCase() === 'all') {
          console.log(`[readDocument] Full article content retrieved:`, result.fullContent);
        } else {
          console.log(`[readDocument] Search for "${query}" found ${result.totalFound} match(es)`);
        }

        return {
          success: true,
          query,
          content: result.matches,
          totalFound: result.totalFound,
          articleLength: result.articleLength,
          fullContent: result.fullContent, // Include full content in response
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
          const warnings: string[] = [];

          // First, locate the target text within the article and capture the actual
          // matched text. Do NOT call renderInlineDiff/changeTracker inside Word.run
          // because renderInlineDiff uses Word.run internally (nested Word.run is unstable).
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
            const capturedOldTexts: string[] = [];

            for (const item of itemsToReplace) {
              context.load(item, 'text');
              await context.sync();
              const actualOldText = item.text;

              replacementCount++;
              capturedOldTexts.push(actualOldText);
            }

            await context.sync();

            return {
              replaced: replacementCount,
              totalFound: searchResults.items.length,
              capturedOldTexts,
            };
          });

          // Second, render the inline diff + record the change OUTSIDE Word.run (best-effort).
          if (changeTracker) {
            const oldTexts = Array.isArray((result as any).capturedOldTexts)
              ? ((result as any).capturedOldTexts as string[])
              : [];

            // Keep behavior similar to before: render a diff per targeted occurrence.
            // Note: renderInlineDiff currently searches by `searchText` and uses the first match,
            // so replaceAll may still behave unexpectedly in documents with repeated matches.
            for (const oldText of oldTexts.length > 0 ? oldTexts : [searchText]) {
              const changeObj: DocumentChange = {
                type: 'edit',
                description: `Replaced "${oldText}" with "${newText}"`,
                oldText: oldText,
                newText: newText,
                searchText: searchText,
                id: `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                timestamp: new Date(),
                applied: false,
                canUndo: true,
              };

              try {
                await renderInlineDiff(changeObj);
              } catch (e) {
                warnings.push(
                  `Edit succeeded, but failed to render inline diff: ${e instanceof Error ? e.message : String(e)
                  }`
                );
              }

              try {
                await changeTracker(changeObj);
              } catch (e) {
                warnings.push(
                  `Edit succeeded, but failed to record change: ${e instanceof Error ? e.message : String(e)
                  }`
                );
              }
            }
          }

          return {
            success: true,
            replaced: result.replaced,
            totalFound: result.totalFound,
            message: `Replaced ${result.replaced} occurrence(s) of "${searchText}" with "${newText}"`,
            ...(warnings.length > 0 ? { warnings } : {}),
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
          const warnings: string[] = [];

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
            let foundRange: Word.Range | null = null;
            let insertAsNewParagraph = false; // Track if we should insert as new paragraph vs inline

            if (location === 'beginning') {
              range = startRange;
              insertLocation = Word.InsertLocation.after;
              targetParagraph = startParagraph;
              insertAsNewParagraph = text.includes('\n'); // Multiline = new paragraph
            } else if (location === 'end') {
              range = endRange;
              insertLocation = Word.InsertLocation.before;
              targetParagraph = endParagraph;
              insertAsNewParagraph = text.includes('\n'); // Multiline = new paragraph
            } else if (location === 'before' || location === 'after' || location === 'inline') {
              if (!searchText) {
                throw new Error('searchText is required when location is "before", "after", or "inline"');
              }

              // Search only within article range
              // Try multiple search strategies to find the text
              // Normalize search text: remove extra whitespace, handle line breaks
              const normalizeSearchText = (text: string): string => {
                return text.replace(/\s+/g, ' ').trim();
              };

              const normalizedSearch = normalizeSearchText(searchText);
              let searchResults = articleRange.search(normalizedSearch, {
                matchCase: false,
                matchWholeWord: false,
              });

              context.load(searchResults, 'items');
              await context.sync();

              const searchAttempts: string[] = [normalizedSearch];

              // If not found, try original (might have different whitespace)
              if (searchResults.items.length === 0 && searchText !== normalizedSearch) {
                searchAttempts.push(searchText);
                searchResults = articleRange.search(searchText, {
                  matchCase: false,
                  matchWholeWord: false,
                });
                context.load(searchResults, 'items');
                await context.sync();
              }

              // If not found, try with different whitespace handling
              if (searchResults.items.length === 0) {
                const trimmedSearch = normalizedSearch.trim();
                if (trimmedSearch !== normalizedSearch) {
                  searchAttempts.push(trimmedSearch);
                  searchResults = articleRange.search(trimmedSearch, {
                    matchCase: false,
                    matchWholeWord: false,
                  });
                  context.load(searchResults, 'items');
                  await context.sync();
                }
              }

              // Try without punctuation at the end
              if (searchResults.items.length === 0 && normalizedSearch && /[.:;]/.test(normalizedSearch)) {
                const withoutPunct = normalizedSearch.replace(/[.:;]+$/, '').trim();
                if (withoutPunct && withoutPunct !== normalizedSearch && !searchAttempts.includes(withoutPunct)) {
                  searchAttempts.push(withoutPunct);
                  searchResults = articleRange.search(withoutPunct, {
                    matchCase: false,
                    matchWholeWord: false,
                  });
                  context.load(searchResults, 'items');
                  await context.sync();
                }
              }

              // Try with punctuation added if original didn't have it
              if (searchResults.items.length === 0 && normalizedSearch && !/[.:;]$/.test(normalizedSearch)) {
                const withPunct = normalizedSearch + ':';
                if (!searchAttempts.includes(withPunct)) {
                  searchAttempts.push(withPunct);
                  searchResults = articleRange.search(withPunct, {
                    matchCase: false,
                    matchWholeWord: false,
                  });
                  context.load(searchResults, 'items');
                  await context.sync();
                }
              }

              // Try partial match - just the key words if it's a phrase
              if (searchResults.items.length === 0 && normalizedSearch) {
                const words = normalizedSearch.trim().split(/\s+/);
                if (words.length > 2) {
                  // Try last 2-3 words (e.g., "Construction Manager shall" from "The Construction Manager shall:")
                  const partialSearch = words.slice(-3).join(' ');
                  if (partialSearch && !searchAttempts.includes(partialSearch)) {
                    searchAttempts.push(partialSearch);
                    searchResults = articleRange.search(partialSearch, {
                      matchCase: false,
                      matchWholeWord: false,
                    });
                    context.load(searchResults, 'items');
                    await context.sync();
                  }
                }
                // Also try first few words if it's a longer phrase
                if (searchResults.items.length === 0 && words.length > 3) {
                  const firstWords = words.slice(0, 3).join(' ');
                  if (firstWords && !searchAttempts.includes(firstWords)) {
                    searchAttempts.push(firstWords);
                    searchResults = articleRange.search(firstWords, {
                      matchCase: false,
                      matchWholeWord: false,
                    });
                    context.load(searchResults, 'items');
                    await context.sync();
                  }
                }
              }

              // Fallback: search paragraph by paragraph if range search fails
              if (searchResults.items.length === 0) {
                // Get all paragraphs in the article
                const articleParagraphs: Word.Paragraph[] = [];
                for (let i = articleBoundaries.startParagraphIndex; i <= articleBoundaries.endParagraphIndex; i++) {
                  articleParagraphs.push(paragraphs.items[i]);
                }

                // Load all paragraph texts first
                for (const para of articleParagraphs) {
                  context.load(para, 'text');
                }
                await context.sync();

                // Search each paragraph individually using Word's search API
                for (const para of articleParagraphs) {
                  const paraRange = para.getRange('Whole');

                  // Try normalized search first
                  let paraSearchResults = paraRange.search(normalizedSearch, {
                    matchCase: false,
                    matchWholeWord: false,
                  });
                  context.load(paraSearchResults, 'items');
                  await context.sync();

                  // If not found, try original search text
                  if (paraSearchResults.items.length === 0 && searchText !== normalizedSearch) {
                    paraSearchResults = paraRange.search(searchText, {
                      matchCase: false,
                      matchWholeWord: false,
                    });
                    context.load(paraSearchResults, 'items');
                    await context.sync();
                  }

                  // If still not found, try searching in the paragraph text directly
                  if (paraSearchResults.items.length === 0) {
                    const paraText = para.text || '';
                    const searchLower = normalizedSearch.toLowerCase();
                    const paraTextLower = paraText.toLowerCase();

                    if (paraTextLower.includes(searchLower)) {
                      // Found in text, now try to get the range using a substring search
                      // Try to find a unique substring that includes our search text
                      const searchIndex = paraTextLower.indexOf(searchLower);
                      if (searchIndex >= 0) {
                        // Get a range that includes the found text
                        // Use a slightly longer search string that includes context
                        const startPos = Math.max(0, searchIndex - 5);
                        const endPos = Math.min(paraText.length, searchIndex + normalizedSearch.length + 5);
                        const contextSearch = paraText.substring(startPos, endPos);

                        // Try searching for this context string
                        paraSearchResults = paraRange.search(contextSearch, {
                          matchCase: false,
                          matchWholeWord: false,
                        });
                        context.load(paraSearchResults, 'items');
                        await context.sync();

                        // If that doesn't work, try just the core words
                        if (paraSearchResults.items.length === 0) {
                          const words = normalizedSearch.split(/\s+/).filter(w => w.length > 2);
                          if (words.length > 0) {
                            const coreSearch = words.join(' ');
                            paraSearchResults = paraRange.search(coreSearch, {
                              matchCase: false,
                              matchWholeWord: false,
                            });
                            context.load(paraSearchResults, 'items');
                            await context.sync();
                          }
                        }
                      }
                    }
                  }

                  if (paraSearchResults.items.length > 0) {
                    searchResults = paraSearchResults;
                    break;
                  }
                }
              }

              // ALWAYS set up fallback paragraph BEFORE trying to use Word search results
              // This way we have a backup if Word search returns invalid results
              let fallbackParagraph: Word.Paragraph | null = null;
              const articleParagraphs: Word.Paragraph[] = [];
              for (let i = articleBoundaries.startParagraphIndex; i <= articleBoundaries.endParagraphIndex; i++) {
                articleParagraphs.push(paragraphs.items[i]);
              }

              for (const para of articleParagraphs) {
                context.load(para, 'text');
              }
              await context.sync();

              const normalizedNeedle = normalizeSearchText(searchText).toLowerCase();
              for (const para of articleParagraphs) {
                const normalizedHaystack = normalizeSearchText(para.text || '').toLowerCase();
                if (normalizedHaystack.includes(normalizedNeedle)) {
                  fallbackParagraph = para;
                  break;
                }
              }

              if (searchResults.items.length === 0 && !fallbackParagraph) {
                // Get article content snippet for better error message
                context.load(articleRange, 'text');
                await context.sync();
                const articlePreview = articleRange.text.substring(0, 1000);
                const searchedTerms = searchAttempts.join('", "');
                throw new Error(`Search text "${searchText}" not found in ARTICLE ${articleName}. Tried: "${searchedTerms}". Article content preview: ${articlePreview.substring(0, 500)}... Please use readDocument first to find the exact text in the article.`);
              }

              // Log the search for debugging
              console.log(`[insertText] Searching for: "${searchText}", found ${searchResults.items.length} match(es) in ARTICLE ${articleName}`);

              // Try to use Word search results, but fall back to paragraph-based if it fails
              let useWordSearch = false;
              if (searchResults.items.length > 0) {
                try {
                  // Use the first match (most relevant)
                  foundRange = searchResults.items[0];
                  targetParagraph = foundRange.paragraphs.getFirst();
                  context.load(targetParagraph, ['listItem', 'list', 'text', 'style']);

                  // Get paragraph text to check context
                  context.load(targetParagraph, 'text');
                  await context.sync();

                  // If we got here, Word search worked
                  useWordSearch = true;
                } catch (error) {
                  // Word search returned results but we can't use them - fall back to paragraph search
                  console.warn(`[insertText] Word search found results but couldn't access them, using fallback:`, error);
                  useWordSearch = false;
                }
              }

              if (useWordSearch && targetParagraph) {
                const paragraphText = targetParagraph.text || '';

                // Check if the found text is at the very beginning of the paragraph
                // (allowing for minimal whitespace)
                // Use the actual found text from the range, not searchText (which might have different punctuation)
                context.load(foundRange!, 'text');
                await context.sync();
                const actualFoundText = foundRange!.text || '';
                const foundTextStart = paragraphText.toLowerCase().indexOf(actualFoundText.toLowerCase());
                const textBeforeMatch = foundTextStart >= 0 ? paragraphText.substring(0, foundTextStart).trim() : '';

                if (location === 'inline') {
                  // Inline: insert right after the found text (within the sentence)
                  range = foundRange!;
                  insertLocation = Word.InsertLocation.after;
                  insertAsNewParagraph = false;
                } else if (location === 'before') {
                  // "Before" means: insert right before the found text
                  // Check if the text is at the very start of the paragraph
                  if (textBeforeMatch.length === 0) {
                    // Text is at paragraph start - insert as new paragraph before this paragraph
                    range = targetParagraph.getRange('Start');
                    insertLocation = Word.InsertLocation.before;
                    insertAsNewParagraph = true;
                  } else {
                    // Text is in the middle or end of paragraph
                    // For multiline text, insert as new paragraph; otherwise inline
                    if (text.includes('\n')) {
                      range = targetParagraph.getRange('Start');
                      insertLocation = Word.InsertLocation.before;
                      insertAsNewParagraph = true;
                    } else {
                      range = foundRange!;
                      insertLocation = Word.InsertLocation.before;
                      insertAsNewParagraph = false;
                    }
                  }
                } else {
                  // For "after", check if we should insert after paragraph or after text
                  // If text is at end of paragraph, insert after paragraph; otherwise after text
                  const textAfterMatch = paragraphText.substring(foundTextStart + foundRange!.text.length).trim();
                  if (textAfterMatch.length === 0 || textAfterMatch.length < 5) {
                    // Text is at or near end of paragraph - insert after paragraph
                    range = targetParagraph.getRange('End');
                    insertLocation = Word.InsertLocation.after;
                    insertAsNewParagraph = true;
                  } else {
                    // Text is in middle
                    // For multiline text, insert as new paragraph; otherwise inline
                    if (text.includes('\n')) {
                      range = targetParagraph.getRange('End');
                      insertLocation = Word.InsertLocation.after;
                      insertAsNewParagraph = true;
                    } else {
                      range = foundRange!;
                      insertLocation = Word.InsertLocation.after;
                      insertAsNewParagraph = false;
                    }
                  }
                }
              } else if (fallbackParagraph) {
                targetParagraph = fallbackParagraph;

                if (location === 'inline') {
                  throw new Error('Unable to locate inline insertion point via Word search. Please use a longer, more specific searchText from readDocument.');
                } else if (location === 'before') {
                  range = targetParagraph.getRange('Start');
                  insertLocation = Word.InsertLocation.before;
                  insertAsNewParagraph = text.includes('\n'); // Multiline = new paragraph
                } else {
                  range = targetParagraph.getRange('End');
                  insertLocation = Word.InsertLocation.after;
                  insertAsNewParagraph = text.includes('\n'); // Multiline = new paragraph
                }
              }
            } else {
              throw new Error(`Invalid location: ${location}`);
            }

            await context.sync();

            // Insert the text intelligently based on location and context
            // Use the insertAsNewParagraph flag we set earlier to determine insertion method
            let insertedRange: Word.Range;

            if (insertAsNewParagraph && targetParagraph) {
              // Insert as new paragraph(s) - split by newlines to preserve formatting
              // Determine the correct InsertLocation based on the location parameter
              const initialInsertLocation = 
                location === 'before' || location === 'end' 
                  ? Word.InsertLocation.before 
                  : Word.InsertLocation.after;
              
              // Split text by newlines to create multiple paragraphs if needed
              const textLines = text.split('\n');
              let firstParagraph: Word.Paragraph | null = null;
              let lastParagraph: Word.Paragraph = targetParagraph;
              
              for (let i = 0; i < textLines.length; i++) {
                const lineText = textLines[i];
                // Skip empty lines at the start/end but preserve them in the middle
                if (i === 0 && lineText.trim() === '' && textLines.length > 1) continue;
                if (i === textLines.length - 1 && lineText.trim() === '' && textLines.length > 1) continue;
                
                // First paragraph uses initial location, subsequent ones always use 'after' to maintain order
                const paragraphInsertLocation = i === 0 ? initialInsertLocation : Word.InsertLocation.after;
                const newParagraph = lastParagraph.insertParagraph(lineText, paragraphInsertLocation);
                context.load(newParagraph, ['style']);
                await context.sync();

                // Preserve paragraph style if it exists
                if (targetParagraph.style && targetParagraph.style !== 'Normal') {
                  newParagraph.style = targetParagraph.style;
                  await context.sync();
                }
                
                if (firstParagraph === null) {
                  firstParagraph = newParagraph;
                }
                lastParagraph = newParagraph;
              }
              
              // Use the range of all inserted paragraphs
              if (firstParagraph) {
                insertedRange = firstParagraph.getRange().expandTo(lastParagraph.getRange());
              } else {
                // Fallback if no paragraphs were created
                const newParagraph = targetParagraph.insertParagraph(text, initialInsertLocation);
                context.load(newParagraph, ['style']);
                await context.sync();
                if (targetParagraph.style && targetParagraph.style !== 'Normal') {
                  newParagraph.style = targetParagraph.style;
                  await context.sync();
                }
                insertedRange = newParagraph.getRange();
              }
            } else if (location === 'inline') {
              // Inline insertion: insert text directly after found text
              // Convert newlines to spaces for inline insertion
              const textToInsert = (text.startsWith(' ') || text.startsWith('\n') ? text : ' ' + text)
                .replace(/\n/g, ' ');
              insertedRange = range.insertText(textToInsert, Word.InsertLocation.after);
            } else if (location === 'after') {
              // Inserting after found text - handle newlines properly
              if (text.includes('\n')) {
                // If text has newlines, insert as paragraphs
                const textLines = text.split('\n');
                let firstParagraph: Word.Paragraph | null = null;
                let lastParagraph: Word.Paragraph = targetParagraph;
                
                for (let i = 0; i < textLines.length; i++) {
                  const lineText = textLines[i];
                  if (i === 0 && lineText.trim() === '' && textLines.length > 1) continue;
                  if (i === textLines.length - 1 && lineText.trim() === '' && textLines.length > 1) continue;
                  
                  const newParagraph = lastParagraph.insertParagraph(lineText, Word.InsertLocation.after);
                  context.load(newParagraph, ['style']);
                  await context.sync();
                  
                  if (targetParagraph.style && targetParagraph.style !== 'Normal') {
                    newParagraph.style = targetParagraph.style;
                    await context.sync();
                  }
                  
                  if (firstParagraph === null) {
                    firstParagraph = newParagraph;
                  }
                  lastParagraph = newParagraph;
                }
                
                if (firstParagraph) {
                  insertedRange = firstParagraph.getRange().expandTo(lastParagraph.getRange());
                } else {
                  const textToInsert = text.startsWith(' ') || text.startsWith('\n') ? text : ' ' + text;
                  insertedRange = range.insertText(textToInsert, Word.InsertLocation.after);
                }
              } else {
                // No newlines - regular inline insertion
                const textToInsert = text.startsWith(' ') || text.startsWith('\n') ? text : ' ' + text;
                insertedRange = range.insertText(textToInsert, Word.InsertLocation.after);
              }
            } else {
              // Regular text insertion (before found text, or beginning/end of article)
              // Handle newlines by splitting into paragraphs
              if (text.includes('\n')) {
                const textLines = text.split('\n');
                let firstParagraph: Word.Paragraph | null = null;
                let lastParagraph: Word.Paragraph = targetParagraph;
                
                // Determine initial insert location based on location parameter
                const initialInsertLocation = 
                  location === 'before' || location === 'end' 
                    ? Word.InsertLocation.before 
                    : Word.InsertLocation.after;
                
                for (let i = 0; i < textLines.length; i++) {
                  const lineText = textLines[i];
                  if (i === 0 && lineText.trim() === '' && textLines.length > 1) continue;
                  if (i === textLines.length - 1 && lineText.trim() === '' && textLines.length > 1) continue;
                  
                  // First paragraph uses initial location, subsequent ones use 'after'
                  const paragraphInsertLocation = i === 0 ? initialInsertLocation : Word.InsertLocation.after;
                  const newParagraph = lastParagraph.insertParagraph(lineText, paragraphInsertLocation);
                  context.load(newParagraph, ['style']);
                  await context.sync();
                  
                  if (targetParagraph.style && targetParagraph.style !== 'Normal') {
                    newParagraph.style = targetParagraph.style;
                    await context.sync();
                  }
                  
                  if (firstParagraph === null) {
                    firstParagraph = newParagraph;
                  }
                  lastParagraph = newParagraph;
                }
                
                if (firstParagraph) {
                  insertedRange = firstParagraph.getRange().expandTo(lastParagraph.getRange());
                } else {
                  insertedRange = range.insertText(text, insertLocation);
                }
              } else {
                insertedRange = range.insertText(text, insertLocation);
              }
            }

            await context.sync();

            // Apply green color to inserted text immediately
            insertedRange.font.color = '#89d185';
            await context.sync();

            return {
              inserted: true,
            };
          });

          // IMPORTANT: Do NOT call renderInlineDiff/changeTracker inside Word.run.
          // renderInlineDiff uses Word.run internally; nesting Word.run often causes
          // opaque Office errors like "We couldn't find the item you requested".
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

            // Best-effort diff rendering + tracking. Insert itself already happened.
            try {
              await renderInlineDiff(changeObj);
            } catch (e) {
              warnings.push(
                `Inserted text successfully, but failed to render inline diff: ${e instanceof Error ? e.message : String(e)
                }`
              );
            }

            try {
              await changeTracker(changeObj);
            } catch (e) {
              warnings.push(
                `Inserted text successfully, but failed to record change: ${e instanceof Error ? e.message : String(e)
                }`
              );
            }
          }

          return {
            success: true,
            message: `Text inserted successfully at ${location}`,
            ...(warnings.length > 0 ? { warnings } : {}),
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

    // Log the article content that was extracted
    console.log(`[executeArticleInstructionsHybrid] Extracted ARTICLE ${articleName}:`);
    console.log(`  Start paragraph: ${articleBoundaries.startParagraphIndex}`);
    console.log(`  End paragraph: ${articleBoundaries.endParagraphIndex}`);
    console.log(`  Content length: ${articleBoundaries.articleContent.length} characters`);
    console.log(`  Full content:`, articleBoundaries.articleContent);

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

CRITICAL: PRESERVE FORMATTING - When extracting text from the user's instruction to insert, you MUST preserve all newlines (\\n), line breaks, and indentation exactly as provided. Do NOT normalize, trim, or modify the formatting of the text to insert. The text parameter should contain the exact formatting including newline characters.

CURRENT ARTICLE CONTENT (for reference):
${articleContentPreview}

AVAILABLE TOOLS:
- readDocument: SEARCH tool - Search ARTICLE ${articleName} for a query and return contextual snippets around each match. MANDATORY: Call this FIRST before any insert/edit/delete. Returns matches with snippets showing context. Use the matchText from results as searchText. IMPORTANT: If you call readDocument with query "*" or "all", it will return the FULL article content. You MUST call this at the start to get the full article content and return it to the user.
- editDocument: Find and replace text within ARTICLE ${articleName} only. Requires searchText from readDocument results.
- insertText: Insert new text within ARTICLE ${articleName} only. MANDATORY: Requires searchText from readDocument results. If user says "before X", you MUST find X via readDocument first, then use that matchText as searchText with location: "before". IMPORTANT: When extracting the text to insert from the user's instruction, preserve ALL newlines (\\n) and formatting exactly as provided. The text parameter must include newline characters where the user has line breaks.
- deleteText: Delete text from ARTICLE ${articleName} only. Requires searchText from readDocument results.

MANDATORY WORKFLOW - FOLLOW THIS EXACTLY:
1. FIRST STEP - GET FULL ARTICLE CONTENT:
   - IMMEDIATELY call readDocument with query "*" or "all" to get the FULL article content
   - This will return the complete article text
   - You MUST return this full article content to the user in your response
   - Log it: "Here is the full content of ARTICLE ${articleName} that I found: [full content]"

2. UNDERSTAND the user's instruction. If they say "before 'The Construction Manager shall:'", you MUST find that exact text or a close variation.

3. ALWAYS call readDocument with the specific text the user specified:
   - If user says "before 'The Construction Manager shall:'", search for "The Construction Manager shall" (with or without colon)
   - Review the readDocument results - you MUST see matches before proceeding
   - If no matches found, try variations: "Construction Manager shall", "The Construction Manager", etc.
   - DO NOT proceed to insert/edit until you've found the location via readDocument

4. Once readDocument returns matches:
   - Use the EXACT matchText from the results as searchText for insertText
   - If multiple matches, use the first one (or the one that makes sense in context)
   - Call insertText with location: "before" and the matchText as searchText

5. CRITICAL: If readDocument doesn't find the text after multiple searches:
   - Report what you searched for and that it wasn't found
   - DO NOT default to "beginning" or "end" - that's wrong!
   - Ask the user for clarification or an alternative search term

6. NEVER insert at "beginning" or "end" unless the user explicitly asks for that. If user says "before X", you MUST find X first.

7. Use ONE tool call at a time. Wait for the tool result before deciding the next action.

8. For insertText: location "before"/"after"/"inline" requires searchText from readDocument results.

9. FINAL STEP: After completing all edits, return a summary that includes:
   - The full article content you retrieved at the start
   - What changes you made
   - Confirmation that edits were applied

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
