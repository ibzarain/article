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

type ScopedReadState = {
  hasFreshRead: boolean;
  lastQuery?: string;
};
type ScopedReadGuard = {
  requiredTokens: string[];
};

/**
 * Extracts all relevant context tokens from the instruction to lock readDocument queries
 * Extracts numbered paragraphs, quoted strings, and key phrases mentioned in the instruction
 */
function extractInstructionContext(instruction: string): string[] {
  const tokens: string[] = [];
  
  // Extract numbered paragraphs (e.g., "1.3", "1.2")
  const numbered = instruction.match(/\b\d+\.\d+\b/g) || [];
  tokens.push(...numbered);
  
  // Extract quoted strings (both double and single quotes)
  const doubleQuoted = instruction.match(/"([^"]+)"/g) || [];
  const singleQuoted = instruction.match(/'([^']+)'/g) || [];
  [...doubleQuoted, ...singleQuoted].forEach(q => {
    const content = q.replace(/["']/g, '').trim();
    // Only add meaningful quoted content (at least 3 characters)
    if (content.length >= 3) {
      tokens.push(content);
      // Also add key words from quoted content
      const words = content.split(/\s+/).filter(w => w.length >= 4);
      tokens.push(...words);
    }
  });
  
  // Extract numbered paragraphs after action words (Delete paragraph 1.3, Substitute 1.2, etc.)
  const actionPatterns = [
    /(?:Delete|Substitute|Replace|Insert|Add)\s+(?:paragraph\s+)?(\d+\.\d+)/gi,
    /paragraph\s+(\d+\.\d+)/gi,
  ];
  actionPatterns.forEach(pattern => {
    let match;
    while ((match = pattern.exec(instruction)) !== null) {
      if (match[1]) tokens.push(match[1]);
    }
  });
  
  // Extract text after "before" or "after" keywords
  const beforeAfterPattern = /(?:before|after)\s+["']?([^"'\n]{3,})["']?/gi;
  let match;
  while ((match = beforeAfterPattern.exec(instruction)) !== null) {
    if (match[1]) {
      const text = match[1].trim();
      tokens.push(text);
      // Add key words from the text
      const words = text.split(/\s+/).filter(w => w.length >= 4);
      tokens.push(...words);
    }
  }
  
  // Extract substitution text patterns (e.g., "substitute the following: 1.3 commence...")
  const substitutePattern = /substitute[^:]*:\s*(\d+\.\d+)/gi;
  let subMatch;
  while ((subMatch = substitutePattern.exec(instruction)) !== null) {
    if (subMatch[1]) tokens.push(subMatch[1]);
  }
  
  // Extract key words from substitution text (the text that will be inserted)
  // Pattern: "substitute [as follows|the following]: [text]"
  const substituteTextPattern = /substitute[^:]*:\s*(.+?)(?:\n|$)/gi;
  let subTextMatch;
  while ((subTextMatch = substituteTextPattern.exec(instruction)) !== null) {
    if (subTextMatch[1]) {
      const subText = subTextMatch[1].trim();
      // Extract key phrases (3+ words) and important words (4+ chars) from substitution text
      const words = subText.split(/\s+/).filter(w => w.length >= 4);
      tokens.push(...words);
      // Extract key phrases (e.g., "except as expressly provided", "Contract Documents")
      const phrases = subText.match(/\b\w{4,}(?:\s+\w{4,}){1,2}\b/g) || [];
      tokens.push(...phrases);
    }
  }
  
  // Extract key phrases from common instruction patterns
  // "Article A-1" or "A-1" references
  const articleRef = instruction.match(/\b([A-Z]-\d+)\b/gi) || [];
  tokens.push(...articleRef);
  
  // Normalize and deduplicate tokens
  const normalized = tokens
    .map(t => t.toLowerCase().trim())
    .filter(t => t.length > 0);
  
  return Array.from(new Set(normalized));
}

/**
 * Uses AI to semantically find relevant text chunks in the article
 */
async function findSemanticMatches(
  query: string,
  articleText: string,
  apiKey: string,
  model: string,
  contextChars: number,
  maxMatches?: number
): Promise<Array<{
  matchText: string;
  snippet: string;
  matchStart: number;
  matchEnd: number;
  snippetStart: number;
  snippetEnd: number;
}>> {
  // Split article into semantic chunks (prioritize numbered paragraphs, then regular paragraphs, then sentences)
  const chunks: Array<{ text: string; start: number; end: number }> = [];
  
  // First, try to split by numbered paragraphs (lines starting with "X.Y" where X and Y are digits)
  const numberedParaPattern = /(^\d+\.\d+\s+[^\n]+(?:\n(?!\d+\.\d+)[^\n]+)*)/gm;
  let match;
  let lastEnd = 0;
  
  while ((match = numberedParaPattern.exec(articleText)) !== null) {
    // Add any text before this match
    if (match.index > lastEnd) {
      const beforeText = articleText.substring(lastEnd, match.index).trim();
      if (beforeText.length > 0) {
        chunks.push({ text: beforeText, start: lastEnd, end: match.index });
      }
    }
    
    const paraText = match[0].trim();
    if (paraText.length > 0) {
      chunks.push({ text: paraText, start: match.index, end: match.index + match[0].length });
    }
    lastEnd = match.index + match[0].length;
  }
  
  // Add remaining text
  if (lastEnd < articleText.length) {
    const remaining = articleText.substring(lastEnd).trim();
    if (remaining.length > 0) {
      chunks.push({ text: remaining, start: lastEnd, end: articleText.length });
    }
  }
  
  // If no numbered paragraphs found, split by regular paragraphs
  if (chunks.length === 0) {
    const paragraphs = articleText.split(/\n\s*\n|\r\n\s*\r\n/);
    let currentPos = 0;
    
    for (const para of paragraphs) {
      if (para.trim().length > 0) {
        const start = currentPos;
        const end = start + para.length;
        chunks.push({ text: para.trim(), start, end });
        currentPos = end + (articleText.substring(end).match(/^\s*\n/) || [''])[0].length;
      } else {
        currentPos += para.length + 1;
      }
    }
  }
  
  // If still no chunks, split by sentences
  if (chunks.length === 0) {
    const sentences = articleText.split(/(?<=[.!?])\s+/);
    let currentPos = 0;
    for (const sent of sentences) {
      if (sent.trim().length > 10) {
        const start = currentPos;
        const end = start + sent.length;
        chunks.push({ text: sent.trim(), start, end });
        currentPos = end + 1;
      } else {
        currentPos += sent.length + 1;
      }
    }
  }
  
  // Use AI to score each chunk's relevance to the query
  const scoringPrompt = `You are a text search assistant. Given a search query and a list of text chunks from a document, identify which chunks are semantically relevant to the query.

IMPORTANT: If the query contains a numbered paragraph reference (like "1.2" or "1.3"), prioritize chunks that start with that exact number pattern. For example, if query is "1.2", look for chunks starting with "1.2 " or "1.2".

Query: "${query}"

Text chunks:
${chunks.map((chunk, i) => `[${i}] ${chunk.text.substring(0, 300)}${chunk.text.length > 300 ? '...' : ''}`).join('\n\n')}

Return a JSON array of chunk indices (0-based) that are semantically relevant to the query, ordered by relevance (most relevant first). Only include chunks that are actually relevant. Format: [0, 3, 5]`;

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: model,
        messages: [
          { role: 'system', content: 'You are a helpful assistant that identifies semantically relevant text chunks. Return only a JSON array of indices.' },
          { role: 'user', content: scoringPrompt },
        ],
        temperature: 0.1,
        max_tokens: 100,
      }),
    });

    if (!response.ok) {
      throw new Error(`API error: ${response.status}`);
    }

    const data = await response.json();
    const content = data.choices[0]?.message?.content || '[]';
    const relevantIndices: number[] = JSON.parse(content.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim());
    
    // Build matches from relevant chunks
    const matches: Array<{
      matchText: string;
      snippet: string;
      matchStart: number;
      matchEnd: number;
      snippetStart: number;
      snippetEnd: number;
    }> = [];
    
    for (const idx of relevantIndices.slice(0, maxMatches || 20)) {
      if (idx >= 0 && idx < chunks.length) {
        const chunk = chunks[idx];
        const snippetStart = Math.max(0, chunk.start - contextChars);
        const snippetEnd = Math.min(articleText.length, chunk.end + contextChars);
        
        matches.push({
          matchText: chunk.text,
          snippet: articleText.slice(snippetStart, snippetEnd),
          matchStart: chunk.start,
          matchEnd: chunk.end,
          snippetStart,
          snippetEnd,
        });
      }
    }
    
    return matches;
  } catch (error) {
    console.error('Error in semantic search:', error);
    // Fallback: return empty matches
    return [];
  }
}

/**
 * Creates a scoped readDocument tool that only reads content within article boundaries
 * Uses AI semantic search instead of regex
 */
function createScopedReadDocumentTool(
  articleBoundaries: ArticleBoundaries,
  readState: ScopedReadState,
  readGuard: ScopedReadGuard,
  apiKey: string,
  model: string
) {
  return {
    description: 'Search ARTICLE content using AI semantic search. This tool uses AI to find semantically relevant text chunks matching your query, not exact text matching. Returns matches with snippets showing context. If query is "*" or "all", returns the full article content.',
    parameters: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Semantic query to search for in the article. The AI will find semantically relevant text chunks. Use "*" or "all" to get the full article content.',
        },
        contextChars: {
          type: 'number',
          description: 'Number of characters of context to include before and after each match. Default: 800',
        },
        maxMatches: {
          type: 'number',
          description: 'Optional cap on number of snippets returned',
        },
      },
      required: ['query'],
    },
    execute: async ({ query, contextChars = 800, maxMatches }: {
      query: string;
      contextChars?: number;
      maxMatches?: number;
    }) => {
      try {
        if (readGuard.requiredTokens.length > 0) {
          const normalizedQuery = query.toLowerCase().trim();
          
          // Allow wildcard queries only if explicitly requested
          if (query === '*' || normalizedQuery === 'all') {
            return {
              success: false,
              error: `Wildcard queries ("*" or "all") are not allowed. You must search for specific content mentioned in the instruction. Allowed search terms: ${readGuard.requiredTokens.slice(0, 10).join(', ')}${readGuard.requiredTokens.length > 10 ? '...' : ''}`,
            };
          }
          
          // Check if query matches any required token (with flexible matching)
          const allowed = readGuard.requiredTokens.some(token => {
            const normalizedToken = token.toLowerCase();
            // Exact match or substring match
            if (normalizedQuery.includes(normalizedToken) || normalizedToken.includes(normalizedQuery)) {
              return true;
            }
            // For numbered paragraphs, allow variations (e.g., "1.2" matches "1.2 " or "1.2")
            if (/^\d+\.\d+$/.test(normalizedToken)) {
              // Allow "1.2", "1.2 ", "1.2.", etc.
              const numberPattern = normalizedToken.replace(/\./g, '\\.');
              const variations = [
                new RegExp(`^${numberPattern}\\s*`), // "1.2 " or "1.2"
                new RegExp(`${numberPattern}\\s+`), // "1.2 " followed by text
                new RegExp(`\\b${numberPattern}\\b`), // Word boundary match
              ];
              return variations.some(pattern => pattern.test(normalizedQuery));
            }
            // For multi-word tokens, check if key words match
            if (normalizedToken.includes(' ')) {
              const tokenWords = normalizedToken.split(/\s+/).filter(w => w.length >= 4);
              return tokenWords.some(word => normalizedQuery.includes(word));
            }
            return false;
          });
          
          if (!allowed) {
            return {
              success: false,
              error: `readDocument query must include content from the current instruction. Your query "${query}" does not match any content mentioned in the instruction. Allowed search terms: ${readGuard.requiredTokens.slice(0, 10).join(', ')}${readGuard.requiredTokens.length > 10 ? '...' : ''}. Do NOT search for content from article preview or previous steps.`,
            };
          }
        }

        // Get article text first
        const articleTextResult = await Word.run(async (context) => {
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'items');
          await context.sync();
          const startParagraph = paragraphs.items[articleBoundaries.startParagraphIndex];
          const endParagraph = paragraphs.items[articleBoundaries.endParagraphIndex];

          const startRange = startParagraph.getRange('Start');
          const endRange = endParagraph.getRange('End');
          const articleRange = startRange.expandTo(endRange);

          context.load(articleRange, 'text');
          await context.sync();

          return {
            text: articleRange.text || '',
            length: articleRange.text?.length || 0,
          };
        });

        const text = articleTextResult.text;

        // If query is "*" or "all", return full article content
        if (query === '*' || query.toLowerCase() === 'all') {
          console.log(`[readDocument] Returning full article content (${text.length} characters)`);
          readState.hasFreshRead = true;
          readState.lastQuery = query;
          return {
            success: true,
            query,
            content: [{
              matchText: 'FULL ARTICLE CONTENT',
              snippet: text,
              matchStart: 0,
              matchEnd: text.length,
              snippetStart: 0,
              snippetEnd: text.length,
            }],
            totalFound: 1,
            articleLength: text.length,
            fullContent: text,
          };
        }

        // Use AI semantic search instead of regex
        const safeContextChars = Math.max(0, Math.floor(contextChars || 0));
        const safeMaxMatches = typeof maxMatches === 'number' && maxMatches > 0 ? Math.floor(maxMatches) : undefined;
        
        const semanticMatches = await findSemanticMatches(
          query,
          text,
          apiKey,
          model,
          safeContextChars,
          safeMaxMatches
        );

        console.log(`[readDocument] Semantic search for "${query}" found ${semanticMatches.length} match(es)`);

        readState.hasFreshRead = true;
        readState.lastQuery = query;

        return {
          success: true,
          query,
          content: semanticMatches,
          totalFound: semanticMatches.length,
          articleLength: text.length,
          fullContent: undefined,
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
function createScopedEditTools(
  articleBoundaries: ArticleBoundaries,
  articleName: string,
  readState: ScopedReadState
) {
  const normalizeLeadingWhitespace = (value: string) => value.replace(/^\s+/, '');
  const isNumberedParagraphLabel = (value: string) => /^\s*\d+\.\d+\s*$/.test(value);
  const extractNumberedLabel = (value: string) => {
    const match = value.match(/\d+\.\d+/);
    return match ? match[0] : '';
  };
  const paragraphStartsWithLabel = (paragraphText: string, label: string) => {
    const normalizedParagraph = normalizeLeadingWhitespace(paragraphText || '');
    return normalizedParagraph.startsWith(`${label} `) || normalizedParagraph === label;
  };
  const ensureFreshRead = (action: string) => {
    if (!readState.hasFreshRead) {
      throw new Error(
        `Must call readDocument before ${action}. Each step must re-read the article and never reuse prior locations.`
      );
    }
  };

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
          ensureFreshRead('editDocument');
          const warnings: string[] = [];
          const labelOnly = isNumberedParagraphLabel(searchText);
          const label = extractNumberedLabel(searchText);
          const shouldMatchWholeWord = typeof matchWholeWord === 'boolean' ? matchWholeWord : labelOnly;

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
              matchWholeWord: shouldMatchWholeWord,
            });

            context.load(searchResults, 'items');
            await context.sync();

            if (searchResults.items.length === 0) {
              throw new Error(`Text "${searchText}" not found in article`);
            }

            let itemsToReplace = replaceAll ? searchResults.items : [searchResults.items[0]];
            if (label) {
              const filtered: Word.Range[] = [];
              for (const item of searchResults.items) {
                const paragraph = item.paragraphs.getFirst();
                context.load(paragraph, 'text');
                await context.sync();
                if (paragraphStartsWithLabel(paragraph.text, label)) {
                  filtered.push(item);
                }
              }
              if (filtered.length === 0) {
                throw new Error(`Paragraph "${label}" not found at paragraph start`);
              }
              itemsToReplace = replaceAll ? filtered : [filtered[0]];
            }
            let replacementCount = 0;
            const capturedOldTexts: string[] = [];

            for (const item of itemsToReplace) {
              if (labelOnly) {
                const paragraph = item.paragraphs.getFirst();
                context.load(paragraph, 'text');
                await context.sync();
                const paragraphText = paragraph.text || '';
                const actualOldText = paragraphText;
                paragraph.insertText(newText, Word.InsertLocation.replace);
                replacementCount++;
                capturedOldTexts.push(actualOldText);
              } else {
                context.load(item, 'text');
                await context.sync();
                const actualOldText = item.text;
                item.insertText(newText, Word.InsertLocation.replace);
                replacementCount++;
                capturedOldTexts.push(actualOldText);
              }
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

          readState.hasFreshRead = false;

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
          ensureFreshRead('insertText');
          const warnings: string[] = [];
          const labelOnly = searchText ? isNumberedParagraphLabel(searchText) : false;
          const label = searchText ? extractNumberedLabel(searchText) : '';

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
              let searchResults: Word.RangeCollection;
              const searchAttempts: string[] = [];

              if (labelOnly && label) {
                searchAttempts.push(label);
                searchResults = articleRange.search(label, {
                  matchCase: false,
                  matchWholeWord: true,
                });
                context.load(searchResults, 'items');
                await context.sync();

                if (searchResults.items.length > 0) {
                  const filtered: Word.Range[] = [];
                  for (const item of searchResults.items) {
                    const paragraph = item.paragraphs.getFirst();
                    context.load(paragraph, 'text');
                    await context.sync();
                    if (paragraphStartsWithLabel(paragraph.text, label)) {
                      filtered.push(item);
                    }
                  }
                  if (filtered.length === 0) {
                    searchResults = { items: [] } as Word.RangeCollection;
                  } else {
                    (searchResults as any).items = filtered;
                  }
                }
              } else {
                searchAttempts.push(normalizedSearch);
                searchResults = articleRange.search(normalizedSearch, {
                  matchCase: false,
                  matchWholeWord: false,
                });

                context.load(searchResults, 'items');
                await context.sync();

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
              }

              // Fallback: search paragraph by paragraph if range search fails
              if (searchResults.items.length === 0) {
                if (labelOnly) {
                  throw new Error(`Paragraph "${label}" not found at paragraph start`);
                }
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

              if (!labelOnly) {
                const normalizedNeedle = normalizeSearchText(searchText).toLowerCase();
                for (const para of articleParagraphs) {
                  const normalizedHaystack = normalizeSearchText(para.text || '').toLowerCase();
                  if (normalizedHaystack.includes(normalizedNeedle)) {
                    fallbackParagraph = para;
                    break;
                  }
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

          readState.hasFreshRead = false;

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
      description: 'Delete text from the article. Only works within the current article section. Use this to delete paragraphs or text. If searchText is a paragraph number like "1.2", it will delete the entire paragraph starting with that number.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to find and delete from the article. Can be a paragraph number (e.g., "1.2") or text from readDocument results.' },
          deleteAll: { type: 'boolean', description: 'If true, deletes all occurrences. If false, deletes only the first occurrence.' },
          matchCase: { type: 'boolean', description: 'Whether the search should be case-sensitive' },
          matchWholeWord: { type: 'boolean', description: 'Whether to match whole words only' },
        },
        required: ['searchText'],
      },
      execute: async ({ searchText, deleteAll, matchCase, matchWholeWord }: any) => {
        try {
          ensureFreshRead('deleteText');
          const labelOnly = isNumberedParagraphLabel(searchText);
          const label = extractNumberedLabel(searchText);
          const shouldMatchWholeWord = typeof matchWholeWord === 'boolean' ? matchWholeWord : labelOnly;
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
              matchWholeWord: shouldMatchWholeWord,
            });

            context.load(searchResults, 'items');
            await context.sync();

            if (searchResults.items.length === 0) {
              throw new Error(`Text "${searchText}" not found in article`);
            }

            let itemsToDelete = deleteAll ? searchResults.items : [searchResults.items[0]];
            if (label) {
              const filtered: Word.Range[] = [];
              for (const item of searchResults.items) {
                const paragraph = item.paragraphs.getFirst();
                context.load(paragraph, 'text');
                await context.sync();
                if (paragraphStartsWithLabel(paragraph.text, label)) {
                  filtered.push(item);
                }
              }
              if (filtered.length === 0) {
                throw new Error(`Paragraph "${label}" not found at paragraph start`);
              }
              itemsToDelete = deleteAll ? filtered : [filtered[0]];
            }
            let deletionCount = 0;

            for (const item of itemsToDelete) {
              if (labelOnly) {
                const paragraph = item.paragraphs.getFirst();
                context.load(paragraph, 'text');
                await context.sync();
                const deletedText = paragraph.text || '';
                paragraph.delete();
                deletionCount++;

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

                  await renderInlineDiff(changeObj);
                  await changeTracker(changeObj);
                }
              } else {
                context.load(item, 'text');
                await context.sync();
                const deletedText = item.text;

                item.delete();
                deletionCount++;

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

                  await renderInlineDiff(changeObj);
                  await changeTracker(changeObj);
                }
              }
            }

            await context.sync();

            return {
              deleted: deletionCount,
              totalFound: searchResults.items.length,
            };
          });

          readState.hasFreshRead = false;

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
    substituteText: {
      description: 'Delete existing text and insert new text in one operation. Use this when the instruction says "substitute" or "replace". This tool does both delete and insert atomically.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to find and delete. Use matchText from readDocument results.' },
          newText: { type: 'string', description: 'The new text to insert in place of the deleted text. Preserve all newlines and formatting exactly as provided in the instruction.' },
          matchCase: { type: 'boolean', description: 'Whether the search should be case-sensitive' },
          matchWholeWord: { type: 'boolean', description: 'Whether to match whole words only' },
        },
        required: ['searchText', 'newText'],
      },
      execute: async ({ searchText, newText, matchCase, matchWholeWord }: any) => {
        try {
          ensureFreshRead('substituteText');
          const warnings: string[] = [];
          const labelOnly = isNumberedParagraphLabel(searchText);
          const label = extractNumberedLabel(searchText);
          const shouldMatchWholeWord = typeof matchWholeWord === 'boolean' ? matchWholeWord : labelOnly;

          // Delete and insert in one Word.run call
          const result = await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            context.load(paragraphs, 'items');
            await context.sync();
            const startParagraph = paragraphs.items[articleBoundaries.startParagraphIndex];
            const endParagraph = paragraphs.items[articleBoundaries.endParagraphIndex];
            const startRange = startParagraph.getRange('Start');
            const endRange = endParagraph.getRange('End');
            const articleRange = startRange.expandTo(endRange);

            const searchResults = articleRange.search(searchText, {
              matchCase: matchCase || false,
              matchWholeWord: shouldMatchWholeWord,
            });

            context.load(searchResults, 'items');
            await context.sync();

            if (searchResults.items.length === 0) {
              throw new Error(`Text "${searchText}" not found in article`);
            }

            let itemsToDelete = [searchResults.items[0]];
            if (label) {
              const filtered: Word.Range[] = [];
              for (const item of searchResults.items) {
                const paragraph = item.paragraphs.getFirst();
                context.load(paragraph, 'text');
                await context.sync();
                if (paragraphStartsWithLabel(paragraph.text, label)) {
                  filtered.push(item);
                }
              }
              if (filtered.length === 0) {
                throw new Error(`Paragraph "${label}" not found at paragraph start`);
              }
              itemsToDelete = [filtered[0]];
            }

            let deletedText = '';
            let insertParagraph: Word.Paragraph | null = null;

            for (const item of itemsToDelete) {
              if (labelOnly) {
                const paragraph = item.paragraphs.getFirst();
                context.load(paragraph, 'text');
                await context.sync();
                deletedText = paragraph.text || '';
                // Find paragraph index by searching through paragraphs
                let paraIndex = -1;
                for (let i = articleBoundaries.startParagraphIndex; i <= articleBoundaries.endParagraphIndex; i++) {
                  if (paragraphs.items[i] === paragraph) {
                    paraIndex = i;
                    break;
                  }
                }
                paragraph.delete();
                await context.sync();
                // Get the paragraph that's now at this position (or the one before)
                if (paraIndex > articleBoundaries.startParagraphIndex) {
                  insertParagraph = paragraphs.items[paraIndex - 1];
                } else if (paraIndex < articleBoundaries.endParagraphIndex) {
                  insertParagraph = paragraphs.items[paraIndex];
                } else {
                  insertParagraph = startParagraph;
                }
              } else {
                context.load(item, 'text');
                await context.sync();
                deletedText = item.text;
                // Get paragraph for insertion
                insertParagraph = item.paragraphs.getFirst();
                item.delete();
                await context.sync();
              }
            }

            // Insert new text
            if (insertParagraph) {
              context.load(insertParagraph, 'text');
              await context.sync();
              
              if (newText.includes('\n')) {
                // Multi-line: insert as paragraphs
                const textLines = newText.split('\n');
                let lastPara = insertParagraph;
                for (let i = 0; i < textLines.length; i++) {
                  const lineText = textLines[i];
                  if (i === 0 && lineText.trim() === '' && textLines.length > 1) continue;
                  if (i === textLines.length - 1 && lineText.trim() === '' && textLines.length > 1) continue;

                  const newPara = lastPara.insertParagraph(lineText, Word.InsertLocation.after);
                  context.load(newPara, ['style']);
                  await context.sync();
                  if (insertParagraph.style && insertParagraph.style !== 'Normal') {
                    newPara.style = insertParagraph.style;
                    await context.sync();
                  }
                  lastPara = newPara;
                }
              } else {
                // Single line: insert as text after the paragraph
                const paraRange = insertParagraph.getRange('End');
                paraRange.insertText(newText, Word.InsertLocation.after);
              }
              await context.sync();
            }

            return { deletedText };
          });

          readState.hasFreshRead = false;

          if (changeTracker) {
            const changeObj: DocumentChange = {
              type: 'edit',
              description: `Substituted "${result.deletedText}" with "${newText}"`,
              oldText: result.deletedText,
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
              warnings.push(`Substitution succeeded, but failed to render inline diff: ${e instanceof Error ? e.message : String(e)}`);
            }

            try {
              await changeTracker(changeObj);
            } catch (e) {
              warnings.push(`Substitution succeeded, but failed to record change: ${e instanceof Error ? e.message : String(e)}`);
            }
          }

          return {
            success: true,
            message: `Substituted "${searchText}" with new text`,
            ...(warnings.length > 0 ? { warnings } : {}),
          };
        } catch (error) {
          return {
            success: false,
            error: error instanceof Error ? error.message : 'Unknown error substituting text',
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
    // Intentionally do NOT log full article content (too verbose / may contain sensitive text)

    // Create scoped tools that only work within the article
    const readState: ScopedReadState = { hasFreshRead: false };
    const requiredTokens = extractInstructionContext(instruction);
    const readGuard: ScopedReadGuard = { requiredTokens };
    const scopedReadDocument = createScopedReadDocumentTool(articleBoundaries, readState, readGuard, apiKey, model);
    const scopedEditTools = createScopedEditTools(articleBoundaries, articleName, readState);

    // Create a scoped agent with only article content
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

CRITICAL: INSTRUCTION-ONLY SEARCHES - You MUST ONLY search for content explicitly mentioned in the current instruction. Do NOT search for content from article previews, previous steps, or any other context. Every readDocument query MUST include content from the current instruction only.

AVAILABLE TOOLS:
- readDocument: SEARCH tool - Search ARTICLE ${articleName} using AI semantic search. MANDATORY: Call this BEFORE any insert/edit/delete to find the exact location text (use matchText as searchText). CRITICAL: You can ONLY search for content explicitly mentioned in the current instruction. Wildcard queries ("*" or "all") are NOT allowed.
- editDocument: Find and replace text within ARTICLE ${articleName} only. Requires searchText from readDocument results.
- insertText: Insert new text within ARTICLE ${articleName} only. MANDATORY: Requires searchText from readDocument results. If user says "before X", you MUST find X via readDocument first, then use that matchText as searchText with location: "before". IMPORTANT: When extracting the text to insert from the user's instruction, preserve ALL newlines (\\n) and formatting exactly as provided. The text parameter must include newline characters where the user has line breaks.
- deleteText: Delete text from ARTICLE ${articleName} only. Requires searchText from readDocument results. If searchText is a paragraph number like "1.2", it will delete the entire paragraph.
- substituteText: Delete existing text and insert new text in ONE operation. USE THIS when the instruction says "substitute" or "replace". This is the PREFERRED tool for substitution operations as it does both delete and insert atomically. Requires searchText from readDocument results and newText from the instruction.

MANDATORY WORKFLOW - FOLLOW THIS EXACTLY:
1. UNDERSTAND the user's instruction. Extract the specific content mentioned (numbered paragraphs, quoted text, key phrases).
   - If instruction says "Delete paragraph X and substitute", you MUST complete BOTH operations: delete the existing paragraph, then insert the new content.

2. ALWAYS call readDocument with content FROM THE INSTRUCTION ONLY. Try multiple search strategies if the first fails:
   - Strategy 1: If instruction mentions "1.3" or "1.2", search for "1.3 " (with trailing space) or "1.2 " to find the paragraph label
   - Strategy 2: If that fails, search for key words from the substitution text (e.g., if substituting "1.2 except as expressly provided", search for "except as expressly provided" or "Contract Documents")
   - Strategy 3: If that fails, search for the paragraph number without space: "1.3" or "1.2"
   - Strategy 4: If instruction mentions quoted text like "The Construction Manager shall", search for that exact phrase
   - DO NOT search for content not mentioned in the instruction
   - DO NOT use article preview or previous step content
   - Review the readDocument results - you MUST see matches before proceeding
   - DO NOT proceed to insert/edit until you've found the location via readDocument
   - If one search fails, try the next strategy - DO NOT give up after one failed search

3. Once readDocument returns matches:
   - Use the EXACT matchText from the results as searchText for the appropriate tool
   - If multiple matches, use the first one (or the one that makes sense in context)
   - For "Delete and substitute" or "substitute" instructions: USE substituteText tool (it does both delete and insert in one operation)
   - DO NOT just read and stop - you MUST call the appropriate tool (deleteText, insertText, editDocument, or substituteText)
   - If instruction says "substitute", use substituteText tool with searchText from readDocument and newText from instruction
   - Call the appropriate tool immediately after finding matches - DO NOT stop after reading

4. CRITICAL: If readDocument doesn't find the text after trying multiple strategies:
   - Try searching for parts of the paragraph content (e.g., if looking for "1.2", try searching for words that appear after "1.2" in the document)
   - Report what you searched for and that it wasn't found
   - DO NOT default to "beginning" or "end" - that's wrong!
   - DO NOT search for unrelated content - stick to what's in the instruction
   - DO NOT give up - try different search strategies

5. NEVER insert at "beginning" or "end" unless the user explicitly asks for that. If user says "before X", you MUST find X first via readDocument.

6. Use ONE tool call at a time. Wait for the tool result before deciding the next action. NEVER reuse a prior location; always re-read before each edit.
   - Each step is independent - do NOT use locations from previous steps
   - Each edit requires a fresh readDocument call with content from the current instruction
   - Complete ALL operations in the instruction (e.g., both delete AND substitute)

7. For insertText: location "before"/"after"/"inline" requires searchText from readDocument results.

8. COMPLETION: You MUST complete ALL operations mentioned in the instruction. 
   - If the instruction says "substitute" or "replace", you MUST call substituteText tool (NOT just readDocument)
   - If the instruction says "Delete X and substitute Y", you MUST call substituteText tool
   - DO NOT just read and stop - you MUST actually execute the operations
   - After calling readDocument and finding matches, you MUST immediately call the appropriate tool (deleteText, insertText, editDocument, or substituteText)
   - Your job is to EXECUTE the instruction, not just read the document

9. FINAL RESPONSE (KEEP IT MINIMAL):
   - Do NOT paste full article content.
   - Do NOT write a detailed summary or "Changes Made".
   - Respond with a single short sentence, e.g. "Done." or "Proposed changes below."

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
