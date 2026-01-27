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
 * Uses AI to semantically find relevant paragraph(s) in the article.
 * Note: This does NOT do exact matching; it ranks paragraphs by relevance.
 */
async function findSemanticMatches(
  query: string,
  paragraphs: string[],
  apiKey: string,
  model: string,
  contextChars: number,
  maxMatches: number
): Promise<Array<{
  matchText: string;
  snippet: string;
  matchStart: number;
  matchEnd: number;
  snippetStart: number;
  snippetEnd: number;
}>> {
  const safeMax = Math.max(1, Math.min(20, Math.floor(maxMatches || 5)));
  const chunkLines = paragraphs
    .map((p, i) => {
      const preview = (p || '').replace(/\s+/g, ' ').trim().slice(0, 240);
      return `[${i}] ${preview}`;
    })
    .join('\n');

  const scoringPrompt = `Query: "${query}"

Paragraphs:
${chunkLines}

Return ONLY a JSON array of paragraph indices (0-based) that best match the query, ordered by relevance. Example: [3, 7]`;

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      messages: [
        { role: 'system', content: 'Return only valid JSON: an array of integers.' },
        { role: 'user', content: scoringPrompt },
      ],
      temperature: 0.0,
      max_tokens: 120,
    }),
  });

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Semantic search API error: ${response.status} ${text}`.trim());
  }

  const data = await response.json();
  const raw = (data?.choices?.[0]?.message?.content ?? '').toString().trim();
  const normalized = raw.replace(/```json\s*/i, '').replace(/```/g, '').trim();

  let indices: number[] = [];
  try {
    const parsed = JSON.parse(normalized);
    if (Array.isArray(parsed)) {
      indices = parsed.filter((n) => Number.isFinite(n)).map((n) => Math.floor(n));
    }
  } catch (e) {
    throw new Error(`Semantic search returned non-JSON: ${raw}`);
  }

  const results: Array<{
    matchText: string;
    snippet: string;
    matchStart: number;
    matchEnd: number;
    snippetStart: number;
    snippetEnd: number;
  }> = [];

  const neighborCount = (() => {
    const c = Math.max(0, Math.floor(contextChars || 0));
    if (c >= 1600) return 3;
    if (c >= 800) return 2;
    return 1;
  })();

  for (const idx of indices.slice(0, safeMax)) {
    if (idx < 0 || idx >= paragraphs.length) continue;
    const parts: string[] = [];
    for (let i = Math.max(0, idx - neighborCount); i <= Math.min(paragraphs.length - 1, idx + neighborCount); i++) {
      const t = paragraphs[i] || '';
      if (t) parts.push(t);
    }
    const cur = paragraphs[idx] || '';
    const snippet = parts.join('\n');

    results.push({
      matchText: cur,
      snippet,
      matchStart: idx,
      matchEnd: idx,
      snippetStart: idx,
      snippetEnd: idx,
    });
  }

  return results;
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

        // Read article paragraphs (NOT regex matching).
        const articleData = await Word.run(async (context) => {
          const paragraphs = context.document.body.paragraphs;
          context.load(paragraphs, 'items');
          await context.sync();

          const startIdx = articleBoundaries.startParagraphIndex;
          const endIdx = articleBoundaries.endParagraphIndex;
          const slice = paragraphs.items.slice(startIdx, endIdx + 1);

          for (const p of slice) {
            context.load(p, 'text');
            const listItem = (p as any).listItemOrNullObject ? (p as any).listItemOrNullObject : (p as any).listItem;
            if (listItem) {
              context.load(listItem, 'listString');
            }
          }
          await context.sync();

          const texts = slice.map((p) => p.text || '');
          const listStrings = slice.map((p) => {
            const listItem = (p as any).listItemOrNullObject ? (p as any).listItemOrNullObject : (p as any).listItem;
            if (listItem && !(listItem as any).isNullObject) {
              return (listItem as any).listString || '';
            }
            return '';
          });

          return { texts, listStrings };
        });

        const paragraphsText: string[] = articleData.texts || [];
        const listStrings: string[] = articleData.listStrings || [];

        // If query is "*" or "all", return full article content
        if (query === '*' || query.toLowerCase() === 'all') {
          const fullText = paragraphsText.join('\n');
          console.log(`[readDocument] Returning full article content (${fullText.length} characters)`);
          readState.hasFreshRead = true;
          readState.lastQuery = query;
          return {
            success: true,
            query,
            content: [{
              matchText: 'FULL ARTICLE CONTENT',
              snippet: fullText,
              matchStart: 0,
              matchEnd: fullText.length,
              snippetStart: 0,
              snippetEnd: fullText.length,
            }],
            totalFound: 1,
            articleLength: fullText.length,
            fullContent: fullText,
          };
        }

        // Special-case: list labels like "1.2" are NOT part of paragraph.text in Word.
        // If the query looks like a list label, try to match listItem.listString first.
        const q = (query || '').trim();
        const looksLikeListLabel = (() => {
          if (!q.includes('.')) return false;
          const parts = q.split('.');
          if (parts.some((p) => p.length === 0)) return false;
          return parts.every((p) => {
            for (let i = 0; i < p.length; i++) {
              const c = p.charCodeAt(i);
              if (c < 48 || c > 57) return false;
            }
            return true;
          });
        })();

        if (looksLikeListLabel) {
          const idx = listStrings.findIndex((s) => (s || '').trim() === q);
          if (idx >= 0) {
            const prev = idx > 0 ? paragraphsText[idx - 1] : '';
            const cur = paragraphsText[idx] || '';
            const next = idx + 1 < paragraphsText.length ? paragraphsText[idx + 1] : '';
            const snippet = [prev, cur, next].filter(Boolean).join('\n');

            readState.hasFreshRead = true;
            readState.lastQuery = query;
            return {
              success: true,
              query,
              content: [{
                matchText: cur,
                snippet,
                matchStart: idx,
                matchEnd: idx,
                snippetStart: idx,
                snippetEnd: idx,
              }],
              totalFound: 1,
              articleLength: paragraphsText.join('\n').length,
              fullContent: undefined,
            };
          }
        }

        // Use AI semantic search on paragraph texts (no regex searching).
        const safeContextChars = Math.max(0, Math.floor(contextChars || 0));
        const safeMaxMatches = typeof maxMatches === 'number' && maxMatches > 0 ? Math.floor(maxMatches) : 5;
        const semanticMatches = await findSemanticMatches(query, paragraphsText, apiKey, model, safeContextChars, safeMaxMatches);
        console.log(`[readDocument] Semantic search for "${query}" found ${semanticMatches.length} match(es)`);

        readState.hasFreshRead = true;
        readState.lastQuery = query;

        const fullLength = paragraphsText.join('\n').length;
        return {
          success: true,
          query,
          content: semanticMatches,
          totalFound: semanticMatches.length,
          articleLength: fullLength,
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
  const ensureFreshRead = (action: string, searchText?: string) => {
    // For numbered paragraphs like "1.2"/"1.3" (Word list numbering),
    // allow direct operations without a prior readDocument.
    if (searchText && isNumberedParagraphLabel(searchText)) {
      return;
    }
    if (!readState.hasFreshRead) {
      throw new Error(
        `Must call readDocument before ${action}. Each step must re-read the article and never reuse prior locations.`
      );
    }
  };

  const stripLeadingLabel = (label: string, value: string) => {
    const trimmed = (value || '').trimStart();
    if (!label) return value;
    if (trimmed.startsWith(label)) {
      // Remove label and following whitespace/punctuation to avoid "1.2 1.2 ..." duplication
      return trimmed.slice(label.length).replace(/^[\s.:;,-]+/, '');
    }
    return value;
  };

  const findParagraphByNumberLabel = async (
    context: Word.RequestContext,
    paragraphs: Word.ParagraphCollection,
    label: string
  ): Promise<{ paragraph: Word.Paragraph; isListItem: boolean; listString?: string } | null> => {
    const startIdx = articleBoundaries.startParagraphIndex;
    const endIdx = articleBoundaries.endParagraphIndex;
    const slice = paragraphs.items.slice(startIdx, endIdx + 1);

    for (const p of slice) {
      context.load(p, 'text');
      const listItem = (p as any).listItemOrNullObject ? (p as any).listItemOrNullObject : (p as any).listItem;
      if (listItem) {
        context.load(listItem, 'listString');
      }
    }
    await context.sync();

    // Prefer Word list numbering label (listString) match
    for (const p of slice) {
      const listItem = (p as any).listItemOrNullObject ? (p as any).listItemOrNullObject : (p as any).listItem;
      const listString = listItem && !(listItem as any).isNullObject ? ((listItem as any).listString || '') : '';
      if ((listString || '').trim() === label) {
        return { paragraph: p, isListItem: true, listString };
      }
    }

    // Fallback: literal label at paragraph start (non-list documents)
    for (const p of slice) {
      const text = p.text || '';
      if (paragraphStartsWithLabel(text, label)) {
        return { paragraph: p, isListItem: false };
      }
    }

    return null;
  };

  return {
    editDocument: {
      description: 'Edit or replace text in the article. REPLACE keeps the same bullet/number: it changes only the wording inside that item (green new, red struck-through old). For numbered sections (e.g., "1.2"), operates on that section until the next sibling (e.g. "1.3"). newText must be the COMPLETE replacement—do not truncate or abbreviate. Only edits within the current article section.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text or paragraph label to find (e.g. "1.2" or "1.3"). For labels like "1.2", replaces that entire section until the next sibling like "1.3" or parent like "2."—all within the same bullet/number.' },
          newText: { type: 'string', description: 'The COMPLETE new text to replace the found text with. Include the full replacement—do not cut off or abbreviate. Preserve line breaks (\\n) if the section has multiple lines.' },
          replaceAll: { type: 'boolean', description: 'If true, replaces all occurrences. If false, replaces only the first occurrence.' },
          matchCase: { type: 'boolean', description: 'Whether the search should be case-sensitive' },
          matchWholeWord: { type: 'boolean', description: 'Whether to match whole words only' },
        },
        required: ['searchText', 'newText'],
      },
      execute: async ({ searchText, newText, matchCase, matchWholeWord }: any) => {
        try {
          ensureFreshRead('editDocument', searchText);
          const warnings: string[] = [];
          const labelOnly = isNumberedParagraphLabel(searchText);
          const label = extractNumberedLabel(searchText);
          const shouldMatchWholeWord = typeof matchWholeWord === 'boolean' ? matchWholeWord : labelOnly;

          /**
           * Helper: Parse a numbered label like "1.2" or "1.2.3" into its numeric parts.
           * Returns array like [1, 2] for "1.2" or [1, 2, 3] for "1.2.3"
           */
          const parseLabelParts = (lbl: string): number[] => {
            return lbl.split('.').map(p => parseInt(p, 10)).filter(n => !isNaN(n));
          };

          /**
           * Helper: Check if a label represents a sibling or ancestor of the target label.
           * For target "1.2": siblings are "1.3", "1.4", etc.; ancestors are "2.", "3.", etc.
           * For target "1.2.3": siblings are "1.2.4", "1.2.5"; ancestors are "1.3", "2.", etc.
           */
          const isSiblingOrAncestor = (targetParts: number[], candidateParts: number[]): boolean => {
            if (candidateParts.length === 0) return false;

            // Check if candidate is at same level or higher (fewer parts = higher level)
            if (candidateParts.length < targetParts.length) {
              // Candidate is at higher level (e.g., target is 1.2.3, candidate is 1.3 or 2.)
              // Check if it's the same parent path or a later one
              for (let i = 0; i < candidateParts.length - 1; i++) {
                if (candidateParts[i] < targetParts[i]) return false;
                if (candidateParts[i] > targetParts[i]) return true;
              }
              // Same parent prefix, check if last part is greater
              const lastIdx = candidateParts.length - 1;
              return candidateParts[lastIdx] > targetParts[lastIdx];
            }

            if (candidateParts.length === targetParts.length) {
              // Same level - check if it's a later sibling
              for (let i = 0; i < candidateParts.length - 1; i++) {
                if (candidateParts[i] !== targetParts[i]) {
                  // Different parent path - only count if greater
                  return candidateParts[i] > targetParts[i];
                }
              }
              // Same parent, compare last part
              return candidateParts[candidateParts.length - 1] > targetParts[targetParts.length - 1];
            }

            // Candidate has more parts (deeper nesting) - not a sibling/ancestor
            // Unless it's under a different parent that comes after
            for (let i = 0; i < targetParts.length; i++) {
              if (candidateParts[i] < targetParts[i]) return false;
              if (candidateParts[i] > targetParts[i]) return true;
            }
            return false;
          };

          // Locate target text but DO NOT mutate the document yet.
          // We render a proposed inline diff (red old + green new) and only finalize on Accept.
          const located = await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            context.load(paragraphs, 'items');
            await context.sync();

            const startParagraph = paragraphs.items[articleBoundaries.startParagraphIndex];
            const endParagraph = paragraphs.items[articleBoundaries.endParagraphIndex];
            const startRange = startParagraph.getRange('Start');
            const endRange = endParagraph.getRange('End');
            const articleRange = startRange.expandTo(endRange);

            if (labelOnly && label) {
              // For numbered list items, locate by listItem.listString when available.
              const startIdx = articleBoundaries.startParagraphIndex;
              const endIdx = articleBoundaries.endParagraphIndex;
              const slice = paragraphs.items.slice(startIdx, endIdx + 1);

              // Load all paragraph data including list info
              for (const p of slice) {
                context.load(p, 'text');
                const listItem = (p as any).listItemOrNullObject ? (p as any).listItemOrNullObject : (p as any).listItem;
                if (listItem) context.load(listItem, 'listString');
              }
              await context.sync();

              // Build array of paragraph info for analysis
              const paragraphInfos: Array<{
                index: number;
                text: string;
                listString: string;
                paragraph: Word.Paragraph;
              }> = [];

              for (let i = startIdx; i <= endIdx; i++) {
                const p = paragraphs.items[i];
                const listItem = (p as any).listItemOrNullObject ? (p as any).listItemOrNullObject : (p as any).listItem;
                const listString = listItem && !(listItem as any).isNullObject ? ((listItem as any).listString || '') : '';
                paragraphInfos.push({
                  index: i,
                  text: p.text || '',
                  listString: listString.trim(),
                  paragraph: p,
                });
              }

              // Find the starting paragraph with the target label
              let targetStartRelIdx: number | null = null;
              let isListItem = false;

              // First try to match by listString
              for (let i = 0; i < paragraphInfos.length; i++) {
                if (paragraphInfos[i].listString === label) {
                  targetStartRelIdx = i;
                  isListItem = true;
                  break;
                }
              }

              // Fallback: match by text starting with label
              if (targetStartRelIdx === null) {
                for (let i = 0; i < paragraphInfos.length; i++) {
                  if (paragraphStartsWithLabel(paragraphInfos[i].text, label)) {
                    targetStartRelIdx = i;
                    isListItem = false;
                    break;
                  }
                }
              }

              if (targetStartRelIdx === null) {
                throw new Error(`Paragraph "${label}" not found in ARTICLE ${articleName}`);
              }

              // Find the end of this section (where the next sibling or ancestor starts)
              const targetParts = parseLabelParts(label);
              let targetEndRelIdx = targetStartRelIdx;

              for (let i = targetStartRelIdx + 1; i < paragraphInfos.length; i++) {
                const info = paragraphInfos[i];
                const candidateLabel = info.listString || '';

                // Check if this paragraph has a numbered label
                if (candidateLabel && /^\d+(\.\d+)*$/.test(candidateLabel)) {
                  const candidateParts = parseLabelParts(candidateLabel);

                  if (isSiblingOrAncestor(targetParts, candidateParts)) {
                    // Found the boundary - don't include this paragraph
                    break;
                  }
                }

                // Also check for inline labels in text (for non-list documents)
                const textLabel = extractNumberedLabel(info.text);
                if (textLabel && paragraphStartsWithLabel(info.text, textLabel)) {
                  const textParts = parseLabelParts(textLabel);
                  if (isSiblingOrAncestor(targetParts, textParts)) {
                    break;
                  }
                }

                // This paragraph is part of the section to replace
                targetEndRelIdx = i;
              }

              // Collect all text from the section
              const sectionTexts: string[] = [];
              for (let i = targetStartRelIdx; i <= targetEndRelIdx; i++) {
                const text = paragraphInfos[i].text.trim();
                if (text) {
                  sectionTexts.push(text);
                }
              }
              const fullSectionText = sectionTexts.join('\n');

              // Normalize the new text (strip leading label if it's a list item)
              const normalizedNewText = isListItem ? stripLeadingLabel(label, newText) : newText;

              console.log(`[editDocument] Replacing section "${label}": paragraphs ${targetStartRelIdx} to ${targetEndRelIdx} (${targetEndRelIdx - targetStartRelIdx + 1} paragraph(s))`);
              console.log(`[editDocument] Old content (${fullSectionText.length} chars): ${fullSectionText.substring(0, 200)}...`);
              console.log(`[editDocument] New content (${normalizedNewText.length} chars): ${normalizedNewText.substring(0, 200)}...`);

              return {
                oldText: fullSectionText,
                newText: normalizedNewText,
                targetParagraphIndex: paragraphInfos[targetStartRelIdx].index,
                targetEndParagraphIndex: paragraphInfos[targetEndRelIdx].index,
                paragraphCount: targetEndRelIdx - targetStartRelIdx + 1,
              };
            }

            // General edit: find within article range
            const searchResults = articleRange.search(searchText, {
              matchCase: matchCase || false,
              matchWholeWord: shouldMatchWholeWord,
            });
            context.load(searchResults, 'items');
            await context.sync();

            if (searchResults.items.length === 0) {
              throw new Error(`Text "${searchText}" not found in article`);
            }

            const target = searchResults.items[0];
            context.load(target, 'text');
            await context.sync();

            return {
              oldText: target.text,
              newText: newText,
              targetParagraphIndex: undefined as number | undefined,
              targetEndParagraphIndex: undefined as number | undefined,
              paragraphCount: 1,
            };
          });

          const changeObj: DocumentChange = {
            type: 'edit',
            description: `Replaced section "${searchText}" (${located.paragraphCount} paragraph(s))`,
            oldText: located.oldText,
            newText: located.newText,
            // IMPORTANT: use oldText as the searchText so the inline diff targets the full existing content.
            searchText: located.oldText,
            id: `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
            timestamp: new Date(),
            applied: false,
            canUndo: true,
            articleName: articleName,
            articleStartParagraphIndex: articleBoundaries.startParagraphIndex,
            articleEndParagraphIndex: articleBoundaries.endParagraphIndex,
            ...(typeof located.targetParagraphIndex === 'number'
              ? { targetParagraphIndex: located.targetParagraphIndex }
              : {}),
            ...(typeof located.targetEndParagraphIndex === 'number'
              ? { targetEndParagraphIndex: located.targetEndParagraphIndex }
              : {}),
          };

          try {
            await renderInlineDiff(changeObj);
          } catch (e) {
            warnings.push(
              `Failed to render inline diff: ${e instanceof Error ? e.message : String(e)}`
            );
          }

          if (changeTracker) {
            try {
              await changeTracker(changeObj);
            } catch (e) {
              warnings.push(
                `Failed to record change: ${e instanceof Error ? e.message : String(e)}`
              );
            }
          }

          readState.hasFreshRead = false;

          return {
            success: true,
            replaced: located.paragraphCount,
            totalFound: 1,
            message: `Proposed replacement for section "${searchText}" (${located.paragraphCount} paragraph(s))`,
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
      description: 'Delete text from the article. For numbered sections (e.g., "1.2"), deletes the ENTIRE section including all content until the next sibling section. Only works within the current article section.',
      parameters: {
        type: 'object',
        properties: {
          searchText: { type: 'string', description: 'The text to find and delete from the article. For numbered paragraphs like "1.2", deletes the entire section (all content until the next sibling like "1.3" or parent like "2.")' },
          deleteAll: { type: 'boolean', description: 'If true, deletes all occurrences. If false, deletes only the first occurrence.' },
          matchCase: { type: 'boolean', description: 'Whether the search should be case-sensitive' },
          matchWholeWord: { type: 'boolean', description: 'Whether to match whole words only' },
        },
        required: ['searchText'],
      },
      execute: async ({ searchText, matchCase, matchWholeWord }: any) => {
        try {
          ensureFreshRead('deleteText', searchText);
          const labelOnly = isNumberedParagraphLabel(searchText);
          const label = extractNumberedLabel(searchText);
          const shouldMatchWholeWord = typeof matchWholeWord === 'boolean' ? matchWholeWord : labelOnly;

          /**
           * Helper: Parse a numbered label like "1.2" or "1.2.3" into its numeric parts.
           * Returns array like [1, 2] for "1.2" or [1, 2, 3] for "1.2.3"
           */
          const parseLabelParts = (lbl: string): number[] => {
            return lbl.split('.').map(p => parseInt(p, 10)).filter(n => !isNaN(n));
          };

          /**
           * Helper: Check if a label represents a sibling or ancestor of the target label.
           * For target "1.2": siblings are "1.3", "1.4", etc.; ancestors are "2.", "3.", etc.
           * For target "1.2.3": siblings are "1.2.4", "1.2.5"; ancestors are "1.3", "2.", etc.
           */
          const isSiblingOrAncestor = (targetParts: number[], candidateParts: number[]): boolean => {
            if (candidateParts.length === 0) return false;

            // Check if candidate is at same level or higher (fewer parts = higher level)
            if (candidateParts.length < targetParts.length) {
              // Candidate is at higher level (e.g., target is 1.2.3, candidate is 1.3 or 2.)
              // Check if it's the same parent path or a later one
              for (let i = 0; i < candidateParts.length - 1; i++) {
                if (candidateParts[i] < targetParts[i]) return false;
                if (candidateParts[i] > targetParts[i]) return true;
              }
              // Same parent prefix, check if last part is greater
              const lastIdx = candidateParts.length - 1;
              return candidateParts[lastIdx] > targetParts[lastIdx];
            }

            if (candidateParts.length === targetParts.length) {
              // Same level - check if it's a later sibling
              for (let i = 0; i < candidateParts.length - 1; i++) {
                if (candidateParts[i] !== targetParts[i]) {
                  // Different parent path - only count if greater
                  return candidateParts[i] > targetParts[i];
                }
              }
              // Same parent, compare last part
              return candidateParts[candidateParts.length - 1] > targetParts[targetParts.length - 1];
            }

            // Candidate has more parts (deeper nesting) - not a sibling/ancestor
            // Unless it's under a different parent that comes after
            for (let i = 0; i < targetParts.length; i++) {
              if (candidateParts[i] < targetParts[i]) return false;
              if (candidateParts[i] > targetParts[i]) return true;
            }
            return false;
          };

          // Locate target text but DO NOT delete yet. We only mark it red/strikethrough as a proposal.
          const located = await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            context.load(paragraphs, 'items');
            await context.sync();

            const startParagraph = paragraphs.items[articleBoundaries.startParagraphIndex];
            const endParagraph = paragraphs.items[articleBoundaries.endParagraphIndex];
            const startRange = startParagraph.getRange('Start');
            const endRange = endParagraph.getRange('End');
            const articleRange = startRange.expandTo(endRange);

            if (labelOnly && label) {
              const startIdx = articleBoundaries.startParagraphIndex;
              const endIdx = articleBoundaries.endParagraphIndex;
              const slice = paragraphs.items.slice(startIdx, endIdx + 1);

              // Load all paragraph data including list info
              for (const p of slice) {
                context.load(p, 'text');
                const listItem = (p as any).listItemOrNullObject ? (p as any).listItemOrNullObject : (p as any).listItem;
                if (listItem) context.load(listItem, 'listString');
              }
              await context.sync();

              // Build array of paragraph info for analysis
              const paragraphInfos: Array<{
                index: number;
                text: string;
                listString: string;
                paragraph: Word.Paragraph;
              }> = [];

              for (let i = startIdx; i <= endIdx; i++) {
                const p = paragraphs.items[i];
                const listItem = (p as any).listItemOrNullObject ? (p as any).listItemOrNullObject : (p as any).listItem;
                const listString = listItem && !(listItem as any).isNullObject ? ((listItem as any).listString || '') : '';
                paragraphInfos.push({
                  index: i,
                  text: p.text || '',
                  listString: listString.trim(),
                  paragraph: p,
                });
              }

              // Find the starting paragraph with the target label
              let targetStartRelIdx: number | null = null;

              // First try to match by listString
              for (let i = 0; i < paragraphInfos.length; i++) {
                if (paragraphInfos[i].listString === label) {
                  targetStartRelIdx = i;
                  break;
                }
              }

              // Fallback: match by text starting with label
              if (targetStartRelIdx === null) {
                for (let i = 0; i < paragraphInfos.length; i++) {
                  if (paragraphStartsWithLabel(paragraphInfos[i].text, label)) {
                    targetStartRelIdx = i;
                    break;
                  }
                }
              }

              if (targetStartRelIdx === null) {
                throw new Error(`Paragraph "${label}" not found in ARTICLE ${articleName}`);
              }

              // Find the end of this section (where the next sibling or ancestor starts)
              const targetParts = parseLabelParts(label);
              let targetEndRelIdx = targetStartRelIdx;

              for (let i = targetStartRelIdx + 1; i < paragraphInfos.length; i++) {
                const info = paragraphInfos[i];
                const candidateLabel = info.listString || '';

                // Check if this paragraph has a numbered label
                if (candidateLabel && /^\d+(\.\d+)*$/.test(candidateLabel)) {
                  const candidateParts = parseLabelParts(candidateLabel);

                  if (isSiblingOrAncestor(targetParts, candidateParts)) {
                    // Found the boundary - don't include this paragraph
                    break;
                  }
                }

                // Also check for inline labels in text (for non-list documents)
                const textLabel = extractNumberedLabel(info.text);
                if (textLabel && paragraphStartsWithLabel(info.text, textLabel)) {
                  const textParts = parseLabelParts(textLabel);
                  if (isSiblingOrAncestor(targetParts, textParts)) {
                    break;
                  }
                }

                // This paragraph is part of the section to delete
                targetEndRelIdx = i;
              }

              // Collect all text from the section
              const sectionTexts: string[] = [];
              for (let i = targetStartRelIdx; i <= targetEndRelIdx; i++) {
                const text = paragraphInfos[i].text.trim();
                if (text) {
                  sectionTexts.push(text);
                }
              }
              const fullSectionText = sectionTexts.join('\n');

              console.log(`[deleteText] Deleting section "${label}": paragraphs ${targetStartRelIdx} to ${targetEndRelIdx} (${targetEndRelIdx - targetStartRelIdx + 1} paragraph(s))`);
              console.log(`[deleteText] Section content (${fullSectionText.length} chars): ${fullSectionText.substring(0, 200)}...`);

              return {
                oldText: fullSectionText,
                targetParagraphIndex: paragraphInfos[targetStartRelIdx].index,
                targetEndParagraphIndex: paragraphInfos[targetEndRelIdx].index,
                paragraphCount: targetEndRelIdx - targetStartRelIdx + 1,
              };
            }

            // Non-label search: just find the specific text
            const searchResults = articleRange.search(searchText, {
              matchCase: matchCase || false,
              matchWholeWord: shouldMatchWholeWord,
            });
            context.load(searchResults, 'items');
            await context.sync();
            if (searchResults.items.length === 0) {
              throw new Error(`Text "${searchText}" not found in article`);
            }
            const target = searchResults.items[0];
            context.load(target, 'text');
            await context.sync();

            return {
              oldText: target.text,
              targetParagraphIndex: undefined as number | undefined,
              targetEndParagraphIndex: undefined as number | undefined,
              paragraphCount: 1,
            };
          });

          const changeObj: DocumentChange = {
            type: 'delete',
            description: `Deleted "${searchText}" section (${located.paragraphCount} paragraph(s))`,
            oldText: located.oldText,
            // IMPORTANT: use oldText as searchText so the inline diff marks the actual content.
            searchText: located.oldText,
            id: `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
            timestamp: new Date(),
            applied: false,
            canUndo: true,
            articleName: articleName,
            articleStartParagraphIndex: articleBoundaries.startParagraphIndex,
            articleEndParagraphIndex: articleBoundaries.endParagraphIndex,
            ...(typeof located.targetParagraphIndex === 'number'
              ? { targetParagraphIndex: located.targetParagraphIndex }
              : {}),
            ...(typeof located.targetEndParagraphIndex === 'number'
              ? { targetEndParagraphIndex: located.targetEndParagraphIndex }
              : {}),
          };

          try {
            await renderInlineDiff(changeObj);
          } catch (e) {
            // best-effort
          }
          if (changeTracker) {
            try {
              await changeTracker(changeObj);
            } catch {
              // best-effort
            }
          }

          readState.hasFreshRead = false;

          return {
            success: true,
            deleted: located.paragraphCount,
            totalFound: 1,
            message: `Proposed deletion for section "${searchText}" (${located.paragraphCount} paragraph(s))`,
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

CRITICAL: FULL REPLACEMENT TEXT - For editDocument, newText MUST contain the COMPLETE replacement. Do NOT truncate, abbreviate, or cut off the text. Include every word and line the user specified for the replacement.

CRITICAL: INSTRUCTION-ONLY SEARCHES - You MUST ONLY search for content explicitly mentioned in the current instruction. Do NOT search for content from article previews, previous steps, or any other context. Every readDocument query MUST include content from the current instruction only.

AVAILABLE TOOLS:
- readDocument: SEARCH tool - Search ARTICLE ${articleName} for a query and return contextual snippets around each match. MANDATORY: Call this BEFORE any insert/edit/delete to find the exact location text (use matchText as searchText). CRITICAL: You can ONLY search for content explicitly mentioned in the current instruction. Wildcard queries ("*" or "all") are NOT allowed.
- editDocument: Find and replace text within ARTICLE ${articleName} only. REPLACE keeps the same bullet/number—only the wording changes (green new, red old). For "1.2" or "1.3", set searchText to that label and set newText to the COMPLETE replacement; do not truncate newText. Requires searchText from readDocument results.
- insertText: Insert new text within ARTICLE ${articleName} only. MANDATORY: Requires searchText from readDocument results. If user says "before X", you MUST find X via readDocument first, then use that matchText as searchText with location: "before". IMPORTANT: When extracting the text to insert from the user's instruction, preserve ALL newlines (\\n) and formatting exactly as provided. The text parameter must include newline characters where the user has line breaks.
- deleteText: Delete text from ARTICLE ${articleName} only. For numbered sections like "1.2", this deletes the ENTIRE section (all paragraphs from "1.2" until the next sibling "1.3" or parent "2."). Requires searchText from readDocument results.

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
   - Use the EXACT matchText from the results as searchText for insertText/editDocument/deleteText
   - If multiple matches, use the first one (or the one that makes sense in context)
   - For "Delete and substitute"/"Replace"/"Substitute" instructions:
     - Use editDocument (NOT deleteText + insertText). This keeps the same bullet/number and shows green new + red struck-through old in that item.
     - If the instruction references numbered paragraphs (e.g., "1.2", "1.3"), set searchText to that label (e.g., "1.3") and set newText to the COMPLETE replacement content. Do NOT truncate or cut off newText—include the full replacement.
     - editDocument operates on that section (e.g. all content from "1.3" until the next sibling "1.4" or parent "2.") within the same list item.
     - Only use insertText for true additions (e.g., "Add the following paragraph before ...") where no existing text is being replaced.
   - Call the appropriate tool with the matchText as searchText

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

8. COMPLETION: You MUST complete ALL operations mentioned in the instruction. If the instruction says "Delete X and substitute Y", you MUST do both. Do not stop after one operation.

8. FINAL RESPONSE (KEEP IT MINIMAL):
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
