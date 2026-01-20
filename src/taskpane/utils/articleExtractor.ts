/* global Word */

/**
 * Extracts article boundaries from the document using simple text search (Ctrl+F style)
 * Finds the start and end of an article by searching for article headers
 */
export interface ArticleBoundaries {
  startIndex: number;
  endIndex: number;
  articleContent: string;
  startParagraphIndex: number;
  endParagraphIndex: number;
}

/**
 * Extracts a specific article from the document
 * @param articleName - The article name to find (e.g., "ARTICLE A-1", "A-1")
 * @returns The article boundaries and content, or null if not found
 */
export async function extractArticle(
  articleName: string
): Promise<ArticleBoundaries | null> {
  try {
    const result = await Word.run(async (context) => {
      // Normalize article name - handle variations like "ARTICLE A-1", "A-1", etc.
      const normalizedName = articleName.toUpperCase().trim();
      let articleId = normalizedName.replace(/^ARTICLE\s*/i, '').trim();
      
      // Build search patterns
      const articlePatterns = [
        `ARTICLE ${articleId}`,
        `ARTICLE ${articleId} –`,
        `ARTICLE ${articleId} -`,
        `ARTICLE ${articleId}:`,
      ];
      
      // Get all paragraphs to find article boundaries
      const paragraphs = context.document.body.paragraphs;
      context.load(paragraphs, 'text');
      await context.sync();
      
      let articleStartIndex = -1;
      let articleEndIndex = -1;
      
      // Find the article start by searching through paragraphs
      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraphText = paragraphs.items[i].text.trim().toUpperCase();
        
        // Check if this paragraph matches any article pattern
        for (const pattern of articlePatterns) {
          if (paragraphText.startsWith(pattern.toUpperCase())) {
            articleStartIndex = i;
            break;
          }
        }
        
        if (articleStartIndex !== -1) {
          break;
        }
      }
      
      if (articleStartIndex === -1) {
        return null;
      }
      
      // Find the article end by looking for the next article header
      // Articles typically follow the pattern: ARTICLE X-Y where X is a letter and Y is a number
      const articleHeaderRegex = /^ARTICLE\s+[A-Z]-\d+/i;
      
      // Start searching from the paragraph after the article start
      for (let i = articleStartIndex + 1; i < paragraphs.items.length; i++) {
        const paragraphText = paragraphs.items[i].text.trim();
        
        // Check if this is another article header
        if (articleHeaderRegex.test(paragraphText)) {
          // Found the next article, so the current article ends before this
          articleEndIndex = i - 1;
          break;
        }
      }
      
      // If no next article found, the current article goes to the end of the document
      if (articleEndIndex === -1) {
        articleEndIndex = paragraphs.items.length - 1;
      }
      
      // Extract the article content by getting ranges
      const startParagraph = paragraphs.items[articleStartIndex];
      const endParagraph = paragraphs.items[articleEndIndex];
      
      const startRange = startParagraph.getRange('Start');
      const endRange = endParagraph.getRange('End');
      const articleRange = startRange.expandTo(endRange);
      
      context.load(articleRange, 'text');
      await context.sync();
      
      return {
        startIndex: articleStartIndex,
        endIndex: articleEndIndex,
        articleContent: articleRange.text,
        startParagraphIndex: articleStartIndex,
        endParagraphIndex: articleEndIndex,
      };
    });
    
    return result;
  } catch (error) {
    console.error('Error extracting article:', error);
    return null;
  }
}

/**
 * Parses the article name from user instruction
 * @param instruction - User instruction text
 * @returns The article name (e.g., "A-1") or null if not found
 */
export function parseArticleName(instruction: string): string | null {
  // Match patterns like "ARTICLE A-1", "ARTICLE A-1 –", "A-1", etc.
  const patterns = [
    /ARTICLE\s+([A-Z]-\d+)/i,
    /^([A-Z]-\d+)/i,
  ];
  
  for (const pattern of patterns) {
    const match = instruction.match(pattern);
    if (match) {
      return match[1].toUpperCase();
    }
  }
  
  return null;
}
