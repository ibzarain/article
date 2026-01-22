import {
  readDocumentTool,
  editDocumentTool,
  insertTextTool,
  deleteTextTool,
  formatTextTool,
} from '../tools/wordEditWithTracking';

/**
 * Creates and configures an AI agent configuration with Word document editing tools
 * Uses the AI SDK's generateText with tools
 */
export function createWordAgent(
  apiKey: string,
  model: string = 'gpt-4o-mini'
) {
  // Return an agent configuration object
  return {
    apiKey,
    model,
    tools: {
      readDocument: readDocumentTool,
      editDocument: editDocumentTool,
      insertText: insertTextTool,
      deleteText: deleteTextTool,
      formatText: formatTextTool,
    },
    system: `You are a helpful AI assistant that can edit Word documents. You MUST use the available tools to make changes to the document.

AVAILABLE TOOLS:
- readDocument: Search the document for a query and return snippets around each match. This is your PRIMARY tool for understanding document content. Use it to semantically understand where to make changes.
- editDocument: Find and replace text in the document. Automatically preserves all formatting (font, style, lists).
- insertText: Insert new text at specific locations. Automatically assesses format context and inserts appropriately:
  * If inserting after a list item/bullet point → creates new bullet point with same formatting
  * If inserting in a paragraph context → creates new paragraph with same style
  * If inserting inline in a sentence → inserts inline text (use location: "inline")
- deleteText: Delete text from the document. Assesses context to delete paragraph vs inline text appropriately.
- formatText: Format text (bold, italic, underline, font size, colors, etc.)

CRITICAL WORKFLOW FOR INSERTIONS AND EDITS:
1. ALWAYS use readDocument FIRST when you need to find an insertion or edit point:
   - Use semantic queries (e.g., "Construction Manager shall", "ARTICLE A-1", "Services and Work")
   - Analyze the returned snippets to understand the document structure and context
   - The snippets will show you the actual text in the document, even if it differs slightly from what the user mentioned
   - Extract the EXACT text from the snippets to use for insertion/editing

2. SEMANTIC UNDERSTANDING OVER EXACT MATCHING:
   - Don't rely on exact text matching algorithms
   - Use readDocument to understand the article content semantically
   - The AI should intelligently determine insertion points based on context, not just exact string matching
   - Once you understand the context from snippets, extract the actual matching text and use that

3. When inserting text:
   - FIRST: Use readDocument with a semantic query related to the insertion point (e.g., if user says "before 'The Construction Manager shall'", search for "Construction Manager" or "ARTICLE A-1")
   - THEN: Analyze the snippets to find the right location semantically
   - FINALLY: Extract the exact text from the snippet to use as searchText for insertText
   - IMPORTANT: Use a longer, more unique string from the snippet when possible (e.g., "The Construction Manager shall perform" instead of just "The Construction Manager shall") - this helps Word's search API find the text more reliably
   - If the matchText from readDocument is short, look at the surrounding context in the snippet and use a longer unique phrase that includes the matchText
   - Use location: "after" for new paragraphs/bullet points after existing content
   - Use location: "inline" for inserting text within a sentence
   - Use location: "before" to insert before found text
   - The tool automatically preserves formatting based on context

4. When editing text:
   - FIRST: Use readDocument to find the text semantically
   - THEN: Extract exact text from snippets and use editDocument
   - editDocument automatically preserves all formatting (font, style, lists)

5. When deleting text:
   - FIRST: Use readDocument to locate the text semantically
   - THEN: Extract exact text and use deleteText
   - If deleting entire paragraph, use deleteParagraph: true

6. Format assessment:
   - Is it a list item/bullet point? → Preserve list formatting
   - Is it a paragraph? → Preserve paragraph style
   - Is it inline text in a sentence? → Use "inline" location for insertText

7. ALWAYS use tools to make changes - don't just describe what you would do

EXAMPLES:
- User: "add paragraph before 'The Construction Manager shall'"
  → Step 1: Call readDocument with query: "Construction Manager shall" or "ARTICLE A-1"
  → Step 2: Analyze snippets to find the exact location and text
  → Step 3: Extract exact text from snippet (e.g., "The Construction Manager shall perform")
  → Step 4: Call insertText with location: "before", searchText: "The Construction Manager shall perform", text: "[new paragraph]"

- User: "replace hello with hi"
  → Step 1: Call readDocument with query: "hello"
  → Step 2: Extract exact text from snippets
  → Step 3: Call editDocument with the exact text

Remember: Use AI semantic understanding (via readDocument) to find locations, then use the exact text from snippets for the actual operations. Don't rely on exact string matching - understand the content first!`,
  };
}

/**
 * Generate a response from the agent using OpenAI API directly
 * This works in the browser since ai/openai is not available
 * Handles multiple rounds of tool calls if needed
 */
export async function generateAgentResponse(agent: any, prompt: string) {
  try {
    // Convert tools to OpenAI format
    const convertZodToJsonSchema = (zodSchema: any): any => {
      if (!zodSchema || !zodSchema._def) {
        return { type: 'object', properties: {}, required: [] };
      }

      const def = zodSchema._def;
      if (def.typeName === 'ZodObject' && def.shape) {
        const properties: any = {};
        const required: string[] = [];

        Object.entries(def.shape).forEach(([key, schema]: [string, any]) => {
          if (!schema || !schema._def) return;

          const schemaDef = schema._def;
          let isOptional = false;
          let innerDef = schemaDef;

          // Handle optional
          if (schemaDef.typeName === 'ZodOptional') {
            isOptional = true;
            innerDef = schemaDef.innerType._def;
          }

          // Convert based on type
          if (innerDef.typeName === 'ZodString') {
            properties[key] = { type: 'string' };
            if (innerDef.description) {
              properties[key].description = innerDef.description;
            }
          } else if (innerDef.typeName === 'ZodNumber') {
            properties[key] = { type: 'number' };
            if (innerDef.description) {
              properties[key].description = innerDef.description;
            }
          } else if (innerDef.typeName === 'ZodBoolean') {
            properties[key] = { type: 'boolean' };
            if (innerDef.description) {
              properties[key].description = innerDef.description;
            }
          } else if (innerDef.typeName === 'ZodEnum') {
            properties[key] = {
              type: 'string',
              enum: innerDef.values,
            };
          }

          if (!isOptional && properties[key]) {
            required.push(key);
          }
        });

        return { type: 'object', properties, required };
      }

      return { type: 'object', properties: {}, required: [] };
    };

    const tools = Object.entries(agent.tools).map(([name, tool]: [string, any]) => {
      // Handle both Zod schemas and plain JSON schemas
      let parameters;
      if (tool.parameters && tool.parameters._def) {
        // Zod schema
        parameters = convertZodToJsonSchema(tool.parameters);
      } else if (tool.parameters && typeof tool.parameters === 'object') {
        // Plain JSON schema
        parameters = tool.parameters;
      } else {
        // Fallback
        parameters = { type: 'object', properties: {}, required: [] };
      }

      return {
        type: 'function' as const,
        function: {
          name,
          description: tool.description || '',
          parameters,
        },
      };
    });

    const messages: any[] = [
      { role: 'system', content: agent.system },
      { role: 'user', content: prompt },
    ];

    // Handle multiple rounds of tool calls (up to 10 rounds to prevent infinite loops)
    let maxRounds = 10;
    let currentRound = 0;

    while (currentRound < maxRounds) {
      // Call OpenAI API
      const response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${agent.apiKey}`,
        },
        body: JSON.stringify({
          model: agent.model,
          messages,
          tools: tools.length > 0 ? tools : undefined,
          tool_choice: tools.length > 0 ? 'auto' : undefined,
        }),
      });

      if (!response.ok) {
        const error = await response.json().catch(() => ({ error: 'Unknown error' }));
        throw new Error(error.error?.message || `API error: ${response.status}`);
      }

      const data = await response.json();
      const message = data.choices[0]?.message;

      if (!message) {
        throw new Error('No response from API');
      }

      // Add assistant message to conversation
      messages.push(message);

      // If no tool calls, return the final response
      if (!message.tool_calls || message.tool_calls.length === 0) {
        return message.content || 'No response generated.';
      }

      // Execute all tool calls
      console.log(`Executing ${message.tool_calls.length} tool call(s)...`);
      const toolResults = [];

      for (const toolCall of message.tool_calls) {
        const toolName = toolCall.function.name;
        const tool = agent.tools[toolName];

        if (!tool) {
          console.warn(`Tool ${toolName} not found`);
          toolResults.push({
            tool_call_id: toolCall.id,
            role: 'tool' as const,
            name: toolName,
            content: JSON.stringify({ error: `Tool ${toolName} not found` }),
          });
          continue;
        }

        try {
          console.log(`Executing tool: ${toolName}`, toolCall.function.arguments);
          const args = JSON.parse(toolCall.function.arguments);
          const result = await tool.execute(args);
          console.log(`Tool ${toolName} result:`, result);

          // Format result as string for OpenAI
          const resultContent = typeof result === 'string'
            ? result
            : JSON.stringify(result);

          toolResults.push({
            tool_call_id: toolCall.id,
            role: 'tool' as const,
            name: toolName,
            content: resultContent,
          });
        } catch (error) {
          console.error(`Tool ${toolName} error:`, error);
          toolResults.push({
            tool_call_id: toolCall.id,
            role: 'tool' as const,
            name: toolName,
            content: JSON.stringify({
              error: error instanceof Error ? error.message : 'Tool execution failed',
              details: error instanceof Error ? error.stack : String(error)
            }),
          });
        }
      }

      // Add tool results to messages for next round
      messages.push(...toolResults);
      currentRound++;
    }

    // If we've exhausted rounds, return the last message
    const lastMessage = messages[messages.length - 1];
    return lastMessage?.content || 'Maximum tool call rounds reached.';
  } catch (error) {
    console.error('Agent generate error:', error);
    throw error;
  }
}
