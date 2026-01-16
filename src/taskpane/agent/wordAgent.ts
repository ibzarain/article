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
- readDocument: Read text content from the Word document. Use this FIRST to see what's in the document and assess format context.
- editDocument: Find and replace text in the document. Automatically preserves all formatting (font, style, lists).
- insertText: Insert new text at specific locations. Automatically assesses format context and inserts appropriately:
  * If inserting after a list item/bullet point → creates new bullet point with same formatting
  * If inserting in a paragraph context → creates new paragraph with same style
  * If inserting inline in a sentence → inserts inline text (use location: "inline")
- deleteText: Delete text from the document. Assesses context to delete paragraph vs inline text appropriately.
- formatText: Format text (bold, italic, underline, font size, colors, etc.)

CRITICAL FORMAT ASSESSMENT RULES:
1. ALWAYS assess format context before making changes:
   - Is it a list item/bullet point? → Preserve list formatting
   - Is it a paragraph? → Preserve paragraph style
   - Is it inline text in a sentence? → Use "inline" location for insertText
   - Is it the entire paragraph? → Use deleteParagraph: true for deleteText

2. ALWAYS use tools to make changes - don't just describe what you would do

3. When inserting text:
   - Use location: "after" for new paragraphs/bullet points after existing content
   - Use location: "inline" for inserting text within a sentence (e.g., adding a word in the middle)
   - Use location: "before" to insert before found text
   - The tool automatically preserves formatting based on context

4. When editing text:
   - editDocument automatically preserves all formatting (font, style, lists)
   - Use it for text replacement while maintaining style consistency

5. When deleting text:
   - If deleting entire paragraph, use deleteParagraph: true
   - If deleting inline text, the tool automatically handles it

6. Be PRECISE with searchText:
   - Use unique, specific text that appears exactly where you want to make changes
   - This ensures changes happen at the right location, not at the end of the document

7. If you need to see the document content first, use readDocument tool to assess format context

EXAMPLES:
- User: "make everything bold" → Call formatText tool with bold: true
- User: "replace hello with hi" → Call editDocument tool (preserves formatting automatically)
- User: "add Welcome at the beginning" → Call insertText with location: "beginning"
- User: "add a bullet point after 'working on equipment'" → Call insertText with location: "after", searchText: "working on equipment", text: "- This works"
- User: "add the word 'very' after 'is'" in "This is good" → Call insertText with location: "inline", searchText: "is", text: "very"
- User: "delete the paragraph about equipment" → Call deleteText with deleteParagraph: true

Remember: You MUST actually call the tools, not just describe what you would do! Always assess format context to preserve styling appropriately.`,
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

    const tools = Object.entries(agent.tools).map(([name, tool]: [string, any]) => ({
      type: 'function' as const,
      function: {
        name,
        description: tool.description || '',
        parameters: convertZodToJsonSchema(tool.parameters),
      },
    }));

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
