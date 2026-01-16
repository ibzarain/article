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
    system: `You are a helpful AI assistant that can edit Word documents. You have access to tools that let you:
- Read text from the document
- Edit/replace text in the document
- Insert new text at various locations
- Delete text from the document
- Format text (bold, italic, colors, etc.)

When a user asks you to edit the document:
1. First, read the relevant parts of the document to understand the current content
2. Then, use the appropriate tools to make the requested changes
3. Be precise with your edits - only change what the user asks for
4. If you need to find specific text, use the readDocument tool first to locate it
5. When replacing text, try to preserve the context and meaning
6. Always confirm what changes you've made

You are working directly with a live Word document, so be careful and precise with your edits.`,
  };
}

/**
 * Generate a response from the agent using OpenAI API directly
 * This works in the browser since ai/openai is not available
 */
export async function generateAgentResponse(agent: any, prompt: string) {
  try {
    // Convert tools to OpenAI format
    // Use zod-to-json-schema conversion helper
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

    // Call OpenAI API directly
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${agent.apiKey}`,
      },
      body: JSON.stringify({
        model: agent.model,
        messages: [
          { role: 'system', content: agent.system },
          { role: 'user', content: prompt },
        ],
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

    // Handle tool calls
    if (message.tool_calls && message.tool_calls.length > 0) {
      const toolResults = [];
      
      for (const toolCall of message.tool_calls) {
        const toolName = toolCall.function.name;
        const tool = agent.tools[toolName];
        
        if (!tool) {
          console.warn(`Tool ${toolName} not found`);
          continue;
        }

        try {
          const args = JSON.parse(toolCall.function.arguments);
          const result = await tool.execute(args);
          toolResults.push({
            tool_call_id: toolCall.id,
            role: 'tool' as const,
            name: toolName,
            content: JSON.stringify(result),
          });
        } catch (error) {
          toolResults.push({
            tool_call_id: toolCall.id,
            role: 'tool' as const,
            name: toolName,
            content: JSON.stringify({ error: error instanceof Error ? error.message : 'Tool execution failed' }),
          });
        }
      }

      // Make a follow-up call with tool results
      const followUpResponse = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${agent.apiKey}`,
        },
        body: JSON.stringify({
          model: agent.model,
          messages: [
            { role: 'system', content: agent.system },
            { role: 'user', content: prompt },
            message,
            ...toolResults,
          ],
        }),
      });

      if (!followUpResponse.ok) {
        throw new Error(`Follow-up API error: ${followUpResponse.status}`);
      }

      const followUpData = await followUpResponse.json();
      return followUpData.choices[0]?.message?.content || 'No response generated.';
    }

    return message.content || 'No response generated.';
  } catch (error) {
    console.error('Agent generate error:', error);
    throw error;
  }
}
