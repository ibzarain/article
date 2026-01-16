import { Agent } from 'ai';
import {
  readDocumentTool,
  editDocumentTool,
  insertTextTool,
  deleteTextTool,
  formatTextTool,
} from '../tools/wordEditWithTracking';

/**
 * Creates and configures an AI agent with Word document editing tools
 */
export function createWordAgent(
  apiKey: string,
  model: string = 'gpt-4o-mini'
) {
  const agent = new Agent({
    model: {
      provider: 'openai',
      name: model,
      apiKey,
    },
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
  });

  return agent;
}

/**
 * Generate a response from the agent
 */
export async function generateAgentResponse(agent: Agent, prompt: string) {
  try {
    // Try with object first (newer API)
    const result = await agent.generate({ prompt });
    return result.text || 'No response generated.';
  } catch (error) {
    // Fallback to string if object doesn't work
    try {
      const result = await (agent as any).generate(prompt);
      return result.text || result || 'No response generated.';
    } catch (err2) {
      console.error('Agent generate error:', err2);
      throw err2;
    }
  }
}
