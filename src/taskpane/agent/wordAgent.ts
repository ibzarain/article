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
 * Generate a response from the agent using generateText
 */
export async function generateAgentResponse(agent: any, prompt: string) {
  try {
    const { generateText } = await import('ai');
    const { createOpenAI } = await import('ai/openai');
    
    const openai = createOpenAI({
      apiKey: agent.apiKey,
    });
    
    const result = await generateText({
      model: openai(agent.model),
      tools: agent.tools,
      system: agent.system,
      prompt,
    });
    
    return result.text || 'No response generated.';
  } catch (error) {
    console.error('Agent generate error:', error);
    throw error;
  }
}
