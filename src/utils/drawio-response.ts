import { ChatResponsePayload } from '@/types/workbench';

export const USER_RESPONSE_TYPE = 'user';
export const DRAWIO_RESPONSE_TYPE = 'drawio';

export const isDrawioXml = (content: string) => {
  const trimmedContent = content.trim();

  return trimmedContent.includes('<mxfile') || trimmedContent.includes('<mxGraphModel');
};

export const extractJsonText = (content: string) => {
  const trimmedContent = content.trim();
  const codeBlockMatch = trimmedContent.match(/```(?:json)?\s*([\s\S]*?)```/i);

  if (codeBlockMatch && codeBlockMatch[1]) {
    return codeBlockMatch[1].trim();
  }

  if (trimmedContent.startsWith('{') && trimmedContent.endsWith('}')) {
    return trimmedContent;
  }

  const firstBraceIndex = trimmedContent.indexOf('{');
  const lastBraceIndex = trimmedContent.lastIndexOf('}');

  if (firstBraceIndex >= 0 && lastBraceIndex > firstBraceIndex) {
    return trimmedContent.substring(firstBraceIndex, lastBraceIndex + 1);
  }

  return '';
};

export const normalizeChatResponse = (type: string, content: string): ChatResponsePayload => {
  const responseType = type || USER_RESPONSE_TYPE;

  if (responseType === DRAWIO_RESPONSE_TYPE) {
    return { type: responseType, content };
  }

  const jsonText = extractJsonText(content);

  if (!jsonText) {
    return { type: responseType, content };
  }

  try {
    const parsed = JSON.parse(jsonText) as Partial<ChatResponsePayload>;

    if (parsed.type === DRAWIO_RESPONSE_TYPE && typeof parsed.content === 'string' && isDrawioXml(parsed.content)) {
      return {
        type: DRAWIO_RESPONSE_TYPE,
        content: parsed.content,
      };
    }
  } catch (error) {
    console.warn('Failed to normalize chat response:', error);
  }

  return { type: responseType, content };
};

