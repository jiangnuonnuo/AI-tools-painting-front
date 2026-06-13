import { PptData, ResponseMetadata } from '@/types/api';
import { ChatResponsePayload } from '@/types/workbench';

export const USER_RESPONSE_TYPE = 'user';
export const DRAWIO_RESPONSE_TYPE = 'drawio';
export const PPT_RESPONSE_TYPE = 'ppt';

export const isDrawioXml = (content: string) => {
  const trimmedContent = content.trim();

  return trimmedContent.includes('<mxfile') || trimmedContent.includes('<mxGraphModel');
};

export const extractDrawioXml = (content: unknown) => {
  if (typeof content !== 'string') {
    return '';
  }

  const trimmedContent = content.trim();
  const xmlCodeBlockMatch = trimmedContent.match(/```(?:xml|drawio)?\s*([\s\S]*?)```/i);
  const codeBlockContent = xmlCodeBlockMatch?.[1]?.trim();

  if (codeBlockContent && isDrawioXml(codeBlockContent)) {
    return codeBlockContent;
  }

  if (isDrawioXml(trimmedContent)) {
    const mxfileStart = trimmedContent.indexOf('<mxfile');
    const mxfileEnd = trimmedContent.lastIndexOf('</mxfile>');

    if (mxfileStart >= 0 && mxfileEnd > mxfileStart) {
      return trimmedContent.substring(mxfileStart, mxfileEnd + '</mxfile>'.length);
    }

    const graphStart = trimmedContent.indexOf('<mxGraphModel');
    const graphEnd = trimmedContent.lastIndexOf('</mxGraphModel>');

    if (graphStart >= 0 && graphEnd > graphStart) {
      return trimmedContent.substring(graphStart, graphEnd + '</mxGraphModel>'.length);
    }
  }

  return '';
};

/**
 * description: Checks whether a value is a renderable PPT payload from the backend.
 * params:
 * - content: Input unknown response content.
 * - output: True when content has a slides array.
 */
export const isPptData = (content: unknown): content is PptData => {
  if (!content || typeof content !== 'object') {
    return false;
  }

  const candidate = content as Partial<PptData>;
  return Array.isArray(candidate.slides);
};

export const extractJsonText = (content: unknown) => {
  if (typeof content !== 'string') {
    return '';
  }

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

/**
 * description: Normalizes backend chat response content for draw.io, PPT, and user text renderers.
 * params:
 * - type: Input backend response type.
 * - content: Input backend response content, string or structured object.
 * - metadata: Input optional response metadata.
 * - output: Stable payload consumed by application pages.
 */
export const normalizeChatResponse = (
  type: string,
  content: unknown,
  metadata?: ResponseMetadata,
): ChatResponsePayload => {
  const responseType = type || USER_RESPONSE_TYPE;
  const drawioXml = extractDrawioXml(content);

  if (responseType === DRAWIO_RESPONSE_TYPE && drawioXml) {
    return { type: DRAWIO_RESPONSE_TYPE, content: drawioXml, metadata };
  }

  if (responseType === PPT_RESPONSE_TYPE && isPptData(content)) {
    return { type: PPT_RESPONSE_TYPE, content, metadata };
  }

  const textContent = typeof content === 'string' ? content : JSON.stringify(content ?? '');
  const jsonText = extractJsonText(content);

  if (drawioXml) {
    return { type: DRAWIO_RESPONSE_TYPE, content: drawioXml, metadata };
  }

  if (!jsonText) {
    return { type: USER_RESPONSE_TYPE, content: textContent, metadata };
  }

  try {
    const parsed = JSON.parse(jsonText) as Partial<ChatResponsePayload> & { metadata?: ResponseMetadata };

    if (parsed.type === DRAWIO_RESPONSE_TYPE && typeof parsed.content === 'string' && isDrawioXml(parsed.content)) {
      return {
        type: DRAWIO_RESPONSE_TYPE,
        content: extractDrawioXml(parsed.content) || parsed.content,
        metadata: parsed.metadata || metadata,
      };
    }

    if (parsed.type === PPT_RESPONSE_TYPE && isPptData(parsed.content)) {
      return {
        type: PPT_RESPONSE_TYPE,
        content: parsed.content,
        metadata: parsed.metadata || metadata,
      };
    }
  } catch (error) {
    console.warn('Failed to normalize chat response:', error);
  }

  return { type: USER_RESPONSE_TYPE, content: textContent, metadata };
};
