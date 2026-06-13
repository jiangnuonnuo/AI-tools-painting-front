import type { PptData, ResponseMetadata } from '@/types/api';

export type MessageRole = 'user' | 'agent';

export interface Message {
  id: string;
  role: MessageRole;
  content: string;
  metadata?: ResponseMetadata;
  reasoning?: string;
  thinkingOpen?: boolean;
  steps?: {
    id: string;
    phase: string;
    label: string;
    summary: string;
    content: string;
    author?: string;
    event?: string;
    open?: boolean;
    status: 'running' | 'done' | 'pending';
  }[];
  timestamp: number;
}

export interface Session {
  id: string;
  backendSessionId?: string;
  title: string;
  messages: Message[];
  drawIoXml: string | null;
  pptData?: PptData | null;
  metadata?: ResponseMetadata;
  lastModified: number;
}

export interface ChatResponsePayload {
  type: 'user' | 'drawio' | 'ppt';
  content: string | PptData;
  metadata?: ResponseMetadata;
}
