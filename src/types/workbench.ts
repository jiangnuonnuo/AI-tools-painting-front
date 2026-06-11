export type MessageRole = 'user' | 'agent';

export interface Message {
  id: string;
  role: MessageRole;
  content: string;
  timestamp: number;
}

export interface Session {
  id: string;
  backendSessionId?: string;
  title: string;
  messages: Message[];
  drawIoXml: string | null;
  lastModified: number;
}

export interface ChatResponsePayload {
  type: string;
  content: string;
}

