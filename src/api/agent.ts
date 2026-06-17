import { API_CONFIG } from '@/config/api-config';
import {
    Response,
    AiAgentConfigResponseDTO,
    CreateSessionResponseDTO,
    ChatRequestDTO,
    ChatResponseDTO,
    PptData,
    PptSlide,
    ResponseMetadata,
    StreamEventType,
    StreamPhase,
    PromptRequestDTO,
    PromptResponseDTO,
    PromptStreamEvent
} from '@/types/api';

const handleResponse = async <T>(response: globalThis.Response): Promise<Response<T>> => {
    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
    }
    const data = await response.json();
    if (data.code !== "0000") {
        throw new Error(data.info || 'Unknown API error');
    }
    return data;
};

// Types for streaming drawio events
export interface DrawioNodeChunk {
    type: 'drawio_node';
    id?: string;
    label?: string;
    xml?: string;
    content?: {
        id: string;
        label: string;
        xml: string;
    };
}

export interface DrawioEdgeChunk {
    type: 'drawio_edge';
    id?: string;
    label?: string;
    source?: string;
    target?: string;
    xml?: string;
    content?: {
        id: string;
        label: string;
        source: string;
        target: string;
        xml: string;
    };
}

export interface DrawioDoneChunk {
    type: 'drawio_done';
    content: string;
    metadata?: ResponseMetadata;
}

export interface DrawioLegacyChunk {
    type: 'drawio';
    content: string;
    metadata?: ResponseMetadata;
}

export interface StatusChunk {
    type: 'status';
    content: string;
}

export interface ErrorChunk {
    type: 'error';
    content: string;
}

export interface UserChunk {
    type: 'user';
    content: string;
    metadata?: ResponseMetadata;
}

export interface DoneChunk {
    type: 'done';
}

export interface PptChunk {
    type: 'ppt';
    content: PptData;
    metadata?: ResponseMetadata;
}

// PPT streaming chunks
export interface PptSlideChunk {
    type: 'ppt_slide';
    content?: PptSlide;
    slide?: PptSlide;
    title?: string;
    slideIndex?: number;
}

export interface PptDoneChunk {
    type: 'ppt_done';
    title?: string;
    totalSlides?: number;
}

export interface PptRawChunk {
    type: 'ppt_raw';
    raw: string;
}

export interface TokenChunk {
    type: 'token';
    content: string;
}

export interface AnalysisChunk {
    type: 'analysis';
    content: unknown;
    metadata?: ResponseMetadata;
}

export interface DraftChunk {
    type: 'draft';
    content: unknown;
    metadata?: ResponseMetadata;
}

export type StreamChunk = DrawioNodeChunk | DrawioEdgeChunk | DrawioDoneChunk | DrawioLegacyChunk | StatusChunk | ErrorChunk | UserChunk | DoneChunk | PptChunk | PptSlideChunk | PptDoneChunk | PptRawChunk | TokenChunk | AnalysisChunk | DraftChunk;

export interface StreamEvent {
    seq?: number;
    phase: StreamPhase | 'thinking';
    author?: string;
    event?: StreamEventType;
    renderable?: boolean;
    final?: boolean;
    chunk: StreamChunk;
}

export type StreamEventCallback = (event: StreamEvent) => void;

export const agentApi = {
    /**
     * Query AI Agent Config List
     * Path: /api/v1/query_ai_agent_config_list
     */
    queryAiAgentConfigList: async (): Promise<Response<AiAgentConfigResponseDTO[]>> => {
        const response = await fetch(`${API_CONFIG.BASE_URL}/query_ai_agent_config_list`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
            },
        });
        return handleResponse<AiAgentConfigResponseDTO[]>(response);
    },

    /**
     * Create Session
     * Path: /api/v1/create_session
     */
    createSession: async (agentId: string, userId: string): Promise<Response<CreateSessionResponseDTO>> => {
        const response = await fetch(`${API_CONFIG.BASE_URL}/create_session`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ agentId, userId }),
        });
        return handleResponse<CreateSessionResponseDTO>(response);
    },

    /**
     * Chat (blocking)
     * Path: /api/v1/chat
     */
    chat: async (data: ChatRequestDTO): Promise<Response<ChatResponseDTO>> => {
        const response = await fetch(`${API_CONFIG.BASE_URL}/chat`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(data),
        });
        return handleResponse<ChatResponseDTO>(response);
    },

    /**
     * Chat Stream (SSE-like streaming)
     * Path: /api/v1/chat_stream
     * Receives structured drawio events: nodes, edges, and final XML
     * Uses ReadableStream to process server-sent events line by line
     */
    chatStream: async (
        data: ChatRequestDTO,
        onEvent: StreamEventCallback,
        onError: (error: Error) => void,
        onComplete: () => void
    ): Promise<AbortController> => {
        const controller = new AbortController();

        try {
            const response = await fetch(`${API_CONFIG.BASE_URL}/chat_stream`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(data),
                signal: controller.signal,
            });

            if (!response.ok) {
                const errorText = await response.text();
                onError(new Error(`HTTP error! status: ${response.status}, message: ${errorText}`));
                return controller;
            }

            const reader = response.body!.getReader();
            const decoder = new TextDecoder();
            let buffer = '';

            const processStream = async () => {
                try {
                    while (true) {
                        const { done, value } = await reader.read();
                        if (done) break;

                        buffer += decoder.decode(value, { stream: true });

                        // Process complete lines
                        const lines = buffer.split('\n');
                        buffer = lines.pop() || ''; // Keep incomplete line in buffer

                        for (const line of lines) {
                            const trimmed = line.trim();
                            if (!trimmed) continue;

                            try {
                                const event: StreamEvent = JSON.parse(trimmed);
                                onEvent(event);
                            } catch (parseErr) {
                                // If JSON parse fails, treat as raw status text
                                console.warn('Failed to parse stream event:', trimmed, parseErr);
                            }
                        }
                    }

                    // Process any remaining buffer
                    if (buffer.trim()) {
                        try {
                            const event: StreamEvent = JSON.parse(buffer.trim());
                            onEvent(event);
                        } catch {}
                    }

                    // Only call onComplete if we didn't abort
                    if (!controller.signal.aborted) {
                        onComplete();
                    }
                } catch (err: unknown) {
                    if (err instanceof DOMException && err.name === 'AbortError') {
                        // User cancelled, no error, but call complete to cleanup UI state
                        onComplete();
                        return;
                    }
                    onError(err instanceof Error ? err : new Error(String(err)));
                }
            };

            processStream();
        } catch (err: unknown) {
            if (!(err instanceof DOMException && err.name === 'AbortError')) {
                onError(err instanceof Error ? err : new Error(String(err)));
            }
        }

        return controller;
    },

    /**
     * Prompt Generate (blocking)
     * Path: /api/v1/prompt/generate_prompt
     */
    generatePrompt: async (data: PromptRequestDTO): Promise<Response<PromptResponseDTO>> => {
        const response = await fetch(`${API_CONFIG.BASE_URL}/prompt/generate_prompt`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(data),
        });
        return handleResponse<PromptResponseDTO>(response);
    },

    /**
     * Prompt Generate Stream (NDJSON streaming)
     * Path: /api/v1/prompt/generate_prompt_stream
     */
    generatePromptStream: async (
        data: PromptRequestDTO,
        onEvent: (event: PromptStreamEvent) => void,
        onError: (error: Error) => void,
        onComplete: () => void
    ): Promise<AbortController> => {
        const controller = new AbortController();

        try {
            const response = await fetch(`${API_CONFIG.BASE_URL}/prompt/generate_prompt_stream`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(data),
                signal: controller.signal,
            });

            if (!response.ok) {
                const errorText = await response.text();
                onError(new Error(`HTTP error! status: ${response.status}, message: ${errorText}`));
                return controller;
            }

            const reader = response.body!.getReader();
            const decoder = new TextDecoder();
            let buffer = '';

            const processStream = async () => {
                try {
                    while (true) {
                        const { done, value } = await reader.read();
                        if (done) break;

                        buffer += decoder.decode(value, { stream: true });

                        const lines = buffer.split('\n');
                        buffer = lines.pop() || '';

                        for (const line of lines) {
                            const trimmed = line.trim();
                            if (!trimmed) continue;

                            try {
                                const event: PromptStreamEvent = JSON.parse(trimmed);
                                onEvent(event);
                            } catch {
                                console.warn('Failed to parse prompt stream event:', trimmed);
                            }
                        }
                    }

                    if (buffer.trim()) {
                        try {
                            const event: PromptStreamEvent = JSON.parse(buffer.trim());
                            onEvent(event);
                        } catch {}
                    }

                    if (!controller.signal.aborted) {
                        onComplete();
                    }
                } catch (err: unknown) {
                    if (err instanceof DOMException && err.name === 'AbortError') {
                        onComplete();
                        return;
                    }
                    onError(err instanceof Error ? err : new Error(String(err)));
                }
            };

            processStream();
        } catch (err: unknown) {
            if (!(err instanceof DOMException && err.name === 'AbortError')) {
                onError(err instanceof Error ? err : new Error(String(err)));
            }
        }

        return controller;
    }
};



