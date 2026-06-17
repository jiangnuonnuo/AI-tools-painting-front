export interface Response<T> {
    code: string;
    info: string;
    data: T;
}

export interface AiAgentConfigResponseDTO {
    agentId: string;
    agentName: string;
    agentDesc: string;
}

export interface CreateSessionRequestDTO {
    agentId: string;
    userId: string;
}

export interface CreateSessionResponseDTO {
    sessionId: string;
}

export interface ChatRequestDTO {
  agentId: string;
  userId: string;
  sessionId: string;
  message: string;
  customBaseUrl?: string;
  customApiKey?: string;
  customCompletionsPath?: string;
  customModel?: string;
}

export interface ResponseMetadata {
    summary?: string;
    suggestions?: string[];
    nextActions?: string[];
    progress?: number;
    backendContent?: string;
    [key: string]: unknown;
}

export interface PptSlideElement {
    kind: 'text' | 'table' | 'shape' | 'image' | 'icon' | 'divider' | 'bullet';
    content: string;
    x: number;
    y: number;
    w: number;
    h: number;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    fill?: string;
    align?: 'left' | 'center' | 'right';
    rows?: string[][];
    icon?: string;
    radius?: number;
    shadow?: boolean;
    opacity?: number;
    gradient?: string;
    thickness?: number;
    number?: number;
    fontFace?: string;
    lineSpacing?: number;
    letterSpacing?: number;
}

export interface PptSlide {
    slideIndex: number;
    layout?: string;
    elements: PptSlideElement[];
}

export interface PptData {
    title: string;
    slides: PptSlide[];
}

export interface ChatResponseDTO {
    type: string;
    content: unknown;
    metadata?: ResponseMetadata;
}

// ===== v2.0 流式协议类型 =====

/** 流式事件类型（后端可枚举） */
export type StreamEventType =
    | 'process_delta'
    | 'process_result'
    | 'message'
    | 'render_delta'
    | 'render_result'
    | 'error'
    | 'done';

/** 流式阶段 */
export type StreamPhase =
    | 'analyzing'
    | 'generating'
    | 'drawing'
    | 'reviewing'
    | 'done'
    | 'error';

/** render_delta / render_result chunk.type 枚举 */
export type RenderChunkType =
    | 'ppt_slide'
    | 'ppt'
    | 'drawio_node'
    | 'drawio_edge'
    | 'drawio';

/** 过程 chunk type */
export type ProcessChunkType = 'token' | 'status' | 'analysis' | 'draft' | 'user' | 'done' | 'error';

/** 流式顶层事件 */
export interface StreamEventV2 {
    seq: number;
    phase: StreamPhase;
    author: string;
    event: StreamEventType;
    renderable: boolean;
    final: boolean;
    chunk: StreamChunkV2;
}

/** 新版 chunk 联合类型 */
export interface TokenChunkV2 {
    type: 'token';
    content: string;
}

export interface AnalysisChunkV2 {
    type: 'analysis';
    content: {
        theme?: string;
        purpose?: string;
        audience?: string;
        page_count?: number;
        style?: string;
        tone?: string;
        structure?: string;
        sections?: Array<{
            title: string;
            items?: string[];
        }>;
        design_guidance?: Record<string, unknown>;
    };
}

export interface PptSlideChunkV2 {
    type: 'ppt_slide';
    content: PptSlide;
}

export interface PptResultChunkV2 {
    type: 'ppt';
    content: PptData;
    metadata?: ResponseMetadata;
}

export interface DrawioNodeChunkV2 {
    type: 'drawio_node';
    content: {
        id: string;
        label: string;
        xml: string;
    };
}

export interface DrawioEdgeChunkV2 {
    type: 'drawio_edge';
    content: {
        id: string;
        label: string;
        source: string;
        target: string;
        xml: string;
    };
}

export interface DrawioResultChunkV2 {
    type: 'drawio';
    content: string;
    metadata?: ResponseMetadata;
}

export interface ErrorChunkV2 {
    type: 'error';
    content: string;
}

export interface DoneChunkV2 {
    type: 'done';
}

export type StreamChunkV2 =
    | TokenChunkV2
    | AnalysisChunkV2
    | PptSlideChunkV2
    | PptResultChunkV2
    | DrawioNodeChunkV2
    | DrawioEdgeChunkV2
    | DrawioResultChunkV2
    | ErrorChunkV2
    | DoneChunkV2;
// ===== Prompt Agent 类型 =====

export type PromptMode = "generate" | "rewrite" | "partial_rewrite";

export interface PromptRequestDTO {
    agentId?: string;
    userId: string;
    sessionId?: string;
    mode?: PromptMode;
    taskType?: string;
    goal?: string;
    currentPrompt?: string;
    selectedPromptText?: string;
    editInstruction?: string;
    constraints?: string;
    outputFormat?: string;
    customBaseUrl?: string;
    customApiKey?: string;
    customCompletionsPath?: string;
    customModel?: string;
}

export interface PromptResponseDTO {
    sessionId: string;
    type: string;
    content: string;
    metadata?: ResponseMetadata;
}

export interface PromptStreamChunk {
    type: string;
    content?: string | Record<string, unknown>;
    metadata?: ResponseMetadata;
    raw?: string;
}

export interface PromptStreamEvent {
    seq?: number;
    phase: string;
    author?: string;
    event?: string;
    renderable?: boolean;
    final?: boolean;
    chunk: PromptStreamChunk;
}
