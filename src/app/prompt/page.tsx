'use client';

import {
  Download,
  FileText,
  Loader2,
  PenLine,
  Sparkles,
  Wand2,
} from 'lucide-react';
import { useCallback, useEffect, useRef, useState } from 'react';
import { useRouter } from 'next/navigation';
import { agentApi } from '@/api/agent';
import { ChatSidebar } from '@/components/workbench/ChatSidebar';
import { SessionsSidebar } from '@/components/workbench/SessionsSidebar';
import { WorkbenchHeader } from '@/components/workbench/WorkbenchHeader';
import type {
  AiAgentConfigResponseDTO,
  PromptMode,
  PromptRequestDTO,
  PromptStreamEvent,
  ResponseMetadata,
} from '@/types/api';
import type { Message, Session } from '@/types/workbench';
import { clearUserInfo, getUserInfo } from '@/utils/cookie';

const PROMPT_AGENT_ID = '300002';
const PROMPT_SESSIONS_STORAGE_KEY = 'prompt_sessions';
const INITIAL_AGENT_MESSAGE = '你好！我是 Prompt Forge 助手。选择模式后输入需求，我会生成或改写可直接使用的 Prompt。';

type PromptSession = Session & {
  promptText: string;
};

const MODE_OPTIONS: Array<{ value: PromptMode; label: string; desc: string }> = [
  { value: 'generate', label: '从零生成', desc: '根据目标生成完整 Prompt' },
  { value: 'rewrite', label: '整体改写', desc: '基于当前 Prompt 全文改写' },
  { value: 'partial_rewrite', label: '局部精修', desc: '选中片段后定向改写' },
];

const TASK_TYPES = [
  { value: '', label: '不限类型' },
  { value: 'backend', label: '后端开发' },
  { value: 'frontend', label: '前端开发' },
  { value: 'review', label: '代码审查' },
  { value: 'debug', label: '调试修复' },
  { value: 'architect', label: '架构设计' },
  { value: 'doc', label: '文档撰写' },
];

const quickActions = [
  { label: '后端 Prompt', text: '生成一个 DDD 后端功能开发 Prompt，要求包含分层实现、注释规范和测试验证。' },
  { label: '前端 Prompt', text: '生成一个 React 前端页面开发 Prompt，要求包含交互状态、响应式和视觉规范。' },
  { label: '审查 Prompt', text: '生成一个代码审查 Prompt，要求优先输出问题、风险和缺失测试。' },
];

const createInitialMessage = (): Message => ({
  id: `${Date.now()}`,
  role: 'agent',
  content: INITIAL_AGENT_MESSAGE,
  timestamp: Date.now(),
});

const stringifyStreamContent = (content: unknown): string => {
  if (content === undefined || content === null) {
    return '';
  }

  if (typeof content === 'string') {
    return content;
  }

  try {
    return JSON.stringify(content, null, 2);
  } catch {
    return String(content);
  }
};

const normalizeInlineText = (value: string): string => {
  return value.replace(/```json|```/g, '').replace(/\s+/g, ' ').trim();
};

const getChunkText = (chunk: { content?: unknown; raw?: string }): string => {
  if (typeof chunk.raw === 'string') {
    return chunk.raw;
  }

  return stringifyStreamContent(chunk.content);
};

const createThinkingSummary = (
  chunk: { content?: unknown; metadata?: ResponseMetadata; raw?: string },
  fallback: string,
): string => {
  if (typeof chunk.metadata?.summary === 'string' && chunk.metadata.summary) {
    return normalizeInlineText(chunk.metadata.summary).slice(0, 34);
  }

  if (typeof chunk.metadata?.backendContent === 'string' && chunk.metadata.backendContent) {
    return normalizeInlineText(chunk.metadata.backendContent).slice(0, 34);
  }

  if (chunk.content && typeof chunk.content === 'object' && !Array.isArray(chunk.content)) {
    const record = chunk.content as Record<string, unknown>;
    for (const key of ['theme', 'purpose', 'title', 'summary', 'label', 'message']) {
      if (typeof record[key] === 'string' && record[key]) {
        return normalizeInlineText(record[key]).slice(0, 34);
      }
    }
  }

  const text = normalizeInlineText(getChunkText(chunk));
  return (text || fallback).slice(0, 34);
};

/**
 * description: Provides a Prompt Agent workbench with mode-driven API requests, editable Markdown output, and streamed reasoning slices.
 * params:
 * - input: No component props.
 * - output: Renders the authenticated Prompt Forge page.
 */
export default function PromptPage() {
  const router = useRouter();
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const editorRef = useRef<HTMLTextAreaElement>(null);
  const streamAbortRef = useRef<AbortController | null>(null);
  const chatResizeStartXRef = useRef(0);
  const chatResizeStartWidthRef = useRef(404);

  const [agents, setAgents] = useState<AiAgentConfigResponseDTO[]>([]);
  const [backendSessionId, setBackendSessionId] = useState('');
  const [chatWidth, setChatWidth] = useState(404);
  const [constraints, setConstraints] = useState('');
  const [currentSessionId, setCurrentSessionId] = useState<string | null>(null);
  const [currentUser, setCurrentUser] = useState('');
  const [editInstruction, setEditInstruction] = useState('');
  const [goal, setGoal] = useState('');
  const [inputValue, setInputValue] = useState('');
  const [isChatOpen, setIsChatOpen] = useState(true);
  const [isResizingChat, setIsResizingChat] = useState(false);
  const [isSending, setIsSending] = useState(false);
  const [messages, setMessages] = useState<Message[]>([createInitialMessage()]);
  const [mode, setMode] = useState<PromptMode>('generate');
  const [outputFormat, setOutputFormat] = useState('输出完整可复制 Prompt 文本，保持 Markdown 结构。');
  const [promptText, setPromptText] = useState('');
  const [selectedText, setSelectedText] = useState('');
  const [sessions, setSessions] = useState<PromptSession[]>([]);
  const [showForm, setShowForm] = useState(true);
  const [streamPhase, setStreamPhase] = useState('');
  const [streamProgress, setStreamProgress] = useState('');
  const [taskType, setTaskType] = useState('');

  useEffect(() => {
    const userInfo = getUserInfo();
    if (!userInfo?.user) {
      router.push('/login');
      return;
    }

    // eslint-disable-next-line react-hooks/set-state-in-effect
    setCurrentUser(userInfo.user);
  }, [router]);

  useEffect(() => {
    if (!currentUser) {
      return;
    }

    agentApi.queryAiAgentConfigList().then((response) => {
      setAgents(response.data);
    }).catch(console.error);
  }, [currentUser]);

  useEffect(() => {
    if (!currentUser) {
      return;
    }

    const rawSessions = localStorage.getItem(`${PROMPT_SESSIONS_STORAGE_KEY}_${currentUser}`);
    if (!rawSessions) {
      return;
    }

    try {
      const parsedSessions = JSON.parse(rawSessions) as PromptSession[];
      // eslint-disable-next-line react-hooks/set-state-in-effect
      setSessions(parsedSessions);
    } catch {
      localStorage.removeItem(`${PROMPT_SESSIONS_STORAGE_KEY}_${currentUser}`);
    }
  }, [currentUser]);

  useEffect(() => {
    if (!currentUser || sessions.length === 0) {
      return;
    }

    localStorage.setItem(`${PROMPT_SESSIONS_STORAGE_KEY}_${currentUser}`, JSON.stringify(sessions));
  }, [currentUser, sessions]);

  useEffect(() => {
    if (!isResizingChat) {
      return;
    }

    const handleMouseMove = (event: MouseEvent) => {
      const delta = chatResizeStartXRef.current - event.clientX;
      setChatWidth(Math.min(720, Math.max(340, chatResizeStartWidthRef.current + delta)));
    };

    const handleMouseUp = () => {
      setIsResizingChat(false);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);

    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isResizingChat]);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  const persistActiveSession = useCallback((nextValues: Partial<PromptSession>) => {
    if (!currentSessionId) {
      return;
    }

    setSessions((prev) =>
      prev.map((session) =>
        session.id === currentSessionId
          ? { ...session, ...nextValues, lastModified: Date.now() }
          : session,
      ),
    );
  }, [currentSessionId]);

  const createNewSession = useCallback((backendId = '') => {
    const initialMessage = createInitialMessage();
    const nextSession: PromptSession = {
      id: `${Date.now()}`,
      backendSessionId: backendId,
      title: '新 Prompt',
      messages: [initialMessage],
      drawIoXml: null,
      promptText: '',
      lastModified: Date.now(),
    };

    setSessions((prev) => [nextSession, ...prev]);
    setCurrentSessionId(nextSession.id);
    setBackendSessionId(backendId);
    setMessages([initialMessage]);
    setPromptText('');
    setGoal('');
    setEditInstruction('');
    setSelectedText('');
    setShowForm(true);
  }, []);

  const ensureSession = useCallback(async () => {
    if (backendSessionId) {
      return backendSessionId;
    }

    const response = await agentApi.createSession(PROMPT_AGENT_ID, currentUser);
    const nextBackendSessionId = response.data.sessionId;
    setBackendSessionId(nextBackendSessionId);

    if (!currentSessionId) {
      createNewSession(nextBackendSessionId);
    } else {
      persistActiveSession({ backendSessionId: nextBackendSessionId });
    }

    return nextBackendSessionId;
  }, [backendSessionId, createNewSession, currentSessionId, currentUser, persistActiveSession]);

  const handleNewChat = () => {
    streamAbortRef.current?.abort();
    createNewSession();
  };

  const handleSwitchSession = (targetSessionId: string) => {
    const targetSession = sessions.find((session) => session.id === targetSessionId);
    if (!targetSession) {
      return;
    }

    setCurrentSessionId(targetSession.id);
    setBackendSessionId(targetSession.backendSessionId || '');
    setMessages(targetSession.messages);
    setPromptText(targetSession.promptText);
    setShowForm(false);
  };

  const handleDeleteSession = (event: React.MouseEvent, targetSessionId: string) => {
    event.stopPropagation();
    setSessions((prev) => prev.filter((session) => session.id !== targetSessionId));
    if (currentSessionId === targetSessionId) {
      setCurrentSessionId(null);
      setBackendSessionId('');
      setMessages([createInitialMessage()]);
      setPromptText('');
    }
  };

  const handleRenameSession = (session: Session) => {
    const nextTitle = window.prompt('重命名 Prompt 记录', session.title);
    if (!nextTitle?.trim()) {
      return;
    }

    setSessions((prev) =>
      prev.map((item) =>
        item.id === session.id
          ? { ...item, title: nextTitle.trim(), lastModified: Date.now() }
          : item,
      ),
    );
  };

  const handleEditorSelect = () => {
    const editor = editorRef.current;
    if (!editor) {
      return;
    }

    const nextSelectedText = editor.value.substring(editor.selectionStart, editor.selectionEnd);
    setSelectedText(nextSelectedText);
  };

  const handleStartResize = (event: React.MouseEvent<HTMLButtonElement>) => {
    chatResizeStartXRef.current = event.clientX;
    chatResizeStartWidthRef.current = chatWidth;
    setIsResizingChat(true);
  };

  const handleToggleThinking = (messageId: string) => {
    setMessages((prev) =>
      prev.map((message) =>
        message.id === messageId
          ? { ...message, thinkingOpen: !message.thinkingOpen }
          : message,
      ),
    );
  };

  const handleExportMD = () => {
    if (!promptText.trim()) {
      return;
    }

    const blob = new Blob([promptText], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `prompt_${Date.now()}.md`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const handleLogout = () => {
    streamAbortRef.current?.abort();
    clearUserInfo();
    router.push('/login');
  };

  const updateCurrentPrompt = useCallback((nextPrompt: string) => {
    setPromptText(nextPrompt);
    persistActiveSession({ promptText: nextPrompt });
  }, [persistActiveSession]);

  const buildRequestBody = useCallback(async (messageText: string): Promise<PromptRequestDTO | null> => {
    const activeMode = mode;
    const resolvedGoal = activeMode === 'generate' ? (messageText || goal).trim() : goal.trim();
    const resolvedInstruction = activeMode === 'generate' ? editInstruction.trim() : (messageText || editInstruction).trim();
    const editorSelection = editorRef.current
      ? editorRef.current.value.substring(editorRef.current.selectionStart, editorRef.current.selectionEnd)
      : '';
    const resolvedSelectedText = activeMode === 'partial_rewrite'
      ? (editorSelection || selectedText).trim()
      : '';

    if (activeMode === 'generate' && !resolvedGoal) {
      return null;
    }

    if (activeMode === 'rewrite' && (!promptText.trim() || !resolvedInstruction)) {
      return null;
    }

    if (activeMode === 'partial_rewrite' && (!promptText.trim() || !resolvedSelectedText || !resolvedInstruction)) {
      return null;
    }

    if (resolvedSelectedText) {
      setSelectedText(resolvedSelectedText);
    }

    return {
      agentId: PROMPT_AGENT_ID,
      userId: currentUser,
      sessionId: await ensureSession(),
      mode: activeMode,
      taskType: taskType || undefined,
      goal: activeMode === 'generate' ? resolvedGoal : undefined,
      currentPrompt: activeMode === 'rewrite' || activeMode === 'partial_rewrite' ? promptText : undefined,
      selectedPromptText: activeMode === 'partial_rewrite' ? resolvedSelectedText : undefined,
      editInstruction: activeMode === 'rewrite' || activeMode === 'partial_rewrite' ? resolvedInstruction : undefined,
      constraints: constraints || undefined,
      outputFormat: outputFormat || undefined,
    };
  }, [
    constraints,
    currentUser,
    editInstruction,
    ensureSession,
    goal,
    mode,
    outputFormat,
    promptText,
    selectedText,
    taskType,
  ]);

  const handleSendMessage = useCallback(async () => {
    if (isSending || !currentUser) {
      return;
    }

    const messageText = inputValue.trim();
    const requestBody = await buildRequestBody(messageText);
    if (!requestBody) {
      return;
    }

    const userMessage: Message = {
      id: `${Date.now()}`,
      role: 'user',
      content: messageText || requestBody.goal || requestBody.editInstruction || '执行 Prompt 任务',
      timestamp: Date.now(),
    };
    const agentMessageId = `${Date.now()}-agent`;
    const agentMessage: Message = {
      id: agentMessageId,
      role: 'agent',
      content: '',
      thinkingOpen: false,
      steps: [],
      timestamp: Date.now(),
    };

    setInputValue('');
    setIsSending(true);
    setMessages((prev) => [...prev, userMessage, agentMessage]);
    setStreamPhase('');
    setStreamProgress('');

    const accumulatedSteps: NonNullable<Message['steps']> = [];
    let activeStepKey = '';
    let stepSeq = 0;
    let finalPrompt = '';
    let finalMetadata: ResponseMetadata | undefined;

    const phaseLabel: Record<string, string> = {
      analyzing: '分析需求',
      generating: '生成 Prompt',
      reviewing: '检查优化',
      formatting: '格式化输出',
      done: '完成',
      error: '异常',
    };

    const getStepGroupKey = (phase: string, eventName?: string) => {
      if (eventName === 'error' || eventName === 'message') {
        return `${phase}:${eventName}`;
      }

      if (eventName?.startsWith('process_')) {
        return `${phase}:process`;
      }

      return `${phase}:${eventName || 'process'}`;
    };

    const mergeStepContent = (current: string, next: string, eventName?: string) => {
      if (!current) {
        return next;
      }

      const appendInline = eventName === 'process_delta' && next.length <= 80 && !next.includes('\n');
      return `${current}${appendInline ? '' : '\n\n'}${next}`;
    };

    const updateAgentMessage = () => {
      setMessages((prev) =>
        prev.map((message) =>
          message.id === agentMessageId
            ? { ...message, steps: [...accumulatedSteps], thinkingOpen: message.thinkingOpen ?? true }
            : message,
        ),
      );
    };

    const upsertStep = (
      phase: string,
      eventName: string,
      content: string,
      status: 'running' | 'done',
      author?: string,
      metadata?: ResponseMetadata,
    ) => {
      const phaseText = phaseLabel[phase] || '思考中';
      const stepKey = getStepGroupKey(phase, eventName);
      const activeStep = accumulatedSteps[accumulatedSteps.length - 1];

      if (activeStep && activeStepKey === stepKey) {
        activeStep.content = mergeStepContent(activeStep.content, content, eventName);
        activeStep.event = eventName;
        activeStep.status = status;
        return;
      }

      accumulatedSteps.forEach((step) => {
        if (step.status === 'running') {
          step.status = 'done';
        }
      });
      activeStepKey = stepKey;

      accumulatedSteps.push({
        id: `${Date.now()}-${stepSeq}`,
        phase: stepKey,
        label: author ? `${phaseText} · ${author}` : phaseText,
        summary: createThinkingSummary({ content, metadata }, phaseText),
        content,
        author,
        event: eventName,
        status,
        open: false,
      });
      stepSeq += 1;
    };

    try {
      const controller = await agentApi.generatePromptStream(
        requestBody,
        (streamEvent: PromptStreamEvent) => {
          const phase = streamEvent.phase || 'generating';
          const eventName = streamEvent.event || 'process_delta';
          const chunk = streamEvent.chunk || { type: 'token', content: '' };

          if (phase !== 'done' && phase !== 'error') {
            setStreamPhase(phase);
          }

          if (eventName === 'render_result' && chunk.type === 'prompt') {
            finalPrompt = getChunkText(chunk);
            finalMetadata = chunk.metadata;
            const text = chunk.metadata?.backendContent || chunk.metadata?.summary || 'Prompt 已生成，可继续编辑或导出。';
            upsertStep(phase, eventName, text, 'done', streamEvent.author, chunk.metadata);
            accumulatedSteps.forEach((step) => { step.status = 'done'; });
            updateAgentMessage();
            updateCurrentPrompt(finalPrompt);
            setShowForm(false);
            setStreamProgress('Prompt 已生成');
            return;
          }

          if (eventName === 'process_delta' || eventName === 'process_result') {
            const text = getChunkText(chunk) || '收到过程事件。';
            upsertStep(
              phase,
              eventName,
              text,
              eventName === 'process_result' ? 'done' : 'running',
              streamEvent.author,
              chunk.metadata,
            );
            setStreamProgress(normalizeInlineText(text).slice(0, 56));
            updateAgentMessage();
            return;
          }

          if (eventName === 'message') {
            const text = getChunkText(chunk) || '需要补充信息。';
            upsertStep(phase, eventName, text, 'done', streamEvent.author, chunk.metadata);
            updateAgentMessage();
            return;
          }

          if (eventName === 'error') {
            const text = `生成失败：${getChunkText(chunk) || '后端返回错误。'}`;
            upsertStep(phase, eventName, text, 'done', streamEvent.author, chunk.metadata);
            accumulatedSteps.forEach((step) => { step.status = 'done'; });
            updateAgentMessage();
            setStreamPhase('error');
            setStreamProgress('');
            return;
          }

          if (eventName === 'done') {
            accumulatedSteps.forEach((step) => { step.status = 'done'; });
            updateAgentMessage();
            setStreamPhase('done');
            return;
          }

          const text = getChunkText(chunk) || `收到 ${eventName} 事件。`;
          upsertStep(phase, eventName, text, 'done', streamEvent.author, chunk.metadata);
          updateAgentMessage();
        },
        (error) => {
          console.error('Prompt stream error:', error);
          accumulatedSteps.forEach((step) => { step.status = 'done'; });
          setMessages((prev) =>
            prev.map((message) =>
              message.id === agentMessageId
                ? { ...message, content: `连接异常：${error.message}`, steps: [...accumulatedSteps] }
                : message,
            ),
          );
          setIsSending(false);
          setStreamPhase('error');
        },
        () => {
          accumulatedSteps.forEach((step) => { step.status = 'done'; });
          setMessages((prev) =>
            prev.map((message) =>
              message.id === agentMessageId
                ? {
                    ...message,
                    content: finalPrompt ? '已生成 Prompt，可在编辑器中预览、手动修改或导出。' : '处理完成。',
                    steps: [...accumulatedSteps],
                    metadata: finalMetadata,
                    thinkingOpen: message.thinkingOpen ?? true,
                  }
                : message,
            ),
          );
          persistActiveSession({
            messages: [...messages, userMessage, {
              ...agentMessage,
              content: finalPrompt ? '已生成 Prompt，可在编辑器中预览、手动修改或导出。' : '处理完成。',
              steps: [...accumulatedSteps],
              metadata: finalMetadata,
            }],
            promptText: finalPrompt || promptText,
            metadata: finalMetadata,
            title: userMessage.content.slice(0, 24) || 'Prompt 任务',
          });
          setIsSending(false);
          setStreamPhase('done');
        },
      );

      streamAbortRef.current = controller;
    } catch (error) {
      console.error('Prompt send failed:', error);
      setIsSending(false);
      setStreamPhase('error');
    }
  }, [
    buildRequestBody,
    currentUser,
    inputValue,
    isSending,
    messages,
    persistActiveSession,
    promptText,
    updateCurrentPrompt,
  ]);

  const handleKeyDown = (event: React.KeyboardEvent) => {
    if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
      event.preventDefault();
      handleSendMessage();
    }
  };

  return (
    <div className="workbench-shell">
      <WorkbenchHeader
        currentUser={currentUser}
        isChatOpen={isChatOpen}
        onExport={handleExportMD}
        onHome={() => router.push('/')}
        onLogout={handleLogout}
        onOpenChat={() => setIsChatOpen(true)}
      />

      <main className="workbench-main prompt-workbench-main">
        <SessionsSidebar
          currentSessionId={currentSessionId}
          onDeleteSession={handleDeleteSession}
          onNewChat={handleNewChat}
          onRenameSession={handleRenameSession}
          onSwitchSession={handleSwitchSession}
          sessions={sessions}
        />

        <section className="prompt-center">
          <div className="prompt-mode-bar">
            <div className="prompt-mode-label">Prompt Agent · 300002</div>
            <div className="prompt-mode-tabs">
              {MODE_OPTIONS.map((option) => {
                const isActive = mode === option.value;
                return (
                  <button
                    className={`prompt-mode-tab ${isActive ? 'prompt-mode-tab-active' : ''}`}
                    key={option.value}
                    onClick={() => {
                      setMode(option.value);
                      setShowForm(true);
                    }}
                    type="button"
                  >
                    <Wand2 className="h-4 w-4" />
                    <span>
                      <span className="prompt-mode-tab-title">{option.label}</span>
                      <span className="prompt-mode-tab-desc">{option.desc}</span>
                    </span>
                  </button>
                );
              })}
            </div>
          </div>

          {showForm ? (
            <div className="prompt-form-panel">
              {mode === 'generate' ? (
                <div className="prompt-form-row">
                  <div className="prompt-form-field">
                    <label className="prompt-form-label">生成目标</label>
                    <textarea
                      className="prompt-form-textarea"
                      onChange={(event) => setGoal(event.target.value)}
                      placeholder="生成一个 DDD 后端功能开发 Prompt，要求包含分层架构、注释规范和测试验证。"
                      rows={3}
                      value={goal}
                    />
                  </div>
                </div>
              ) : (
                <div className="prompt-form-row">
                  <div className="prompt-form-field">
                    <label className="prompt-form-label">修改要求</label>
                    <textarea
                      className="prompt-form-textarea"
                      onChange={(event) => setEditInstruction(event.target.value)}
                      placeholder="强化实现计划、文件清单、验证命令和风险说明。"
                      rows={3}
                      value={editInstruction}
                    />
                  </div>
                  {mode === 'partial_rewrite' && (
                    <div className="prompt-form-field">
                      <label className="prompt-form-label">选中片段</label>
                      <div className="prompt-selected-preview">
                        {selectedText || '在编辑器中选中需要局部改写的文本'}
                      </div>
                    </div>
                  )}
                </div>
              )}

              <div className="prompt-form-row prompt-form-row-split">
                <div className="prompt-form-field">
                  <label className="prompt-form-label">任务类型</label>
                  <select className="prompt-form-select" onChange={(event) => setTaskType(event.target.value)} value={taskType}>
                    {TASK_TYPES.map((item) => (
                      <option key={item.value} value={item.value}>{item.label}</option>
                    ))}
                  </select>
                </div>
                <div className="prompt-form-field">
                  <label className="prompt-form-label">输出格式</label>
                  <input
                    className="prompt-form-input"
                    onChange={(event) => setOutputFormat(event.target.value)}
                    value={outputFormat}
                  />
                </div>
              </div>

              <div className="prompt-form-row">
                <div className="prompt-form-field">
                  <label className="prompt-form-label">附加约束</label>
                  <input
                    className="prompt-form-input"
                    onChange={(event) => setConstraints(event.target.value)}
                    placeholder="例如：注释必须包含 description，输出先计划后实现。"
                    value={constraints}
                  />
                </div>
              </div>

              <div className="prompt-form-actions">
                <button className="prompt-form-submit" disabled={isSending} onClick={handleSendMessage} type="button">
                  {isSending ? <Loader2 className="h-4 w-4 animate-spin" /> : <Sparkles className="h-4 w-4" />}
                  <span>{mode === 'generate' ? '开始生成' : '开始改写'}</span>
                </button>
                <button className="prompt-form-collapse" onClick={() => setShowForm(false)} type="button">收起面板</button>
              </div>
            </div>
          ) : (
            <button className="prompt-form-expand" onClick={() => setShowForm(true)} type="button">
              <PenLine className="h-4 w-4" />
              <span>展开控制面板</span>
            </button>
          )}

          {streamPhase && streamPhase !== 'done' && streamPhase !== 'error' && (
            <div className="prompt-progress">
              <Loader2 className="h-3.5 w-3.5 animate-spin" />
              <span>{streamProgress || '处理中...'}</span>
            </div>
          )}

          <div className="prompt-editor-wrap">
            <div className="prompt-editor-head">
              <div className="flex min-w-0 items-center gap-2">
                <FileText className="h-4 w-4 text-[var(--xerina-gold)]" />
                <span className="prompt-editor-title">Prompt Markdown</span>
                {promptText && (
                  <span className="prompt-editor-stats">
                    {promptText.length} 字符 · {promptText.split('\n').length} 行
                  </span>
                )}
              </div>
              <button
                className="prompt-editor-action"
                disabled={!promptText.trim()}
                onClick={handleExportMD}
                title="导出 Markdown"
                type="button"
              >
                <Download className="h-4 w-4" />
                <span>导出 .md</span>
              </button>
            </div>
            <textarea
              className="prompt-editor"
              onChange={(event) => updateCurrentPrompt(event.target.value)}
              onKeyUp={handleEditorSelect}
              onMouseUp={handleEditorSelect}
              placeholder={
                mode === 'generate'
                  ? '生成后的 Prompt 会显示在这里，也可以直接手动修改。'
                  : mode === 'rewrite'
                    ? '粘贴或编辑当前 Prompt，然后输入整体改写要求。'
                    : '粘贴或编辑当前 Prompt，选中局部文本后输入改写要求。'
              }
              ref={editorRef}
              spellCheck={false}
              value={promptText}
            />
          </div>
        </section>

        <ChatSidebar
          agents={agents.filter((agent) => agent.agentId === PROMPT_AGENT_ID)}
          chatWidth={chatWidth}
          inputValue={inputValue}
          isChatOpen={isChatOpen}
          isResizing={isResizingChat}
          isSending={isSending}
          messages={messages}
          messagesEndRef={messagesEndRef}
          onAgentChange={() => undefined}
          onCloseChat={() => setIsChatOpen(false)}
          onInputChange={setInputValue}
          onKeyDown={handleKeyDown}
          onRestartSession={handleNewChat}
          onSendMessage={handleSendMessage}
          onStartResize={handleStartResize}
          onToggleHistoryContext={() => undefined}
          onToggleThinking={handleToggleThinking}
          quickActions={quickActions}
          selectedAgentId={PROMPT_AGENT_ID}
          useHistoryContext={false}
        />
      </main>
    </div>
  );
}
