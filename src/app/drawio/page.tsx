'use client';

import { DrawIoEmbedRef, EventAutoSave, EventExport } from 'react-drawio';
import { useCallback, useEffect, useRef, useState } from 'react';
import { useRouter } from 'next/navigation';
import { agentApi } from '@/api/agent';
import { ChatSidebar } from '@/components/workbench/ChatSidebar';
import { DrawCanvasPanel } from '@/components/workbench/DrawCanvasPanel';
import { ExportPreviewModal } from '@/components/workbench/ExportPreviewModal';
import { RenameSessionModal } from '@/components/workbench/RenameSessionModal';
import { SessionsSidebar } from '@/components/workbench/SessionsSidebar';
import { WorkbenchHeader } from '@/components/workbench/WorkbenchHeader';
import { AiAgentConfigResponseDTO, ResponseMetadata } from '@/types/api';
import { Message, Session } from '@/types/workbench';
import { clearUserInfo, getUserInfo } from '@/utils/cookie';
import { extractDrawioXml } from '@/utils/drawio-response';

type ExportFormat = Parameters<DrawIoEmbedRef['exportDiagram']>[0]['format'];

interface ExportData {
  data: string;
  timestamp: number;
}

const INITIAL_AGENT_MESSAGE = '你好！我是你的智能架构助手。请选择一个智能体开始对话。';
const DRAWIO_SESSIONS_STORAGE_KEY = 'drawio_sessions';
const DRAWIO_AGENT_ID = '300000';

const quickActions = [
  {
    label: '绘制h5端登录流程图',
    text: '请帮我绘制一个H5端的登录流程图，包含用户输入手机号、获取验证码、验证登录等步骤。',
  },
  {
    label: '绘制电商购物流程图',
    text: '请帮我绘制一个电商购物流程图，包含商品浏览、加入购物车、下单、支付、发货等环节。',
  },
];

const createInitialMessage = (): Message => ({
  id: Date.now().toString(),
  role: 'agent',
  content: INITIAL_AGENT_MESSAGE,
  timestamp: Date.now(),
});

const stringifyStreamContent = (content: unknown) => {
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

const getChunkText = (chunk: { content?: unknown; raw?: string }) => {
  if (typeof chunk.raw === 'string') {
    return chunk.raw;
  }

  return stringifyStreamContent(chunk.content);
};

const normalizeInlineText = (value: string): string => {
  return value.replace(/```json|```/g, '').replace(/\s+/g, ' ').trim();
};

const createThinkingSummary = (
  chunk: { content?: unknown; raw?: string; metadata?: ResponseMetadata; label?: string },
  fallback: string,
): string => {
  if (typeof chunk.metadata?.summary === 'string' && chunk.metadata.summary) {
    return normalizeInlineText(chunk.metadata.summary).slice(0, 34);
  }

  if (typeof chunk.metadata?.backendContent === 'string' && chunk.metadata.backendContent) {
    return normalizeInlineText(chunk.metadata.backendContent).slice(0, 34);
  }

  if (typeof chunk.label === 'string' && chunk.label) {
    return normalizeInlineText(chunk.label).slice(0, 34);
  }

  if (chunk.content && typeof chunk.content === 'object' && !Array.isArray(chunk.content)) {
    const record = chunk.content as Record<string, unknown>;
    const keys = ['theme', 'purpose', 'title', 'summary', 'structure', 'label', 'message'];

    for (const key of keys) {
      if (typeof record[key] === 'string' && record[key]) {
        return normalizeInlineText(record[key]).slice(0, 34);
      }
    }
  }

  const text = normalizeInlineText(getChunkText(chunk));
  return (text || fallback).slice(0, 34);
};

/**
 * description: Coordinates AI chat, local sessions, and draw.io rendering for the desktop workbench.
 * params:
 * - input: No component props.
 * - output: Renders the authenticated AI + draw.io workbench page.
 */
export default function Home() {
  const router = useRouter();
  const drawioRef = useRef<DrawIoEmbedRef>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const isExportingForChatRef = useRef(false);
  const isAutosaveRef = useRef(false);
  const pendingMessageRef = useRef('');
  const initialLoadDoneRef = useRef(false);
  const streamAbortRef = useRef<AbortController | null>(null);
  const chatResizeStartXRef = useRef(0);
  const chatResizeStartWidthRef = useRef(404);

  const [imgData, setImgData] = useState<string | null>(null);
  const [currentUser, setCurrentUser] = useState('');
  const [isChatOpen, setIsChatOpen] = useState(true);
  const [messages, setMessages] = useState<Message[]>([createInitialMessage()]);
  const [inputValue, setInputValue] = useState('');
  const [isSending, setIsSending] = useState(false);
  const [chatWidth, setChatWidth] = useState(404);
  const [isResizingChat, setIsResizingChat] = useState(false);
  const [useHistoryContext, setUseHistoryContext] = useState(false);
  const [lastExportedData, setLastExportedData] = useState<ExportData | null>(null);
  const [isDrawIoReady, setIsDrawIoReady] = useState(false);
  const [agents, setAgents] = useState<AiAgentConfigResponseDTO[]>([]);
  const [selectedAgentId, setSelectedAgentId] = useState('');
  const [sessionId, setSessionId] = useState('');
  const [isRenameModalOpen, setIsRenameModalOpen] = useState(false);
  const [renamingSessionId, setRenamingSessionId] = useState<string | null>(null);
  const [newSessionTitle, setNewSessionTitle] = useState('');
  const [sessions, setSessions] = useState<Session[]>([]);
  const [currentSessionId, setCurrentSessionId] = useState<string | null>(null);
  const currentSession = sessions.find((session) => session.id === currentSessionId);

  const currentSessionRef = useRef(currentSessionId);

  useEffect(() => {
    currentSessionRef.current = currentSessionId;
  }, [currentSessionId]);

  useEffect(() => {
    if (!initialLoadDoneRef.current && isDrawIoReady && currentSessionId && sessions.length > 0) {
      const session = sessions.find((item) => item.id === currentSessionId);

      if (session?.drawIoXml && drawioRef.current) {
        drawioRef.current.load({ xml: session.drawIoXml });
      }

      initialLoadDoneRef.current = true;
    }
  }, [isDrawIoReady, currentSessionId, sessions]);

  const createNewSession = useCallback((isInitial = false, backendId = '') => {
    const newSession: Session = {
      id: Date.now().toString(),
      backendSessionId: backendId,
      title: 'New Chat',
      messages: [createInitialMessage()],
      drawIoXml: null,
      metadata: undefined,
      lastModified: Date.now(),
    };

    setSessions((prev) => [newSession, ...prev]);
    setCurrentSessionId(newSession.id);
    setMessages(newSession.messages);
    setSessionId(backendId);

    if (!isInitial && drawioRef.current) {
      drawioRef.current.load({ xml: '' });
    }
  }, []);

  useEffect(() => {
    const savedSessions = localStorage.getItem(DRAWIO_SESSIONS_STORAGE_KEY);

    if (!savedSessions) {
      window.setTimeout(() => createNewSession(true), 0);
      return;
    }

    try {
      const parsed = JSON.parse(savedSessions) as Session[];
      window.setTimeout(() => {
        setSessions(parsed);

        if (parsed.length > 0) {
          const mostRecent = [...parsed].sort((a, b) => b.lastModified - a.lastModified)[0];
          setCurrentSessionId(mostRecent.id);
          setMessages(mostRecent.messages);
          return;
        }

        createNewSession(true);
      }, 0);
    } catch (error) {
      console.error('Failed to parse sessions:', error);
      window.setTimeout(() => createNewSession(true), 0);
    }
  }, [createNewSession]);

  useEffect(() => {
    if (sessions.length > 0) {
      try {
        localStorage.setItem(DRAWIO_SESSIONS_STORAGE_KEY, JSON.stringify(sessions));
      } catch (error) {
        console.error('Failed to save sessions to localStorage:', error);
      }
    }
  }, [sessions]);

  useEffect(() => {
    if (!currentSessionId) {
      return;
    }

    window.setTimeout(() => {
      setSessions((prev) =>
        prev.map((session) => {
          if (session.id !== currentSessionId) {
            return session;
          }

          const firstUserMessage = messages.find((message) => message.role === 'user');

          return {
            ...session,
            messages,
            backendSessionId: sessionId,
            title:
              session.title === 'New Chat' && firstUserMessage
                ? firstUserMessage.content.slice(0, 20) || 'New Chat'
                : session.title,
          };
        }),
      );
    }, 0);
  }, [messages, currentSessionId, sessionId]);

  const loadSession = useCallback(
    (targetSessionId: string) => {
      const session = sessions.find((item) => item.id === targetSessionId);

      if (!session) {
        return;
      }

      setCurrentSessionId(targetSessionId);
      setMessages(session.messages);
      setSessionId(session.backendSessionId || '');

      drawioRef.current?.load({ xml: session.drawIoXml || '' });
    },
    [sessions],
  );

  const handleSwitchSession = (targetSessionId: string) => {
    if (targetSessionId === currentSessionId) {
      return;
    }

    loadSession(targetSessionId);
  };

  const handleDeleteSession = (event: React.MouseEvent, sessionIdToDelete: string) => {
    event.stopPropagation();

    const newSessions = sessions.filter((session) => session.id !== sessionIdToDelete);
    setSessions(newSessions);
    localStorage.setItem(DRAWIO_SESSIONS_STORAGE_KEY, JSON.stringify(newSessions));

    if (currentSessionId !== sessionIdToDelete) {
      return;
    }

    if (newSessions.length > 0) {
      loadSession(newSessions[0].id);
      return;
    }

    createNewSession();
  };

  const handleRenameSession = (session: Session) => {
    setRenamingSessionId(session.id);
    setNewSessionTitle(session.title);
    setIsRenameModalOpen(true);
  };

  const handleRenameSave = () => {
    if (!renamingSessionId || !newSessionTitle.trim()) {
      return;
    }

    setSessions((prev) =>
      prev.map((session) =>
        session.id === renamingSessionId
          ? {
              ...session,
              title: newSessionTitle.trim(),
            }
          : session,
      ),
    );
    setIsRenameModalOpen(false);
    setRenamingSessionId(null);
    setNewSessionTitle('');
  };

  const exportDiagram = () => {
    drawioRef.current?.exportDiagram({
      format: 'xmlsvg',
    });
  };

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isChatOpen]);

  useEffect(() => {
    if (!isResizingChat) {
      return;
    }

    const handleMouseMove = (event: MouseEvent) => {
      const nextWidth = chatResizeStartWidthRef.current + chatResizeStartXRef.current - event.clientX;
      setChatWidth(Math.min(760, Math.max(340, nextWidth)));
    };

    const handleMouseUp = () => {
      setIsResizingChat(false);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };

    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);

    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };
  }, [isResizingChat]);

  useEffect(() => {
    const userInfo = getUserInfo();

    if (!userInfo || !userInfo.user) {
      router.push('/login');
      return;
    }

    window.setTimeout(() => setCurrentUser(userInfo.user), 0);

    const loadAgents = async () => {
      try {
        const res = await agentApi.queryAiAgentConfigList();
        setAgents(res.data || []);

        setSelectedAgentId(DRAWIO_AGENT_ID);
      } catch (error) {
        console.error('Failed to load agents:', error);
        setMessages((prev) => [
          ...prev,
          {
            id: Date.now().toString(),
            role: 'agent',
            content: '加载智能体列表失败，请检查后端服务是否启动。',
            timestamp: Date.now(),
          },
        ]);
      }
    };

    loadAgents();
  }, [router]);

  const handleLogout = () => {
    clearUserInfo();
    router.push('/login');
  };

  const handleAgentChange = () => {
    const newAgentId = DRAWIO_AGENT_ID;
    setSelectedAgentId(newAgentId);
    setSessionId('');
  };

  const finalizeNewChat = useCallback(async () => {
    if (!selectedAgentId || !currentUser) {
      return;
    }

    try {
      const res = await agentApi.createSession(selectedAgentId, currentUser);
      createNewSession(false, res.data.sessionId);
      setInputValue('');
    } catch (error) {
      console.error('Failed to create new session:', error);
    }
  }, [selectedAgentId, currentUser, createNewSession]);

  const handleNewChat = async () => {
    await finalizeNewChat();
  };

  const handleRestartSession = async () => {
    if (!selectedAgentId || !currentUser) {
      return;
    }

    if (!currentSessionId) {
      await finalizeNewChat();
      return;
    }

    try {
      const res = await agentApi.createSession(selectedAgentId, currentUser);
      const newBackendId = res.data.sessionId;
      const initialMsg = createInitialMessage();

      setSessionId(newBackendId);
      setMessages([initialMsg]);
      setInputValue('');

      setSessions((prev) =>
        prev.map((session) => {
          if (session.id === currentSessionId) {
            return {
              ...session,
              backendSessionId: newBackendId,
              messages: [initialMsg],
              lastModified: Date.now(),
            };
          }

          return session;
        }),
      );
    } catch (error) {
      console.error('Failed to restart session:', error);
    }
  };

  const performSendMessage = useCallback(
    async (displayContent: string, apiContent: string) => {
      if (!selectedAgentId) {
        setMessages((prev) => [
          ...prev,
          {
            id: Date.now().toString(),
            role: 'agent',
            content: '请先选择一个智能体。',
            timestamp: Date.now(),
          },
        ]);
        setIsSending(false);
        return;
      }

      const userMsg: Message = {
        id: Date.now().toString(),
        role: 'user',
        content: displayContent,
        timestamp: Date.now(),
      };
      const agentMsgId = `${Date.now()}-agent`;
      const agentMsg: Message = {
        id: agentMsgId,
        role: 'agent',
        content: '',
        thinkingOpen: false,
        steps: [],
        timestamp: Date.now(),
      };
      setMessages((prev) => [...prev, userMsg, agentMsg]);

      try {
        let activeBackendSessionId = sessionId;

        if (!activeBackendSessionId) {
          const sessionRes = await agentApi.createSession(selectedAgentId, currentUser);
          activeBackendSessionId = sessionRes.data.sessionId;
          setSessionId(activeBackendSessionId);
        }

        setSessions((prev) =>
          prev.map((session) => {
            if (session.id === currentSessionId) {
              return {
                ...session,
                lastModified: Date.now(),
              };
            }

            return session;
          }),
        );

        const currentSessionIdRef = currentSessionId;
        const accumulatedSteps: NonNullable<Message['steps']> = [];
        let accumulatedContent = '';
        let renderedXml = '';
        const drawioFragments: string[] = [];
        let stepSeq = 0;

        const phaseLabel: Record<string, string> = {
          analyzing: '分析架构',
          drawing: '绘制图表',
          generating: '生成内容',
          reviewing: '检查优化',
          thinking: '思考中',
          done: '完成',
          error: '异常',
        };

        const updateAgentMessage = () => {
          setMessages((prev) =>
            prev.map((message) =>
              message.id === agentMsgId
                ? {
                    ...message,
                    content: accumulatedContent,
                    steps: [...accumulatedSteps],
                    thinkingOpen: message.thinkingOpen ?? false,
                  }
                : message,
            ),
          );
        };

        const appendStep = (
          phase: string,
          author: string | undefined,
          eventName: string,
          content: string,
          summary?: string,
        ) => {
          accumulatedSteps.forEach((step) => {
            if (step.status === 'running') {
              step.status = 'done';
            }
          });

          const phaseText = phaseLabel[phase] || '处理';
          const cleanContent = normalizeInlineText(content);
          accumulatedSteps.push({
            id: `${Date.now()}-${stepSeq}`,
            phase,
            label: author ? `${phaseText} · ${author}` : phaseText,
            summary: summary || cleanContent.slice(0, 34) || phaseText,
            author,
            event: eventName,
            content,
            status: 'done',
            open: false,
          });
          stepSeq += 1;
        };

        const controller = await agentApi.chatStream(
          {
            agentId: selectedAgentId,
            userId: currentUser,
            sessionId: activeBackendSessionId,
            message: apiContent,
          },
          (streamEvent) => {
            const chunk = streamEvent.chunk as {
              type: string;
              content?: unknown;
              metadata?: ResponseMetadata;
              raw?: string;
              xml?: string;
              label?: string;
            };
            const chunkContent = typeof chunk.content === 'object' && chunk.content !== null
              ? chunk.content as { label?: string; xml?: string }
              : null;
            const eventName = streamEvent.event || (chunk.type === 'drawio' ? 'render_result' : chunk.type === 'error' ? 'error' : chunk.type === 'done' ? 'done' : 'process_delta');
            const isRenderable = streamEvent.renderable ?? eventName.startsWith('render_');

            if (eventName === 'render_result' && isRenderable && chunk.type === 'drawio') {
              const xml = extractDrawioXml(chunk.content) || stringifyStreamContent(chunk.content);
              const metadata = chunk.metadata;
              renderedXml = xml;

              if (xml) {
                drawioRef.current?.load({ xml });
                accumulatedContent = metadata?.backendContent || metadata?.summary || '已将生成结果渲染到 draw.io 画布。';
                appendStep(
                  streamEvent.phase,
                  streamEvent.author,
                  eventName,
                  accumulatedContent,
                  createThinkingSummary(chunk, accumulatedContent),
                );
                accumulatedSteps.forEach((step) => { step.status = 'done'; });
                updateAgentMessage();
                setSessions((prev) =>
                  prev.map((session) =>
                    session.id === currentSessionIdRef
                      ? {
                          ...session,
                          drawIoXml: xml,
                          metadata,
                          lastModified: Date.now(),
                        }
                      : session,
                  ),
                );
              }
              return;
            }

            if (eventName === 'render_delta' && isRenderable && (chunk.type === 'drawio_node' || chunk.type === 'drawio_edge')) {
              const text = chunk.label || chunkContent?.label || getChunkText(chunk);
              const fragmentXml = chunk.xml || chunkContent?.xml || '';
              if (fragmentXml) {
                drawioFragments.push(fragmentXml);
                const incrementalXml = `<mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/>${drawioFragments.join('')}</root></mxGraphModel>`;
                drawioRef.current?.load({ xml: incrementalXml });
              }
              const stepText = text ? `收到可渲染片段：${text}` : '收到可渲染片段。';
              appendStep(
                streamEvent.phase,
                streamEvent.author,
                eventName,
                stepText,
                createThinkingSummary(chunk, stepText),
              );
              updateAgentMessage();
              return;
            }

            if (eventName === 'process_delta' || eventName === 'process_result') {
              const text = getChunkText(chunk) || '收到过程事件。';
              appendStep(
                streamEvent.phase,
                streamEvent.author,
                eventName,
                text,
                createThinkingSummary(chunk, phaseLabel[streamEvent.phase] || '处理'),
              );
              updateAgentMessage();
              return;
            }

            if (eventName === 'message' || chunk.type === 'user') {
              const displayContent = getChunkText(chunk) || '需要补充信息。';
              accumulatedContent += (accumulatedContent ? '\n\n' : '') + displayContent;
              appendStep(
                streamEvent.phase,
                streamEvent.author,
                eventName,
                displayContent,
                createThinkingSummary(chunk, displayContent),
              );
              updateAgentMessage();
              return;
            }

            if (eventName === 'error' || chunk.type === 'error') {
              const errorText = `生成失败：${getChunkText(chunk) || '后端返回错误。'}`;
              accumulatedContent += (accumulatedContent ? '\n\n' : '') + errorText;
              appendStep(
                streamEvent.phase,
                streamEvent.author,
                eventName,
                errorText,
                createThinkingSummary(chunk, errorText),
              );
              accumulatedSteps.forEach((step) => { step.status = 'done'; });
              updateAgentMessage();
              return;
            }

            if (eventName === 'done' || chunk.type === 'done') {
              appendStep(
                streamEvent.phase,
                streamEvent.author,
                eventName,
                '后端事件流已完成。',
                '后端事件流已完成',
              );
              accumulatedSteps.forEach((step) => { step.status = 'done'; });
              updateAgentMessage();
              return;
            }

            const fallbackText = getChunkText(chunk) || `收到 ${eventName} 事件。`;
            appendStep(
              streamEvent.phase,
              streamEvent.author,
              eventName,
              fallbackText,
              createThinkingSummary(chunk, fallbackText),
            );
            updateAgentMessage();
          },
          (error) => {
            console.error('Stream error:', error);
            accumulatedContent += (accumulatedContent ? '\n\n' : '') + `连接异常：${error.message}`;
            accumulatedSteps.forEach((step) => { step.status = 'done'; });
            updateAgentMessage();
            setIsSending(false);
          },
          () => {
            accumulatedSteps.forEach((step) => { step.status = 'done'; });
            if (!accumulatedContent && renderedXml) {
              accumulatedContent = '已将生成结果渲染到 draw.io 画布。';
            }
            updateAgentMessage();
            setIsSending(false);
          },
        );

        streamAbortRef.current = controller;
      } catch (error) {
        console.error('Chat error:', error);
        const errorMsg: Message = {
          id: (Date.now() + 1).toString(),
          role: 'agent',
          content: error instanceof Error ? `Error: ${error.message}` : '发送失败，请重试。',
          timestamp: Date.now(),
        };
        setMessages((prev) => [...prev, errorMsg]);
        setIsSending(false);
      }
    },
    [selectedAgentId, sessionId, currentUser, currentSessionId],
  );

  const handleSendMessage = async () => {
    if (!inputValue.trim() || isSending) {
      return;
    }

    const content = inputValue;
    setInputValue('');
    setIsSending(true);

    if (useHistoryContext && drawioRef.current && isDrawIoReady) {
      isExportingForChatRef.current = true;
      pendingMessageRef.current = content;

      try {
        drawioRef.current.exportDiagram({
          format: 'xml' as ExportFormat,
        });
      } catch (error) {
        console.error('Export failed', error);
        await performSendMessage(content, content);
      }
    } else {
      await performSendMessage(content, content);
    }
  };

  useEffect(() => {
    if (!lastExportedData) {
      return;
    }

    if (isExportingForChatRef.current) {
      isExportingForChatRef.current = false;
      const xml = lastExportedData.data;
      const content = pendingMessageRef.current;
      const apiContent = `[Context: Current Draw.io XML]\n\`\`\`xml\n${xml}\n\`\`\`\n\n${content}`;
      performSendMessage(content, apiContent);
      return;
    }

    if (isAutosaveRef.current) {
      isAutosaveRef.current = false;
      const xml = lastExportedData.data;
      setSessions((prev) =>
        prev.map((session) => {
          if (session.id === currentSessionId) {
            return {
              ...session,
              drawIoXml: xml,
            };
          }

          return session;
        }),
      );
      return;
    }

    setImgData(lastExportedData.data);
  }, [lastExportedData, currentSessionId, performSendMessage]);

  const handleKeyDown = (event: React.KeyboardEvent) => {
    if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
      event.preventDefault();
      handleSendMessage();
    }
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

  const handleStartChatResize = (event: React.MouseEvent<HTMLButtonElement>) => {
    chatResizeStartXRef.current = event.clientX;
    chatResizeStartWidthRef.current = chatWidth;
    setIsResizingChat(true);
  };

  const handleAutoSave = (data: EventAutoSave) => {
    if (!currentSessionId || !isDrawIoReady || isExportingForChatRef.current) {
      return;
    }

    if (data.xml) {
      setSessions((prev) =>
        prev.map((session) => {
          if (session.id === currentSessionId) {
            return {
              ...session,
              drawIoXml: data.xml,
            };
          }

          return session;
        }),
      );
      return;
    }

    isAutosaveRef.current = true;
    drawioRef.current?.exportDiagram({
      format: 'xml' as ExportFormat,
    });
  };

  const handleExport = (data: EventExport) => {
    setLastExportedData({
      data: data.data,
      timestamp: Date.now(),
    });
  };

  return (
    <div className="workbench-shell">
      <WorkbenchHeader
        currentUser={currentUser}
        isChatOpen={isChatOpen}
        onExport={exportDiagram}
        onHome={() => router.push('/')}
        onLogout={handleLogout}
        onOpenChat={() => setIsChatOpen(true)}
      />

      <main className="workbench-main">
        <SessionsSidebar
          currentSessionId={currentSessionId}
          onDeleteSession={handleDeleteSession}
          onNewChat={handleNewChat}
          onRenameSession={handleRenameSession}
          onSwitchSession={handleSwitchSession}
          sessions={sessions}
        />

        <DrawCanvasPanel
          drawIoXml={currentSession?.drawIoXml}
          drawioRef={drawioRef}
          isDrawIoReady={isDrawIoReady}
          metadata={currentSession?.metadata}
          onAutoSave={handleAutoSave}
          onExport={handleExport}
          onLoad={() => setIsDrawIoReady(true)}
        />

        <ChatSidebar
          agents={agents.filter((agent) => agent.agentId === DRAWIO_AGENT_ID)}
          chatWidth={chatWidth}
          inputValue={inputValue}
          isChatOpen={isChatOpen}
          isResizing={isResizingChat}
          isSending={isSending}
          messages={messages}
          messagesEndRef={messagesEndRef}
          onAgentChange={handleAgentChange}
          onCloseChat={() => setIsChatOpen(false)}
          onInputChange={setInputValue}
          onKeyDown={handleKeyDown}
          onRestartSession={handleRestartSession}
          onSendMessage={handleSendMessage}
          onStartResize={handleStartChatResize}
          onToggleHistoryContext={() => setUseHistoryContext((prev) => !prev)}
          onToggleThinking={handleToggleThinking}
          quickActions={quickActions}
          selectedAgentId={selectedAgentId}
          useHistoryContext={useHistoryContext}
        />
      </main>

      <ExportPreviewModal imgData={imgData} onClose={() => setImgData(null)} />

      <RenameSessionModal
        isOpen={isRenameModalOpen}
        newSessionTitle={newSessionTitle}
        onClose={() => setIsRenameModalOpen(false)}
        onSave={handleRenameSave}
        onTitleChange={setNewSessionTitle}
      />
    </div>
  );
}
