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
import { AiAgentConfigResponseDTO } from '@/types/api';
import { Message, Session } from '@/types/workbench';
import { clearUserInfo, getUserInfo } from '@/utils/cookie';
import { DRAWIO_RESPONSE_TYPE, normalizeChatResponse, USER_RESPONSE_TYPE } from '@/utils/drawio-response';

type ExportFormat = Parameters<DrawIoEmbedRef['exportDiagram']>[0]['format'];

interface ExportData {
  data: string;
  timestamp: number;
}

const INITIAL_AGENT_MESSAGE = '你好！我是你的智能架构助手。请选择一个智能体开始对话。';
const DRAWIO_SESSIONS_STORAGE_KEY = 'drawio_sessions';
const LAST_AGENT_STORAGE_KEY = 'ai_agent_last_agent';

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

  const [imgData, setImgData] = useState<string | null>(null);
  const [currentUser, setCurrentUser] = useState('');
  const [isChatOpen, setIsChatOpen] = useState(true);
  const [messages, setMessages] = useState<Message[]>([createInitialMessage()]);
  const [inputValue, setInputValue] = useState('');
  const [isSending, setIsSending] = useState(false);
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

  const currentSessionRef = useRef(currentSessionId);

  useEffect(() => {
    currentSessionRef.current = currentSessionId;
  }, [currentSessionId]);

  useEffect(() => {
    if (!initialLoadDoneRef.current && isDrawIoReady && currentSessionId && sessions.length > 0) {
      const session = sessions.find((item) => item.id === currentSessionId);

      if (session && session.drawIoXml && drawioRef.current) {
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
      createNewSession(true);
      return;
    }

    try {
      const parsed = JSON.parse(savedSessions) as Session[];
      setSessions(parsed);

      if (parsed.length > 0) {
        const mostRecent = [...parsed].sort((a, b) => b.lastModified - a.lastModified)[0];
        setCurrentSessionId(mostRecent.id);
        setMessages(mostRecent.messages);
        return;
      }

      createNewSession(true);
    } catch (error) {
      console.error('Failed to parse sessions:', error);
      createNewSession(true);
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

      if (drawioRef.current && session.drawIoXml) {
        drawioRef.current.load({ xml: session.drawIoXml });
      } else if (drawioRef.current) {
        drawioRef.current.load({ xml: '' });
      }
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
    const userInfo = getUserInfo();

    if (!userInfo || !userInfo.user) {
      router.push('/login');
      return;
    }

    setCurrentUser(userInfo.user);

    const loadAgents = async () => {
      try {
        const res = await agentApi.queryAiAgentConfigList();
        setAgents(res.data || []);

        if (res.data && res.data.length > 0) {
          const lastAgentId = localStorage.getItem(LAST_AGENT_STORAGE_KEY);

          if (lastAgentId && res.data.find((agent) => agent.agentId === lastAgentId)) {
            setSelectedAgentId(lastAgentId);
          } else {
            setSelectedAgentId(res.data[0].agentId);
          }
        }
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

  const handleAgentChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const newAgentId = event.target.value;
    setSelectedAgentId(newAgentId);
    setSessionId('');
    localStorage.setItem(LAST_AGENT_STORAGE_KEY, newAgentId);
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
      setMessages((prev) => [...prev, userMsg]);

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

        const chatRes = await agentApi.chat({
          agentId: selectedAgentId,
          userId: currentUser,
          sessionId: activeBackendSessionId,
          message: apiContent,
        });

        const { type, content } = normalizeChatResponse(chatRes.data.type, chatRes.data.content);

        if (type === USER_RESPONSE_TYPE) {
          const agentMsg: Message = {
            id: (Date.now() + 1).toString(),
            role: 'agent',
            content,
            timestamp: Date.now(),
          };
          setMessages((prev) => [...prev, agentMsg]);
        } else if (type === DRAWIO_RESPONSE_TYPE) {
          setSessions((prev) =>
            prev.map((session) => {
              if (session.id === currentSessionId) {
                return {
                  ...session,
                  drawIoXml: content,
                  lastModified: Date.now(),
                };
              }

              return session;
            }),
          );

          if (drawioRef.current && currentSessionId === currentSessionRef.current) {
            try {
              drawioRef.current.load({
                xml: content,
              });
            } catch (error) {
              console.error('Failed to load diagram:', error);
            }
          }
        }
      } catch (error) {
        console.error('Chat error:', error);
        const errorMsg: Message = {
          id: (Date.now() + 1).toString(),
          role: 'agent',
          content: error instanceof Error ? `Error: ${error.message}` : '发送失败，请重试。',
          timestamp: Date.now(),
        };
        setMessages((prev) => [...prev, errorMsg]);
      } finally {
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
          drawioRef={drawioRef}
          isDrawIoReady={isDrawIoReady}
          onAutoSave={handleAutoSave}
          onExport={handleExport}
          onLoad={() => setIsDrawIoReady(true)}
        />

        <ChatSidebar
          agents={agents}
          inputValue={inputValue}
          isChatOpen={isChatOpen}
          isSending={isSending}
          messages={messages}
          messagesEndRef={messagesEndRef}
          onAgentChange={handleAgentChange}
          onCloseChat={() => setIsChatOpen(false)}
          onInputChange={setInputValue}
          onKeyDown={handleKeyDown}
          onRestartSession={handleRestartSession}
          onSendMessage={handleSendMessage}
          onToggleHistoryContext={() => setUseHistoryContext((prev) => !prev)}
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
