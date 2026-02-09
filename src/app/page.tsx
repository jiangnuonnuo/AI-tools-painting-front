'use client';

import { DrawIoEmbed, DrawIoEmbedRef } from 'react-drawio';
import { useRef, useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { getUserInfo, clearUserInfo } from '@/utils/cookie';
import { agentApi } from '@/api/agent';
import { AiAgentConfigResponseDTO } from '@/types/api';

// Message type definition
type Message = {
  id: string;
  role: 'user' | 'agent';
  content: string;
  timestamp: number;
};

// Elegant SVG Icons with consistent styling
const Icons = {
  Chat: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
    </svg>
  ),
  Close: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <line x1="18" y1="6" x2="6" y2="18"></line>
      <line x1="6" y1="6" x2="18" y2="18"></line>
    </svg>
  ),
  Send: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <line x1="22" y1="2" x2="11" y2="13"></line>
      <polygon points="22 2 15 22 11 13 2 9 22 2"></polygon>
    </svg>
  ),
  User: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
      <circle cx="12" cy="7" r="4"></circle>
    </svg>
  ),
  Bot: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M12 2a2 2 0 0 1 2 2v2a2 2 0 0 1-2 2 2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z"></path>
      <path d="M4 11v6a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2z"></path>
      <path d="M9 22v-3"></path>
      <path d="M15 22v-3"></path>
    </svg>
  ),
  Download: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
      <polyline points="7 10 12 15 17 10"></polyline>
      <line x1="12" y1="15" x2="12" y2="3"></line>
    </svg>
  ),
  Sparkles: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z" />
    </svg>
  ),
  Logout: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path>
      <polyline points="16 17 21 12 16 7"></polyline>
      <line x1="21" y1="12" x2="9" y2="12"></line>
    </svg>
  ),
  Layers: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <polygon points="12 2 2 7 12 12 22 7 12 2"></polygon>
      <polyline points="2 17 12 22 22 17"></polyline>
      <polyline points="2 12 12 17 22 12"></polyline>
    </svg>
  ),
  Loader: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={`animate-spin ${className}`}>
      <path d="M21 12a9 9 0 1 1-6.219-8.56"></path>
    </svg>
  ),
  Plus: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <line x1="12" y1="5" x2="12" y2="19"></line>
      <line x1="5" y1="12" x2="19" y2="12"></line>
    </svg>
  )
};

export default function Home() {
  const router = useRouter();
  const [imgData, setImgData] = useState<string | null>(null);
  const drawioRef = useRef<DrawIoEmbedRef>(null);
  
  // User State
  const [currentUser, setCurrentUser] = useState('');

  // Chat State
  const [isChatOpen, setIsChatOpen] = useState(true);
  const [messages, setMessages] = useState<Message[]>([
    {
      id: '1',
      role: 'agent',
      content: '你好！我是你的智能架构助手。请选择一个智能体开始对话。',
      timestamp: Date.now()
    }
  ]);
  const [inputValue, setInputValue] = useState('');
  const [isSending, setIsSending] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  // Context State
  const [useHistoryContext, setUseHistoryContext] = useState(false);
  const [lastExportedData, setLastExportedData] = useState<{data: string, timestamp: number} | null>(null);
  const isExportingForChatRef = useRef(false);
  const pendingMessageRef = useRef('');

  // Agent State
  const [agents, setAgents] = useState<AiAgentConfigResponseDTO[]>([]);
  const [selectedAgentId, setSelectedAgentId] = useState('');
  const [sessionId, setSessionId] = useState('');

  const exportDiagram = () => {
    if (drawioRef.current) {
      drawioRef.current.exportDiagram({
        format: 'xmlsvg'
      });
    }
  };

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages, isChatOpen]);

  // Check Login & Load Agents
  useEffect(() => {
    const userInfo = getUserInfo();
    if (!userInfo || !userInfo.user) {
      router.push('/login');
      return;
    }
    setCurrentUser(userInfo.user);

    // Load Agents
    const loadAgents = async () => {
      try {
        const res = await agentApi.queryAiAgentConfigList();
        setAgents(res.data || []);
        if (res.data && res.data.length > 0) {
          // Try to restore last agent or default to first
          const lastAgentId = localStorage.getItem('ai_agent_last_agent');
          if (lastAgentId && res.data.find(a => a.agentId === lastAgentId)) {
            setSelectedAgentId(lastAgentId);
          } else {
            setSelectedAgentId(res.data[0].agentId);
          }
        }
      } catch (error) {
        console.error('Failed to load agents:', error);
        setMessages(prev => [...prev, {
          id: Date.now().toString(),
          role: 'agent',
          content: '加载智能体列表失败，请检查后端服务是否启动。',
          timestamp: Date.now()
        }]);
      }
    };
    loadAgents();
  }, [router]);

  const handleLogout = () => {
    clearUserInfo();
    router.push('/login');
  };

  const handleAgentChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newAgentId = e.target.value;
    setSelectedAgentId(newAgentId);
    setSessionId(''); // Reset session when agent changes
    localStorage.setItem('ai_agent_last_agent', newAgentId);
  };

  const handleNewChat = async () => {
    if (!selectedAgentId || !currentUser) return;
    
    try {
        const res = await agentApi.createSession(selectedAgentId, currentUser);
        setSessionId(res.data.sessionId);
        setMessages([
            {
                id: Date.now().toString(),
                role: 'agent',
                content: '已为您开启新的会话。',
                timestamp: Date.now()
            }
        ]);
        setInputValue('');
    } catch (error) {
        console.error('Failed to create new session:', error);
    }
  };

  const performSendMessage = async (displayContent: string, apiContent: string) => {
    if (!selectedAgentId) {
      setMessages(prev => [...prev, {
        id: Date.now().toString(),
        role: 'agent',
        content: '请先选择一个智能体。',
        timestamp: Date.now()
      }]);
      setIsSending(false);
      return;
    }

    const userMsg: Message = {
      id: Date.now().toString(),
      role: 'user',
      content: displayContent,
      timestamp: Date.now()
    };
    setMessages(prev => [...prev, userMsg]);

    try {
      // 1. Ensure Session
      let currentSessionId = sessionId;
      if (!currentSessionId) {
        const sessionRes = await agentApi.createSession(selectedAgentId, currentUser);
        currentSessionId = sessionRes.data.sessionId;
        setSessionId(currentSessionId);
      }

      // 2. Send Message
      const chatRes = await agentApi.chat({
        agentId: selectedAgentId,
        userId: currentUser,
        sessionId: currentSessionId,
        message: apiContent
      });

      const agentMsg: Message = {
        id: (Date.now() + 1).toString(),
        role: 'agent',
        content: chatRes.data.content,
        timestamp: Date.now()
      };
      setMessages(prev => [...prev, agentMsg]);

      // Try to load content as Draw.io XML if it looks like XML
      if (drawioRef.current && chatRes.data.content && (chatRes.data.content.includes('<mxGraphModel') || chatRes.data.content.includes('<mxfile'))) {
        try {
          drawioRef.current.load({
            xml: chatRes.data.content
          });
        } catch (e) {
          console.error('Failed to load diagram:', e);
        }
      }

    } catch (error) {
      console.error('Chat error:', error);
      const errorMsg: Message = {
        id: (Date.now() + 1).toString(),
        role: 'agent',
        content: error instanceof Error ? `Error: ${error.message}` : '发送失败，请重试。',
        timestamp: Date.now()
      };
      setMessages(prev => [...prev, errorMsg]);
    } finally {
      setIsSending(false);
    }
  };

  const handleSendMessage = async () => {
    if (!inputValue.trim() || isSending) return;
    
    const content = inputValue;
    setInputValue('');
    setIsSending(true);

    if (useHistoryContext && drawioRef.current) {
        isExportingForChatRef.current = true;
        pendingMessageRef.current = content;
        try {
            drawioRef.current.exportDiagram({
                 format: 'xmlsvg'
             });
        } catch (e) {
            console.error("Export failed", e);
            performSendMessage(content, content);
        }
    } else {
        performSendMessage(content, content);
    }
  };

  useEffect(() => {
    if (!lastExportedData) return;
    
    if (isExportingForChatRef.current) {
        isExportingForChatRef.current = false;
        const xml = lastExportedData.data;
        const content = pendingMessageRef.current;
        const apiContent = `[Context: Current Draw.io XML]\n\`\`\`xml\n${xml}\n\`\`\`\n\n${content}`;
        performSendMessage(content, apiContent);
    } else {
        setImgData(lastExportedData.data);
    }
  }, [lastExportedData]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  return (
    <div className="flex flex-col h-screen w-full overflow-hidden bg-slate-50 text-slate-900 font-sans">
      {/* Header - Minimal & Clean */}
      <div className="h-14 px-6 bg-white border-b border-slate-200 flex items-center justify-between shrink-0 shadow-sm z-10">
        <div className="flex items-center gap-3">
          <div className="bg-indigo-600 p-1.5 rounded-lg">
            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
               <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
               <circle cx="8.5" cy="8.5" r="1.5"></circle>
               <polyline points="21 15 16 10 5 21"></polyline>
            </svg>
          </div>
          <h1 className="text-lg font-semibold text-slate-800 tracking-tight">ai + draw.io - @小傅哥</h1>
        </div>
        
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-3 py-1.5 bg-slate-50 rounded-lg border border-slate-200">
             <div className="w-2 h-2 rounded-full bg-green-500"></div>
             <span className="text-xs font-medium text-slate-600">{currentUser || 'Guest'}</span>
          </div>

          <button 
            onClick={exportDiagram}
            className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 text-slate-700 rounded-lg hover:bg-slate-50 hover:border-slate-300 hover:text-slate-900 transition-all text-sm font-medium shadow-sm active:scale-95"
          >
            <Icons.Download className="text-slate-500" />
            Export
          </button>

          <button
             onClick={handleLogout}
             className="p-2 text-slate-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors border border-transparent hover:border-red-100"
             title="Logout"
          >
            <Icons.Logout />
          </button>
          
          {!isChatOpen && (
            <button
              onClick={() => setIsChatOpen(true)}
              className="p-2 text-slate-500 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-colors border border-transparent hover:border-indigo-100"
              title="Open Assistant"
            >
              <Icons.Chat />
            </button>
          )}
        </div>
      </div>
      
      {/* Main Layout */}
      <div className="flex flex-1 w-full overflow-hidden relative">
        {/* Draw.io Canvas Area */}
        <div className="flex-1 relative bg-slate-100 h-full flex flex-col">
          <div className="flex-1 m-2 rounded-xl overflow-hidden border border-slate-200 shadow-sm bg-white">
            <DrawIoEmbed 
              ref={drawioRef}
              onExport={(data) => setLastExportedData({ data: data.data, timestamp: Date.now() })}
              urlParameters={{
                ui: 'atlas', // More modern UI theme for draw.io
                spin: true,
                libraries: true,
                saveAndExit: false,
                noSaveBtn: true,
                noExitBtn: true
              }}
            />
          </div>
        </div>

        {/* Chat Sidebar - Modern & Elegant */}
        <div 
          className={`
            border-l border-slate-200 bg-white flex flex-col transition-all duration-300 ease-[cubic-bezier(0.25,0.1,0.25,1)]
            ${isChatOpen ? 'w-[360px] translate-x-0' : 'w-0 translate-x-full opacity-0 overflow-hidden'}
            shadow-xl z-20
          `}
        >
          {/* Chat Header */}
          <div className="h-14 px-5 border-b border-slate-100 flex items-center justify-between shrink-0 bg-white/80 backdrop-blur-sm sticky top-0 z-10">
            <div className="flex items-center gap-2.5 flex-1 min-w-0">
              <div className="flex items-center justify-center w-8 h-8 rounded-full bg-gradient-to-tr from-indigo-500 to-purple-500 text-white shadow-md shadow-indigo-200 shrink-0">
                <Icons.Sparkles className="w-4 h-4" />
              </div>
              <div className="flex-1 min-w-0">
                <select 
                  value={selectedAgentId} 
                  onChange={handleAgentChange}
                  className="w-full bg-transparent text-sm font-semibold text-slate-800 focus:outline-none cursor-pointer truncate"
                >
                  {agents.length === 0 && <option value="">Loading agents...</option>}
                  {agents.map(agent => (
                    <option key={agent.agentId} value={agent.agentId}>
                      {agent.agentName}
                    </option>
                  ))}
                </select>
                <span className="block text-xs text-green-500 font-medium leading-tight">● Online</span>
              </div>
            </div>
            <div className="flex items-center gap-1">
                <button 
                  onClick={handleNewChat}
                  className="p-1.5 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-md transition-all shrink-0"
                  title="New Chat"
                >
                  <Icons.Plus className="w-5 h-5" />
                </button>
                <button 
                  onClick={() => setIsChatOpen(false)}
                  className="p-1.5 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-md transition-all shrink-0"
                >
                  <Icons.Close className="w-5 h-5" />
                </button>
            </div>
          </div>

          {/* Messages Area */}
          <div className="flex-1 overflow-y-auto p-5 space-y-6 bg-slate-50/50 scrollbar-thin scrollbar-thumb-slate-200 scrollbar-track-transparent">
            {messages.map((msg) => (
              <div 
                key={msg.id} 
                className={`flex gap-3 ${msg.role === 'user' ? 'flex-row-reverse' : 'flex-row'}`}
              >
                <div className={`
                  shrink-0 w-8 h-8 rounded-full flex items-center justify-center shadow-sm mt-1
                  ${msg.role === 'user' 
                    ? 'bg-indigo-100 text-indigo-600' 
                    : 'bg-white text-indigo-500 border border-slate-100'
                  }
                `}>
                  {msg.role === 'user' ? <Icons.User className="w-5 h-5" /> : <Icons.Bot className="w-5 h-5" />}
                </div>
                
                <div className="flex flex-col max-w-[85%]">
                    <span className={`text-[10px] mb-1 font-medium ${msg.role === 'user' ? 'text-right text-slate-400' : 'text-left text-slate-400'}`}>
                        {msg.role === 'user' ? 'You' : 'Agent'}
                    </span>
                    <div 
                    className={`
                        p-3.5 text-sm leading-relaxed shadow-sm whitespace-pre-wrap
                        ${msg.role === 'user' 
                        ? 'bg-indigo-600 text-white rounded-2xl rounded-tr-sm' 
                        : 'bg-white border border-slate-200 text-slate-700 rounded-2xl rounded-tl-sm'
                        }
                    `}
                    >
                    {msg.content}
                    </div>
                </div>
              </div>
            ))}
            <div ref={messagesEndRef} />
          </div>

          {/* Input Area */}
          <div className="p-4 bg-white border-t border-slate-100 shrink-0">
            {/* Context Toolbar */}
            <div className="flex items-center gap-2 mb-2 px-1">
                <button
                    onClick={() => setUseHistoryContext(!useHistoryContext)}
                    className={`
                        flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium transition-all border
                        ${useHistoryContext 
                            ? 'bg-indigo-50 text-indigo-600 border-indigo-200' 
                            : 'bg-white text-slate-500 border-slate-200 hover:bg-slate-50'
                        }
                    `}
                >
                    <Icons.Layers className={`w-3.5 h-3.5 ${useHistoryContext ? 'text-indigo-500' : 'text-slate-400'}`} />
                    携带当前画布
                </button>
            </div>
            <div className="relative flex items-end gap-2 bg-slate-50 p-1.5 rounded-xl border border-slate-200 focus-within:border-indigo-300 focus-within:ring-2 focus-within:ring-indigo-100 focus-within:bg-white transition-all">
              <textarea
                value={inputValue}
                onChange={(e) => setInputValue(e.target.value)}
                onKeyDown={handleKeyDown}
                placeholder={isSending ? "服务端正在思考中，请稍候..." : "输入您的问题..."}
                disabled={isSending}
                className="flex-1 px-3 py-2 bg-transparent border-none focus:ring-0 text-sm text-slate-800 placeholder:text-slate-400 resize-none max-h-32 min-h-[44px]"
                rows={1}
                style={{ height: 'auto', minHeight: '44px' }}
              />
              <button
                onClick={handleSendMessage}
                disabled={!inputValue.trim() || isSending}
                className={`
                  p-2.5 rounded-lg mb-0.5 transition-all duration-200 flex-shrink-0
                  ${inputValue.trim() && !isSending
                    ? 'bg-indigo-600 text-white shadow-md shadow-indigo-200 hover:bg-indigo-700 hover:scale-105 active:scale-95' 
                    : 'bg-slate-200 text-slate-400 cursor-not-allowed'
                  }
                `}
              >
                {isSending ? <Icons.Loader className="w-4 h-4" /> : <Icons.Send className="w-4 h-4" />}
              </button>
            </div>
            <div className="text-center mt-2">
                <span className="text-[10px] text-slate-400">
                  {isSending ? '服务端正在思考处理，暂时无法输入...' : 'AI can make mistakes. Please verify important info.'}
                </span>
            </div>
          </div>
        </div>
      </div>

      {/* Export Modal - Polished */}
      {imgData && (
        <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-6 animate-in fade-in duration-200">
            <div className="bg-white p-0 rounded-2xl shadow-2xl max-h-[90vh] flex flex-col w-full max-w-4xl overflow-hidden animate-in zoom-in-95 duration-200 border border-white/20">
                <div className="flex justify-between items-center px-6 py-4 border-b border-slate-100 bg-slate-50/50">
                    <div className="flex items-center gap-3">
                        <div className="p-2 bg-green-100 text-green-600 rounded-lg">
                            <Icons.Download className="w-5 h-5" />
                        </div>
                        <div>
                            <h2 className="text-lg font-bold text-slate-800">Export Ready</h2>
                            <p className="text-xs text-slate-500">Your diagram has been successfully converted</p>
                        </div>
                    </div>
                    <button 
                        onClick={() => setImgData(null)}
                        className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full transition-colors"
                    >
                        <Icons.Close className="w-5 h-5" />
                    </button>
                </div>
                
                <div className="flex-1 overflow-auto bg-slate-50/50 p-8 flex items-center justify-center min-h-[400px]">
                    <div className="bg-white p-2 rounded shadow-sm border border-slate-200">
                        <img src={imgData} alt="Exported diagram" className="max-w-full h-auto object-contain" />
                    </div>
                </div>
                
                <div className="px-6 py-4 border-t border-slate-100 bg-white flex justify-end gap-3">
                   <button 
                        onClick={() => setImgData(null)}
                        className="px-5 py-2.5 text-slate-600 font-medium hover:bg-slate-100 rounded-lg transition-colors text-sm"
                    >
                        Close Preview
                    </button>
                    <a 
                        href={imgData} 
                        download="diagram.svg"
                        className="px-5 py-2.5 bg-indigo-600 text-white font-medium rounded-lg hover:bg-indigo-700 shadow-lg shadow-indigo-200 hover:shadow-indigo-300 transition-all text-sm flex items-center gap-2"
                    >
                        <Icons.Download className="w-4 h-4" />
                        Download File
                    </a>
                </div>
            </div>
        </div>
      )}
    </div>
  );
}
