'use client';

import { DrawIoEmbed, DrawIoEmbedRef } from 'react-drawio';
import { useRef, useState, useEffect } from 'react';

// Message type definition
type Message = {
  id: string;
  role: 'user' | 'agent';
  content: string;
  timestamp: number;
};

// Simple SVG Icons
const Icons = {
  Chat: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
    </svg>
  ),
  Close: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <line x1="18" y1="6" x2="6" y2="18"></line>
      <line x1="6" y1="6" x2="18" y2="18"></line>
    </svg>
  ),
  Send: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <line x1="22" y1="2" x2="11" y2="13"></line>
      <polygon points="22 2 15 22 11 13 2 9 22 2"></polygon>
    </svg>
  ),
  User: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="bg-blue-100 p-1 rounded-full text-blue-600">
      <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
      <circle cx="12" cy="7" r="4"></circle>
    </svg>
  ),
  Bot: () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="bg-green-100 p-1 rounded-full text-green-600">
      <path d="M12 2a2 2 0 0 1 2 2v2a2 2 0 0 1-2 2 2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z"></path>
      <path d="M4 11v6a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2z"></path>
      <path d="M9 22v-3"></path>
      <path d="M15 22v-3"></path>
    </svg>
  )
};

export default function Home() {
  const [imgData, setImgData] = useState<string | null>(null);
  const drawioRef = useRef<DrawIoEmbedRef>(null);
  
  // Chat State
  const [isChatOpen, setIsChatOpen] = useState(true);
  const [messages, setMessages] = useState<Message[]>([
    {
      id: '1',
      role: 'agent',
      content: '你好！我是你的智能助手。有什么我可以帮你的吗？',
      timestamp: Date.now()
    }
  ]);
  const [inputValue, setInputValue] = useState('');
  const messagesEndRef = useRef<HTMLDivElement>(null);

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

  const handleSendMessage = () => {
    if (!inputValue.trim()) return;

    const userMsg: Message = {
      id: Date.now().toString(),
      role: 'user',
      content: inputValue,
      timestamp: Date.now()
    };

    setMessages(prev => [...prev, userMsg]);
    setInputValue('');

    // Simulate Agent Reply
    setTimeout(() => {
      const agentMsg: Message = {
        id: (Date.now() + 1).toString(),
        role: 'agent',
        content: `我收到了你的消息："${userMsg.content}"。这是一个模拟回复。`,
        timestamp: Date.now()
      };
      setMessages(prev => [...prev, agentMsg]);
    }, 1000);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  return (
    <div className="flex flex-col h-screen w-full overflow-hidden">
      {/* Header */}
      <div className="h-14 px-4 flex gap-4 bg-gray-100 border-b border-gray-200 items-center justify-between shrink-0">
        <h1 className="text-xl font-bold text-black">React Draw.io Demo</h1>
        <div className="flex items-center gap-2">
          <button 
            onClick={exportDiagram}
            className="px-4 py-1.5 bg-blue-500 text-white rounded hover:bg-blue-600 text-sm font-medium transition-colors"
          >
            Export
          </button>
          {!isChatOpen && (
            <button
              onClick={() => setIsChatOpen(true)}
              className="p-2 text-gray-600 hover:bg-gray-200 rounded transition-colors"
              title="Open Chat"
            >
              <Icons.Chat />
            </button>
          )}
        </div>
      </div>
      
      {/* Main Content */}
      <div className="flex flex-1 w-full overflow-hidden relative">
        {/* Draw.io Container */}
        <div className="flex-1 relative bg-white h-full">
          <DrawIoEmbed 
            ref={drawioRef}
            onExport={(data) => setImgData(data.data)} 
          />
        </div>

        {/* Chat Sidebar */}
        <div 
          className={`
            border-l border-gray-200 bg-gray-50 flex flex-col transition-all duration-300 ease-in-out
            ${isChatOpen ? 'w-80 translate-x-0' : 'w-0 translate-x-full opacity-0 overflow-hidden'}
          `}
        >
          {/* Chat Header */}
          <div className="h-12 px-4 border-b border-gray-200 flex items-center justify-between bg-white shrink-0">
            <span className="font-semibold text-gray-700 flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-green-500"></span>
              AI Assistant
            </span>
            <button 
              onClick={() => setIsChatOpen(false)}
              className="p-1 text-gray-500 hover:text-gray-700 hover:bg-gray-100 rounded"
            >
              <Icons.Close />
            </button>
          </div>

          {/* Messages Area */}
          <div className="flex-1 overflow-y-auto p-4 space-y-4">
            {messages.map((msg) => (
              <div 
                key={msg.id} 
                className={`flex gap-3 ${msg.role === 'user' ? 'flex-row-reverse' : 'flex-row'}`}
              >
                <div className="shrink-0 mt-1">
                  {msg.role === 'user' ? <Icons.User /> : <Icons.Bot />}
                </div>
                <div 
                  className={`
                    max-w-[80%] p-3 rounded-lg text-sm leading-relaxed
                    ${msg.role === 'user' 
                      ? 'bg-blue-500 text-white rounded-tr-none' 
                      : 'bg-white border border-gray-200 text-gray-800 rounded-tl-none shadow-sm'
                    }
                  `}
                >
                  {msg.content}
                </div>
              </div>
            ))}
            <div ref={messagesEndRef} />
          </div>

          {/* Input Area */}
          <div className="p-4 bg-white border-t border-gray-200 shrink-0">
            <div className="flex gap-2">
              <input
                type="text"
                value={inputValue}
                onChange={(e) => setInputValue(e.target.value)}
                onKeyDown={handleKeyDown}
                placeholder="Type a message..."
                className="flex-1 px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm text-black"
              />
              <button
                onClick={handleSendMessage}
                disabled={!inputValue.trim()}
                className={`
                  p-2 rounded-md text-white transition-colors
                  ${inputValue.trim() 
                    ? 'bg-blue-500 hover:bg-blue-600' 
                    : 'bg-gray-300 cursor-not-allowed'
                  }
                `}
              >
                <Icons.Send />
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Export Modal */}
      {imgData && (
        <div className="absolute inset-0 bg-black/50 flex items-center justify-center z-50 p-10 backdrop-blur-sm">
            <div className="bg-white p-6 rounded-xl shadow-2xl max-h-[90vh] overflow-hidden flex flex-col gap-4 max-w-4xl w-full">
                <div className="flex justify-between items-center border-b pb-4">
                    <h2 className="text-lg font-bold text-black">Exported Diagram</h2>
                    <button 
                        onClick={() => setImgData(null)}
                        className="p-2 hover:bg-gray-100 rounded-full transition-colors"
                    >
                        <Icons.Close />
                    </button>
                </div>
                <div className="flex-1 overflow-auto bg-gray-50 p-4 rounded border border-gray-200 flex justify-center">
                    <img src={imgData} alt="Exported diagram" className="max-w-full h-auto object-contain shadow-sm" />
                </div>
                <div className="flex justify-end pt-2">
                   <button 
                        onClick={() => setImgData(null)}
                        className="px-4 py-2 bg-gray-800 text-white rounded hover:bg-gray-900 transition-colors"
                    >
                        Close
                    </button>
                </div>
            </div>
        </div>
      )}
    </div>
  );
}
