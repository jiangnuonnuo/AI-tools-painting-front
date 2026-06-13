'use client';

import {
  Bot,
  ChevronDown,
  Layers3,
  Loader2,
  RotateCcw,
  Send,
  User,
  X,
} from 'lucide-react';
import { AiAgentConfigResponseDTO } from '@/types/api';
import { Message } from '@/types/workbench';

interface QuickAction {
  label: string;
  text: string;
}

interface ChatSidebarProps {
  agents: AiAgentConfigResponseDTO[];
  selectedAgentId: string;
  messages: Message[];
  inputValue: string;
  isSending: boolean;
  isChatOpen: boolean;
  useHistoryContext: boolean;
  quickActions: QuickAction[];
  chatWidth?: number;
  isResizing?: boolean;
  messagesEndRef: React.RefObject<HTMLDivElement | null>;
  onAgentChange: (event: React.ChangeEvent<HTMLSelectElement>) => void;
  onInputChange: (value: string) => void;
  onSendMessage: () => void;
  onRestartSession: () => void;
  onToggleHistoryContext: () => void;
  onStartResize?: (event: React.MouseEvent<HTMLButtonElement>) => void;
  onToggleThinking?: (messageId: string) => void;
  onCloseChat: () => void;
  onKeyDown: (event: React.KeyboardEvent) => void;
}

/**
 * description: Displays AI agent selection, messages, context controls, and prompt input.
 * params:
 * - agents: Input AI agent options loaded from the backend.
 * - selectedAgentId: Input selected agent id.
 * - messages: Input chat message list for the active session.
 * - inputValue: Input controlled prompt value.
 * - isSending: Input sending state for disabling controls.
 * - isChatOpen: Input visibility state for desktop sidebar sizing.
 * - useHistoryContext: Input context export toggle state.
 * - quickActions: Input starter prompts.
 * - messagesEndRef: Input ref used by the page container for auto-scroll.
 * - callbacks: Output callbacks for agent changes, sending, restart, context toggle, close, and keyboard handling.
 */
export function ChatSidebar({
  agents,
  selectedAgentId,
  messages,
  inputValue,
  isSending,
  isChatOpen,
  useHistoryContext,
  quickActions,
  chatWidth,
  isResizing,
  messagesEndRef,
  onAgentChange,
  onInputChange,
  onSendMessage,
  onRestartSession,
  onToggleHistoryContext,
  onStartResize,
  onToggleThinking,
  onCloseChat,
  onKeyDown,
}: ChatSidebarProps) {
  return (
    <aside
      className={`chat-panel relative ${isChatOpen ? 'chat-panel-open' : 'chat-panel-closed'}`}
      style={chatWidth && isChatOpen ? { width: `${chatWidth}px` } : undefined}
    >
      {isChatOpen && onStartResize && (
        <button
          aria-label="调整对话面板宽度"
          className={`absolute left-0 top-0 z-20 h-full w-2 -translate-x-1 cursor-col-resize ${isResizing ? 'bg-[var(--workbench-accent)]/40' : 'bg-transparent hover:bg-[var(--workbench-accent)]/20'}`}
          onMouseDown={onStartResize}
          type="button"
        />
      )}
      <div className="chat-heading">
        <div className="flex min-w-0 items-center gap-3">
          <div className="assistant-mark">
            <Bot className="h-4 w-4" />
          </div>
          <div className="min-w-0 flex-1">
            <div className="relative">
              <select className="agent-select" onChange={onAgentChange} value={selectedAgentId}>
                {agents.length === 0 && <option value="">Loading agents...</option>}
                {agents.map((agent) => (
                  <option key={agent.agentId} value={agent.agentId}>
                    {agent.agentName}
                  </option>
                ))}
              </select>
              <ChevronDown className="pointer-events-none absolute right-0 top-1/2 h-4 w-4 -translate-y-1/2 text-[var(--workbench-muted-2)]" />
            </div>
            <div className="mt-1 flex items-center gap-1.5 text-[11px] text-[var(--workbench-muted-2)]">
              <span className="h-1.5 w-1.5 rounded-full bg-emerald-300" />
              <span>AI Assistant Online</span>
            </div>
          </div>
        </div>
        <button className="icon-button" onClick={onCloseChat} title="Close Assistant" type="button">
          <X className="h-4 w-4" />
        </button>
      </div>

      <div className="chat-messages">
        {messages.map((message) => (
          <div className={`message-row ${message.role === 'user' ? 'message-row-user' : ''}`} key={message.id}>
            <div className={`message-avatar ${message.role === 'user' ? 'message-avatar-user' : ''}`}>
              {message.role === 'user' ? <User className="h-4 w-4" /> : <Bot className="h-4 w-4" />}
            </div>
            <div className="min-w-0 max-w-[86%]">
              <div className={`message-meta ${message.role === 'user' ? 'text-right' : 'text-left'}`}>
                {message.role === 'user' ? 'You' : 'Agent'}
              </div>
              {message.role === 'agent' && (message.steps?.length || message.reasoning) ? (
                <div className="mb-2">
                  <button
                    className="thinking-toggle"
                    onClick={() => onToggleThinking?.(message.id)}
                    type="button"
                  >
                    <Loader2 className={`h-3.5 w-3.5 ${message.steps?.some(step => step.status === 'running') ? 'animate-spin' : ''}`} />
                    <span>{message.thinkingOpen ? '收起思考过程' : `思考过程 · ${message.steps?.length || 0} 个阶段`}</span>
                  </button>
                  {message.thinkingOpen && (
                    <div className="thinking-drawer">
                      {message.steps?.length ? (
                        message.steps.map((step, index) => (
                          <details className="thinking-step" key={step.id || `${message.id}-${step.phase}-${index}`}>
                            <summary className="thinking-step-summary">
                              <span>{step.status === 'running' ? '进行中' : '完成'}</span>
                              <strong>{step.label}</strong>
                              <small>{step.summary || step.content.replace(/\s+/g, ' ').slice(0, 34)}</small>
                              {step.event && <em>{step.event}</em>}
                            </summary>
                            {step.content && <pre>{step.content}</pre>}
                          </details>
                        ))
                      ) : (
                        <pre>{message.reasoning}</pre>
                      )}
                    </div>
                  )}
                </div>
              ) : null}
              {message.content && (
                <div className={`message-bubble ${message.role === 'user' ? 'message-bubble-user' : ''}`}>
                  {message.content}
                </div>
              )}
            </div>
          </div>
        ))}
        <div ref={messagesEndRef} />
      </div>

      <div className="chat-composer">
        {messages.length <= 1 && (
          <div className="mb-3 flex flex-wrap gap-2">
            {quickActions.map((action) => (
              <button className="quick-action" key={action.label} onClick={() => onInputChange(action.text)} type="button">
                {action.label}
              </button>
            ))}
          </div>
        )}

        <div className="mb-2 flex items-center gap-2">
          <button
            className={`context-toggle ${useHistoryContext ? 'context-toggle-active' : ''}`}
            onClick={onToggleHistoryContext}
            type="button"
          >
            <Layers3 className="h-3.5 w-3.5" />
            <span>携带画布上下文</span>
          </button>
          <span className="ml-auto text-[10px] text-[var(--workbench-muted-2)]">Ctrl/Command + Enter</span>
        </div>

        <div className="composer-box">
          <textarea
            className="composer-input"
            disabled={isSending}
            onChange={(event) => onInputChange(event.target.value)}
            onKeyDown={onKeyDown}
            placeholder={isSending ? 'AI 正在思考中...' : '输入您的问题，描述您的需求...'}
            rows={3}
            value={inputValue}
          />
          <div className="flex shrink-0 flex-col gap-2">
            <button
              className="send-button"
              disabled={!inputValue.trim() || isSending}
              onClick={onSendMessage}
              title="Send"
              type="button"
            >
              {isSending ? <Loader2 className="h-4 w-4 animate-spin" /> : <Send className="h-4 w-4" />}
            </button>
            <button className="composer-icon-button" onClick={onRestartSession} title="Restart Session" type="button">
              <RotateCcw className="h-4 w-4" />
            </button>
          </div>
        </div>

        <p className="mt-2 text-center text-[10px] text-[var(--workbench-muted-2)]">
          {isSending ? 'AI is generating response...' : 'AI can make mistakes. Please verify important info.'}
        </p>
      </div>
    </aside>
  );
}
