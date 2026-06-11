'use client';

import { MessageSquareText, Plus, Trash2 } from 'lucide-react';
import type { MouseEvent } from 'react';
import { Session } from '@/types/workbench';

interface SessionsSidebarProps {
  sessions: Session[];
  currentSessionId: string | null;
  onNewChat: () => void;
  onSwitchSession: (targetSessionId: string) => void;
  onDeleteSession: (event: MouseEvent, sessionIdToDelete: string) => void;
  onRenameSession: (session: Session) => void;
}

/**
 * description: Shows saved diagram conversations and session-level actions.
 * params:
 * - sessions: Input list of local draw.io sessions.
 * - currentSessionId: Input active session id.
 * - onNewChat: Output callback that creates a new backend/local chat session.
 * - onSwitchSession: Output callback that loads the selected session.
 * - onDeleteSession: Output callback that removes a session.
 * - onRenameSession: Output callback that opens the rename dialog.
 */
export function SessionsSidebar({
  sessions,
  currentSessionId,
  onNewChat,
  onSwitchSession,
  onDeleteSession,
  onRenameSession,
}: SessionsSidebarProps) {
  const sortedSessions = [...sessions].sort((a, b) => b.lastModified - a.lastModified);

  return (
    <aside className="sessions-panel">
      <div className="panel-heading">
        <div>
          <p className="panel-kicker">History</p>
          <h2 className="panel-title">
            <MessageSquareText className="h-4 w-4 text-[var(--workbench-accent)]" />
            绘图记录
          </h2>
        </div>
        <button className="icon-button icon-button-active" onClick={onNewChat} title="New Chat" type="button">
          <Plus className="h-4 w-4" />
        </button>
      </div>

      <div className="flex-1 overflow-y-auto px-3 py-3">
        <div className="space-y-2">
          {sortedSessions.map((session) => {
            const isActive = currentSessionId === session.id;

            return (
              <div
                className={`session-item group ${isActive ? 'session-item-active' : ''}`}
                key={session.id}
                onClick={() => onSwitchSession(session.id)}
                onDoubleClick={(event) => {
                  event.stopPropagation();
                  onRenameSession(session);
                }}
                role="button"
                tabIndex={0}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    onSwitchSession(session.id);
                  }
                }}
              >
                <span className="session-line" />
                <span className="min-w-0 flex-1 text-left">
                  <span className="block truncate text-sm font-medium">{session.title}</span>
                  <span className="mt-1 block text-[11px] text-[var(--workbench-muted-2)]">
                    {new Date(session.lastModified).toLocaleDateString()}{' '}
                    {new Date(session.lastModified).toLocaleTimeString([], {
                      hour: '2-digit',
                      minute: '2-digit',
                    })}
                  </span>
                </span>
                <button
                  className="session-delete"
                  onClick={(event) => onDeleteSession(event, session.id)}
                  title="Delete"
                  type="button"
                >
                  <Trash2 className="h-4 w-4" />
                </button>
              </div>
            );
          })}
        </div>

        {sessions.length === 0 && (
          <div className="empty-state">
            <MessageSquareText className="h-5 w-5" />
            <span>No history yet</span>
          </div>
        )}
      </div>
    </aside>
  );
}
