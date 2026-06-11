'use client';

import { Bot, Download, LogOut, MessageCircle, Network } from 'lucide-react';

interface WorkbenchHeaderProps {
  currentUser: string;
  isChatOpen: boolean;
  onExport: () => void;
  onLogout: () => void;
  onOpenChat: () => void;
}

/**
 * description: Displays the desktop workbench identity, user state, and global actions.
 * params:
 * - currentUser: The current logged-in user name displayed in the status pill.
 * - isChatOpen: Whether the assistant sidebar is currently visible.
 * - onExport: Output callback that starts draw.io export.
 * - onLogout: Output callback that clears login and navigates away.
 * - onOpenChat: Output callback that reopens the assistant sidebar.
 */
export function WorkbenchHeader({
  currentUser,
  isChatOpen,
  onExport,
  onLogout,
  onOpenChat,
}: WorkbenchHeaderProps) {
  return (
    <header className="workbench-topbar">
      <div className="flex min-w-0 items-center gap-4">
        <div className="brand-mark">
          <Network className="h-5 w-5" />
        </div>
        <div className="min-w-0">
          <div className="flex items-center gap-3">
            <h1 className="truncate text-[17px] font-semibold tracking-[0.01em] text-[var(--workbench-text)]">
              AI + draw.io Workbench
            </h1>
            <span className="rounded-full border border-[var(--workbench-border)] bg-white/[0.04] px-2 py-0.5 text-[11px] text-[var(--workbench-muted)]">
              @小傅哥
            </span>
          </div>
          <p className="mt-0.5 text-[11px] text-[var(--workbench-muted-2)]">
            Agent conversation, XML rendering, and diagram editing in one desktop surface.
          </p>
        </div>
      </div>

      <div className="flex shrink-0 items-center gap-2">
        <div className="status-pill">
          <span className="h-2 w-2 rounded-full bg-emerald-300 shadow-[0_0_16px_rgba(110,231,183,0.85)]" />
          <Bot className="h-4 w-4 text-[var(--workbench-accent)]" />
          <span className="max-w-32 truncate">{currentUser || 'Guest'}</span>
        </div>

        <button className="toolbar-button toolbar-button-primary" onClick={onExport} type="button">
          <Download className="h-4 w-4" />
          <span>Export</span>
        </button>

        <button className="icon-button" onClick={onLogout} title="Logout" type="button">
          <LogOut className="h-4 w-4" />
        </button>

        {!isChatOpen && (
          <button className="icon-button icon-button-active" onClick={onOpenChat} title="Open Assistant" type="button">
            <MessageCircle className="h-4 w-4" />
          </button>
        )}
      </div>
    </header>
  );
}

