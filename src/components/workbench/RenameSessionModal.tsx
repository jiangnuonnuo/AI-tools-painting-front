'use client';

import { Check, X } from 'lucide-react';

interface RenameSessionModalProps {
  isOpen: boolean;
  newSessionTitle: string;
  onTitleChange: (value: string) => void;
  onClose: () => void;
  onSave: () => void;
}

/**
 * description: Edits the title of the active local workbench session.
 * params:
 * - isOpen: Input modal visibility state.
 * - newSessionTitle: Input controlled session title value.
 * - onTitleChange: Output callback that updates the controlled title.
 * - onClose: Output callback that closes the modal without saving.
 * - onSave: Output callback that persists the new title.
 */
export function RenameSessionModal({
  isOpen,
  newSessionTitle,
  onTitleChange,
  onClose,
  onSave,
}: RenameSessionModalProps) {
  if (!isOpen) {
    return null;
  }

  return (
    <div className="modal-backdrop">
      <div className="rename-modal">
        <div className="modal-header">
          <div>
            <h2 className="text-base font-semibold text-[var(--workbench-text)]">Rename Session</h2>
            <p className="text-xs text-[var(--workbench-muted-2)]">Update the local title shown in 绘图记录.</p>
          </div>
          <button className="icon-button" onClick={onClose} type="button">
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="p-5">
          <label className="field-label" htmlFor="session-title">
            Session Name
          </label>
          <input
            autoFocus
            className="dark-input mt-2"
            id="session-title"
            onChange={(event) => onTitleChange(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                onSave();
              }
            }}
            placeholder="Enter new name"
            type="text"
            value={newSessionTitle}
          />
        </div>

        <div className="modal-actions">
          <button className="toolbar-button" onClick={onClose} type="button">
            Cancel
          </button>
          <button className="toolbar-button toolbar-button-primary" onClick={onSave} type="button">
            <Check className="h-4 w-4" />
            Save Changes
          </button>
        </div>
      </div>
    </div>
  );
}

