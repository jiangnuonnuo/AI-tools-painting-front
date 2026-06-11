'use client';

import Image from 'next/image';
import { Download, X } from 'lucide-react';

interface ExportPreviewModalProps {
  imgData: string | null;
  onClose: () => void;
}

/**
 * description: Shows the exported draw.io SVG preview and download action.
 * params:
 * - imgData: Input exported image data URL.
 * - onClose: Output callback that closes the preview modal.
 */
export function ExportPreviewModal({ imgData, onClose }: ExportPreviewModalProps) {
  if (!imgData) {
    return null;
  }

  return (
    <div className="modal-backdrop">
      <div className="export-modal">
        <div className="modal-header">
          <div className="flex items-center gap-3">
            <div className="modal-icon">
              <Download className="h-5 w-5" />
            </div>
            <div>
              <h2 className="text-base font-semibold text-[var(--workbench-text)]">Export Ready</h2>
              <p className="text-xs text-[var(--workbench-muted-2)]">Your diagram has been converted to SVG.</p>
            </div>
          </div>
          <button className="icon-button" onClick={onClose} type="button">
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="export-preview">
          <div className="export-preview-surface">
            <Image
              alt="Exported diagram"
              className="h-auto max-h-[60vh] w-auto max-w-full object-contain"
              height={900}
              src={imgData}
              unoptimized
              width={1200}
            />
          </div>
        </div>

        <div className="modal-actions">
          <button className="toolbar-button" onClick={onClose} type="button">
            Close Preview
          </button>
          <a className="toolbar-button toolbar-button-primary" download="diagram.svg" href={imgData}>
            <Download className="h-4 w-4" />
            Download File
          </a>
        </div>
      </div>
    </div>
  );
}

