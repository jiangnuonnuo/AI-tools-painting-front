'use client';

import { DrawIoEmbed, DrawIoEmbedRef, EventAutoSave, EventExport } from 'react-drawio';
import { RefObject } from 'react';
import { CircuitBoard, Loader2 } from 'lucide-react';

interface DrawCanvasPanelProps {
  drawioRef: RefObject<DrawIoEmbedRef | null>;
  isDrawIoReady: boolean;
  onAutoSave: (data: EventAutoSave) => void;
  onLoad: () => void;
  onExport: (data: EventExport) => void;
}

/**
 * description: Wraps the draw.io iframe in the desktop workbench canvas shell.
 * params:
 * - drawioRef: Input/output ref used by the page container to call draw.io actions.
 * - isDrawIoReady: Input readiness state displayed in the canvas status strip.
 * - onAutoSave: Output callback that receives draw.io autosave XML.
 * - onLoad: Output callback triggered after draw.io finishes loading.
 * - onExport: Output callback that receives exported draw.io data.
 */
export function DrawCanvasPanel({
  drawioRef,
  isDrawIoReady,
  onAutoSave,
  onLoad,
  onExport,
}: DrawCanvasPanelProps) {
  return (
    <section className="canvas-panel">
      <div className="canvas-statusbar">
        <div className="flex items-center gap-2">
          <CircuitBoard className="h-4 w-4 text-[var(--workbench-accent)]" />
          <span className="font-medium text-[var(--workbench-text)]">draw.io Canvas</span>
          <span className="text-[var(--workbench-muted-2)]">XML rendering surface</span>
        </div>
        <div className="flex items-center gap-2 text-[var(--workbench-muted)]">
          {isDrawIoReady ? (
            <span className="h-2 w-2 rounded-full bg-emerald-300 shadow-[0_0_12px_rgba(110,231,183,0.8)]" />
          ) : (
            <Loader2 className="h-3.5 w-3.5 animate-spin text-[var(--workbench-accent)]" />
          )}
          <span>{isDrawIoReady ? 'Ready' : 'Loading'}</span>
        </div>
      </div>

      <div className="canvas-frame">
        <DrawIoEmbed
          ref={drawioRef}
          autosave={true}
          onAutoSave={onAutoSave}
          onLoad={onLoad}
          onExport={onExport}
          urlParameters={{
            ui: 'atlas',
            spin: true,
            libraries: true,
            saveAndExit: false,
            noSaveBtn: true,
            noExitBtn: true,
          }}
        />
      </div>
    </section>
  );
}

