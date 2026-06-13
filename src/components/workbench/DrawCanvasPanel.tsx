'use client';

import { DrawIoEmbed, DrawIoEmbedRef, EventAutoSave, EventExport } from 'react-drawio';
import { RefObject, useEffect } from 'react';
import { CircuitBoard, Loader2 } from 'lucide-react';
import type { ResponseMetadata } from '@/types/api';

interface DrawCanvasPanelProps {
  drawioRef: RefObject<DrawIoEmbedRef | null>;
  drawIoXml?: string | null;
  isDrawIoReady: boolean;
  metadata?: ResponseMetadata;
  onAutoSave: (data: EventAutoSave) => void;
  onLoad: () => void;
  onExport: (data: EventExport) => void;
}

/**
 * description: Wraps the draw.io iframe in the desktop workbench canvas shell.
 * params:
 * - drawioRef: Input/output ref used by the page container to call draw.io actions.
 * - drawIoXml: Input XML loaded into the draw.io iframe when the active session changes.
 * - isDrawIoReady: Input readiness state displayed in the canvas status strip.
 * - metadata: Input optional backend result metadata displayed as diagram notes.
 * - onAutoSave: Output callback that receives draw.io autosave XML.
 * - onLoad: Output callback triggered after draw.io finishes loading.
 * - onExport: Output callback that receives exported draw.io data.
 */
export function DrawCanvasPanel({
  drawioRef,
  drawIoXml,
  isDrawIoReady,
  metadata,
  onAutoSave,
  onLoad,
  onExport,
}: DrawCanvasPanelProps) {
  const suggestions = Array.isArray(metadata?.suggestions) ? metadata.suggestions : [];
  const nextActions = Array.isArray(metadata?.nextActions) ? metadata.nextActions : [];
  const hasMetadata = Boolean(metadata?.summary || suggestions.length > 0 || nextActions.length > 0);

  useEffect(() => {
    if (!isDrawIoReady || !drawioRef.current) {
      return;
    }

    drawioRef.current.load({ xml: drawIoXml || '' });
  }, [drawIoXml, drawioRef, isDrawIoReady]);

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
          xml={drawIoXml || undefined}
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

      {hasMetadata && (
        <aside className="canvas-metadata">
          {metadata?.summary && (
            <div>
              <p className="canvas-metadata-label">summary</p>
              <p className="canvas-metadata-text">{metadata.summary}</p>
            </div>
          )}

          {suggestions.length > 0 && (
            <div>
              <p className="canvas-metadata-label">suggestions</p>
              <div className="canvas-metadata-list">
                {suggestions.map((item) => (
                  <span key={item}>{item}</span>
                ))}
              </div>
            </div>
          )}

          {nextActions.length > 0 && (
            <div>
              <p className="canvas-metadata-label">next actions</p>
              <div className="canvas-metadata-list">
                {nextActions.map((item) => (
                  <span key={item}>{item}</span>
                ))}
              </div>
            </div>
          )}
        </aside>
      )}
    </section>
  );
}
