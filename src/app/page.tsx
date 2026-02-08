'use client';

import { DrawIoEmbed, DrawIoEmbedRef } from 'react-drawio';
import { useRef, useState } from 'react';

export default function Home() {
  const [imgData, setImgData] = useState<string | null>(null);
  const drawioRef = useRef<DrawIoEmbedRef>(null);

  const exportDiagram = () => {
    if (drawioRef.current) {
      drawioRef.current.exportDiagram({
        format: 'xmlsvg'
      });
    }
  };

  return (
    <div className="flex flex-col h-screen w-full">
      <div className="p-4 flex gap-4 bg-gray-100 border-b items-center justify-between">
        <h1 className="text-xl font-bold text-black">React Draw.io Demo</h1>
        <button 
          onClick={exportDiagram}
          className="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600"
        >
          Export
        </button>
      </div>
      
      <div className="flex-1 w-full relative bg-white">
        <DrawIoEmbed 
          ref={drawioRef}
          onExport={(data) => setImgData(data.data)} 
        />
      </div>

      {imgData && (
        <div className="absolute inset-0 bg-black/50 flex items-center justify-center z-50 p-10">
            <div className="bg-white p-6 rounded shadow-lg max-h-full overflow-auto relative flex flex-col gap-4">
                <div className="flex justify-between items-center">
                    <h2 className="text-lg font-bold text-black">Exported Image</h2>
                    <button 
                        onClick={() => setImgData(null)}
                        className="px-3 py-1 bg-gray-200 rounded hover:bg-gray-300 text-black"
                    >
                        Close
                    </button>
                </div>
                <div className="border border-gray-200 p-2 rounded">
                    <img src={imgData} alt="Exported diagram" className="max-w-full" />
                </div>
            </div>
        </div>
      )}
    </div>
  );
}
