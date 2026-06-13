'use client';

import { useRef, useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { getUserInfo, clearUserInfo } from '@/utils/cookie';
import { agentApi } from '@/api/agent';
import {
  AiAgentConfigResponseDTO,
  PptData,
  PptSlide,
  PptSlideElement,
  ResponseMetadata,
} from '@/types/api';
import pptxgen from 'pptxgenjs';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';

// Message type definition
type MessageStep = {
  id: string;
  phase: string;
  label: string;
  summary: string;
  content: string;
  author?: string;
  event?: string;
  open?: boolean;
  status: 'running' | 'done' | 'pending';
};

type Message = {
  id: string;
  role: 'user' | 'agent';
  content: string;
  reasoning?: string;
  steps?: MessageStep[];
  thinkingOpen?: boolean;
  timestamp: number;
};

// Icons
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
  ),
  Trash: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <polyline points="3 6 5 6 21 6"></polyline>
      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
    </svg>
  ),
  MessageSquare: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
    </svg>
  ),
  Square: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="currentColor" className={className}>
      <rect x="6" y="6" width="12" height="12" rx="2" ry="2"></rect>
    </svg>
  ),
  FilePresentation: ({ className }: { className?: string }) => (
    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect>
      <line x1="8" y1="21" x2="16" y2="21"></line>
      <line x1="12" y1="17" x2="12" y2="21"></line>
    </svg>
  ),
};

interface Session {
  id: string;
  backendSessionId?: string;
  title: string;
  messages: Message[];
  pptData: PptData | null;
  metadata?: ResponseMetadata;
  lastModified: number;
}

export interface CustomModelConfig {
  id: string;
  name: string;
  baseUrl: string;
  apiKey: string;
  model: string;
  completionsPath: string;
  enabled: boolean;
}

// ===== Multi-Theme System =====
interface PptTheme {
  id: string;
  name: string;
  primary: string;       // Main color - header/cover background
  primaryMid: string;    // Slightly lighter primary
  primaryLight: string;  // Light accent
  accent: string;        // Highlight stripe color
  titleColor: string;    // Text on dark bg
  bodyColor: string;     // Body text
  subColor: string;      // Subtitles
  lightGray: string;
  white: string;
  offWhite: string;
  // Cover layout proportions
  coverNavyHeight: number;    // How much of the slide the top color covers (in PPT inches, out of 7.5)
  contentHeaderHeight: number; // Content page header height
}

const THEMES: PptTheme[] = [
  {
    id: 'navy',
    name: '深海蓝',
    primary: '1F3864',
    primaryMid: '2E5090',
    primaryLight: '4472C4',
    accent: 'D4560A',
    titleColor: 'FFFFFF',
    bodyColor: '333333',
    subColor: '666666',
    lightGray: 'AAAAAA',
    white: 'FFFFFF',
    offWhite: 'F2F4F7',
    coverNavyHeight: 3.4,
    contentHeaderHeight: 1.3,
  },
  {
    id: 'emerald',
    name: '翡翠绿',
    primary: '1B5E3A',
    primaryMid: '2E7D50',
    primaryLight: '4CAF6E',
    accent: 'F9A825',
    titleColor: 'FFFFFF',
    bodyColor: '2E3B2E',
    subColor: '5A6B5A',
    lightGray: 'A0AEA0',
    white: 'FFFFFF',
    offWhite: 'F0F5F0',
    coverNavyHeight: 3.4,
    contentHeaderHeight: 1.3,
  },
  {
    id: 'burgundy',
    name: '酒红金',
    primary: '6B1D2A',
    primaryMid: '8E2D3E',
    primaryLight: 'B84056',
    accent: 'C9A84C',
    titleColor: 'FFFFFF',
    bodyColor: '3B2020',
    subColor: '6B4A4A',
    lightGray: 'B09090',
    white: 'FFFFFF',
    offWhite: 'F7F0F0',
    coverNavyHeight: 3.4,
    contentHeaderHeight: 1.3,
  },
  {
    id: 'charcoal',
    name: '极简灰',
    primary: '2C2C2C',
    primaryMid: '4A4A4A',
    primaryLight: '6A6A6A',
    accent: 'E85D3A',
    titleColor: 'FFFFFF',
    bodyColor: '333333',
    subColor: '666666',
    lightGray: 'AAAAAA',
    white: 'FFFFFF',
    offWhite: 'F5F5F5',
    coverNavyHeight: 3.4,
    contentHeaderHeight: 1.3,
  },
  {
    id: 'ocean',
    name: '海洋蓝',
    primary: '0D47A1',
    primaryMid: '1565C0',
    primaryLight: '42A5F5',
    accent: 'FF6D00',
    titleColor: 'FFFFFF',
    bodyColor: '263238',
    subColor: '546E7A',
    lightGray: '90A4AE',
    white: 'FFFFFF',
    offWhite: 'E3F2FD',
    coverNavyHeight: 3.4,
    contentHeaderHeight: 1.3,
  },
];

const PPT_AGENT_ID = '300001';

// Default theme (used as fallback)
const DEFAULT_THEME = THEMES[0];

// Shape elements are now fully skipped (theme handles all decorations)

// Generate PPTX from PptData with professional theme
// Smart layout inference — ensure visual diversity even when AI doesn't specify layout
const inferLayout = (slideData: PptSlide, slideIdx: number, totalSlides: number): string => {
  let layout = slideData.layout || '';
  // Map old names to new names just in case
  if (layout === 'title_slide') layout = 'title_classic';
  if (layout === 'content_slide') layout = 'content_classic';

  const isTitleSlide = layout.startsWith('title_') || slideIdx === 0;
  const isEndSlide = layout === 'end_slide' || slideIdx === totalSlides - 1;
  
  if (!layout) {
    if (isTitleSlide) layout = 'title_classic';
    else if (isEndSlide) layout = 'end_slide';
    else {
      const diverseLayouts = ['content_top', 'content_classic', 'card_3col', 'card_2col', 'comparison', 'timeline', 'data_highlight', 'quote_slide'];
      layout = diverseLayouts[(slideIdx - 1) % diverseLayouts.length];
    }
  }
  return layout;
};

/**
 * Determine if a point (x, y) is on a dark-colored area of the layout decoration.
 * Returns true if the point falls within a primary-color (dark) decorative region.
 * This replaces the fragile layout-type-based color logic with position-aware detection.
 */
const isOnDarkArea = (x: number, y: number, layout: string, theme: PptTheme): boolean => {
  // title_classic / end_slide: dark cover area from y=0 to y=coverNavyHeight, dark bottom bar y=7.15..7.5
  if (layout === 'title_classic' || layout === 'end_slide') {
    return y < theme.coverNavyHeight || y >= 7.15;
  }
  // title_center: dark bars at top 0..0.25 and bottom 7.25..7.5
  if (layout === 'title_center') {
    return y < 0.25 || y >= 7.25;
  }
  // title_split: left half is dark
  if (layout === 'title_split') {
    return x < 6.66;
  }
  // content_classic: left band is dark (x=0..4.5), bottom bar
  if (layout === 'content_classic') {
    return x < 4.5 || y >= 7.15;
  }
  // content_top: dark header y=0..1.2, bottom bar
  if (layout === 'content_top') {
    return y < 1.2 || y >= 7.15;
  }
  // card_3col: dark header y=0..0.9, bottom bar
  if (layout === 'card_3col') {
    return y < 0.9 || y >= 7.15;
  }
  // card_2col: dark header y=0..0.9, bottom bar
  if (layout === 'card_2col') {
    return y < 0.9 || y >= 7.15;
  }
  // comparison: dark header y=0..0.9, bottom bar
  if (layout === 'comparison') {
    return y < 0.9 || y >= 7.15;
  }
  // timeline: dark left bar x=0..0.5, bottom bar
  if (layout === 'timeline') {
    return x < 0.5 || y >= 7.15;
  }
  // data_highlight: dark header y=0..0.7, bottom bar
  if (layout === 'data_highlight') {
    return y < 0.7 || y >= 7.15;
  }
  // quote_slide: dark left bar x=0..0.8
  if (layout === 'quote_slide') {
    return x < 0.8;
  }
  // Default fallback
  return false;
};

/**
 * Get the safe content area for a layout (elements should be placed within this area).
 * Returns { x, y, w, h } in slide coordinates.
 */
const getSafeContentArea = (layout: string, theme: PptTheme): { x: number; y: number; w: number; h: number } => {
  const W = 13.33;
  const H = 7.5;
  const bottomBar = 0.35; // bottom bar height

  if (layout === 'title_classic' || layout === 'end_slide') {
    return { x: 0.5, y: 0.8, w: W - 1.0, h: theme.coverNavyHeight - 0.8 };
  }
  if (layout === 'title_center') {
    return { x: 1.0, y: 0.8, w: W - 2.0, h: H - 1.6 };
  }
  if (layout === 'title_split') {
    return { x: 7.2, y: 0.8, w: 5.5, h: H - 1.6 }; // Right side content area
  }
  if (layout === 'content_classic') {
    return { x: 5.0, y: 0.5, w: W - 5.5, h: H - bottomBar - 0.5 };
  }
  if (layout === 'content_top') {
    return { x: 0.5, y: 1.5, w: W - 1.0, h: H - bottomBar - 1.5 };
  }
  if (layout === 'card_3col') {
    return { x: 0.5, y: 1.6, w: W - 1.0, h: H - bottomBar - 1.6 };
  }
  if (layout === 'card_2col') {
    return { x: 1.0, y: 1.6, w: W - 2.0, h: H - bottomBar - 1.6 };
  }
  if (layout === 'comparison') {
    return { x: 0.5, y: 1.5, w: W - 1.0, h: H - bottomBar - 1.5 };
  }
  if (layout === 'timeline') {
    return { x: 0.8, y: 0.5, w: W - 1.3, h: H - bottomBar - 0.5 };
  }
  if (layout === 'data_highlight') {
    return { x: 0.8, y: 1.0, w: W - 1.3, h: H - bottomBar - 1.0 };
  }
  if (layout === 'quote_slide') {
    return { x: 1.5, y: 1.0, w: W - 2.0, h: H - bottomBar - 1.0 };
  }
  // Default: entire slide minus small margins
  return { x: 0.5, y: 0.5, w: W - 1.0, h: H - bottomBar - 0.5 };
};

// --- PPT Data Normalization (shared) ---
type JsonRecord = Record<string, unknown>;

const asRecord = (value: unknown): JsonRecord | null => {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? value as JsonRecord
    : null;
};

const asString = (value: unknown, fallback = ''): string => {
  return typeof value === 'string' ? value : fallback;
};

const asNumber = (value: unknown, fallback: number): number => {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
};

const asBoolean = (value: unknown): boolean | undefined => {
  return typeof value === 'boolean' ? value : undefined;
};

const normalizeRows = (value: unknown): string[][] | undefined => {
  if (!Array.isArray(value)) {
    return undefined;
  }

  return value
    .filter((row): row is unknown[] => Array.isArray(row) && row.length > 0)
    .map((row) => row.map((cell) => String(cell ?? '')));
};

const normalizePptElement = (value: unknown): PptSlideElement => {
  const element = asRecord(value) || {};
  const metaKeys = new Set(['kind','x','y','w','h','fontSize','color','bold','fill','align','rows','type','layout','slideIndex','fontFace','italic','underline','icon','number','radius','shadow','opacity','gradient','thickness','lineSpacing','letterSpacing']);

  let content = asString(element.content);
  if (!content) {
    content = asString(element.text)
      || asString(element.value)
      || asString(element.label)
      || asString(element.body)
      || asString(element.title)
      || asString(element.message);
  }

  if (!content) {
    for (const [key, candidate] of Object.entries(element)) {
      if (!metaKeys.has(key) && typeof candidate === 'string' && candidate.length > 0) {
        content = candidate;
        break;
      }
    }
  }

  const rows = normalizeRows(element.rows);
  const icon = asString(element.icon);
  let kind = asString(element.kind) as PptSlideElement['kind'];

  if (!kind) {
    if (rows) {
      kind = 'table';
    } else if (icon && !content) {
      kind = 'icon';
    } else if (content && content.startsWith('http') && /\.(png|jpg|jpeg|gif|svg|webp)/i.test(content)) {
      kind = 'image';
    } else if (content) {
      kind = 'text';
    } else if (element.fill) {
      kind = 'shape';
    } else {
      kind = 'text';
    }
  }

  if (kind === 'icon' && !content && icon) {
    content = icon;
  }

  return {
    kind,
    content,
    x: asNumber(element.x, 0),
    y: asNumber(element.y, 0),
    w: asNumber(element.w, 4),
    h: asNumber(element.h, 1),
    fontSize: asNumber(element.fontSize, 0) || undefined,
    color: asString(element.color) || undefined,
    bold: asBoolean(element.bold),
    fill: asString(element.fill) || undefined,
    align: element.align === 'center' || element.align === 'right' ? element.align : 'left',
    rows,
    icon: icon || undefined,
    radius: asNumber(element.radius, 0) || undefined,
    shadow: asBoolean(element.shadow),
    opacity: asNumber(element.opacity, 0) || undefined,
    gradient: asString(element.gradient) || undefined,
    thickness: asNumber(element.thickness, 0) || undefined,
    number: asNumber(element.number, 0) || undefined,
    fontFace: asString(element.fontFace) || undefined,
    lineSpacing: asNumber(element.lineSpacing, 0) || undefined,
    letterSpacing: asNumber(element.letterSpacing, 0) || undefined,
  };
};

const normalizePptSlide = (slide: unknown): PptSlide => {
  const slideRecord = asRecord(slide) || {};

  return {
    slideIndex: asNumber(slideRecord.slideIndex, 0),
    layout: asString(slideRecord.layout) || undefined,
    elements: Array.isArray(slideRecord.elements)
      ? slideRecord.elements.map(normalizePptElement)
      : [],
  };
};

const normalizePptData = (data: PptData): PptData => {
  return {
    ...data,
    slides: data.slides.map(normalizePptSlide),
  };
};

const stringifyStreamContent = (content: unknown): string => {
  if (content === undefined || content === null) {
    return '';
  }

  if (typeof content === 'string') {
    return content;
  }

  try {
    return JSON.stringify(content, null, 2);
  } catch {
    return String(content);
  }
};

const getChunkText = (chunk: { content?: unknown; raw?: string; metadata?: ResponseMetadata }): string => {
  if (typeof chunk.raw === 'string') {
    return chunk.raw;
  }

  return stringifyStreamContent(chunk.content);
};

const normalizeInlineText = (value: string): string => {
  return value.replace(/```json|```/g, '').replace(/\s+/g, ' ').trim();
};

const createThinkingSummary = (
  chunk: { type?: string; content?: unknown; raw?: string; metadata?: ResponseMetadata; label?: string },
  fallback: string,
): string => {
  if (typeof chunk.metadata?.summary === 'string' && chunk.metadata.summary) {
    return normalizeInlineText(chunk.metadata.summary).slice(0, 34);
  }

  if (typeof chunk.metadata?.backendContent === 'string' && chunk.metadata.backendContent) {
    return normalizeInlineText(chunk.metadata.backendContent).slice(0, 34);
  }

  if (typeof chunk.label === 'string' && chunk.label) {
    return normalizeInlineText(chunk.label).slice(0, 34);
  }

  if (chunk.content && typeof chunk.content === 'object' && !Array.isArray(chunk.content)) {
    const record = chunk.content as Record<string, unknown>;
    const keys = ['theme', 'purpose', 'title', 'summary', 'structure', 'label', 'message'];

    for (const key of keys) {
      if (typeof record[key] === 'string' && record[key]) {
        return normalizeInlineText(record[key]).slice(0, 34);
      }
    }
  }

  const text = normalizeInlineText(getChunkText(chunk));
  return (text || fallback).slice(0, 34);
};

const formatProcessChunkText = (chunk: { type?: string; content?: unknown; raw?: string; metadata?: ResponseMetadata }): string => {
  const text = getChunkText(chunk);

  if (!text) {
    return '';
  }

  if (chunk.type === 'analysis' || chunk.type === 'draft') {
    return `\`\`\`json\n${text}\n\`\`\``;
  }

  return text;
};

const generatePptx = (data: PptData, theme: PptTheme) => {
  const pres = new pptxgen();
  pres.title = data.title;
  pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5

  data.slides.forEach((slideData, slideIdx) => {
    const slide = pres.addSlide();
    const layout = inferLayout(slideData, slideIdx, data.slides.length);
    const isTitleSlide = layout.startsWith('title_') || slideIdx === 0;
    const isEndSlide = layout === 'end_slide' || slideIdx === data.slides.length - 1;

    // === STEP 1: Layout-specific theme decoration ===
    // Helper: add rectangle shape (with optional gradient)
    const addRect = (x: number, y: number, w: number, h: number, color: string, gradient?: { from: string; to: string }) => {
      if (gradient) {
        slide.addShape(pres.ShapeType.rect, {
          x, y, w, h,
          fill: { color: gradient.from },
          line: { width: 0 },
        });
        // Overlay with semi-transparent gradient effect using second shape
        slide.addShape(pres.ShapeType.rect, {
          x, y, w, h,
          fill: { color: gradient.to, transparency: 50 },
          line: { width: 0 },
        });
      } else {
        slide.addShape(pres.ShapeType.rect, {
          x, y, w, h, fill: { color }, line: { width: 0 },
        });
      }
    };
    // Helper: add circle shape
    const addCircle = (x: number, y: number, w: number, h: number, color: string, transparency = 0) => {
      slide.addShape(pres.ShapeType.ellipse, {
        x, y, w, h, fill: { color, transparency }, line: { width: 0 },
      });
    };

    // Gradient shortcut
    const grad = { from: theme.primary, to: theme.primaryMid };

    if (layout === 'title_classic' || layout === 'end_slide') {
      addRect(0, 0, 13.33, theme.coverNavyHeight, theme.primary, grad);
      addRect(0, theme.coverNavyHeight, 13.33, 0.12, theme.accent);
      addRect(0, 7.15, 13.33, 0.35, theme.primary, grad);
      addCircle(10.5, 0.6, 2.2, 2.2, theme.primaryLight, 70);
      addCircle(0.3, 4.8, 1.4, 1.4, theme.primaryLight, 85);
    } else if (layout === 'title_center') {
      addRect(0, 0, 13.33, 0.25, theme.primary, grad);
      addRect(0, 7.25, 13.33, 0.25, theme.primary, grad);
      addRect(0, 0.25, 13.33, 0.05, theme.accent);
      addRect(0, 7.2, 13.33, 0.05, theme.accent);
      addCircle(5.16, 2.25, 3.0, 3.0, theme.offWhite, 60);
    } else if (layout === 'title_split') {
      addRect(0, 0, 6.66, 7.5, theme.primary, grad);
      addRect(6.66, 0, 6.67, 7.5, theme.offWhite);
      addRect(6.66, 0, 0.1, 7.5, theme.accent);
      addCircle(1.0, 5.0, 1.8, 1.8, theme.primaryLight, 80);
    } else if (layout === 'card_3col') {
      addRect(0, 0, 13.33, 0.9, theme.primary, grad);
      addRect(0, 0.9, 13.33, 0.06, theme.accent);
      addRect(0.5, 1.6, 3.8, 5.0, theme.offWhite);
      addRect(4.75, 1.6, 3.8, 5.0, theme.offWhite);
      addRect(9.0, 1.6, 3.8, 5.0, theme.offWhite);
      addRect(0.5, 1.6, 3.8, 0.12, theme.primary, grad);
      addRect(4.75, 1.6, 3.8, 0.12, theme.primary, grad);
      addRect(9.0, 1.6, 3.8, 0.12, theme.primary, grad);
      addRect(0, 7.15, 13.33, 0.35, theme.primary, grad);
    } else if (layout === 'comparison') {
      addRect(0, 0, 13.33, 0.9, theme.primary, grad);
      addRect(0, 0.9, 13.33, 0.06, theme.accent);
      addRect(0.5, 1.5, 5.9, 5.2, theme.offWhite);
      addRect(0.5, 1.5, 5.9, 0.1, theme.primary, grad);
      addRect(6.9, 1.5, 5.9, 5.2, theme.offWhite);
      addRect(6.9, 1.5, 5.9, 0.1, theme.accent);
      addRect(6.55, 1.5, 0.2, 5.2, theme.primaryLight, { from: theme.primaryLight, to: theme.primaryMid });
      addRect(0, 7.15, 13.33, 0.35, theme.primary, grad);
    } else if (layout === 'timeline') {
      addRect(0, 0, 0.5, 7.5, theme.primary, grad);
      addRect(0, 7.15, 13.33, 0.35, theme.primary, grad);
      addRect(1.2, 3.5, 11.5, 0.12, theme.primaryLight);
      addCircle(2.5, 3.2, 0.7, 0.7, theme.primary);
      addCircle(5.5, 3.2, 0.7, 0.7, theme.primary);
      addCircle(8.5, 3.2, 0.7, 0.7, theme.primary);
      addCircle(11.5, 3.2, 0.7, 0.7, theme.accent);
    } else if (layout === 'data_highlight') {
      addRect(0, 0, 13.33, 0.7, theme.primary, grad);
      addRect(0, 0.7, 13.33, 0.06, theme.accent);
      addRect(0, 6.0, 13.33, 1.5, theme.offWhite);
      addRect(0, 7.15, 13.33, 0.35, theme.primary, grad);
      addRect(0.5, 1.2, 0.08, 4.5, theme.primary);
    } else if (layout === 'quote_slide') {
      addRect(0, 0, 13.33, 7.5, theme.offWhite);
      addRect(0, 0, 0.8, 7.5, theme.primary, grad);
      addRect(0.8, 2.8, 0.12, 1.8, theme.accent);
      addRect(0, 7.15, 13.33, 0.35, theme.primary, grad);
    } else if (layout === 'card_2col') {
      addRect(0, 0, 13.33, 0.9, theme.primary, grad);
      addRect(0, 0.9, 13.33, 0.06, theme.accent);
      addRect(1.0, 1.6, 5.0, 5.0, theme.offWhite);
      addRect(7.33, 1.6, 5.0, 5.0, theme.offWhite);
      addRect(1.0, 1.6, 5.0, 0.15, theme.primary, grad);
      addRect(7.33, 1.6, 5.0, 0.15, theme.primary, grad);
      addRect(0, 7.15, 13.33, 0.35, theme.primary, grad);
    } else if (layout === 'content_top') {
      addRect(0, 0, 13.33, 1.2, theme.primary, grad);
      addRect(0, 1.2, 13.33, 0.08, theme.accent);
      addRect(0, 7.15, 13.33, 0.35, theme.primary, grad);
    } else {
      addRect(0, 0, 4.5, 7.15, theme.primary, grad);
      addRect(4.5, 0, 0.08, 7.15, theme.accent);
      addRect(0, 7.15, 13.33, 0.35, theme.primary, grad);
    }

    // === STEP 2: Render AI content elements (skip all shapes) ===
    slideData.elements.forEach((el) => {
      if (el.kind === 'shape') return;

      // Safe content area clamping (shared by all element types)
      const safe = getSafeContentArea(layout, theme);

      switch (el.kind) {
        case 'text': {
          const textFontSize = el.fontSize || 18;
          const elX = Math.max(safe.x, Math.min(el.x || 0, safe.x + safe.w - (el.w || 2)));
          const elY = Math.max(safe.y, Math.min(el.y || 0, safe.y + safe.h - (el.h || 1)));
          // Position-aware color detection
          const elCenterX = elX + (el.w || 4) / 2;
          const elCenterY = elY + (el.h || 1) / 2;
          const onDark = isOnDarkArea(elCenterX, elCenterY, layout, theme);
          let textColor = onDark ? theme.white : theme.bodyColor;

          // Title text (fontSize >= 24) gets primary color on light areas
          if (textFontSize >= 24) {
            textColor = onDark ? theme.white : theme.primary;
          }
          // Large decorative text (big numbers etc): allow AI color choice
          if (el.color && textFontSize >= 30) {
            textColor = el.color;
          }

          const lines = (el.content || '').split('\n');
          const textParts = lines.map(line => {
            const trimmed = line.trim();
            const isBullet = trimmed.startsWith('\u2022') || trimmed.startsWith('-') || trimmed.startsWith('\u00b7');
            return {
              text: isBullet ? '  ' + trimmed : trimmed,
              options: {
                fontSize: textFontSize,
                color: textColor,
                bold: el.bold || false,
                breakType: 'none' as const,
                paraSpaceAfter: textFontSize < 22 ? 8 : 4,
                paraSpaceBefore: 2,
              },
            };
          });

          slide.addText(textParts, {
            x: elX,
            y: elY,
            w: el.w,
            h: el.h,
            fill: el.fill ? { color: el.fill } : undefined,
            align: el.align || 'left',
            valign: textFontSize >= 30 ? 'middle' : 'top',
            lineSpacingMultiple: textFontSize < 22 ? 1.5 : 1.2,
            fontFace: textFontSize >= 24 ? 'Microsoft YaHei' : 'Microsoft YaHei',
          });
          break;
        }

        case 'table': {
          try {
            const tblX = Math.max(safe.x, el.x || 0);
            const tblY = Math.max(safe.y, el.y || 0);
            const safeRows = Array.isArray(el.rows)
              ? el.rows.filter((row): row is string[] => Array.isArray(row) && row.length > 0)
              : [];
            if (safeRows.length > 0 && (el.w || 0) > 0) {
              const colCount = safeRows[0].length || 1;
              const tableRows = safeRows.map((row, rowIdx) =>
                row.map((cell) => ({
                  text: String(cell ?? ''),
                  options: {
                    fontSize: 12,
                    color: rowIdx === 0 ? theme.white : theme.bodyColor,
                    align: 'center' as const,
                    valign: 'middle' as const,
                    fill: { color: rowIdx === 0 ? theme.primary : (rowIdx % 2 === 0 ? theme.offWhite : theme.white) },
                    bold: rowIdx === 0,
                    border: { pt: 0.5, color: 'C0C8D4' },
                  },
                }))
              );
              slide.addTable(tableRows, {
                x: tblX, y: tblY, w: el.w, h: el.h || 2,
                border: { pt: 1, color: 'C0C8D4' },
                colW: el.w / colCount,
                rowH: 0.45,
                autoPage: true,
              });
            }
          } catch (e) {
            // Skip invalid table elements gracefully
            console.warn('Skipping invalid table element:', e);
          }
          break;
        }

        case 'image':
          try {
            const imgX = Math.max(safe.x, el.x || 0);
            const imgY = Math.max(safe.y, el.y || 0);
            slide.addImage({ path: el.content, x: imgX, y: imgY, w: el.w, h: el.h });
          } catch { /* skip */ }
          break;

        case 'icon': {
          const iconFontSize = el.fontSize || 32;
          const iconX = Math.max(safe.x, el.x || 0);
          const iconY = Math.max(safe.y, el.y || 0);
          slide.addText(el.icon || el.content || '●', {
            x: iconX, y: iconY, w: el.w, h: el.h,
            fontSize: iconFontSize,
            color: el.color || theme.primary,
            align: 'center',
            valign: 'middle',
          });
          break;
        }

        case 'divider': {
          const isH = (el.w || 0) > (el.h || 0);
          const divX = Math.max(safe.x, el.x || 0);
          const divY = Math.max(safe.y, el.y || 0);
          if (isH) {
            addRect(divX, divY, el.w || 1, el.thickness || 0.04, el.color || 'AAAAAA');
          } else {
            addRect(divX, divY, el.thickness || 0.04, el.h || 1, el.color || 'AAAAAA');
          }
          break;
        }

        case 'bullet': {
          const bulletFontSize = el.fontSize || 16;
          const numSize = Math.max(10, bulletFontSize * 0.65);
          const circleW = 0.45;
          const bulX = Math.max(safe.x, el.x || 0);
          const bulY = Math.max(safe.y, el.y || 0);
          addCircle(bulX, bulY, circleW, circleW, el.fill || theme.primary);
          slide.addText(String(el.number ?? 1), {
            x: bulX, y: bulY, w: circleW, h: circleW,
            fontSize: numSize,
            color: theme.white,
            bold: true,
            align: 'center',
            valign: 'middle',
          });
          slide.addText(el.content || '', {
            x: bulX + circleW + 0.15, y: bulY, w: (el.w || 5) - circleW - 0.15, h: el.h || 0.6,
            fontSize: bulletFontSize,
            color: el.color || theme.bodyColor,
            bold: el.bold || false,
            align: 'left',
            valign: 'middle',
            lineSpacingMultiple: 1.4,
          });
          break;
        }
      }
    });

    // === STEP 3: Page number ===
    if (!isTitleSlide && !isEndSlide) {
      slide.addText(`${slideIdx + 1} / ${data.slides.length}`, {
        x: 11.5, y: 7.15, w: 1.5, h: 0.35,
        fontSize: 10, color: theme.white, align: 'right', valign: 'middle',
      });
    }
  });

  return pres;
};

export default function PptPage() {
  const router = useRouter();
  const [currentUser, setCurrentUser] = useState('');

  // (inferLayout defined at module level above)

  // Chat State
  const [isChatOpen, setIsChatOpen] = useState(true);
  const [messages, setMessages] = useState<Message[]>([
    {
      id: '1',
      role: 'agent',
      content: '你好！我是你的 PPT 智能助手。请告诉我你想制作什么主题的演示文稿？',
      timestamp: Date.now(),
    },
  ]);
  const [inputValue, setInputValue] = useState('');
  const [isSending, setIsSending] = useState(false);
  const [chatWidth, setChatWidth] = useState(420);
  const [isResizingChat, setIsResizingChat] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const chatResizeStartXRef = useRef(0);
  const chatResizeStartWidthRef = useRef(420);

  // Stream State
  const [streamPhase, setStreamPhase] = useState<string>('');
  const [streamProgress, setStreamProgress] = useState<string>('');
  const streamAbortRef = useRef<AbortController | null>(null);

  // Agent State
  const [agents, setAgents] = useState<AiAgentConfigResponseDTO[]>([]);
  const [selectedAgentId, setSelectedAgentId] = useState('');
  const [sessionId, setSessionId] = useState('');

  // PPT Preview State
  const [pptData, setPptData] = useState<PptData | null>(null);
  const [responseMetadata, setResponseMetadata] = useState<ResponseMetadata | undefined>();
  const [currentSlideIndex, setCurrentSlideIndex] = useState(0);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [selectedThemeId, setSelectedThemeId] = useState('navy');
  const [isStylePanelOpen, setIsStylePanelOpen] = useState(false);
  const [selectedStyle, setSelectedStyle] = useState('professional');
  const [selectedStructure, setSelectedStructure] = useState('auto');
  const [selectedTone, setSelectedTone] = useState('neutral');
  const [selectedScene, setSelectedScene] = useState('general');
  const [selectedLayouts, setSelectedLayouts] = useState<string[]>(['auto']);
  const [expandedCategory, setExpandedCategory] = useState<string | null>('business');
  const [leftTab, setLeftTab] = useState<'templates' | 'slides'>('templates');

  // ===== PPT Template Library =====
  const PPT_TEMPLATES = [
    {
      category: 'business',
      icon: '💼',
      label: '商务办公',
      templates: [
        { id: 'quarterly-report', name: '季度工作汇报', desc: '项目进展、数据分析、下季规划', prompt: '请制作一份2026年Q2季度工作汇报PPT，共9页。\n\n第1页 封面：标题「2026年Q2季度工作汇报」，副标题「技术研发部」，右下角演讲者姓名和日期。\n第2页 本季概览：用3个数据高亮展示核心指标——「项目完成率 92%」「线上事故 0次」「团队满意度 4.8/5」。\n第3页 项目进展：用bullet列表展示4个重点项目——①用户中心重构（已完成80%，预计7月上线）②支付系统升级（已完成，交易成功率99.9%）③数据中台建设（进行中，完成数据采集层）④移动端性能优化（启动速度提升40%）。\n第4页 技术亮点：左右对比布局，左侧「架构优化」列出微服务拆分、缓存策略、CDN加速3个要点，右侧「工程效能」列出CI/CD流水线、自动化测试覆盖率85%、代码审查机制3个要点。\n第5页 数据回顾：用表格展示Q1与Q2核心指标对比（响应时间、可用性、日活、转化率），Q2数据用绿色高亮。\n第6页 团队成果：三列卡片，分别展示「技术突破」（3项专利申请）、「团队成长」（2人晋升、3人入职）、「开源贡献」（5个开源项目贡献）。\n第7页 问题与挑战：用bullet列出3个挑战——①历史技术债积压 ②跨团队协作效率待提升 ③监控体系不够完善，每个挑战后附一句话解决思路。\n第8页 下季规划：时间轴布局，7月完成用户中心上线、8月启动推荐系统、9月完成监控体系2.0。\n第9页 结尾页：感谢语+联系方式。', scene: 'report', style: 'professional' },
        { id: 'annual-review', name: '年度总结报告', desc: '全年回顾、亮点成果、展望未来', prompt: '请制作一份2026年度工作总结报告PPT，共10页。\n\n第1页 封面：「2026年度总结报告」，副标题「产品研发中心」，日期2026年12月。\n第2页 年度关键词：用3个大数字展示——「32个项目交付」「98.5%系统可用性」「150+团队成员」。\n第3页 里程碑时间轴：按季度展示4个里程碑——Q1产品2.0上线、Q2海外市场开拓、Q3AI功能发布、Q4用户突破1000万。\n第4页 核心业务成果：三列卡片——「产品侧」3个核心功能上线、「技术侧」架构升级与性能优化、「运营侧」用户增长与留存提升。\n第5页 技术架构演进：左右对比，左侧「年初架构」单体+MySQL，右侧「年末架构」微服务+云原生，中间用对比布局突出变化。\n第6页 数据看板：用表格展示全年12个月核心指标趋势（DAU、收入、NPS、故障数）。\n第7页 团队建设：三列卡片——「人才发展」培训体系+晋升通道、「文化凝聚」团建活动+价值观、「组织效能」OKR+复盘机制。\n第8页 挑战与反思：bullet列出3个反思——①过度设计导致部分项目延期 ②技术选型需更谨慎 ③跨部门沟通仍需加强。\n第9页 2027展望：3个战略方向——①AI原生产品化 ②全球化技术架构 ③开发者生态建设。\n第10页 结尾：感谢+座右铭。', scene: 'report', style: 'professional' },
        { id: 'project-proposal', name: '项目立项方案', desc: '项目背景、目标、计划、预算', prompt: '请制作一份智能客服系统项目立项方案PPT，共9页。\n\n第1页 封面：「智能客服系统项目立项方案」，副标题「产品技术部」，日期。\n第2页 项目背景：左侧色带放标题「为什么做」，右侧用bullet列出——①当前人工客服成本占运营支出35% ②平均响应时间8分钟，用户满意度仅72% ③竞品已上线AI客服，我们落后2个季度。\n第3页 项目目标：3个数据高亮——「响应时间 <30秒」「人工成本降低50%」「用户满意度提升至90%」。\n第4页 方案设计：三列卡片——「对话引擎」基于大模型的意图识别+多轮对话，「知识库」产品FAQ+工单历史自动学习，「人工兜底」复杂问题无缝转人工。\n第5页 技术架构：左右对比，左侧「前端」Web+App+小程序多端接入，右侧「后端」LLM服务+向量检索+规则引擎三层架构。\n第6页 实施计划：时间轴——第1月需求确认+技术选型、第2-3月核心开发、第4月内部测试、第5月灰度发布、第6月全量上线。\n第7页 资源需求：表格展示人员配置（2前端+3后端+1算法+1产品+1测试）、硬件资源（GPU服务器2台）、外部依赖（LLM API费用）。\n第8页 风险评估：bullet列出3个风险及应对——①大模型幻觉风险→人工兜底+审核机制 ②数据隐私风险→本地化部署+脱敏 ③用户接受度风险→渐进式上线+反馈收集。\n第9页 结尾：项目slogan+联系方式。', scene: 'report', style: 'professional' },
        { id: 'team-intro', name: '团队介绍', desc: '团队成员、职能分工、协作模式', prompt: '请制作一份技术研发团队介绍PPT，共7页。\n\n第1页 封面：「技术研发团队介绍」，副标题「创新驱动·技术赋能」，团队Logo区域。\n第2页 团队概览：3个数据高亮——「50+工程师」「8个技术方向」「3地协作办公」。\n第3页 组织架构：三列卡片——「平台工程组」负责基础架构和DevOps、「业务研发组」负责核心产品开发、「算法团队」负责AI和推荐系统。\n第4页 核心能力：用bullet列出4个能力——①云原生架构设计与实施 ②高并发系统性能优化 ③大模型应用开发 ④全栈工程能力（前端+后端+移动端）。\n第5页 协作方式：左右对比——左侧「敏捷开发」2周迭代+每日站会+迭代演示，右侧「质量保障」代码审查+自动化测试+持续部署。\n第6页 技术成果：三列卡片展示年度亮点——「开源贡献」5个开源项目+2000+ GitHub Stars、「技术专利」12项发明专利、「技术分享」40+场内部分享+10+场外部演讲。\n第7页 结尾：团队slogan+招聘邮箱。', scene: 'general', style: 'professional' },
      ],
    },
    {
      category: 'roadshow',
      icon: '🚀',
      label: '路演融资',
      templates: [
        { id: 'startup-pitch', name: '创业融资路演', desc: '痛点、方案、市场、团队、融资', prompt: '请制作一份AI+教育赛道创业融资路演PPT，共11页。\n\n第1页 封面：「智学AI——让每个孩子都有专属AI老师」，公司名+Logo区+融资轮次（Pre-A轮）。\n第2页 痛点洞察：三列卡片——「教育资源不均」一线城市与三四线师资差距3倍、「学习效率低」传统一对一每小时300元、「个性化缺失」千人一面无法因材施教。\n第3页 解决方案：左右对比布局——左侧「传统方式」固定课程+统一进度+被动学习，右侧「智学AI」自适应学习路径+实时反馈+主动探索，中间用divider和VS标识。\n第4页 产品展示：用bullet列出3个核心功能——①AI学情诊断：5分钟精准定位知识薄弱点 ②自适应练习：基于能力动态调整题目难度 ③AI答疑助手：24小时个性化答疑，解题思路而非答案。\n第5页 商业模式：三列卡片——「B2C」个人订阅￥99/月、「B2B」学校合作￥5万/年/校、「B2G」教育局采购区域授权。\n第6页 市场规模：3个数据高亮——「TAM 6000亿」中国K12教育市场、「SAM 1200亿」在线教育细分、「SOM 50亿」AI+教育可触达市场。\n第7页 增长数据：数据高亮页——「注册用户 50万+」「月活增长 180%」「续费率 78%」「NPS 72」。\n第8页 竞争优势：左右对比我方vs竞品——我方「自适应算法领先」「全学科覆盖」「公立校渠道」，竞品「题库驱动」「单科突破」「纯线上C端」。\n第9页 团队介绍：三列卡片——「创始人」前字节跳动AI Lab负责人，清华CS博士、「CTO」前好未来技术总监，10年教育科技经验、「教研负责人」省特级教师，20年一线教学经验。\n第10页 融资计划：bullet列出——融资金额2000万元、估值1亿元、资金用途（40%研发、30%市场、20%团队、10%运营）、18个月跑道。\n第11页 结尾：公司愿景「让AI成为每个孩子的专属老师」+联系方式。', scene: 'pitch', style: 'creative' },
        { id: 'investor-deck', name: '投资人推介', desc: '核心数据、增长曲线、竞品分析', prompt: '请制作一份面向投资人的数据驱动型推介PPT，共9页。\n\n第1页 封面：「数据驱动的增长故事」，公司名+核心Slogan。\n第2页 核心指标：4个数据高亮——「ARR ¥8000万」「YoY增长 156%」「净收入留存 125%」「LTV/CAC 4.2x」。\n第3页 增长飞轮：三列卡片展示——「更多用户」→数据积累→「更好产品」→口碑传播→循环加速。\n第4页 收入模型：用表格展示3条产品线收入（基础版/专业版/企业版）的定价、用户数、收入占比。\n第5页 单位经济：数据高亮——「CAC ¥1200」「LTV ¥5100」「回本周期 8个月」「毛利率 78%」。\n第6页 市场定位：左右对比——左侧「我们」产品驱动增长+自助式上手+高净留存，右侧「竞品」销售驱动+重实施+低留存。\n第7页 增长策略：bullet列出3个增长引擎——①产品自增长：免费版→付费版转化漏斗优化 ②渠道拓展：行业解决方案+ISV生态 ③国际化：东南亚市场先发优势。\n第8页 未来预测：表格展示未来3年财务预测（收入、毛利率、用户数、市场份额）。\n第9页 结尾：融资需求摘要+联系方式。', scene: 'pitch', style: 'professional' },
      ],
    },
    {
      category: 'education',
      icon: '📚',
      label: '培训教学',
      templates: [
        { id: 'tech-sharing', name: '技术分享', desc: '技术方案、架构、最佳实践', prompt: '请制作一份微服务架构演进技术分享PPT，共10页。\n\n第1页 封面：「从单体到微服务——架构演进实战分享」，演讲者姓名+职位+日期。\n第2页 目录：用bullet列出4个章节——①为什么需要微服务 ②架构演进路径 ③踩过的坑与经验 ④未来展望。\n第3页 背景：左侧色带标题「业务痛点」，右侧bullet列出——①代码仓库10万+行，编译需15分钟 ②一个模块故障全站挂 ③团队8人同时改代码冲突频繁 ④新功能上线周期2周+。\n第4页 演进路径：时间轴——Phase1服务拆分（按业务域拆4个服务）→Phase2基础设施（注册中心+配置中心+网关）→Phase3可观测性（链路追踪+指标监控+日志聚合）→Phase4云原生（容器化+K8s+自动扩缩容）。\n第5页 服务拆分策略：左右对比——左侧「按技术层拆」（❌反面教材）导致分布式单体，右侧「按业务域拆」（✅正确做法）用户域/订单域/支付域/商品域各自独立。\n第6页 核心组件：三列卡片——🔧「服务注册与发现」Nacos+健康检查+负载均衡，⚡「API网关」统一入口+鉴权+限流+熔断，📊「可观测性」SkyWalking+Prometheus+ELK。\n第7页 踩坑经验：bullet列出3个血泪教训——①分布式事务：最终一致性>强一致性，用Saga模式 ②服务间循环依赖：引入事件总线解耦 ③配置管理混乱：统一配置中心+灰度发布。\n第8页 成果数据：3个数据高亮——「部署频率 从月更→日更」「故障恢复 从4小时→15分钟」「系统可用性 从99.5%→99.99%」。\n第9页 未来方向：三列卡片——「Service Mesh」Istio+Envoy下沉基础设施、「Serverless」事件驱动+按需弹缩、「AI赋能」智能容量预测+异常检测。\n第10页 结尾：Q&A+参考资源链接。', scene: 'training', style: 'academic' },
        { id: 'course-lecture', name: '课程讲义', desc: '知识体系、核心概念、案例', prompt: '请制作一份「认知心理学导论」课程讲义PPT，共10页。\n\n第1页 封面：「认知心理学导论——理解人类思维的科学」，教师姓名+课程编号+学期。\n第2页 课程概览：3个数据高亮展示课程框架——「3大核心理论」「6个经典实验」「12个应用案例」。\n第3页 核心概念：三列卡片——🧠「认知负荷理论」工作记忆容量7±2，教学需降低外在负荷，🔬「建构主义」学习者主动构建知识，教师是引导者，🎯「最近发展区」学生独立水平与潜在水平的差距是教学最佳区。\n第4页 信息加工模型：左侧色带标题「人类信息加工」，右侧用bullet列出流程——感觉登记（<1秒）→注意选择（过滤机制）→工作记忆（7±2组块）→长时记忆（无限容量）。每个阶段加emoji图标。\n第5页 经典实验：左右对比——左侧「Stroop效应」色词干扰实验说明自动化加工，右侧「Miller定律」7±2组块理论说明记忆容量限制。\n第6页 记忆原理：三列卡片——「编码」深度加工>浅层加工（Craik & Lockhart），「存储」间隔重复>集中学习（Ebbinghaus遗忘曲线），「提取」情境依赖+状态依赖影响回忆效果。\n第7页 注意力机制：bullet列出3个要点——①选择性注意：鸡尾酒会效应 ②持续性注意：警戒任务中的衰减 ③分配性注意：多任务处理的代价。\n第8页 应用案例：三列卡片——「教育领域」基于认知负荷的教学设计、「UI设计」符合注意规律的界面布局、「人工智能」认知启发的AI架构。\n第9页 本章小结：bullet编号列表——①认知心理学研究人类信息加工过程 ②工作记忆是瓶颈，7±2是关键 ③深度加工促进长时记忆 ④注意力资源有限需合理分配。\n第10页 结尾：思考题+下节预告+参考文献。', scene: 'training', style: 'academic' },
      ],
    },
    {
      category: 'data',
      icon: '📊',
      label: '数据报告',
      templates: [
        { id: 'data-dashboard', name: '数据看板报告', desc: '关键指标、趋势分析、预警', prompt: '请制作一份电商平台月度数据看板报告PPT，共8页。\n\n第1页 封面：「2026年5月电商平台数据月报」，副标题「数据智能部」。\n第2页 核心指标：4个数据高亮——「GMV ¥2.8亿」「日均订单 12.5万」「客单价 ¥224」「转化率 4.8%」。\n第3页 流量分析：三列卡片——「UV 5800万」环比+12%，「PV 3.2亿」环比+8%，「人均浏览5.5页」环比持平。\n第4页 用户分析：左右对比——左侧「新用户」占比35%，获客成本¥45，首单转化率28%；右侧「老用户」占比65%，复购率42%，客单价高出新用户60%。\n第5页 品类表现：用表格展示Top5品类（服饰/3C/食品/美妆/家居）的GMV、订单量、增长率、毛利率。\n第6页 营销效果：bullet列出3个重点活动——①5.1大促GMV¥4500万，ROI 1:8.5 ②会员日活动带动老客复购率+15% ③短视频渠道新增用户占比提升至22%。\n第7页 风险预警：三列卡片——⚠️「退货率上升」服饰类退货率18%需关注，📉「搜索转化下降」首页改版后搜索转化降0.3%，🔄「库存周转放缓」部分品类周转天数>30天。\n第8页 下月重点：bullet列出——①优化搜索算法提升转化 ②服饰品类退货率专项治理 ③短视频渠道加大投入 ④会员体系2.0上线。', scene: 'report', style: 'professional' },
        { id: 'competitive-analysis', name: '竞品分析', desc: '竞品对比、差异化、策略建议', prompt: '请制作一份AI编程助手竞品分析PPT，共8页。\n\n第1页 封面：「AI编程助手竞品分析报告」，副标题「产品战略部·2026年6月」。\n第2页 市场格局：三列卡片——「海外选手」GitHub Copilot+Cursor+Codeium，「国内选手」通义灵码+百度Comate+CodeGeeX，「新兴玩家」Windsurf+Augment+Cody。\n第3页 功能对比：用表格展示6个竞品在代码补全/对话式编程/多文件编辑/项目理解/私有化部署5个维度的支持情况（✅/⚠️/❌）。\n第4页 核心差异：左右对比——左侧「Copilot优势」生态整合+代码质量+用户基数，右侧「国内产品优势」中文理解+私有化+合规+价格。\n第5页 用户口碑：三列卡片——⭐「易用性」Copilot领先，即装即用，📊「准确率」Cursor的Agent模式代码通过率最高，💰「性价比」Codeium免费版功能最全。\n第6页 SWOT分析：bullet列出我方——S:国内合规+中文优化 W:技术积累不足 O:企业私有化需求爆发 T:Copilot全球扩张。\n第7页 策略建议：bullet列出3个方向——①差异化：深耕企业私有化场景 ②技术追赶：Agent模式+多文件编辑能力 ③生态建设：IDE插件+API开放平台。\n第8页 结尾：总结+下一步行动计划。', scene: 'report', style: 'professional' },
      ],
    },
    {
      category: 'creative',
      icon: '🎨',
      label: '创意设计',
      templates: [
        { id: 'brand-story', name: '品牌故事', desc: '品牌起源、价值观、愿景', prompt: '请制作一份品牌故事PPT，品牌名「茶里物语」——新中式茶饮品牌，共7页。\n\n第1页 封面：「茶里物语——一杯茶里的东方故事」，品牌Slogan「以茶为媒，链接古今」。\n第2页 品牌起源：左侧色带标题「缘起」，右侧讲述——创始人走访中国20座茶山，发现好茶只出口不内销，决心让国人喝到中国好茶。\n第3页 品牌理念：三列卡片——🍃「寻源」每款茶可溯源至具体茶山和制茶师，🫖「传承」古法制茶+现代萃取技术，🌸「创新」茶+花+果的东方风味融合。\n第4页 产品哲学：左右对比——左侧「传统茶饮」仪式感强但门槛高，右侧「茶里物语」保留仪式感+降低门槛+年轻人友好。\n第5页 空间体验：bullet列出3个设计理念——①新中式空间：竹+木+石材的东方美学 ②互动体验：手作茶饮+茶艺表演 ③社交场景：适合拍照分享的角落设计。\n第6页 品牌愿景：数据高亮——「3年100店」「覆盖20城」「年服务1000万杯」。\n第7页 结尾：品牌Slogan+招商联系方式。', scene: 'general', style: 'creative' },
        { id: 'event-plan', name: '活动策划方案', desc: '活动主题、流程、预算', prompt: '请制作一份「2026公司年度技术峰会」活动策划方案PPT，共8页。\n\n第1页 封面：「TechForward 2026——年度技术峰会策划方案」，主办方+日期+地点。\n第2页 活动定位：3个数据高亮——「500+参会者」「20+演讲嘉宾」「2天沉浸体验」。\n第3页 活动亮点：三列卡片——🎤「顶级嘉宾」CTO/技术VP/开源维护者，🛠️「动手工坊」AI开发+云原生+前端实战，🤝「社交场景」技术社区晚宴+闪电演讲。\n第4页 议程安排：时间轴展示Day1——09:00开场→10:00主题演讲（3场）→14:00分论坛（AI/架构/前端3个track）→17:00圆桌对话→19:00晚宴。\n第5页 场地规划：左右对比——左侧「主会场」容纳500人+直播设备+同声传译，右侧「分会场」3个200人分会场+工坊区。\n第6页 宣传策略：bullet列出——①提前60天启动倒计时海报 ②技术社区KOL合作推广 ③往届精彩回顾短视频 ④早鸟票+团队票组合优惠。\n第7页 预算概览：用表格展示各项费用（场地/设备/嘉宾/餐饮/宣传/物料）的预算金额和占比。\n第8页 结尾：策划团队+联系方式+报名二维码占位。', scene: 'general', style: 'creative' },
      ],
    },
    {
      category: 'personal',
      icon: '🌟',
      label: '个人成长',
      templates: [
        { id: 'resume', name: '个人简历', desc: '教育、经历、技能、项目', prompt: '请制作一份高级Java工程师个人简历PPT，共6页。\n\n第1页 封面：姓名「张三」，职位「高级Java工程师 / 技术负责人」，联系方式（邮箱+手机+GitHub），一句话定位「8年Java架构经验，专注高并发系统设计与微服务治理」。\n第2页 个人简介：左侧色带放头像占位区和姓名，右侧——5年一线互联网经验+3年技术管理经验，主导过日活千万级系统架构设计，带领15人团队完成3次重大架构升级。\n第3页 工作经历：时间轴布局——2024-至今「某大厂」技术负责人，负责推荐系统架构升级；2021-2024「某独角兽」高级工程师，主导订单系统微服务化；2018-2021「某上市公司」Java工程师，参与核心交易系统开发。\n第4页 核心技能：三列卡片——💻「技术栈」Java/Spring Boot/MySQL/Redis/Kafka/ES/Docker/K8s，🏗️「架构能力」微服务设计/分布式系统/高并发方案/DDD领域驱动，🚀「工程效能」CI/CD/自动化测试/代码质量/性能调优。\n第5页 项目成果：数据高亮——「系统可用性 99.99%」「QPS峰值 50万+」「响应时间 P99<100ms」「团队效率提升 3倍」。\n第6页 结尾：教育背景（XX大学·计算机科学·硕士）+一句话「代码改变世界，架构成就未来」。', scene: 'general', style: 'minimal' },
        { id: 'year-review', name: '个人年度回顾', desc: '成就、学习、新年计划', prompt: '请制作一份2026年个人年度回顾PPT，共7页。\n\n第1页 封面：「我的2026——年度回顾与展望」，姓名+年份。\n第2页 年度关键词：3个数据高亮——「12本书」「6个新技能」「3次突破」。\n第3页 职业成长：时间轴——Q1完成架构师认证、Q2主导核心系统重构、Q3在技术大会做演讲、Q4晋升技术总监。\n第4页 技能收获：三列卡片——📖「技术能力」掌握K8s+Service Mesh+LLM应用开发，🎯「管理能力」团队从8人扩展到20人+建立OKR体系，💡「软技能」公众演讲+跨部门协作+向上管理。\n第5页 年度好书：bullet列出3本——①《凤凰项目》理解DevOps本质 ②《系统设计面试》架构思维升级 ③《纳瓦尔宝典》重新理解财富与幸福。\n第6页 2027计划：三列卡片——🚀「职业」技术VP目标+建立技术影响力，📚「学习」深度学习+系统设计+英语提升，💪「生活」健身100次+旅行3座城市+读15本书。\n第7页 结尾：年度座右铭「Stay hungry, stay foolish」。', scene: 'general', style: 'creative' },
      ],
    },
  ];

  // Style Options
  const STYLE_OPTIONS = [
    { id: 'professional', label: '👔 商务' },
    { id: 'creative', label: '🎨 创意' },
    { id: 'academic', label: '🎓 学术' },
    { id: 'minimal', label: '✨ 极简' },
  ];
  const STRUCTURE_OPTIONS = [
    { id: 'auto', label: '🤖 自动' },
    { id: 'title-only', label: '📋 纯标题' },
    { id: 'full', label: '📖 完整' },
  { id: 'concise', label: '⚡ 精简' },
  ];
  const TONE_OPTIONS = [
    { id: 'neutral', label: '⚖️ 中性' },
    { id: 'warm', label: '🔥 暖色' },
    { id: 'cool', label: '❄️ 冷色' },
    { id: 'vivid', label: '🌈 鲜艳' },
  ];
  const SCENE_OPTIONS = [
    { id: 'general', label: '📋 通用' },
    { id: 'report', label: '📊 汇报' },
    { id: 'pitch', label: '🚀 路演' },
    { id: 'training', label: '📚 培训' },
  ];
  const LAYOUT_OPTIONS = [
    { id: 'auto', label: '自动', desc: 'AI自动选择' },
    { id: 'title_classic', label: '经典封面', desc: '上下结构封面' },
    { id: 'title_center', label: '居中封面', desc: '简约边框居中' },
    { id: 'title_split', label: '左右封面', desc: '50/50双色划分' },
    { id: 'content_classic', label: '经典内容', desc: '左侧深色侧边栏' },
    { id: 'content_top', label: '顶栏内容', desc: '顶部横向深色条' },
    { id: 'card_3col', label: '三栏卡片', desc: '3列信息块' },
    { id: 'card_2col', label: '双栏卡片', desc: '2列大信息块' },
    { id: 'comparison', label: '对比布局', desc: '左右对称比较' },
    { id: 'data_highlight', label: '数据高亮', desc: '突出核心数据' },
    { id: 'timeline', label: '时间线', desc: '流程/时间节点' },
  ];
  const activeTheme = THEMES.find(t => t.id === selectedThemeId) || DEFAULT_THEME;

  // Session State
  const [sessions, setSessions] = useState<Session[]>([]);
  const [currentSessionId, setCurrentSessionId] = useState<string | null>(null);

  // Custom API Config State
  const [showApiConfig, setShowApiConfig] = useState(false);
  const [customModels, setCustomModels] = useState<CustomModelConfig[]>([]);
  const [selectedCustomModelId, setSelectedCustomModelId] = useState<string>('default');
  const [editingModel, setEditingModel] = useState<CustomModelConfig | null>(null);

  // Rename State (missing in original, adding for consistency if needed, but wait, if it's not used I won't add it to avoid errors)

  // Auto-scroll
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  useEffect(() => {
    if (!isResizingChat) {
      return;
    }

    const handleMouseMove = (event: MouseEvent) => {
      const nextWidth = chatResizeStartWidthRef.current + chatResizeStartXRef.current - event.clientX;
      setChatWidth(Math.min(720, Math.max(340, nextWidth)));
    };

    const handleMouseUp = () => {
      setIsResizingChat(false);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };

    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);

    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };
  }, [isResizingChat]);

  // Check Login & Load Agents
  useEffect(() => {
    const userInfo = getUserInfo();
    if (!userInfo || !userInfo.user) {
      router.push('/login');
      return;
    }
    setCurrentUser(userInfo.user);

    // Load Custom Models
    const savedModels = localStorage.getItem('ai_agent_custom_models');
    if (savedModels) {
      try {
        setCustomModels(JSON.parse(savedModels));
      } catch (e) {}
    }
    const savedSelected = localStorage.getItem('ai_agent_selected_model');
    if (savedSelected) {
      setSelectedCustomModelId(savedSelected);
    }

    const loadAgents = async () => {
      try {
        const res = await agentApi.queryAiAgentConfigList();
        setAgents(res.data || []);
        setSelectedAgentId(PPT_AGENT_ID);
      } catch (error) {
        console.error('Failed to load agents:', error);
      }
    };
    loadAgents();
  }, [router]);

  // Load sessions from localStorage
  useEffect(() => {
    const savedSessions = localStorage.getItem('ppt_sessions');
    if (savedSessions) {
      try {
        const parsed = JSON.parse(savedSessions);
        setSessions(parsed);
        if (parsed.length > 0) {
          const mostRecent = parsed.sort((a: Session, b: Session) => b.lastModified - a.lastModified)[0];
          setCurrentSessionId(mostRecent.id);
          setMessages(mostRecent.messages);
          if (mostRecent.pptData) {
            setPptData(mostRecent.pptData);
          }
          setResponseMetadata(mostRecent.metadata);
        } else {
          createNewSession(true);
        }
      } catch {
        createNewSession(true);
      }
    } else {
      createNewSession(true);
    }
  }, []);

  // Save sessions
  useEffect(() => {
    if (sessions.length > 0) {
      localStorage.setItem('ppt_sessions', JSON.stringify(sessions));
    }
  }, [sessions]);

  // Update session messages
  useEffect(() => {
    if (currentSessionId) {
      setSessions((prev) =>
        prev.map((session) => {
          if (session.id === currentSessionId) {
            return {
              ...session,
              messages,
              backendSessionId: sessionId,
              title:
                session.title === 'New PPT' && messages.find((m) => m.role === 'user')
                  ? messages.find((m) => m.role === 'user')?.content.slice(0, 20) || 'New PPT'
                  : session.title,
            };
          }
          return session;
        })
      );
    }
  }, [messages, currentSessionId, sessionId]);

  const createNewSession = (isInitial = false, backendId = '') => {
    const newSession: Session = {
      id: Date.now().toString(),
      backendSessionId: backendId,
      title: 'New PPT',
      messages: [
        {
          id: Date.now().toString(),
          role: 'agent',
          content: '你好！我是你的 PPT 智能助手。请告诉我你想制作什么主题的演示文稿？',
          timestamp: Date.now(),
        },
      ],
      pptData: null,
      metadata: undefined,
      lastModified: Date.now(),
    };

    setSessions((prev) => [newSession, ...prev]);
    setCurrentSessionId(newSession.id);
    setMessages(newSession.messages);
    setSessionId(backendId);
    setPptData(null);
    setResponseMetadata(undefined);
    setCurrentSlideIndex(0);
  };

  const handleSwitchSession = (targetSessionId: string) => {
    if (targetSessionId === currentSessionId) return;
    const session = sessions.find((s) => s.id === targetSessionId);
    if (session) {
      setCurrentSessionId(targetSessionId);
      setMessages(session.messages);
      setSessionId(session.backendSessionId || '');
      setPptData(session.pptData);
      setResponseMetadata(session.metadata);
      setCurrentSlideIndex(0);
    }
  };

  const saveCustomModels = (models: CustomModelConfig[]) => {
    setCustomModels(models);
    localStorage.setItem('ai_agent_custom_models', JSON.stringify(models));
  };

  const handleAddNewModel = () => {
    setEditingModel({
      id: Date.now().toString(),
      name: '新模型',
      baseUrl: 'https://api.openai.com',
      apiKey: '',
      model: 'gpt-4o',
      completionsPath: 'v1/chat/completions',
      enabled: true
    });
  };

  const handleSaveEditingModel = () => {
    if (!editingModel) return;
    const exists = customModels.some(m => m.id === editingModel.id);
    let newModels;
    if (exists) {
      newModels = customModels.map(m => m.id === editingModel.id ? editingModel : m);
    } else {
      newModels = [...customModels, editingModel];
    }
    saveCustomModels(newModels);
    setSelectedCustomModelId(editingModel.id);
    localStorage.setItem('ai_agent_selected_model', editingModel.id);
    setEditingModel(null);
  };

  const handleDeleteModel = (id: string) => {
    const newModels = customModels.filter(m => m.id !== id);
    saveCustomModels(newModels);
    if (selectedCustomModelId === id) {
      setSelectedCustomModelId('default');
      localStorage.setItem('ai_agent_selected_model', 'default');
    }
  };

  const handleDeleteSession = (e: React.MouseEvent, sessionIdToDelete: string) => {
    e.stopPropagation();
    const newSessions = sessions.filter((s) => s.id !== sessionIdToDelete);
    setSessions(newSessions);
    localStorage.setItem('ppt_sessions', JSON.stringify(newSessions));
    if (currentSessionId === sessionIdToDelete) {
      if (newSessions.length > 0) {
        handleSwitchSession(newSessions[0].id);
      } else {
        createNewSession();
      }
    }
  };

  const handleLogout = () => {
    clearUserInfo();
    router.push('/login');
  };

  const handleAgentChange = () => {
    const newAgentId = PPT_AGENT_ID;
    setSelectedAgentId(newAgentId);
    setSessionId('');
  };

  const handleToggleThinking = (messageId: string) => {
    setMessages((prev) =>
      prev.map((message) =>
        message.id === messageId
          ? { ...message, thinkingOpen: !message.thinkingOpen }
          : message
      )
    );
  };

  const handleStartChatResize = (event: React.MouseEvent<HTMLButtonElement>) => {
    chatResizeStartXRef.current = event.clientX;
    chatResizeStartWidthRef.current = chatWidth;
    setIsResizingChat(true);
  };

  const handleStopStream = () => {
    if (streamAbortRef.current) {
      streamAbortRef.current.abort();
      streamAbortRef.current = null;
    }
    setIsSending(false);
    setStreamPhase('');
    setStreamProgress('');
    setMessages(prev => [...prev, {
      id: Date.now().toString(),
      role: 'agent',
      content: '⚠️ 已停止生成。',
      timestamp: Date.now()
    }]);
  };

  const handleSendMessage = async () => {
    if (!inputValue.trim() || isSending) return;

    const content = inputValue;
    setInputValue('');
    setIsSending(true);
    
    const textarea = document.querySelector('textarea');
    if (textarea) textarea.style.height = '80px';

    // Build style hints from control panel selections
    const styleHints: string[] = [];
    const styleLabel = STYLE_OPTIONS.find(o => o.id === selectedStyle)?.label || selectedStyle;
    const structureLabel = STRUCTURE_OPTIONS.find(o => o.id === selectedStructure)?.label || selectedStructure;
    const toneLabel = TONE_OPTIONS.find(o => o.id === selectedTone)?.label || selectedTone;
    const sceneLabel = SCENE_OPTIONS.find(o => o.id === selectedScene)?.label || selectedScene;
    styleHints.push(`风格=${styleLabel}`);
    styleHints.push(`结构=${structureLabel}`);
    styleHints.push(`色调=${toneLabel}`);
    styleHints.push(`场景=${sceneLabel}`);
    if (!selectedLayouts.includes('auto') && selectedLayouts.length > 0) {
      const layoutLabels = selectedLayouts.map(id => LAYOUT_OPTIONS.find(o => o.id === id)?.label || id).join('、');
      styleHints.push(`布局=${layoutLabels}`);
    }
    const enrichedContent = `[设计指令: ${styleHints.join(', ')}]

${content}`;

    if (!selectedAgentId) {
      setMessages((prev) => [
        ...prev,
        {
          id: Date.now().toString(),
          role: 'agent',
          content: '请先选择一个智能体。',
          timestamp: Date.now(),
        },
      ]);
      setIsSending(false);
      return;
    }

    const userMsg: Message = {
      id: Date.now().toString(),
      role: 'user',
      content,
      timestamp: Date.now(),
    };
    
    const agentMsgId = Date.now().toString() + '-agent';
    const initialAgentMsg: Message = {
      id: agentMsgId,
      role: 'agent',
      content: '',
      reasoning: '',
      steps: [],
      timestamp: Date.now()
    };
    
    setMessages((prev) => [...prev, userMsg, initialAgentMsg]);

    try {
      // Ensure session
      let activeBackendSessionId = sessionId;
      if (!activeBackendSessionId) {
        const sessionRes = await agentApi.createSession(selectedAgentId, currentUser);
        activeBackendSessionId = sessionRes.data.sessionId;
        setSessionId(activeBackendSessionId);
      }
      
      setSessions(prev => prev.map(session => {
        if (session.id === currentSessionId) {
          return { ...session, lastModified: Date.now() };
        }
        return session;
      }));

      let renderedSlideCount = 0;
      const currentSessionIdRef = currentSessionId;
      let workingPpt: PptData = { title: 'PPT', slides: [] };
      let accumulatedContent = '';
      const accumulatedSteps: MessageStep[] = [];
      let stepSeq = 0;
      let activeStepKey = '';

      setStreamPhase('analyzing');
      setStreamProgress('正在分析需求...');
      
      const phaseLabel: Record<string, string> = {
        analyzing: '分析需求',
        drawing: '绘制内容',
        generating: '生成内容',
        reviewing: '检查优化',
        thinking: '思考中',
        done: '完成',
        error: '异常',
      };

      const updateAgentMessage = () => {
        setMessages(prev => prev.map(m => m.id === agentMsgId ? {
          ...m,
          content: accumulatedContent,
          steps: [...accumulatedSteps],
          thinkingOpen: m.thinkingOpen ?? false,
        } : m));
      };

      const updateSessionPpt = (nextPpt: PptData, metadata?: ResponseMetadata) => {
        setSessions((prev) =>
          prev.map((session) =>
            session.id === currentSessionIdRef
              ? {
                  ...session,
                  pptData: nextPpt,
                  metadata,
                  lastModified: Date.now(),
                }
              : session
          )
        );
      };

      const getStepGroupKey = (phaseStr: string, eventName?: string) => {
          if (eventName === 'error' || eventName === 'message') {
              return `${phaseStr}:${eventName}`;
          }

          if (eventName?.startsWith('render_')) {
              return `${phaseStr}:render`;
          }

          if (eventName?.startsWith('process_')) {
              return `${phaseStr}:process`;
          }

          return phaseStr || eventName || 'thinking';
      };

      const mergeStepContent = (current: string, next: string, eventName?: string) => {
          if (!current) {
              return next;
          }

          const shouldAppendInline =
              eventName === 'process_delta' &&
              next.length <= 80 &&
              !next.includes('\n') &&
              !next.includes('```') &&
              !current.endsWith('```');

          return `${current}${shouldAppendInline ? '' : '\n\n'}${next}`;
      };

      const upsertStep = (
        phaseStr: string,
        phaseText: string,
        content: string,
        status: MessageStep['status'] = 'running',
        author?: string,
        eventName?: string,
        summary?: string,
      ) => {
          const stepKey = getStepGroupKey(phaseStr, eventName);
          const activeStep = accumulatedSteps[accumulatedSteps.length - 1];

          if (activeStep && activeStepKey === stepKey) {
              activeStep.content = mergeStepContent(activeStep.content, content, eventName);
              activeStep.event = eventName;
              activeStep.status = status;
              if (summary && (!activeStep.summary || activeStep.summary === phaseText)) {
                  activeStep.summary = summary;
              }
              return;
          }

          accumulatedSteps.forEach(s => { if (s.status === 'running') s.status = 'done'; });
          activeStepKey = stepKey;

          const cleanContent = normalizeInlineText(content);
          accumulatedSteps.push({
              id: `${Date.now()}-${stepSeq}`,
              phase: stepKey,
              label: author ? `${phaseText} · ${author}` : phaseText,
              summary: summary || cleanContent.slice(0, 34) || phaseText,
              author,
              event: eventName,
              content,
              status,
              open: false,
          });
          stepSeq += 1;
      };

      const inferEventName = (chunkType: string) => {
        if (chunkType === 'ppt') return 'render_result';
        if (chunkType === 'ppt_slide') return 'render_delta';
        if (chunkType === 'user') return 'message';
        if (chunkType === 'error') return 'error';
        if (chunkType === 'done' || chunkType === 'ppt_done') return 'done';
        return 'process_delta';
      };

      const upsertSlide = (slideContent: unknown) => {
        const normalizedSlide = normalizePptSlide(slideContent);
        const slideIndex = normalizedSlide.slideIndex || workingPpt.slides.length + 1;
        const nextSlide = { ...normalizedSlide, slideIndex };
        const nextSlides = [...workingPpt.slides];
        const existingIndex = nextSlides.findIndex(slide => slide.slideIndex === slideIndex);

        if (existingIndex >= 0) {
          nextSlides[existingIndex] = nextSlide;
        } else {
          nextSlides.push(nextSlide);
        }

        nextSlides.sort((a, b) => a.slideIndex - b.slideIndex);
        workingPpt = { ...workingPpt, slides: nextSlides };
        renderedSlideCount = nextSlides.length;
        setPptData(workingPpt);
        setLeftTab('slides');
        setStreamProgress(`已渲染 ${renderedSlideCount} 页`);
        updateSessionPpt(workingPpt, responseMetadata);
      };

      const activeModelConfig = customModels.find(m => m.id === selectedCustomModelId && m.enabled);

      const controller = await agentApi.chatStream(
        {
          agentId: selectedAgentId,
          userId: currentUser,
          sessionId: activeBackendSessionId,
          message: enrichedContent,
          customBaseUrl: activeModelConfig?.baseUrl || undefined,
          customApiKey: activeModelConfig?.apiKey || undefined,
          customCompletionsPath: activeModelConfig?.completionsPath || undefined,
          customModel: activeModelConfig?.model || undefined
        },
        (streamEvent) => {
          const { phase, chunk } = streamEvent;
          const chunkRecord = chunk as {
            type: string;
            content?: unknown;
            metadata?: ResponseMetadata;
            raw?: string;
            slide?: PptSlide;
          };
          const eventName = streamEvent.event || inferEventName(chunkRecord.type);
          const isRenderable = streamEvent.renderable ?? eventName.startsWith('render_');
          const currentPhaseLabel = phaseLabel[phase] || phaseLabel.thinking;
          
          if (phase !== 'done' && phase !== 'error') {
            setStreamPhase(phase);
          }

          if (eventName === 'render_result' && isRenderable && chunkRecord.type === 'ppt') {
            const normalizedPpt = normalizePptData(chunkRecord.content as PptData);
            const metadata = chunkRecord.metadata;
            const backendContent = metadata?.backendContent || metadata?.summary || `PPT 已生成，共 ${normalizedPpt.slides.length} 页，可以预览或导出。`;
            workingPpt = normalizedPpt;
            renderedSlideCount = normalizedPpt.slides.length;
            setPptData(normalizedPpt);
            setResponseMetadata(metadata);
            setCurrentSlideIndex(0);
            setLeftTab('slides');
            accumulatedContent += (accumulatedContent ? '\n\n' : '') + backendContent;
            upsertStep(
              phase,
              currentPhaseLabel,
              backendContent,
              'done',
              streamEvent.author,
              eventName,
              createThinkingSummary(chunkRecord, backendContent),
            );
            accumulatedSteps.forEach(s => { s.status = 'done'; });
            updateAgentMessage();
            updateSessionPpt(normalizedPpt, metadata);
            setStreamProgress(`PPT 已生成 ${normalizedPpt.slides.length} 页`);
            return;
          }

          if (eventName === 'render_delta' && isRenderable && chunkRecord.type === 'ppt_slide') {
            upsertSlide(chunkRecord.content || chunkRecord.slide);
            const stepText = `已接收第 ${renderedSlideCount} 页可渲染数据。`;
            upsertStep(
              phase,
              currentPhaseLabel,
              stepText,
              'running',
              streamEvent.author,
              eventName,
              stepText,
            );
            updateAgentMessage();
            return;
          }

          if (eventName === 'process_delta' || eventName === 'process_result') {
            const text = formatProcessChunkText(chunkRecord) || '收到过程事件。';
            upsertStep(
              phase,
              currentPhaseLabel,
              text,
              eventName === 'process_result' ? 'done' : 'running',
              streamEvent.author,
              eventName,
              createThinkingSummary(chunkRecord, currentPhaseLabel),
            );
            setStreamProgress(text.replace(/\s+/g, ' ').slice(0, 56));
            updateAgentMessage();
            return;
          }

          if (eventName === 'message' || chunkRecord.type === 'user') {
            const displayContent = getChunkText(chunkRecord) || '需要补充信息。';
            accumulatedContent += (accumulatedContent ? '\n\n' : '') + displayContent;
            upsertStep(
              phase,
              currentPhaseLabel,
              displayContent,
              'done',
              streamEvent.author,
              eventName,
              createThinkingSummary(chunkRecord, displayContent),
            );
            updateAgentMessage();
            return;
          }

          if (eventName === 'error' || chunkRecord.type === 'error') {
            const errorText = `生成失败：${getChunkText(chunkRecord) || '后端返回错误。'}`;
            accumulatedContent += (accumulatedContent ? '\n\n' : '') + errorText;
            upsertStep(
              phase,
              currentPhaseLabel,
              errorText,
              'done',
              streamEvent.author,
              eventName,
              createThinkingSummary(chunkRecord, errorText),
            );
            accumulatedSteps.forEach(s => { s.status = 'done'; });
            updateAgentMessage();
            setStreamPhase('error');
            setStreamProgress('');
            return;
          }

          if (eventName === 'done' || chunkRecord.type === 'done') {
            accumulatedSteps.forEach(s => { s.status = 'done'; });
            updateAgentMessage();
            setStreamPhase('done');
            return;
          }

          const fallbackText = getChunkText(chunkRecord) || `收到 ${eventName} 事件。`;
          upsertStep(
            phase,
            currentPhaseLabel,
            fallbackText,
            'done',
            streamEvent.author,
            eventName,
            createThinkingSummary(chunkRecord, fallbackText),
          );
          updateAgentMessage();
        },
        (error: Error) => {
          console.error('Stream error:', error);
          if (error.name !== 'AbortError' && !accumulatedContent) {
              accumulatedContent += (accumulatedContent ? '\n\n' : '') + `连接异常: ${error.message}`;
          }
          accumulatedSteps.forEach(s => { s.status = 'done'; });
          updateAgentMessage();
          setIsSending(false);
          setStreamPhase('');
          setStreamProgress('');
        },
        () => {
          setIsSending(false);
          setStreamPhase('');
          setStreamProgress('');
          accumulatedSteps.forEach(s => { s.status = 'done'; });
          if (!accumulatedContent && renderedSlideCount > 0) {
            accumulatedContent = `PPT 已生成 ${renderedSlideCount} 页，可以预览或导出。`;
            setCurrentSlideIndex(0);
          }
          updateAgentMessage();
        }
      );
      
      streamAbortRef.current = controller;

    } catch (error) {
      console.error('Chat error:', error);
      setMessages((prev) => [
        ...prev,
        {
          id: (Date.now() + 1).toString(),
          role: 'agent',
          content: error instanceof Error ? `Error: ${error.message}` : '发送失败，请重试。',
          timestamp: Date.now(),
        },
      ]);
      setIsSending(false);
      setStreamPhase('');
      setStreamProgress('');
    }
  };
  const handleDownloadPptx = () => {
    if (!pptData) return;
    try {
      const activeTheme = THEMES.find(t => t.id === selectedThemeId) || DEFAULT_THEME;
      const pres = generatePptx(pptData, activeTheme);
      pres.writeFile({ fileName: `${pptData.title || 'AI-Generated'}.pptx` });
    } catch (err) {
      console.error('Failed to generate PPTX:', err);
    }
  };

  const handleNewChat = async () => {
    if (!selectedAgentId || !currentUser) return;
    try {
      const res = await agentApi.createSession(selectedAgentId, currentUser);
      createNewSession(false, res.data.sessionId);
    } catch (error) {
      console.error('Failed to create new session:', error);
    }
  };

  const handleRestartSession = async () => {
    if (!selectedAgentId || !currentUser) return;
    try {
      const res = await agentApi.createSession(selectedAgentId, currentUser);
      const newBackendId = res.data.sessionId;
      const initialMsg: Message = {
        id: Date.now().toString(),
        role: 'agent',
        content: '你好！我是你的 PPT 智能助手。请告诉我你想制作什么主题的演示文稿？',
        timestamp: Date.now(),
      };
      setSessionId(newBackendId);
      setMessages([initialMsg]);
      setPptData(null);
      setResponseMetadata(undefined);
      setCurrentSlideIndex(0);

      if (currentSessionId) {
        setSessions((prev) =>
          prev.map((session) => {
            if (session.id === currentSessionId) {
              return { ...session, backendSessionId: newBackendId, messages: [initialMsg], pptData: null, metadata: undefined, lastModified: Date.now() };
            }
            return session;
          })
        );
      }
    } catch (error) {
      console.error('Failed to restart session:', error);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  const quickActions = [
    { label: '产品汇报 PPT', text: '请帮我制作一份产品季度汇报PPT，包含业绩数据、重点项目进展、下季度规划，大约8页' },
    { label: '技术方案 PPT', text: '请帮我制作一份微服务架构技术方案PPT，包含架构设计、技术选型、部署方案，大约6页' },
  ];

  // Helper: render layout-specific mini decorations for thumbnails
  const renderMiniDecor = (layout: string, slideIdx: number, t: PptTheme) => {
    const coverNavyPct = (t.coverNavyHeight / 7.5) * 100;
    const shadowMini = '0 1px 3px rgba(0,0,0,0.1)';

    if (layout === 'title_classic' || layout === 'end_slide') {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: `${coverNavyPct}%`, background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute inset-x-0" style={{ top: `${coverNavyPct}%`, height: '2%', backgroundColor: `#${t.accent}` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute rounded-full" style={{ right:'6%', top:'8%', width:'16%', height:'16%', backgroundColor:`#${t.primaryLight}`, opacity:0.3 }} />
      </>;
    } else if (layout === 'title_center') {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: '4%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '4%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute inset-x-0" style={{ top: '4%', height: '1%', backgroundColor: `#${t.accent}` }} />
        <div className="absolute inset-x-0" style={{ bottom: '4%', height: '1%', backgroundColor: `#${t.accent}` }} />
        <div className="absolute rounded-full" style={{ left:'38%', top:'26%', width:'24%', height:'24%', backgroundColor:`#${t.offWhite}`, opacity:0.4, boxShadow: shadowMini }} />
      </>;
    } else if (layout === 'title_split') {
      return <>
        <div className="absolute left-0 top-0 bottom-0 w-1/2" style={{ background: `linear-gradient(180deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute right-0 top-0 bottom-0 w-1/2" style={{ backgroundColor: `#${t.offWhite}` }} />
        <div className="absolute" style={{ left:'50%', top:0, width:'1%', height:'100%', backgroundColor:`#${t.accent}` }} />
      </>;
    } else if (layout === 'card_3col') {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: '12%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
        {[0, 34, 68].map((leftPct, i) => (
          <div key={i} className="absolute rounded-sm" style={{ left:`${leftPct}%`, top:'28%', width:'30%', height:'50%', backgroundColor:`#${t.offWhite}`, borderTop:`2px solid #${t.primary}`, boxShadow: shadowMini }} />
        ))}
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
      </>;
    } else if (layout === 'comparison') {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: '12%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute rounded-sm" style={{ left:'4%', top:'22%', width:'44%', height:'60%', backgroundColor:`#${t.offWhite}`, borderTop:`2px solid #${t.primary}`, boxShadow: shadowMini }} />
        <div className="absolute rounded-sm" style={{ left:'52%', top:'22%', width:'44%', height:'60%', backgroundColor:`#${t.offWhite}`, borderTop:`2px solid #${t.accent}`, boxShadow: shadowMini }} />
        <div className="absolute" style={{ left:'49.5%', top:'22%', width:'1%', height:'60%', backgroundColor:`#${t.primaryLight}`, opacity:0.5 }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
      </>;
    } else if (layout === 'timeline') {
      return <>
        <div className="absolute" style={{ left:0, top:0, width:'4%', height:'100%', background: `linear-gradient(180deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute" style={{ left:'10%', top:'46%', width:'85%', height:'2%', backgroundColor:`#${t.primaryLight}`, opacity:0.5 }} />
        {['16%','39%','62%','85%'].map((leftPct, i) => (
          <div key={i} className="absolute rounded-full" style={{ left:leftPct, top:'40%', width:'12%', height:'12%', backgroundColor:`#${i<3?t.primary:t.accent}`, boxShadow: shadowMini }} />
        ))}
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
      </>;
    } else if (layout === 'data_highlight') {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: '10%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '22%', backgroundColor: `#${t.offWhite}` }} />
        <div className="absolute" style={{ left:'4%', top:'16%', width:'0.5%', height:'60%', backgroundColor:`#${t.primary}`, borderRadius:'4px' }} />
        <div className="absolute inset-x-0" style={{ bottom: '5%', height: '5%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
      </>;
    } else if (layout === 'quote_slide') {
      return <>
        <div className="absolute inset-x-0" style={{ backgroundColor: `#${t.offWhite}`, inset:'0' }} />
        <div className="absolute" style={{ left:0, top:0, width:'6%', height:'100%', background: `linear-gradient(180deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute" style={{ left:'6%', top:'37%', width:'1%', height:'24%', backgroundColor:`#${t.accent}`, borderRadius:'4px' }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
      </>;
    } else if (layout === 'card_2col') {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: '12%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute rounded-sm" style={{ left:'8%', top:'28%', width:'38%', height:'50%', backgroundColor:`#${t.offWhite}`, borderTop:`2px solid #${t.primary}`, boxShadow: shadowMini }} />
        <div className="absolute rounded-sm" style={{ left:'54%', top:'28%', width:'38%', height:'50%', backgroundColor:`#${t.offWhite}`, borderTop:`2px solid #${t.primary}`, boxShadow: shadowMini }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
      </>;
    } else if (layout === 'content_top') {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: '16%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute inset-x-0" style={{ top: '16%', height: '2%', backgroundColor: `#${t.accent}` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
      </>;
    } else {
      // Default content_classic
      return <>
        <div className="absolute" style={{ left:0, top:0, width:'34%', height:'95%', background: `linear-gradient(180deg, #${t.primary}, #${t.primaryMid})` }} />
        <div className="absolute" style={{ left:'34%', top:0, width:'0.5%', height:'95%', backgroundColor:`#${t.accent}` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', background: `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})` }} />
      </>;
    }
  };

  // Render mini slide for thumbnail panel (simplified, very small)
  const renderMiniSlide = (slideData: PptSlide, slideIdx: number) => {
    const layout = inferLayout(slideData, slideIdx, pptData?.slides.length ?? 1);
    const t = activeTheme;

    return (
      <div className="w-full h-full relative bg-white" style={{ fontSize: '2px' }}>
        {/* Theme decorations (mini) — layout-specific */}
        {renderMiniDecor(layout, slideIdx, t)}
        {/* Actual text content in thumbnail */}
        {slideData.elements
          .filter(el => el.kind === 'text')
          .slice(0, 3)
          .map((el, i) => {
            const isTitle = (el.fontSize || 0) >= 24;
            // Use safe-area clamped position for color detection
            const safe = getSafeContentArea(layout, t);
            const elX = Math.max(safe.x, Math.min(el.x || 0, safe.x + safe.w - (el.w || 2)));
            const elY = Math.max(safe.y, Math.min(el.y || 0, safe.y + safe.h - (el.h || 1)));
            const elCenterX = elX + (el.w || 4) / 2;
            const elCenterY = elY + (el.h || 1) / 2;
            const onDark = isOnDarkArea(elCenterX, elCenterY, layout, t);
            const yPct = (el.y || 0) / 7.5 * 100;
            const hPct = Math.max(8, (el.h || 1) / 7.5 * 100);
            // Position-aware color for thumbnails
            let miniColor = onDark ? `#${t.white}` : `#${t.bodyColor}`;
            if (isTitle && !onDark) miniColor = `#${t.primary}`;
            return (
              <div
                key={i}
                className="absolute overflow-hidden"
                style={{
                  left: '6%',
                  right: '6%',
                  top: `${Math.min(yPct, 85)}%`,
                  height: `${hPct}%`,
                  color: miniColor,
                  fontSize: isTitle ? '3px' : '2px',
                  fontWeight: isTitle ? 'bold' : 'normal',
                  lineHeight: 1.2,
                  whiteSpace: 'nowrap',
                  textOverflow: 'ellipsis',
                }}
              >
                {(el.content || '').replace(/[•\-]/g, '').trim().slice(0, 30)}
              </div>
            );
          })}
      </div>
    );
  };

  // Render slide content elements (shared between main preview and fullscreen)
  const renderSlideContent = (slideData: PptSlide, slideIdx: number) => {
    const layout = inferLayout(slideData, slideIdx, pptData?.slides.length ?? 1);
    const isTitleSlide = layout === 'title_slide' || slideIdx === 0;
    const isEndSlide = layout === 'end_slide' || slideIdx === (pptData?.slides.length ?? 1) - 1;

    // Filter: remove ALL shape elements (decorative), only keep text, table, image
    const filteredElements = slideData.elements.filter(el => {
      if (el.kind === 'shape') return false; // Always skip shapes - theme handles decor
      return true;
    });

    return filteredElements.map((el, idx) => {

      // Clamp element positions to safe content area to prevent overlap with decorations
      const safe = getSafeContentArea(layout, activeTheme);
      const elX = Math.max(safe.x, Math.min(el.x || 0, safe.x + safe.w - (el.w || 2)));
      const elY = Math.max(safe.y, Math.min(el.y || 0, safe.y + safe.h - (el.h || 1)));

      const xPct = (elX / 13.33) * 100;
      const yPct = (elY / 7.5) * 100;
      const wPct = ((el.w || 4) / 13.33) * 100;
      const hPct = ((el.h || 1) / 7.5) * 100;

      if (el.kind === 'text') {
        const fontSizeNum = el.fontSize || 18;
        const fontSize = Math.max(8, fontSizeNum * 0.7);
        // Position-aware color detection: check if element center is on a dark area
        const elCenterX = elX + (el.w || 4) / 2;
        const elCenterY = elY + (el.h || 1) / 2;
        const onDark = isOnDarkArea(elCenterX, elCenterY, layout, activeTheme);
        let textColor = onDark ? `#${activeTheme.white}` : `#${activeTheme.bodyColor}`;

        // Title text (fontSize >= 24) gets special treatment for emphasis
        if (fontSizeNum >= 24) {
          if (onDark) {
            textColor = `#${activeTheme.white}`;
          } else {
            // On light area: use primary color for titles
            textColor = `#${activeTheme.primary}`;
          }
        }
        // Large decorative text (big numbers etc): allow AI color choice
        if (el.color && fontSizeNum >= 30) {
          textColor = `#${el.color}`;
        }

        return (
          <div
            key={idx}
            className="absolute overflow-hidden"
            style={{
              left: `${xPct}%`,
              top: `${yPct}%`,
              width: `${wPct}%`,
              height: `${hPct}%`,
              fontSize: `${fontSize}px`,
              color: textColor,
              fontWeight: el.bold ? 'bold' : 'normal',
              backgroundColor: el.fill ? `#${el.fill}` : 'transparent',
              textAlign: el.align || 'left',
              display: 'flex',
              alignItems: el.fontSize && el.fontSize >= 30 ? 'center' : (el.icon ? 'center' : 'flex-start'),
              lineHeight: fontSizeNum >= 30 ? 1.2 : 1.6,
              letterSpacing: fontSizeNum >= 30 ? '0.5px' : '0.2px',
              whiteSpace: 'pre-wrap',
              paddingTop: fontSizeNum < 22 ? '2px' : '0',
              textShadow: (fontSizeNum >= 30 && textColor === `#${activeTheme.white}`) ? '0 1px 3px rgba(0,0,0,0.2)' : 'none',
              gap: el.icon ? '4px' : '0',
            }}
          >
            {el.icon && <span style={{ fontSize: `${fontSize * 1.1}px`, flexShrink: 0 }}>{el.icon}</span>}
            <span dangerouslySetInnerHTML={{ __html: (el.content || '').replace(/\n/g, '<br/>') }} />
          </div>
        );
      }

      if (el.kind === 'shape' && el.w >= 0.6 && el.h >= 0.25) {
        return (
          <div
            key={idx}
            className="absolute"
            style={{
              left: `${xPct}%`,
              top: `${yPct}%`,
              width: `${wPct}%`,
              height: `${hPct}%`,
              backgroundColor: el.fill ? `#${el.fill}` : '#2E5090',
              borderRadius: el.radius ? `${el.radius}px` : '0',
              boxShadow: el.shadow ? '0 2px 8px rgba(0,0,0,0.12)' : 'none',
              opacity: el.opacity ?? 1,
            }}
          />
        );
      }

      // Icon element: renders emoji/icon at a specific position
      if (el.kind === 'icon') {
        const iconSize = Math.max(16, (el.fontSize || 32) * 0.7);
        return (
          <div
            key={idx}
            className="absolute flex items-center justify-center"
            style={{
              left: `${xPct}%`,
              top: `${yPct}%`,
              width: `${wPct}%`,
              height: `${hPct}%`,
              fontSize: `${iconSize}px`,
              color: el.color ? `#${el.color}` : `#${activeTheme.primary}`,
            }}
          >
            {el.icon || el.content || '●'}
          </div>
        );
      }

      // Divider line element
      if (el.kind === 'divider') {
        const isHorizontal = (el.w || 0) > (el.h || 0);
        return (
          <div
            key={idx}
            className="absolute"
            style={{
              left: `${xPct}%`,
              top: `${yPct}%`,
              width: isHorizontal ? `${wPct}%` : `${(el.thickness || 0.04) / 13.33 * 100}%`,
              height: isHorizontal ? `${(el.thickness || 0.04) / 7.5 * 100}%` : `${hPct}%`,
              backgroundColor: el.color ? `#${el.color}` : `#${activeTheme.accent}`,
              borderRadius: '2px',
              opacity: el.opacity ?? 0.8,
            }}
          />
        );
      }

      // Numbered bullet element: circle with number + text
      if (el.kind === 'bullet') {
        const bulletSize = Math.max(20, (el.fontSize || 14) * 0.7);
        const numSize = Math.max(10, bulletSize * 0.55);
        return (
          <div
            key={idx}
            className="absolute"
            style={{
              left: `${xPct}%`,
              top: `${yPct}%`,
              width: `${wPct}%`,
              height: `${hPct}%`,
              display: 'flex',
              alignItems: 'flex-start',
              gap: '6px',
            }}
          >
            {/* Number circle */}
            <div
              style={{
                minWidth: `${bulletSize}px`,
                height: `${bulletSize}px`,
                borderRadius: '50%',
                backgroundColor: el.fill ? `#${el.fill}` : `#${activeTheme.primary}`,
                color: `#${activeTheme.white}`,
                fontSize: `${numSize}px`,
                fontWeight: 'bold',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                flexShrink: 0,
                boxShadow: '0 2px 8px rgba(0,0,0,0.1)',
              }}
            >
              {el.number ?? 1}
            </div>
            {/* Text content */}
            <div
              style={{
                fontSize: `${bulletSize * 0.72}px`,
                color: el.color ? `#${el.color}` : `#${activeTheme.bodyColor}`,
                lineHeight: 1.5,
                fontWeight: el.bold ? 'bold' : 'normal',
              }}
              dangerouslySetInnerHTML={{ __html: (el.content || '').replace(/\n/g, '<br/>') }}
            />
          </div>
        );
      }

      if (el.kind === 'table' && Array.isArray(el.rows) && el.rows.length > 0) {
        const safeRows = el.rows.filter((row): row is string[] => Array.isArray(row) && row.length > 0);
        if (safeRows.length === 0) return null;
        return (
          <div
            key={idx}
            className="absolute overflow-auto"
            style={{
              left: `${xPct}%`,
              top: `${yPct}%`,
              width: `${wPct}%`,
              height: `${hPct}%`,
            }}
          >
            <table className="w-full border-collapse" style={{ fontSize: '10px' }}>
              <tbody>
                {safeRows.map((row, ri) => (
                  <tr key={ri}>
                    {row.map((cell, ci) => (
                      <td
                        key={ci}
                        className="border px-0.5 py-0.5 text-center"
                        style={{
                          borderColor: '#C0C8D4',
                          backgroundColor: ri === 0 ? `#${activeTheme.primary}` : (ri % 2 === 0 ? `#${activeTheme.offWhite}` : 'transparent'),
                          color: ri === 0 ? 'white' : `#${activeTheme.bodyColor}`,
                          fontWeight: ri === 0 ? 'bold' : 'normal',
                        }}
                      >
                        {cell ?? ''}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );
      }

      if (el.kind === 'image') {
        return (
          <div
            key={idx}
            className="absolute bg-slate-100 overflow-hidden"
            style={{
              left: `${xPct}%`,
              top: `${yPct}%`,
              width: `${wPct}%`,
              height: `${hPct}%`,
              borderRadius: el.radius ? `${el.radius}px` : '4px',
              boxShadow: el.shadow ? '0 6px 24px rgba(0,0,0,0.08), 0 2px 8px rgba(0,0,0,0.04)' : 'none',
              opacity: el.opacity ?? 1,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
            }}
          >
            {el.content ? (
              <img src={el.content} alt="slide image" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
            ) : (
              <Icons.FilePresentation className="w-8 h-8 text-slate-300" />
            )}
          </div>
        );
      }

      return null;
    });
  };

  // Render current slide detail
  // Helper: render a positioned div for theme decoration in preview
  // Enhanced with gradient, shadow, rounded corners support
  const renderDecorDiv = (
    x: number, y: number, w: number, h: number,
    color: string, key: string,
    opts?: {
      isCircle?: boolean;
      radius?: number;      // border-radius in px
      shadow?: boolean;
      gradient?: string;    // CSS gradient string, e.g. 'linear-gradient(135deg, #1F3864, #2E5090)'
      opacity?: number;     // 0-1
    }
  ) => {
    const { isCircle = false, radius = 0, shadow = false, gradient, opacity = 1 } = opts || {};
    return (
      <div
        key={key}
        className="absolute"
        style={{
          left: `${(x / 13.33) * 100}%`,
          top: `${(y / 7.5) * 100}%`,
          width: `${(w / 13.33) * 100}%`,
          height: `${(h / 7.5) * 100}%`,
          background: gradient || `#${color}`,
          borderRadius: isCircle ? '50%' : (radius ? `${radius}px` : '0'),
          boxShadow: shadow ? '0 4px 16px rgba(0,0,0,0.06), 0 1px 4px rgba(0,0,0,0.04)' : 'none',
          opacity,
        }}
      />
    );
  };

  // Render theme decorations for preview — layout-specific skeletons (ENHANCED)
  const renderThemeDecor = (slideData: PptSlide, slideIdx: number) => {
    const layout = inferLayout(slideData, slideIdx, pptData?.slides.length ?? 1);
    const t = activeTheme;
    const elements: React.ReactNode[] = [];

    const pctX = (v: number) => (v / 13.33) * 100;
    const pctY = (v: number) => (v / 7.5) * 100;
    const pctW = (v: number) => (v / 13.33) * 100;
    const pctH = (v: number) => (v / 7.5) * 100;

    // Common gradient helpers
    const primaryGrad = `linear-gradient(135deg, #${t.primary}, #${t.primaryMid})`;
    const primaryVertGrad = `linear-gradient(180deg, #${t.primary}, #${t.primaryMid})`;

    if (layout === 'title_classic' || layout === 'end_slide') {
      // Cover / End Classic — gradient cover + soft accent
      elements.push(renderDecorDiv(0, 0, 13.33, t.coverNavyHeight, t.primary, 't-top', {
        gradient: primaryGrad, radius: 0,
      }));
      elements.push(renderDecorDiv(0, t.coverNavyHeight, 13.33, 0.12, t.accent, 't-stripe'));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom', {
        gradient: primaryGrad,
      }));
      // Decorative circle with transparency
      elements.push(renderDecorDiv(10.5, 0.6, 2.2, 2.2, t.primaryLight, 't-circle', {
        isCircle: true, opacity: 0.3,
      }));
      // Secondary subtle circle
      elements.push(renderDecorDiv(0.3, 4.8, 1.4, 1.4, t.primaryLight, 't-circle2', {
        isCircle: true, opacity: 0.15,
      }));
    } else if (layout === 'title_center') {
      // Cover Center — gradient bars + soft center circle
      elements.push(renderDecorDiv(0, 0, 13.33, 0.25, t.primary, 't-top', {
        gradient: primaryGrad,
      }));
      elements.push(renderDecorDiv(0, 7.25, 13.33, 0.25, t.primary, 't-bottom', {
        gradient: primaryGrad,
      }));
      elements.push(renderDecorDiv(0, 0.25, 13.33, 0.05, t.accent, 't-top-accent'));
      elements.push(renderDecorDiv(0, 7.2, 13.33, 0.05, t.accent, 't-bot-accent'));
      elements.push(renderDecorDiv(5.16, 2.25, 3.0, 3.0, t.offWhite, 't-circle', {
        isCircle: true, opacity: 0.4, shadow: true,
      }));
    } else if (layout === 'title_split') {
      // Cover Split — gradient left + subtle divider
      elements.push(renderDecorDiv(0, 0, 6.66, 7.5, t.primary, 't-left', {
        gradient: primaryVertGrad,
      }));
      elements.push(renderDecorDiv(6.66, 0, 6.67, 7.5, t.offWhite, 't-right'));
      elements.push(renderDecorDiv(6.66, 0, 0.1, 7.5, t.accent, 't-divider', {
        radius: 2,
      }));
      // Decorative circle on left panel
      elements.push(renderDecorDiv(1.0, 5.0, 1.8, 1.8, t.primaryLight, 't-circle1', {
        isCircle: true, opacity: 0.2,
      }));
    } else if (layout === 'card_3col') {
      // 3-Column Cards — shadow cards with rounded top accent
      elements.push(renderDecorDiv(0, 0, 13.33, 0.9, t.primary, 't-header', {
        gradient: primaryGrad,
      }));
      elements.push(renderDecorDiv(0, 0.9, 13.33, 0.06, t.accent, 't-accent'));
      elements.push(renderDecorDiv(0.5, 1.6, 3.8, 5.0, t.offWhite, 't-card1bg', {
        shadow: true, radius: 8,
      }));
      elements.push(renderDecorDiv(4.75, 1.6, 3.8, 5.0, t.offWhite, 't-card2bg', {
        shadow: true, radius: 8,
      }));
      elements.push(renderDecorDiv(9.0, 1.6, 3.8, 5.0, t.offWhite, 't-card3bg', {
        shadow: true, radius: 8,
      }));
      elements.push(renderDecorDiv(0.5, 1.6, 3.8, 0.12, t.primary, 't-card1top', {
        gradient: primaryGrad, radius: 8,
      }));
      elements.push(renderDecorDiv(4.75, 1.6, 3.8, 0.12, t.primary, 't-card2top', {
        gradient: primaryGrad, radius: 8,
      }));
      elements.push(renderDecorDiv(9.0, 1.6, 3.8, 0.12, t.primary, 't-card3top', {
        gradient: primaryGrad, radius: 8,
      }));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom', {
        gradient: primaryGrad,
      }));
    } else if (layout === 'comparison') {
      // Comparison: shadow cards with accent divider
      elements.push(renderDecorDiv(0, 0, 13.33, 0.9, t.primary, 't-header', {
        gradient: primaryGrad,
      }));
      elements.push(renderDecorDiv(0, 0.9, 13.33, 0.06, t.accent, 't-accent'));
      elements.push(renderDecorDiv(0.5, 1.5, 5.9, 5.2, t.offWhite, 't-leftblock', {
        shadow: true, radius: 8,
      }));
      elements.push(renderDecorDiv(0.5, 1.5, 5.9, 0.1, t.primary, 't-lefttop', {
        gradient: primaryGrad, radius: 8,
      }));
      elements.push(renderDecorDiv(6.9, 1.5, 5.9, 5.2, t.offWhite, 't-rightblock', {
        shadow: true, radius: 8,
      }));
      elements.push(renderDecorDiv(6.9, 1.5, 5.9, 0.1, t.accent, 't-righttop', {
        radius: 8,
      }));
      elements.push(renderDecorDiv(6.55, 1.5, 0.2, 5.2, t.primaryLight, 't-divider', {
        radius: 2, opacity: 0.5,
      }));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom', {
        gradient: primaryGrad,
      }));
    } else if (layout === 'timeline') {
      // Timeline: gradient left bar + soft dots
      elements.push(renderDecorDiv(0, 0, 0.5, 7.5, t.primary, 't-leftbar', {
        gradient: primaryVertGrad,
      }));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom', {
        gradient: primaryGrad,
      }));
      elements.push(renderDecorDiv(1.2, 3.5, 11.5, 0.12, t.primaryLight, 't-line', {
        opacity: 0.5,
      }));
      elements.push(renderDecorDiv(2.5, 3.2, 0.7, 0.7, t.primary, 't-dot1', {
        isCircle: true, shadow: true,
      }));
      elements.push(renderDecorDiv(5.5, 3.2, 0.7, 0.7, t.primary, 't-dot2', {
        isCircle: true, shadow: true,
      }));
      elements.push(renderDecorDiv(8.5, 3.2, 0.7, 0.7, t.primary, 't-dot3', {
        isCircle: true, shadow: true,
      }));
      elements.push(renderDecorDiv(11.5, 3.2, 0.7, 0.7, t.accent, 't-dot4', {
        isCircle: true, shadow: true,
      }));
    } else if (layout === 'data_highlight') {
      // Data Highlight: gradient header + soft bottom band
      elements.push(renderDecorDiv(0, 0, 13.33, 0.7, t.primary, 't-header', {
        gradient: primaryGrad,
      }));
      elements.push(renderDecorDiv(0, 0.7, 13.33, 0.06, t.accent, 't-accent'));
      elements.push(renderDecorDiv(0, 6.0, 13.33, 1.5, t.offWhite, 't-bottomband', {
        radius: 0,
      }));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom', {
        gradient: primaryGrad,
      }));
      elements.push(renderDecorDiv(0.5, 1.2, 0.08, 4.5, t.primary, 't-leftbar', {
        radius: 4,
      }));
    } else if (layout === 'quote_slide') {
      // Quote: gradient left bar + soft bg
      elements.push(renderDecorDiv(0, 0, 13.33, 7.5, t.offWhite, 't-bg'));
      elements.push(renderDecorDiv(0, 0, 0.8, 7.5, t.primary, 't-leftbar', {
        gradient: primaryVertGrad, radius: 0,
      }));
      elements.push(renderDecorDiv(0.8, 2.8, 0.12, 1.8, t.accent, 't-accent', {
        radius: 4,
      }));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom', {
        gradient: primaryGrad,
      }));
    } else if (layout === 'card_2col') {
      // 2-Column Cards — shadow + rounded
      elements.push(renderDecorDiv(0, 0, 13.33, 0.9, t.primary, 't-header', {
        gradient: primaryGrad,
      }));
      elements.push(renderDecorDiv(0, 0.9, 13.33, 0.06, t.accent, 't-accent'));
      elements.push(renderDecorDiv(1.0, 1.6, 5.0, 5.0, t.offWhite, 't-card1bg', {
        shadow: true, radius: 8,
      }));
      elements.push(renderDecorDiv(7.33, 1.6, 5.0, 5.0, t.offWhite, 't-card2bg', {
        shadow: true, radius: 8,
      }));
      elements.push(renderDecorDiv(1.0, 1.6, 5.0, 0.15, t.primary, 't-card1top', {
        gradient: primaryGrad, radius: 8,
      }));
      elements.push(renderDecorDiv(7.33, 1.6, 5.0, 0.15, t.primary, 't-card2top', {
        gradient: primaryGrad, radius: 8,
      }));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom', {
        gradient: primaryGrad,
      }));
    } else if (layout === 'content_top') {
      // Content Top Bar — gradient header
      elements.push(renderDecorDiv(0, 0, 13.33, 1.2, t.primary, 't-header', {
        gradient: primaryGrad,
      }));
      elements.push(renderDecorDiv(0, 1.2, 13.33, 0.08, t.accent, 't-accent', {
        radius: 0,
      }));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom', {
        gradient: primaryGrad,
      }));
    } else {
      // Default content_classic: gradient left band + accent divider
      elements.push(renderDecorDiv(0, 0, 4.5, 7.15, t.primary, 't-leftband', {
        gradient: primaryVertGrad,
      }));
      elements.push(renderDecorDiv(4.5, 0, 0.08, 7.15, t.accent, 't-divider', {
        radius: 4,
      }));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom', {
        gradient: primaryGrad,
      }));
    }
    return elements;
  };


  return (
    <div className="flex flex-col h-screen w-full overflow-hidden bg-slate-50 text-slate-900 font-sans">
      {/* ===== Header Bar ===== */}
      <header className="ppt-atelier-header">
        <div className="flex items-center gap-3">
          <div className="ppt-atelier-mark">
            <Icons.FilePresentation className="text-white w-5 h-5" />
          </div>
          <div>
            <p className="text-[10px] font-semibold uppercase tracking-[0.28em] text-slate-400">Presentation Atelier</p>
            <h1 className="text-lg font-bold text-slate-800 tracking-tight">AI PPT <span className="xerina-inline">xerina</span></h1>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={() => router.push('/')}
            className="flex items-center gap-2 px-4 py-1.5 bg-white border border-slate-200 text-slate-600 rounded-lg hover:bg-slate-50 hover:border-slate-300 hover:text-slate-900 transition-all text-sm font-medium shadow-sm active:scale-95"
          >
            返回主页
          </button>

          <a
            href="https://bugstack.cn/md/project/ai-agent-scaffold/ai-agent-scaffold.html"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-500 hover:text-slate-800 bg-white border border-slate-200 hover:border-slate-300 rounded-lg transition-all shadow-sm"
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>
            课程&源码
          </a>

          <div className="h-6 w-px bg-slate-200 mx-1"></div>

          <div className="flex items-center gap-2 px-3 py-1.5 bg-slate-50 rounded-full border border-slate-200 shadow-sm">
             <div className="w-2 h-2 rounded-full bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.4)]"></div>
             <span className="text-xs font-semibold text-slate-600">{currentUser || 'Guest'}</span>
          </div>

          <div className="h-6 w-px bg-slate-200 mx-1"></div>

          <button
            onClick={handleDownloadPptx}
            disabled={!pptData}
            className={`flex items-center gap-1.5 px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-sm font-medium transition-all shadow-sm active:scale-95 ${
              pptData
                ? 'text-slate-600 hover:bg-slate-50 hover:border-slate-300'
                : 'text-slate-400 cursor-not-allowed'
            }`}
          >
            <Icons.Download className="w-4 h-4" />
            下载 PPTX
          </button>

          {pptData && (
            <button
              onClick={() => setIsFullscreen(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 hover:border-slate-300 rounded-lg text-sm font-medium transition-all shadow-sm active:scale-95"
            >
              <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M8 3H5a2 2 0 0 0-2 2v3"></path>
                <path d="M21 8V5a2 2 0 0 0-2-2h-3"></path>
                <path d="M3 16v3a2 2 0 0 0 2 2h3"></path>
                <path d="M16 21h3a2 2 0 0 0 2-2v-3"></path>
              </svg>
              全屏演示
            </button>
          )}

          <button
            onClick={() => setIsStylePanelOpen(prev => !prev)}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm font-medium transition-all shadow-sm active:scale-95 border ${
              isStylePanelOpen ? 'bg-indigo-50 border-indigo-200 text-indigo-700' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50 hover:border-slate-300'
            }`}
            title="风格设置"
          >
            🎛️
          </button>

          <button onClick={handleLogout} className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors" title="退出登录">
            <Icons.Logout className="w-4 h-4" />
          </button>

          {!isChatOpen && (
            <button onClick={() => setIsChatOpen(true)} className="p-1.5 text-indigo-600 bg-indigo-50 hover:bg-indigo-100 rounded-md transition-colors" title="打开助手">
              <Icons.Chat className="w-4 h-4" />
            </button>
          )}
        </div>
      </header>

      {(responseMetadata?.summary || responseMetadata?.suggestions?.length || responseMetadata?.nextActions?.length) && (
        <section className="ppt-metadata-strip">
          {responseMetadata?.summary && (
            <p><span>summary</span>{responseMetadata.summary}</p>
          )}
          {responseMetadata?.suggestions?.slice(0, 2).map((item) => (
            <p key={item}><span>suggest</span>{item}</p>
          ))}
          {responseMetadata?.nextActions?.slice(0, 2).map((item) => (
            <p key={item}><span>next</span>{item}</p>
          ))}
        </section>
      )}

      {/* ===== Main 3-Column Layout ===== */}
      <div className="flex flex-1 w-full overflow-hidden">

        {/* ===== Left: Slide Thumbnails / Template Library ===== */}
        <aside className="w-[180px] bg-white border-r border-slate-100/60 flex flex-col shrink-0 z-30">
          {/* Tab Switcher: Templates / Slides */}
          <div className="px-2 pt-2 pb-1 border-b border-slate-100 flex gap-1">
            <button
              onClick={() => setLeftTab('templates')}
              className={`flex-1 text-[10px] font-medium py-1 rounded transition ${leftTab === 'templates' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-400 hover:bg-slate-50'}`}
            >
              📋 模板
            </button>
            <button
              onClick={() => setLeftTab('slides')}
              className={`flex-1 text-[10px] font-medium py-1 rounded transition ${leftTab === 'slides' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-400 hover:bg-slate-50'}`}
            >
              🖼️ 页面
            </button>
          </div>

          {/* Theme Selector (always visible) */}
          <div className="px-3 pt-2 pb-2 border-b border-slate-100">
            <div className="text-[10px] text-slate-400 font-medium mb-1.5">🎨 主题</div>
            <div className="flex flex-wrap gap-1.5">
              {THEMES.map(t => (
                <button
                  key={t.id}
                  onClick={() => setSelectedThemeId(t.id)}
                  className={`w-6 h-6 rounded-full border-2 transition-all flex items-center justify-center ${
                    selectedThemeId === t.id
                      ? 'border-indigo-500 scale-110 shadow-md'
                      : 'border-slate-200 hover:border-slate-400 hover:scale-105'
                  }`}
                  title={t.name}
                >
                  <div
                    className="w-4 h-4 rounded-full overflow-hidden"
                    style={{
                      background: `linear-gradient(135deg, #${t.primary} 50%, #${t.accent} 50%)`,
                    }}
                  />
                </button>
              ))}
            </div>
          </div>

          {/* Content area */}
          <div className="flex-1 overflow-y-auto p-2 space-y-1 scrollbar-thin scrollbar-thumb-slate-200">
            {leftTab === 'templates' ? (
              /* Template Library */
              <div className="space-y-0.5">
                {PPT_TEMPLATES.map((cat) => (
                  <div key={cat.category}>
                    <button
                      onClick={() => setExpandedCategory(expandedCategory === cat.category ? null : cat.category)}
                      className="w-full flex items-center gap-1.5 px-1.5 py-1.5 rounded-md hover:bg-slate-50 transition text-left"
                    >
                      <span className="text-xs">{cat.icon}</span>
                      <span className="text-[11px] font-medium text-slate-600 flex-1">{cat.label}</span>
                      <span className="text-[9px] text-slate-300">{cat.templates.length}</span>
                      <svg
                        className={`w-3 h-3 text-slate-400 transition-transform ${expandedCategory === cat.category ? 'rotate-90' : ''}`}
                        fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}
                      >
                        <path strokeLinecap="round" strokeLinejoin="round" d="M9 5l7 7-7 7" />
                      </svg>
                    </button>
                    {expandedCategory === cat.category && (
                      <div className="ml-3 space-y-0.5 mb-1">
                        {cat.templates.map((tpl) => (
                          <button
                            key={tpl.id}
                            onClick={() => {
                              setInputValue(tpl.prompt);
                              if (tpl.style) setSelectedStyle(tpl.style);
                              if (tpl.scene) setSelectedScene(tpl.scene);
                            }}
                            className="w-full text-left px-2 py-1.5 rounded-md hover:bg-indigo-50 hover:border-indigo-200 border border-transparent transition group"
                          >
                            <div className="text-[11px] font-medium text-slate-700 group-hover:text-indigo-700">{tpl.name}</div>
                            <div className="text-[9px] text-slate-400 mt-0.5 leading-tight">{tpl.desc}</div>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            ) : pptData && pptData.slides.length > 0 ? (
              pptData.slides.map((slide, idx) => (
                <div
                  key={idx}
                  onClick={() => setCurrentSlideIndex(idx)}
                  className={`cursor-pointer rounded-md transition-all ${
                    currentSlideIndex === idx
                      ? 'ring-2 ring-indigo-500 ring-offset-2 shadow-md'
                      : 'ring-1 ring-slate-200 hover:ring-slate-300 hover:shadow-sm'
                  }`}
                >
                  <div className="text-[10px] text-slate-400 mb-1 pl-0.5 font-medium">{idx + 1}</div>
                  <div className="w-full aspect-[16/9] bg-white rounded overflow-hidden relative shadow-sm">
                    {renderMiniSlide(slide, idx)}
                  </div>
                </div>
              ))
            ) : (
              <div className="text-center py-8">
                <Icons.FilePresentation className="w-8 h-8 mx-auto text-slate-200 mb-2" />
                <p className="text-[11px] text-slate-400">暂无幻灯片</p>
                <p className="text-[9px] text-slate-300 mt-1">在右侧输入需求或选择模板</p>
              </div>
            )
          }
          </div>

          {/* Session list */}
          <div className="border-t border-slate-100">
            <div className="h-9 px-3 flex items-center justify-between">
              <span className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">历史</span>
              <button onClick={handleNewChat} className="p-1 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded transition" title="新建">
                <Icons.Plus className="w-3.5 h-3.5" />
              </button>
            </div>
            <div className="max-h-[120px] overflow-y-auto px-2 pb-2 space-y-0.5">
              {[...sessions].sort((a, b) => b.lastModified - a.lastModified).slice(0, 8).map((session) => (
                <div
                  key={session.id}
                  onClick={() => handleSwitchSession(session.id)}
                  className={`group flex items-center gap-1 px-2 py-1 rounded cursor-pointer transition ${
                    currentSessionId === session.id
                      ? 'bg-indigo-50 text-indigo-700'
                      : 'hover:bg-slate-50 text-slate-500 hover:text-slate-700'
                  }`}
                >
                  <span className="text-[11px] truncate flex-1">{session.title}</span>
                  <button
                    onClick={(e) => handleDeleteSession(e, session.id)}
                    className="p-0.5 rounded opacity-0 group-hover:opacity-100 hover:bg-red-100 text-slate-300 hover:text-red-500 transition"
                  >
                    <Icons.Trash className="w-3 h-3" />
                  </button>
                </div>
              ))}
            </div>
          </div>
        </aside>

        {/* ===== Center: Main Slide Preview ===== */}
        <main className="flex-1 flex flex-col bg-slate-50 h-full overflow-hidden">
          <div
            className="flex-1 flex items-center justify-center p-6 overflow-hidden"
            tabIndex={0}
            onKeyDown={(e) => {
              if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
                setCurrentSlideIndex(Math.max(0, currentSlideIndex - 1));
              } else if (e.key === 'ArrowRight' || e.key === 'ArrowDown' || e.key === ' ') {
                e.preventDefault();
                if (pptData) setCurrentSlideIndex(Math.min(pptData.slides.length - 1, currentSlideIndex + 1));
              } else if (e.key === 'Escape') {
                setIsFullscreen(false);
              }
            }}
          >
            {!pptData || !pptData.slides[currentSlideIndex] ? (
              <div className="text-center">
                <div className="w-24 h-24 mx-auto mb-6 bg-slate-100 rounded-2xl flex items-center justify-center">
                  <Icons.FilePresentation className="w-12 h-12 text-slate-300" />
                </div>
                <p className="text-base text-slate-500 mb-2 font-medium">AI 生成 PPT 后将在此预览</p>
                <p className="text-sm text-slate-400">在右侧对话区描述你的需求</p>
              </div>
            ) : (
              <div
                className="bg-white rounded-lg shadow-xl border border-slate-200 overflow-hidden relative"
                style={{
                  width: '100%',
                  maxWidth: '960px',
                  aspectRatio: '16/9',
                  maxHeight: 'calc(100vh - 160px)',
                }}
              >
                <div className="w-full h-full relative" style={{ fontFamily: "\"PingFang SC\", \"Microsoft YaHei\", system-ui, -apple-system, sans-serif", letterSpacing: "0.2px" }}>
                  {renderThemeDecor(pptData.slides[currentSlideIndex], currentSlideIndex)}
                  {renderSlideContent(pptData.slides[currentSlideIndex], currentSlideIndex)}
                </div>
              </div>
            )}
          </div>

          {/* Slide navigation bar */}
          {pptData && pptData.slides.length > 0 && (
            <div className="h-11 px-4 bg-white border-t border-slate-200 flex items-center justify-center gap-4 shrink-0">
              <button
                onClick={() => setCurrentSlideIndex(Math.max(0, currentSlideIndex - 1))}
                disabled={currentSlideIndex === 0}
                className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-md disabled:opacity-30 disabled:cursor-not-allowed transition"
              >
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="15 18 9 12 15 6"></polyline>
                </svg>
              </button>
              <span className="text-sm text-slate-500 font-mono tabular-nums min-w-[60px] text-center">
                {currentSlideIndex + 1} / {pptData.slides.length}
              </span>
              <button
                onClick={() => setCurrentSlideIndex(Math.min(pptData.slides.length - 1, currentSlideIndex + 1))}
                disabled={currentSlideIndex === pptData.slides.length - 1}
                className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-md disabled:opacity-30 disabled:cursor-not-allowed transition"
              >
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="9 18 15 12 9 6"></polyline>
                </svg>
              </button>
              <div className="h-4 w-px bg-slate-200 mx-2"></div>
              <span className="text-xs text-slate-400">← → 翻页</span>
            </div>
          )}
        </main>

        {/* ===== Right: Style Panel + Chat Panel ===== */}
        <div className="flex shrink-0 z-20">
          {/* Style Control Panel */}
          <div
            className={`border-l border-slate-100/60 bg-white flex flex-col transition-all duration-300 ease-[cubic-bezier(0.25,0.1,0.25,1)] ${
              isStylePanelOpen ? 'w-[220px] translate-x-0' : 'w-0 translate-x-full opacity-0 overflow-hidden'
            }`}
          >
            <div className="h-12 px-4 border-b border-slate-100 flex items-center justify-between shrink-0">
              <span className="text-sm font-bold text-slate-800">🎛️ 风格设置</span>
              <button onClick={() => setIsStylePanelOpen(false)} className="p-1.5 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-md transition shrink-0">
                <Icons.Close className="w-4 h-4" />
              </button>
            </div>
            <div className="flex-1 overflow-y-auto p-4 space-y-5 text-sm">
              {/* Style */}
              <div>
                <div className="text-xs text-slate-400 font-medium mb-2">风格</div>
                <div className="grid grid-cols-2 gap-1.5">
                  {STYLE_OPTIONS.map(opt => (
                    <button
                      key={opt.id}
                      onClick={() => setSelectedStyle(opt.id)}
                      className={`px-2 py-1.5 rounded-md text-xs transition ${
                        selectedStyle === opt.id ? 'bg-indigo-100 text-indigo-700 border border-indigo-300' : 'bg-slate-50 text-slate-600 border border-transparent hover:bg-slate-100'
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>
              {/* Structure */}
              <div>
                <div className="text-xs text-slate-400 font-medium mb-2">结构</div>
                <div className="grid grid-cols-2 gap-1.5">
                  {STRUCTURE_OPTIONS.map(opt => (
                    <button
                      key={opt.id}
                      onClick={() => setSelectedStructure(opt.id)}
                      className={`px-2 py-1.5 rounded-md text-xs transition ${
                        selectedStructure === opt.id ? 'bg-indigo-100 text-indigo-700 border border-indigo-300' : 'bg-slate-50 text-slate-600 border border-transparent hover:bg-slate-100'
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>
              {/* Tone */}
              <div>
                <div className="text-xs text-slate-400 font-medium mb-2">色调</div>
                <div className="grid grid-cols-2 gap-1.5">
                  {TONE_OPTIONS.map(opt => (
                    <button
                      key={opt.id}
                      onClick={() => setSelectedTone(opt.id)}
                      className={`px-2 py-1.5 rounded-md text-xs transition ${
                        selectedTone === opt.id ? 'bg-indigo-100 text-indigo-700 border border-indigo-300' : 'bg-slate-50 text-slate-600 border border-transparent hover:bg-slate-100'
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>
              {/* Scene */}
              <div>
                <div className="text-xs text-slate-400 font-medium mb-2">场景</div>
                <div className="grid grid-cols-2 gap-1.5">
                  {SCENE_OPTIONS.map(opt => (
                    <button
                      key={opt.id}
                      onClick={() => setSelectedScene(opt.id)}
                      className={`px-2 py-1.5 rounded-md text-xs transition ${
                        selectedScene === opt.id ? 'bg-indigo-100 text-indigo-700 border border-indigo-300' : 'bg-slate-50 text-slate-600 border border-transparent hover:bg-slate-100'
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>
              {/* Layout Preference */}
              <div>
                <div className="text-xs text-slate-400 font-medium mb-2">布局偏好</div>
                <div className="space-y-1">
                  {LAYOUT_OPTIONS.map(opt => (
                    <button
                      key={opt.id}
                      onClick={() => {
                        if (opt.id === 'auto') {
                          setSelectedLayouts(['auto']);
                        } else {
                          setSelectedLayouts(prev => {
                            const without = prev.filter(x => x !== 'auto');
                            return without.includes(opt.id) ? without.filter(x => x !== opt.id) : [...without, opt.id];
                          });
                        }
                      }}
                      className={`w-full text-left px-2 py-1.5 rounded-md text-xs transition flex items-center justify-between ${
                        selectedLayouts.includes(opt.id) ? 'bg-indigo-100 text-indigo-700 border border-indigo-300' : 'bg-slate-50 text-slate-600 border border-transparent hover:bg-slate-100'
                      }`}
                    >
                      <span>{opt.label}</span>
                      <span className="text-[10px] text-slate-400">{opt.desc}</span>
                    </button>
                  ))}
                </div>
              </div>
            </div>
          </div>

          {/* Chat Sidebar - Modern & Elegant */}
        <div 
          className={`
            relative border-l border-slate-100/60 bg-white flex flex-col transition-all duration-300 ease-[cubic-bezier(0.25,0.1,0.25,1)]
            ${isChatOpen ? 'translate-x-0' : 'w-0 translate-x-full opacity-0 overflow-hidden'}
            shadow-xl z-20
          `}
          style={{ width: isChatOpen ? `${chatWidth}px` : 0 }}
        >
          {isChatOpen && (
            <button
              aria-label="调整对话面板宽度"
              className={`absolute left-0 top-0 z-30 h-full w-2 -translate-x-1 cursor-col-resize transition-colors ${isResizingChat ? 'bg-indigo-300/70' : 'bg-transparent hover:bg-indigo-200/60'}`}
              onMouseDown={handleStartChatResize}
              type="button"
            />
          )}
          {/* Chat Header */}
          <div className="h-14 px-5 border-b border-slate-100 flex items-center justify-between shrink-0 bg-white/80 backdrop-blur-sm sticky top-0 z-10">
            <div className="flex items-center gap-3 flex-1 min-w-0">
              <div className="flex items-center justify-center w-8 h-8 rounded-full bg-gradient-to-br from-indigo-500 to-purple-600 text-white shadow-md shadow-indigo-200 shrink-0 ring-2 ring-white">
                <Icons.Sparkles className="w-4 h-4" />
              </div>
              <div className="flex-1 min-w-0">
                <select 
                  value={selectedAgentId} 
                  onChange={handleAgentChange}
                  className="w-full bg-transparent text-sm font-bold text-slate-800 focus:outline-none cursor-pointer truncate appearance-none pr-4"
                  style={{ backgroundImage: 'none' }}
                >
                  {agents.filter((agent) => agent.agentId === PPT_AGENT_ID).length === 0 && (
                    <option value={PPT_AGENT_ID}>AI交互式PPT生成智能体</option>
                  )}
                  {agents.filter((agent) => agent.agentId === PPT_AGENT_ID).map(agent => (
                    <option key={agent.agentId} value={agent.agentId}>
                      {agent.agentName}
                    </option>
                  ))}
                </select>
                <div className="flex items-center gap-1.5 mt-0.5">
                   <span className="w-1.5 h-1.5 rounded-full bg-green-500 animate-pulse"></span>
                   <span className="text-[10px] text-slate-500 font-medium leading-tight">AI Assistant Online</span>
                </div>
              </div>
            </div>
            <div className="flex items-center gap-1">
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
            {messages.map((msg, index) => {
              return (
                <div 
                  key={`${msg.id}-${index}`} 
                  className={`flex gap-3 ${msg.role === 'user' ? 'flex-row-reverse' : 'flex-row'}`}
                >
                  <div className={`
                    shrink-0 w-8 h-8 rounded-full flex items-center justify-center shadow-sm mt-1 ring-2 ring-white
                    ${msg.role === 'user' 
                      ? 'bg-indigo-100 text-indigo-600' 
                      : 'bg-white text-indigo-500 border border-slate-100'
                    }
                  `}>
                    {msg.role === 'user' ? <Icons.User className="w-5 h-5" /> : <Icons.Bot className="w-5 h-5" />}
                  </div>
                  
                  <div className="flex flex-col max-w-[85%] w-full">
                      <span className={`text-[10px] mb-1.5 font-medium ${msg.role === 'user' ? 'text-right text-slate-400' : 'text-left text-slate-400'}`}>
                          {msg.role === 'user' ? 'You' : 'Agent'}
                      </span>
                      
                      <div className={`flex flex-col gap-2 ${msg.role === 'user' ? 'items-end' : 'items-start'}`}>
                        {/* Steps / Reasoning Block */}
                        {msg.role === 'agent' && ((msg.steps && msg.steps.length > 0) || msg.reasoning) && (
                          <div className="w-full max-w-full">
                            <div className="w-full">
                              <button
                                className="inline-flex max-w-full items-center gap-2 rounded-xl border border-slate-200 bg-white px-3 py-2 text-left text-xs font-medium text-slate-600 shadow-sm transition-all hover:border-indigo-200 hover:bg-indigo-50/40"
                                onClick={() => handleToggleThinking(msg.id)}
                                type="button"
                              >
                                 <Icons.Sparkles className="w-3.5 h-3.5 shrink-0 text-indigo-400" />
                                 <span className="truncate">
                                   {msg.thinkingOpen ? '收起思考过程' : `思考过程 · ${(msg.steps?.length || 0)} 个阶段`}
                                 </span>
                                 {msg.steps?.some(step => step.status === 'running') && (
                                   <Icons.Loader className="w-3.5 h-3.5 shrink-0 text-indigo-500" />
                                 )}
                              </button>
                              {msg.thinkingOpen && (
                              <div className="mt-2 flex max-h-[360px] flex-col gap-2 overflow-y-auto rounded-2xl border border-slate-200 bg-gradient-to-b from-white to-slate-50/80 p-3 text-sm text-slate-600 shadow-sm">
                                 {msg.steps && msg.steps.length > 0 ? (
                                    msg.steps.map((step, idx) => (
                                      <details key={step.id || `${step.phase}-${idx}`} className="group rounded-xl border border-slate-100 bg-white shadow-sm">
                                          <summary className="flex cursor-pointer list-none items-center gap-2 px-3 py-2 text-xs font-semibold text-slate-700 [&::-webkit-details-marker]:hidden">
                                              {step.status === 'running' ? (
                                                 <Icons.Loader className="w-3.5 h-3.5 text-indigo-500" />
                                              ) : (
                                                 <span className="text-green-500">完成</span>
                                              )}
                                              <span className="min-w-0 flex-1 truncate">{step.label} · {step.summary || normalizeInlineText(step.content).slice(0, 34)}</span>
                                              {step.event && (
                                                <span className="shrink-0 rounded-full bg-slate-100 px-2 py-0.5 text-[10px] font-medium text-slate-400">
                                                  {step.event}
                                                </span>
                                              )}
                                          </summary>
                                          {step.content && (
                                              <div className="mx-3 mb-3 border-l-2 border-slate-100 pl-3 text-xs text-slate-500 prose prose-sm prose-slate max-w-none prose-p:my-1 prose-pre:my-2 prose-pre:bg-slate-100 prose-pre:text-slate-700">
                                                <ReactMarkdown remarkPlugins={[remarkGfm]}>{step.content}</ReactMarkdown>
                                            </div>
                                          )}
                                      </details>
                                    ))
                                  ) : (
                                   <div className="p-2 bg-white rounded-lg border border-slate-100 shadow-sm prose prose-sm prose-slate max-w-none prose-p:my-1 prose-pre:my-2 prose-pre:bg-slate-100 prose-pre:text-slate-700">
                                     <ReactMarkdown remarkPlugins={[remarkGfm]}>{msg.reasoning || ''}</ReactMarkdown>
                                   </div>
                                 )}
                              </div>
                              )}
                            </div>
                          </div>
                        )}

                        {/* Content Block */}
                        {msg.content && (
                          <div 
                            className={`
                                p-3.5 text-sm leading-relaxed shadow-sm whitespace-pre-wrap w-fit
                                ${msg.role === 'user' 
                                ? 'bg-indigo-600 text-white rounded-2xl rounded-tr-sm shadow-indigo-200' 
                                : 'bg-white border border-slate-200 text-slate-700 rounded-2xl rounded-tl-sm shadow-sm prose prose-sm prose-slate max-w-none overflow-x-auto prose-p:my-1 prose-pre:my-2 prose-pre:bg-slate-100 prose-pre:text-slate-700'
                                }
                            `}
                          >
                            {msg.role === 'user' ? (
                               msg.content
                            ) : (
                               <ReactMarkdown remarkPlugins={[remarkGfm]}>{msg.content}</ReactMarkdown>
                            )}
                          </div>
                        )}

                        {/* Empty state while generating */}
                        {msg.role === 'agent' && !msg.content && !msg.reasoning && isSending && (
                           <div className="flex gap-1 items-center px-4 py-3 text-sm shadow-sm bg-white border border-indigo-100 text-indigo-600 rounded-2xl rounded-tl-sm">
                             <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: '0ms' }}></span>
                             <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: '150ms' }}></span>
                             <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: '300ms' }}></span>
                           </div>
                        )}
                      </div>
                  </div>
                </div>
              );
            })}

            {/* Stream Progress Indicator */}
            {isSending && streamPhase !== 'done' && (
              <div className="flex gap-3 flex-row animate-in fade-in duration-300">
                <div className="shrink-0 w-8 h-8 rounded-full flex items-center justify-center shadow-sm mt-1 ring-2 ring-white bg-white text-indigo-500 border border-slate-100 opacity-50">
                  <Icons.Bot className="w-5 h-5" />
                </div>
                <div className="flex flex-col max-w-[85%]">
                  <div className="px-4 py-3 text-sm shadow-sm bg-white border border-indigo-100 text-indigo-600 rounded-2xl rounded-tl-sm flex items-center gap-3">
                    <div className="flex gap-1">
                      <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: '0ms' }}></span>
                      <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: '150ms' }}></span>
                      <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-bounce" style={{ animationDelay: '300ms' }}></span>
                    </div>
                  </div>
                </div>
              </div>
            )}

            <div ref={messagesEndRef} />
          </div>

          {/* Input Area */}
          <div className="p-4 bg-white border-t border-slate-100 shrink-0 relative z-20 shadow-[0_-4px_6px_-1px_rgba(0,0,0,0.02)]">
            {/* Quick Actions - Only show when chat is empty (just greeting) */}
            {messages.length <= 1 && (
              <div className="flex flex-wrap gap-2 mb-3 px-1 animate-in fade-in slide-in-from-bottom-2 duration-300">
                {quickActions.map((action, idx) => (
                  <button
                    key={idx}
                    onClick={() => {
                      setInputValue(action.text);
                      setTimeout(() => {
                         const textarea = document.querySelector('textarea');
                         if (textarea) {
                             textarea.style.height = 'auto';
                             textarea.style.height = Math.min(textarea.scrollHeight, 240) + 'px';
                         }
                      }, 10);
                    }}
                    className="text-xs px-3 py-1.5 bg-indigo-50 text-indigo-600 rounded-full hover:bg-indigo-100 transition-colors border border-indigo-100 font-medium shadow-sm text-left"
                  >
                    {action.label}
                  </button>
                ))}
              </div>
            )}

            <div className="relative flex items-end gap-2 bg-slate-50 p-2 rounded-2xl border border-slate-200 focus-within:border-indigo-400 focus-within:ring-4 focus-within:ring-indigo-100 focus-within:bg-white transition-all shadow-sm">
              <textarea
                value={inputValue}
                onChange={(e) => {
                  setInputValue(e.target.value);
                  e.target.style.height = 'auto';
                  e.target.style.height = Math.min(e.target.scrollHeight, 300) + 'px';
                }}
                onKeyDown={handleKeyDown}
                placeholder={isSending ? "AI 正在生成中..." : "输入您的问题，描述您的需求..."}
                disabled={isSending}
                className="flex-1 px-4 py-3 bg-transparent border-none focus:ring-0 text-[15px] text-slate-800 placeholder:text-slate-400 resize-none max-h-[300px] min-h-[80px] scrollbar-thin scrollbar-thumb-slate-200 scrollbar-track-transparent disabled:opacity-50 disabled:cursor-not-allowed outline-none leading-relaxed"
                rows={1}
                style={{ height: 'auto', minHeight: '80px' }}
              />
              <div className="flex gap-1 mb-0.5 shrink-0">
                  {isSending ? (
                    <button
                      onClick={handleStopStream}
                      className="p-2.5 rounded-lg transition-all duration-200 flex items-center justify-center bg-red-100 text-red-600 hover:bg-red-200 shadow-sm"
                      title="停止生成"
                    >
                      <Icons.Square className="w-4 h-4" />
                    </button>
                  ) : (
                    <button
                      onClick={handleSendMessage}
                      disabled={!inputValue.trim()}
                      className={`
                        p-2.5 rounded-lg transition-all duration-200 flex items-center justify-center
                        ${inputValue.trim()
                          ? 'bg-indigo-600 text-white shadow-md shadow-indigo-200 hover:bg-indigo-700 hover:scale-105 active:scale-95' 
                          : 'bg-slate-200 text-slate-400 cursor-not-allowed'
                        }
                      `}
                      title="发送消息"
                    >
                      <Icons.Send className="w-4 h-4" />
                    </button>
                  )}
                  <button
                    onClick={handleRestartSession}
                    disabled={isSending}
                    className="p-2.5 rounded-lg bg-white text-slate-400 hover:bg-slate-50 hover:text-indigo-600 transition-all duration-200 border border-slate-200 hover:border-indigo-100 shadow-sm disabled:opacity-50 disabled:cursor-not-allowed"
                    title="重启对话"
                  >
                    <Icons.Plus className="w-4 h-4" />
                  </button>
              </div>
            </div>

            {/* Model Selection - below textarea */}
            <div className="flex items-center gap-2 mt-2 px-1">
                <div className="relative flex items-center bg-white border border-slate-200 rounded-full shadow-sm hover:border-slate-300 transition-colors">
                    <Icons.Sparkles className={`w-3 h-3 ml-2 ${selectedCustomModelId !== 'default' ? 'text-indigo-500' : 'text-slate-400'}`} />
                    <select
                        value={selectedCustomModelId}
                        onChange={(e) => {
                            if (e.target.value === 'add_new') {
                                setShowApiConfig(true);
                                e.target.value = selectedCustomModelId;
                            } else {
                                setSelectedCustomModelId(e.target.value);
                                localStorage.setItem('ai_agent_selected_model', e.target.value);
                            }
                        }}
                        className="appearance-none bg-transparent border-none text-[11px] font-medium text-slate-600 focus:ring-0 py-1 pl-1 pr-5 cursor-pointer outline-none"
                    >
                        <option value="default">默认模型</option>
                        {customModels.filter(m => m.enabled).map(m => (
                            <option key={m.id} value={m.id}>{m.name || m.model}</option>
                        ))}
                        <option disabled>──────────</option>
                        <option value="add_new">+ 管理模型</option>
                    </select>
                    <div className="pointer-events-none absolute inset-y-0 right-0 flex items-center px-1.5 text-slate-400">
                        <svg className="fill-current h-3 w-3" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path d="M9.293 12.95l.707.707L15.657 8l-1.414-1.414L10 10.828 5.757 6.586 4.343 8z"/></svg>
                    </div>
                </div>
                <span className="text-[10px] text-slate-400 ml-auto hidden sm:inline">
                    <kbd className="px-1 py-0.5 bg-slate-100 border border-slate-200 rounded text-slate-500">Enter</kbd> 发送
                </span>
            </div>
            <div className="text-center mt-1.5">
                <p className="text-[10px] text-slate-400">
                  {isSending ? (streamProgress || '生成中...') : ''}
                </p>
            </div>
          </div>
        </div>
      </div>
    </div>
      {showApiConfig && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-6 animate-in fade-in duration-200">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in-95 duration-200 border border-white/20 flex flex-col max-h-[90vh]">
                <div className="px-6 py-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center shrink-0">
                    <div className="flex items-center gap-2.5">
                      <div className="p-1.5 bg-indigo-50 rounded-lg">
                        <Icons.Sparkles className="w-4 h-4 text-indigo-500" />
                      </div>
                      <h2 className="text-base font-bold text-slate-800">自定义模型配置</h2>
                    </div>
                    <button
                        onClick={() => {
                          setShowApiConfig(false);
                          setEditingModel(null);
                        }}
                        className="p-1.5 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
                    >
                        <Icons.Close className="w-4 h-4" />
                    </button>
                </div>

                <div className="flex flex-1 overflow-hidden">
                    {/* List of Models */}
                    <div className="w-1/3 border-r border-slate-100 bg-slate-50 flex flex-col">
                        <div className="p-3">
                            <button
                                onClick={handleAddNewModel}
                                className="w-full flex items-center justify-center gap-2 py-2 bg-white border border-indigo-200 text-indigo-600 rounded-lg hover:bg-indigo-50 transition-colors shadow-sm text-sm font-medium"
                            >
                                <Icons.Plus className="w-4 h-4" /> 添加模型
                            </button>
                        </div>
                        <div className="flex-1 overflow-y-auto p-3 space-y-2">
                            {customModels.map(model => (
                                <div
                                    key={model.id}
                                    onClick={() => setEditingModel(model)}
                                    className={`p-3 rounded-xl border cursor-pointer transition-all ${editingModel?.id === model.id ? 'bg-indigo-50 border-indigo-200 shadow-sm ring-1 ring-indigo-100' : 'bg-white border-slate-200 hover:border-indigo-100 hover:shadow-sm'}`}
                                >
                                    <div className="flex items-center justify-between mb-1">
                                        <div className="font-semibold text-sm text-slate-800 truncate pr-2">{model.name}</div>
                                        <div className="flex items-center gap-1 shrink-0">
                                            {/* Toggle Switch */}
                                            <button
                                                onClick={(e) => {
                                                    e.stopPropagation();
                                                    const newModels = customModels.map(m => m.id === model.id ? {...m, enabled: !m.enabled} : m);
                                                    saveCustomModels(newModels);
                                                    if (!(!model.enabled) && selectedCustomModelId === model.id) {
                                                        setSelectedCustomModelId('default');
                                                    }
                                                }}
                                                className={`relative inline-flex h-4 w-7 items-center rounded-full transition-colors focus:outline-none ${model.enabled ? 'bg-indigo-500' : 'bg-slate-300'}`}
                                            >
                                                <span className={`inline-block h-3 w-3 transform rounded-full bg-white transition-transform ${model.enabled ? 'translate-x-3.5' : 'translate-x-0.5'}`} />
                                            </button>
                                            <button onClick={(e) => { e.stopPropagation(); handleDeleteModel(model.id); }} className="text-slate-400 hover:text-red-500 ml-1">
                                                <Icons.Trash className="w-3.5 h-3.5" />
                                            </button>
                                        </div>
                                    </div>
                                    <div className="text-[10px] text-slate-500 truncate">{model.model}</div>
                                </div>
                            ))}
                            {customModels.length === 0 && (
                                <div className="text-center text-xs text-slate-400 py-6">
                                    暂无自定义模型<br/>点击上方按钮添加
                                </div>
                            )}
                        </div>
                    </div>

                    {/* Edit Form */}
                    <div className="flex-1 p-6 overflow-y-auto bg-white">
                        {editingModel ? (
                            <div className="space-y-4 animate-in fade-in duration-200">
                                <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">展示名称</label>
                                    <input type="text" value={editingModel.name} onChange={e => setEditingModel({...editingModel, name: e.target.value})} className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all" placeholder="例如：我的GPT-4o" />
                                </div>
                                <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">模型名称 (Model)</label>
                                    <input type="text" value={editingModel.model} onChange={e => setEditingModel({...editingModel, model: e.target.value})} className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all" placeholder="例如：gpt-4o" />
                                </div>
                                <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">Base URL</label>
                                    <input type="text" value={editingModel.baseUrl} onChange={e => setEditingModel({...editingModel, baseUrl: e.target.value})} className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all" placeholder="例如：https://api.openai.com" />
                                </div>
                                <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">API Key</label>
                                    <input type="password" value={editingModel.apiKey} onChange={e => setEditingModel({...editingModel, apiKey: e.target.value})} className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all" placeholder="sk-..." />
                                </div>
                                <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">Completions Path (可选)</label>
                                    <input type="text" value={editingModel.completionsPath} onChange={e => setEditingModel({...editingModel, completionsPath: e.target.value})} className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all" placeholder="默认为 v1/chat/completions" />
                                </div>
                                <div className="pt-2 flex justify-end">
                                    <button onClick={handleSaveEditingModel} className="px-6 py-2 bg-indigo-600 text-white font-medium rounded-lg hover:bg-indigo-700 shadow-sm transition-all text-sm">
                                        保存配置
                                    </button>
                                </div>
                            </div>
                        ) : (
                            <div className="h-full flex flex-col items-center justify-center text-slate-400">
                                <Icons.Sparkles className="w-12 h-12 mb-3 opacity-20" />
                                <p className="text-sm">选择左侧模型进行编辑，或点击添加</p>
                            </div>
                        )}
                    </div>
                </div>
            </div>
        </div>
      )}
{/* ===== Fullscreen Presentation ===== */}
      {isFullscreen && pptData && (
        <div
          className="fixed inset-0 z-50 bg-black flex items-center justify-center"
          onClick={() => setIsFullscreen(false)}
          onKeyDown={(e) => {
            if (e.key === 'Escape' || e.key === 'q') setIsFullscreen(false);
            else if (e.key === 'ArrowRight' || e.key === 'ArrowDown' || e.key === ' ') {
              e.preventDefault();
              setCurrentSlideIndex(Math.min(pptData.slides.length - 1, currentSlideIndex + 1));
            } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
              setCurrentSlideIndex(Math.max(0, currentSlideIndex - 1));
            }
          }}
          tabIndex={0}
          autoFocus
        >
          <div
            className="bg-white overflow-hidden relative"
            style={{
              width: '100vw',
              height: 'calc(100vw * 9 / 16)',
              maxHeight: '100vh',
              maxWidth: 'calc(100vh * 16 / 9)',
            }}
          >
            <div className="w-full h-full relative" style={{ fontFamily: "\"PingFang SC\", \"Microsoft YaHei\", system-ui, -apple-system, sans-serif", letterSpacing: "0.2px" }}>
              {renderThemeDecor(pptData.slides[currentSlideIndex], currentSlideIndex)}
              {renderSlideContent(pptData.slides[currentSlideIndex], currentSlideIndex)}
            </div>
          </div>
          <div className="absolute bottom-4 left-1/2 -translate-x-1/2 text-white/20 text-xs pointer-events-none">
            ESC 退出 · ← → 翻页
          </div>
          <div className="absolute bottom-4 right-6 text-white/20 text-xs font-mono pointer-events-none">
            {currentSlideIndex + 1} / {pptData.slides.length}
          </div>
        </div>
      )}
    </div>
  );
}
