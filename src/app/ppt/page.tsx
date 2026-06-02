'use client';

import { useRef, useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { getUserInfo, clearUserInfo } from '@/utils/cookie';
import { agentApi } from '@/api/agent';
import { AiAgentConfigResponseDTO } from '@/types/api';
import pptxgen from 'pptxgenjs';

// Message type definition
type Message = {
  id: string;
  role: 'user' | 'agent';
  content: string;
  timestamp: number;
};

// PPT Slide data structure from AI Agent
interface PptSlideElement {
  kind: 'text' | 'table' | 'shape' | 'image';
  content: string;
  x: number;
  y: number;
  w: number;
  h: number;
  fontSize?: number;
  color?: string;
  bold?: boolean;
  fill?: string;
  align?: 'left' | 'center' | 'right';
  rows?: string[][];
}

interface PptSlide {
  slideIndex: number;
  layout?: string;
  elements: PptSlideElement[];
}

interface PptData {
  title: string;
  slides: PptSlide[];
}

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
  lastModified: number;
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

// Default theme (used as fallback)
const DEFAULT_THEME = THEMES[0];

// Shape elements are now fully skipped (theme handles all decorations)

// Generate PPTX from PptData with professional theme
// Smart layout inference — ensure visual diversity even when AI doesn't specify layout
const inferLayout = (slideData: PptSlide, slideIdx: number, totalSlides: number): string => {
  let layout = slideData.layout || '';
  const isTitleSlide = layout === 'title_slide' || slideIdx === 0;
  const isEndSlide = layout === 'end_slide' || slideIdx === totalSlides - 1;
  if (!layout && !isTitleSlide && !isEndSlide) {
    const diverseLayouts = ['content_slide', 'card_3col', 'comparison', 'timeline', 'data_highlight', 'quote_slide'];
    layout = diverseLayouts[(slideIdx - 1) % diverseLayouts.length];
  }
  return layout;
};

const generatePptx = (data: PptData, theme: PptTheme) => {
  const pres = new pptxgen();
  pres.title = data.title;
  pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5

  data.slides.forEach((slideData, slideIdx) => {
    const slide = pres.addSlide();
    const layout = inferLayout(slideData, slideIdx, data.slides.length);
    const isTitleSlide = layout === 'title_slide' || slideIdx === 0;
    const isEndSlide = layout === 'end_slide' || slideIdx === data.slides.length - 1;

    // === STEP 1: Layout-specific theme decoration ===
    // Helper: add rectangle shape
    const addRect = (x: number, y: number, w: number, h: number, color: string) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      slide.addShape((pres as any).shapes.RECTANGLE, {
        x, y, w, h, fill: { color }, line: { width: 0 },
      });
    };

    if (isTitleSlide || isEndSlide) {
      // ── Cover / End: Full-width primary top + accent stripe + bottom bar ──
      addRect(0, 0, 13.33, theme.coverNavyHeight, theme.primary);
      addRect(0, theme.coverNavyHeight, 13.33, 0.12, theme.accent);
      addRect(0, 7.15, 13.33, 0.35, theme.primary);
      // Decorative: large circle watermark
      addRect(10.5, 0.6, 2.2, 2.2, theme.primaryLight);
    } else if (layout === 'card_3col') {
      // ── 3-Column Cards: Short header + 3 card background blocks ──
      addRect(0, 0, 13.33, 0.9, theme.primary);
      addRect(0, 0.9, 13.33, 0.06, theme.accent);
      // Three card backgrounds
      addRect(0.5, 1.6, 3.8, 5.0, theme.offWhite);
      addRect(4.75, 1.6, 3.8, 5.0, theme.offWhite);
      addRect(9.0, 1.6, 3.8, 5.0, theme.offWhite);
      // Card top accent bars
      addRect(0.5, 1.6, 3.8, 0.12, theme.primary);
      addRect(4.75, 1.6, 3.8, 0.12, theme.primary);
      addRect(9.0, 1.6, 3.8, 0.12, theme.primary);
      addRect(0, 7.15, 13.33, 0.35, theme.primary);
    } else if (layout === 'comparison') {
      // ── Comparison: Left-right symmetric color blocks ──
      addRect(0, 0, 13.33, 0.9, theme.primary);
      addRect(0, 0.9, 13.33, 0.06, theme.accent);
      // Left block
      addRect(0.5, 1.5, 5.9, 5.2, theme.offWhite);
      addRect(0.5, 1.5, 5.9, 0.1, theme.primary);
      // Right block
      addRect(6.9, 1.5, 5.9, 5.2, theme.offWhite);
      addRect(6.9, 1.5, 5.9, 0.1, theme.accent);
      // Center divider
      addRect(6.55, 1.5, 0.2, 5.2, theme.primaryLight);
      addRect(0, 7.15, 13.33, 0.35, theme.primary);
    } else if (layout === 'timeline') {
      // ── Timeline: Horizontal flow bar + node markers ──
      addRect(0, 0, 0.5, 7.5, theme.primary); // Left vertical bar
      addRect(0, 7.15, 13.33, 0.35, theme.primary);
      // Horizontal timeline bar
      addRect(1.2, 3.5, 11.5, 0.12, theme.primaryLight);
      // Node circles (decorative dots)
      addRect(2.5, 3.2, 0.7, 0.7, theme.primary);
      addRect(5.5, 3.2, 0.7, 0.7, theme.primary);
      addRect(8.5, 3.2, 0.7, 0.7, theme.primary);
      addRect(11.5, 3.2, 0.7, 0.7, theme.accent);
    } else if (layout === 'data_highlight') {
      // ── Data Highlight: Narrow header + big number zone + bottom color band ──
      addRect(0, 0, 13.33, 0.7, theme.primary);
      addRect(0, 0.7, 13.33, 0.06, theme.accent);
      addRect(0, 6.0, 13.33, 1.5, theme.offWhite); // Bottom highlight band
      addRect(0, 7.15, 13.33, 0.35, theme.primary);
      // Left accent vertical bar
      addRect(0.5, 1.2, 0.08, 4.5, theme.primary);
    } else if (layout === 'quote_slide') {
      // ── Quote: Left wide accent bar + soft background ──
      addRect(0, 0, 13.33, 7.5, theme.offWhite);
      addRect(0, 0, 0.8, 7.5, theme.primary); // Wide left bar
      addRect(0.8, 2.8, 0.12, 1.8, theme.accent); // Accent mark
      addRect(0, 7.15, 13.33, 0.35, theme.primary);
    } else {
      // ── Default content_slide: Left wide color band + right content area ──
      addRect(0, 0, 4.5, 7.15, theme.primary); // Left wide band
      addRect(4.5, 0, 0.08, 7.15, theme.accent); // Accent vertical divider
      addRect(0, 7.15, 13.33, 0.35, theme.primary); // Bottom bar
    }

    // === STEP 2: Render AI content elements (skip all shapes) ===
    slideData.elements.forEach((el) => {
      if (el.kind === 'shape') return;

      switch (el.kind) {
        case 'text': {
          let textY = el.y || 0;
          const textFontSize = el.fontSize || 18;
          let textColor = theme.bodyColor; // Default: use theme body color

          // Title color & position logic — varies by layout type
          if (textFontSize >= 24) {
            if (isTitleSlide || isEndSlide) {
              textColor = theme.white; // Cover/End: in primary area
            } else if (layout === 'content_slide') {
              textColor = theme.white; // Left band: title is white
            } else if (layout === 'card_3col' || layout === 'comparison') {
              textY = 0.15; // Short header bar
              textColor = theme.white;
            } else if (layout === 'data_highlight') {
              textY = 0.1; // Narrow header
              textColor = theme.white;
            } else if (layout === 'quote_slide') {
              textColor = theme.primary; // Quote: primary color title
            } else {
              textY = 0.2;
              textColor = theme.white;
            }
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
            x: el.x,
            y: textY,
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

        case 'table':
          if (el.rows && el.rows.length > 0) {
            const tableRows = el.rows.map((row, rowIdx) =>
              row.map((cell) => ({
                text: cell,
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
              x: el.x, y: el.y, w: el.w, h: el.h || 2,
              border: { pt: 1, color: 'C0C8D4' },
              colW: el.w / (el.rows[0]?.length || 1),
              rowH: 0.45,
              autoPage: true,
            });
          }
          break;

        case 'image':
          try {
            slide.addImage({ path: el.content, x: el.x, y: el.y, w: el.w, h: el.h });
          } catch { /* skip */ }
          break;
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
  const messagesEndRef = useRef<HTMLDivElement>(null);

  // Agent State
  const [agents, setAgents] = useState<AiAgentConfigResponseDTO[]>([]);
  const [selectedAgentId, setSelectedAgentId] = useState('');
  const [sessionId, setSessionId] = useState('');

  // PPT Preview State
  const [pptData, setPptData] = useState<PptData | null>(null);
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
        { id: 'quarterly-report', name: '季度工作汇报', desc: '项目进展、数据分析、下季规划', prompt: '请制作一份季度工作汇报PPT，包含项目进展、关键数据指标、团队成果和下季度规划，8-10页', scene: 'report', style: 'professional' },
        { id: 'annual-review', name: '年度总结报告', desc: '全年回顾、亮点成果、展望未来', prompt: '请制作年度总结报告PPT，包含全年核心成果、数据对比、里程碑事件和新年展望，10-12页', scene: 'report', style: 'professional' },
        { id: 'project-proposal', name: '项目立项方案', desc: '项目背景、目标、计划、预算', prompt: '请制作项目立项方案PPT，包含项目背景、目标、实施计划、资源需求和预算，8-10页', scene: 'report', style: 'professional' },
        { id: 'team-intro', name: '团队介绍', desc: '团队成员、职能分工、协作模式', prompt: '请制作团队介绍PPT，包含团队概况、成员介绍、职能分工和协作方式，6-8页', scene: 'general', style: 'professional' },
        { id: 'product-roadmap', name: '产品路线图', desc: '版本规划、里程碑、时间线', prompt: '请制作产品路线图PPT，包含版本规划、功能里程碑、时间线和资源安排，8-10页', scene: 'report', style: 'professional' },
        { id: 'meeting-summary', name: '会议纪要', desc: '议题、决议、行动项', prompt: '请制作会议纪要PPT，包含会议议题、讨论要点、决议事项和行动项，5-7页', scene: 'general', style: 'minimal' },
      ],
    },
    {
      category: 'roadshow',
      icon: '🚀',
      label: '路演融资',
      templates: [
        { id: 'startup-pitch', name: '创业融资路演', desc: '痛点、方案、市场、团队、融资', prompt: '请制作创业融资路演PPT，包含市场痛点、解决方案、商业模式、市场规模、团队介绍和融资需求，10-12页', scene: 'pitch', style: 'creative' },
        { id: 'investor-deck', name: '投资人推介', desc: '核心数据、增长曲线、竞品分析', prompt: '请制作投资人推介PPT，重点展示核心数据、用户增长、收入模型和竞争优势，8-10页', scene: 'pitch', style: 'professional' },
        { id: 'product-launch', name: '新品发布会', desc: '产品亮点、技术突破、市场策略', prompt: '请制作新品发布会PPT，包含产品亮点、技术突破、用户体验和市场策略，8-10页', scene: 'pitch', style: 'creative' },
        { id: 'business-plan', name: '商业计划书', desc: '市场分析、运营策略、财务预测', prompt: '请制作商业计划书PPT，包含市场分析、运营策略、收入模型和财务预测，10-12页', scene: 'pitch', style: 'professional' },
      ],
    },
    {
      category: 'education',
      icon: '📚',
      label: '培训教学',
      templates: [
        { id: 'onboarding', name: '新员工入职培训', desc: '公司文化、制度、流程', prompt: '请制作新员工入职培训PPT，包含公司文化、组织架构、规章制度和工作流程，8-10页', scene: 'training', style: 'professional' },
        { id: 'tech-sharing', name: '技术分享', desc: '技术方案、架构、最佳实践', prompt: '请制作技术分享PPT，包含技术背景、方案设计、架构图和最佳实践，8-10页', scene: 'training', style: 'academic' },
        { id: 'course-lecture', name: '课程讲义', desc: '知识体系、核心概念、案例', prompt: '请制作课程讲义PPT，包含知识体系、核心概念、实例解析和总结，8-12页', scene: 'training', style: 'academic' },
        { id: 'workshop', name: '工作坊指南', desc: '活动流程、互动环节、任务卡', prompt: '请制作工作坊指南PPT，包含活动流程、互动环节设计、小组任务和总结分享，6-8页', scene: 'training', style: 'creative' },
        { id: 'sop-training', name: '标准操作培训', desc: '流程步骤、注意事项、示例', prompt: '请制作标准操作流程培训PPT，包含操作步骤、注意事项、常见问题和示例，6-8页', scene: 'training', style: 'minimal' },
      ],
    },
    {
      category: 'data',
      icon: '📊',
      label: '数据报告',
      templates: [
        { id: 'data-dashboard', name: '数据看板报告', desc: '关键指标、趋势分析、预警', prompt: '请制作数据看板报告PPT，包含核心指标、趋势分析、异常预警和优化建议，8-10页', scene: 'report', style: 'professional' },
        { id: 'market-research', name: '市场调研报告', desc: '行业现状、用户画像、机会点', prompt: '请制作市场调研报告PPT，包含行业现状、用户画像分析、竞品对比和机会点，8-10页', scene: 'report', style: 'academic' },
        { id: 'competitive-analysis', name: '竞品分析', desc: '竞品对比、差异化、策略建议', prompt: '请制作竞品分析PPT，包含竞品功能对比、差异化分析、SWOT和策略建议，6-8页', scene: 'report', style: 'professional' },
        { id: 'user-research', name: '用户研究报告', desc: '用户反馈、行为分析、改进建议', prompt: '请制作用户研究报告PPT，包含用户反馈、行为分析、满意度数据和改进建议，8-10页', scene: 'report', style: 'academic' },
      ],
    },
    {
      category: 'creative',
      icon: '🎨',
      label: '创意设计',
      templates: [
        { id: 'brand-story', name: '品牌故事', desc: '品牌起源、价值观、愿景', prompt: '请制作品牌故事PPT，包含品牌起源、核心价值观、品牌故事和未来愿景，6-8页', scene: 'general', style: 'creative' },
        { id: 'event-plan', name: '活动策划方案', desc: '活动主题、流程、预算', prompt: '请制作活动策划方案PPT，包含活动主题、流程安排、场地布置和预算方案，8-10页', scene: 'general', style: 'creative' },
        { id: 'portfolio', name: '作品集展示', desc: '项目作品、设计理念、成果', prompt: '请制作作品集展示PPT，包含代表项目、设计理念、技术方案和项目成果，8-10页', scene: 'general', style: 'creative' },
        { id: 'proposal', name: '创意提案', desc: '创意概念、视觉呈现、执行方案', prompt: '请制作创意提案PPT，包含创意概念、视觉风格、执行方案和效果预期，6-8页', scene: 'pitch', style: 'creative' },
      ],
    },
    {
      category: 'personal',
      icon: '🌟',
      label: '个人成长',
      templates: [
        { id: 'resume', name: '个人简历', desc: '教育、经历、技能、项目', prompt: '请制作个人简历PPT，包含教育背景、工作经历、技能专长和代表项目，6-8页', scene: 'general', style: 'minimal' },
        { id: 'career-plan', name: '职业规划', desc: '现状分析、目标、发展路径', prompt: '请制作职业规划PPT，包含现状分析、职业目标、发展路径和行动计划，6-8页', scene: 'general', style: 'professional' },
        { id: 'year-review', name: '个人年度回顾', desc: '成就、学习、新年计划', prompt: '请制作个人年度回顾PPT，包含年度成就、学习成长、旅行记录和新年计划，6-8页', scene: 'general', style: 'creative' },
        { id: 'knowledge-map', name: '知识体系梳理', desc: '知识框架、核心要点、关联', prompt: '请制作知识体系梳理PPT，包含知识框架、核心要点、关联关系和学习建议，8-10页', scene: 'training', style: 'academic' },
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
    { id: 'auto', label: '自动', desc: 'AI自动选择布局' },
    { id: 'title_slide', label: '标题', desc: '封面页' },
    { id: 'content_slide', label: '内容', desc: '图文内容页' },
    { id: 'card_3col', label: '三栏', desc: '三栏卡片' },
    { id: 'comparison', label: '对比', desc: '左右对比' },
    { id: 'data_highlight', label: '数据', desc: '数据突出' },
    { id: 'timeline', label: '时间线', desc: '流程/时间线' },
  ];
  const activeTheme = THEMES.find(t => t.id === selectedThemeId) || DEFAULT_THEME;

  // Session State
  const [sessions, setSessions] = useState<Session[]>([]);
  const [currentSessionId, setCurrentSessionId] = useState<string | null>(null);

  // Auto-scroll
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  // Check Login & Load Agents
  useEffect(() => {
    const userInfo = getUserInfo();
    if (!userInfo || !userInfo.user) {
      router.push('/login');
      return;
    }
    setCurrentUser(userInfo.user);

    const loadAgents = async () => {
      try {
        const res = await agentApi.queryAiAgentConfigList();
        setAgents(res.data || []);
        if (res.data && res.data.length > 0) {
          // Prefer PPT agent if available, otherwise first
          const pptAgent = res.data.find((a) => a.agentName?.includes('PPT') || a.agentDesc?.includes('PPT'));
          const lastAgentId = localStorage.getItem('ai_ppt_last_agent');
          if (lastAgentId && res.data.find((a) => a.agentId === lastAgentId)) {
            setSelectedAgentId(lastAgentId);
          } else if (pptAgent) {
            setSelectedAgentId(pptAgent.agentId);
          } else {
            setSelectedAgentId(res.data[0].agentId);
          }
        }
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
      lastModified: Date.now(),
    };

    setSessions((prev) => [newSession, ...prev]);
    setCurrentSessionId(newSession.id);
    setMessages(newSession.messages);
    setSessionId(backendId);
    setPptData(null);
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
      setCurrentSlideIndex(0);
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

  const handleAgentChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newAgentId = e.target.value;
    setSelectedAgentId(newAgentId);
    setSessionId('');
    localStorage.setItem('ai_ppt_last_agent', newAgentId);
  };

  const handleSendMessage = async () => {
    if (!inputValue.trim() || isSending) return;

    const content = inputValue;
    setInputValue('');
    setIsSending(true);

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
    const enrichedContent = `[设计指令: ${styleHints.join(', ')}]\n\n${content}`;

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
    setMessages((prev) => [...prev, userMsg]);

    try {
      // Ensure session
      let activeBackendSessionId = sessionId;
      if (!activeBackendSessionId) {
        const sessionRes = await agentApi.createSession(selectedAgentId, currentUser);
        activeBackendSessionId = sessionRes.data.sessionId;
        setSessionId(activeBackendSessionId);
      }

      // Send message
      const chatRes = await agentApi.chat({
        agentId: selectedAgentId,
        userId: currentUser,
        sessionId: activeBackendSessionId,
        message: enrichedContent,
      });

      const { type, content: resContent } = chatRes.data;

      // Helper: try to parse response as PPT data regardless of type field
      const stripMdCodeBlock = (s: string): string => s.replace(/^```(?:json)?\s*\n?/i, '').replace(/\n?```\s*$/i, '').trim();
      const tryParsePpt = (raw: unknown): PptData | null => {
        try {
          let obj = typeof raw === 'string' ? JSON.parse(stripMdCodeBlock(raw)) : raw;
          // Deep unwrap: handle multiple nesting levels
          // e.g. {type:"ppt", content: "{\"type\":\"ppt\",\"content\":{...}}"} — string-in-object
          // e.g. {type:"ppt", content: {title, slides}} — nested object
          for (let depth = 0; depth < 5; depth++) {
            if (obj === null || obj === undefined) return null;
            if (obj.title && Array.isArray(obj.slides)) return obj as PptData;
            if (Array.isArray(obj.slides) && obj.slides.length > 0) return obj as PptData;
            // Unwrap: try obj.content, obj.data, obj.result (may be string or object)
            const inner = obj.content ?? obj.data ?? obj.result ?? null;
            if (inner === null) return null;
            obj = typeof inner === 'string' ? JSON.parse(stripMdCodeBlock(inner)) : inner;
          }
          return null;
        } catch { return null; }
      };

      // First: try to detect PPT data in any response type
      const detectedPpt = tryParsePpt(resContent);
      if (detectedPpt) {
        setPptData(detectedPpt);
        setCurrentSlideIndex(0);
        setSessions((prev) =>
          prev.map((session) => {
            if (session.id === currentSessionId) {
              return { ...session, pptData: detectedPpt, lastModified: Date.now() };
            }
            return session;
          })
        );
        setMessages((prev) => [
          ...prev,
          {
            id: (Date.now() + 1).toString(),
            role: 'agent',
            content: `✅ PPT 已生成！共 ${detectedPpt.slides?.length || 0} 页，可以预览或下载。`,
            timestamp: Date.now(),
          },
        ]);
      } else if (type === 'user') {
        // AI asks for more info
        setMessages((prev) => [
          ...prev,
          {
            id: (Date.now() + 1).toString(),
            role: 'agent',
            content: resContent,
            timestamp: Date.now(),
          },
        ]);
      } else if (type === 'drawio') {
        // Agent returned drawio type (wrong agent), inform user
        setMessages((prev) => [
          ...prev,
          {
            id: (Date.now() + 1).toString(),
            role: 'agent',
            content: '⚠️ 当前智能体返回了 Draw.io 格式数据，请切换到 Draw.io 工作区使用，或在 PPT 工作区选择 PPT 专用智能体。',
            timestamp: Date.now(),
          },
        ]);
      } else {
        // Fallback - treat as text
        setMessages((prev) => [
          ...prev,
          {
            id: (Date.now() + 1).toString(),
            role: 'agent',
            content: typeof resContent === 'string' ? resContent : JSON.stringify(resContent),
            timestamp: Date.now(),
          },
        ]);
      }
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
    } finally {
      setIsSending(false);
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
      setCurrentSlideIndex(0);

      if (currentSessionId) {
        setSessions((prev) =>
          prev.map((session) => {
            if (session.id === currentSessionId) {
              return { ...session, backendSessionId: newBackendId, messages: [initialMsg], pptData: null, lastModified: Date.now() };
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
    if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
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
    const isTitleSlide = layout === 'title_slide' || slideIdx === 0;
    const isEndSlide = layout === 'end_slide' || slideIdx === (pptData?.slides.length ?? 1) - 1;
    const coverNavyPct = (t.coverNavyHeight / 7.5) * 100;
    const hdrPct = (t.contentHeaderHeight / 7.5) * 100;

    if (isTitleSlide || isEndSlide) {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: `${coverNavyPct}%`, backgroundColor: `#${t.primary}` }} />
        <div className="absolute inset-x-0" style={{ top: `${coverNavyPct}%`, height: '2%', backgroundColor: `#${t.accent}` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', backgroundColor: `#${t.primary}` }} />
      </>;
    } else if (layout === 'card_3col') {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: `${(0.9/7.5)*100}%`, backgroundColor: `#${t.primary}` }} />
        <div className="absolute rounded-sm" style={{ left:'5%', top:'28%', width:'27%', height:'50%', backgroundColor:`#${t.offWhite}`, borderTop:`2px solid #${t.primary}` }} />
        <div className="absolute rounded-sm" style={{ left:'37%', top:'28%', width:'27%', height:'50%', backgroundColor:`#${t.offWhite}`, borderTop:`2px solid #${t.primary}` }} />
        <div className="absolute rounded-sm" style={{ left:'69%', top:'28%', width:'27%', height:'50%', backgroundColor:`#${t.offWhite}`, borderTop:`2px solid #${t.primary}` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', backgroundColor: `#${t.primary}` }} />
      </>;
    } else if (layout === 'comparison') {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: `${(0.9/7.5)*100}%`, backgroundColor: `#${t.primary}` }} />
        <div className="absolute rounded-sm" style={{ left:'4%', top:'22%', width:'44%', height:'60%', backgroundColor:`#${t.offWhite}`, borderTop:`2px solid #${t.primary}` }} />
        <div className="absolute rounded-sm" style={{ left:'52%', top:'22%', width:'44%', height:'60%', backgroundColor:`#${t.offWhite}`, borderTop:`2px solid #${t.accent}` }} />
        <div className="absolute" style={{ left:'49.5%', top:'22%', width:'1%', height:'60%', backgroundColor:`#${t.primaryLight}` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', backgroundColor: `#${t.primary}` }} />
      </>;
    } else if (layout === 'timeline') {
      return <>
        <div className="absolute" style={{ left:0, top:0, width:'4%', height:'100%', backgroundColor:`#${t.primary}` }} />
        <div className="absolute" style={{ left:'10%', top:'46%', width:'85%', height:'2%', backgroundColor:`#${t.primaryLight}` }} />
        <div className="absolute rounded-full" style={{ left:'16%', top:'40%', width:'12%', height:'12%', backgroundColor:`#${t.primary}` }} />
        <div className="absolute rounded-full" style={{ left:'40%', top:'40%', width:'12%', height:'12%', backgroundColor:`#${t.primary}` }} />
        <div className="absolute rounded-full" style={{ left:'64%', top:'40%', width:'12%', height:'12%', backgroundColor:`#${t.primary}` }} />
        <div className="absolute rounded-full" style={{ left:'88%', top:'40%', width:'12%', height:'12%', backgroundColor:`#${t.accent}` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', backgroundColor: `#${t.primary}` }} />
      </>;
    } else if (layout === 'data_highlight') {
      return <>
        <div className="absolute inset-x-0 top-0" style={{ height: `${(0.7/7.5)*100}%`, backgroundColor: `#${t.primary}` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '22%', backgroundColor: `#${t.offWhite}` }} />
        <div className="absolute" style={{ left:'4%', top:'16%', width:'0.5%', height:'60%', backgroundColor:`#${t.primary}` }} />
        <div className="absolute inset-x-0" style={{ bottom: '5%', height: '5%', backgroundColor: `#${t.primary}` }} />
      </>;
    } else if (layout === 'quote_slide') {
      return <>
        <div className="absolute inset-x-0" style={{ backgroundColor: `#${t.offWhite}`, inset:0 }} />
        <div className="absolute" style={{ left:0, top:0, width:'6%', height:'100%', backgroundColor:`#${t.primary}` }} />
        <div className="absolute" style={{ left:'6%', top:'37%', width:'1%', height:'24%', backgroundColor:`#${t.accent}` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', backgroundColor: `#${t.primary}` }} />
      </>;
    } else {
      // Default content_slide: Left wide band
      return <>
        <div className="absolute" style={{ left:0, top:0, width:'34%', height:'95%', backgroundColor:`#${t.primary}` }} />
        <div className="absolute" style={{ left:'34%', top:0, width:'0.5%', height:'95%', backgroundColor:`#${t.accent}` }} />
        <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', backgroundColor: `#${t.primary}` }} />
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
            const yPct = (el.y || 0) / 7.5 * 100;
            const hPct = Math.max(8, (el.h || 1) / 7.5 * 100);
            return (
              <div
                key={i}
                className="absolute overflow-hidden"
                style={{
                  left: '6%',
                  right: '6%',
                  top: `${Math.min(yPct, 85)}%`,
                  height: `${hPct}%`,
                  color: isTitle ? `#${t.titleColor}` : `#${t.bodyColor}`,
                  fontSize: isTitle ? '3px' : '2px',
                  fontWeight: isTitle ? 'bold' : 'normal',
                  lineHeight: 1.2,
                  whiteSpace: 'nowrap',
                  textOverflow: 'ellipsis',
                }}
              >
                {el.content.replace(/[•\-]/g, '').trim().slice(0, 30)}
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

      const xPct = ((el.x || 0) / 13.33) * 100;
      let yPct = ((el.y || 0) / 7.5) * 100;
      const wPct = ((el.w || 4) / 13.33) * 100;
      const hPct = ((el.h || 1) / 7.5) * 100;

      if (el.kind === 'text') {
        const fontSizeNum = el.fontSize || 18;
        let fontSize = Math.max(8, fontSizeNum * 0.7);
        let textColor = `#${activeTheme.bodyColor}`; // Default: use theme body color

        // Title color & position logic — varies by layout type
        if (fontSizeNum >= 24) {
          if (isTitleSlide || isEndSlide) {
            // Cover/End: title in primary color area, white
            textColor = `#${activeTheme.white}`;
          } else if (layout === 'content_slide') {
            // Content slide with left band: title in left band (white), keep AI y position
            textColor = `#${activeTheme.white}`;
          } else if (layout === 'card_3col' || layout === 'comparison') {
            // Card/Comparison: title in short header bar (white)
            yPct = (0.15 / 7.5) * 100;
            textColor = `#${activeTheme.white}`;
          } else if (layout === 'data_highlight') {
            // Data highlight: title in narrow header (white)
            yPct = (0.1 / 7.5) * 100;
            textColor = `#${activeTheme.white}`;
          } else if (layout === 'quote_slide') {
            // Quote: title uses primary color
            textColor = `#${activeTheme.primary}`;
          } else {
            // Fallback: push to top
            yPct = (0.2 / 7.5) * 100;
            textColor = `#${activeTheme.white}`;
          }
        }
        // Large decorative text (big numbers etc): allow AI color choice
        if (el.color && fontSizeNum >= 30) {
          textColor = `#${el.color}`; // Large decorative text: respect AI color choice
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
              alignItems: el.fontSize && el.fontSize >= 30 ? 'center' : 'flex-start',
              lineHeight: 1.5,
              whiteSpace: 'pre-wrap',
              paddingTop: el.fontSize && el.fontSize < 22 ? '2px' : '0',
            }}
            dangerouslySetInnerHTML={{ __html: el.content.replace(/\n/g, '<br/>') }}
          />
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
            }}
          />
        );
      }

      if (el.kind === 'table' && el.rows) {
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
                {el.rows.map((row, ri) => (
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
                        {cell}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );
      }

      return null;
    });
  };

  // Render current slide detail
  // Helper: render a positioned div for theme decoration in preview
  const renderDecorDiv = (x: number, y: number, w: number, h: number, color: string, key: string) => {
    return (
      <div
        key={key}
        className="absolute"
        style={{
          left: `${(x / 13.33) * 100}%`,
          top: `${(y / 7.5) * 100}%`,
          width: `${(w / 13.33) * 100}%`,
          height: `${(h / 7.5) * 100}%`,
          backgroundColor: `#${color}`,
        }}
      />
    );
  };

  // Render theme decorations for preview — layout-specific skeletons
  const renderThemeDecor = (slideData: PptSlide, slideIdx: number) => {
    const layout = inferLayout(slideData, slideIdx, pptData?.slides.length ?? 1);
    const isTitleSlide = layout === 'title_slide' || slideIdx === 0;
    const isEndSlide = layout === 'end_slide' || slideIdx === (pptData?.slides.length ?? 1) - 1;
    const t = activeTheme;
    const elements: React.ReactNode[] = [];

    const pctX = (v: number) => (v / 13.33) * 100;
    const pctY = (v: number) => (v / 7.5) * 100;
    const pctW = (v: number) => (v / 13.33) * 100;
    const pctH = (v: number) => (v / 7.5) * 100;

    if (isTitleSlide || isEndSlide) {
      // Cover / End
      elements.push(renderDecorDiv(0, 0, 13.33, t.coverNavyHeight, t.primary, 't-top'));
      elements.push(renderDecorDiv(0, t.coverNavyHeight, 13.33, 0.12, t.accent, 't-stripe'));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom'));
      elements.push(renderDecorDiv(10.5, 0.6, 2.2, 2.2, t.primaryLight, 't-circle'));
    } else if (layout === 'card_3col') {
      // 3-Column Cards
      elements.push(renderDecorDiv(0, 0, 13.33, 0.9, t.primary, 't-header'));
      elements.push(renderDecorDiv(0, 0.9, 13.33, 0.06, t.accent, 't-accent'));
      elements.push(renderDecorDiv(0.5, 1.6, 3.8, 5.0, t.offWhite, 't-card1bg'));
      elements.push(renderDecorDiv(4.75, 1.6, 3.8, 5.0, t.offWhite, 't-card2bg'));
      elements.push(renderDecorDiv(9.0, 1.6, 3.8, 5.0, t.offWhite, 't-card3bg'));
      elements.push(renderDecorDiv(0.5, 1.6, 3.8, 0.12, t.primary, 't-card1top'));
      elements.push(renderDecorDiv(4.75, 1.6, 3.8, 0.12, t.primary, 't-card2top'));
      elements.push(renderDecorDiv(9.0, 1.6, 3.8, 0.12, t.primary, 't-card3top'));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom'));
    } else if (layout === 'comparison') {
      // Comparison: Left-right
      elements.push(renderDecorDiv(0, 0, 13.33, 0.9, t.primary, 't-header'));
      elements.push(renderDecorDiv(0, 0.9, 13.33, 0.06, t.accent, 't-accent'));
      elements.push(renderDecorDiv(0.5, 1.5, 5.9, 5.2, t.offWhite, 't-leftblock'));
      elements.push(renderDecorDiv(0.5, 1.5, 5.9, 0.1, t.primary, 't-lefttop'));
      elements.push(renderDecorDiv(6.9, 1.5, 5.9, 5.2, t.offWhite, 't-rightblock'));
      elements.push(renderDecorDiv(6.9, 1.5, 5.9, 0.1, t.accent, 't-righttop'));
      elements.push(renderDecorDiv(6.55, 1.5, 0.2, 5.2, t.primaryLight, 't-divider'));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom'));
    } else if (layout === 'timeline') {
      // Timeline: Vertical left bar + horizontal flow + dots
      elements.push(renderDecorDiv(0, 0, 0.5, 7.5, t.primary, 't-leftbar'));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom'));
      elements.push(renderDecorDiv(1.2, 3.5, 11.5, 0.12, t.primaryLight, 't-line'));
      elements.push(renderDecorDiv(2.5, 3.2, 0.7, 0.7, t.primary, 't-dot1'));
      elements.push(renderDecorDiv(5.5, 3.2, 0.7, 0.7, t.primary, 't-dot2'));
      elements.push(renderDecorDiv(8.5, 3.2, 0.7, 0.7, t.primary, 't-dot3'));
      elements.push(renderDecorDiv(11.5, 3.2, 0.7, 0.7, t.accent, 't-dot4'));
    } else if (layout === 'data_highlight') {
      // Data Highlight: Narrow header + bottom band + left accent
      elements.push(renderDecorDiv(0, 0, 13.33, 0.7, t.primary, 't-header'));
      elements.push(renderDecorDiv(0, 0.7, 13.33, 0.06, t.accent, 't-accent'));
      elements.push(renderDecorDiv(0, 6.0, 13.33, 1.5, t.offWhite, 't-bottomband'));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom'));
      elements.push(renderDecorDiv(0.5, 1.2, 0.08, 4.5, t.primary, 't-leftbar'));
    } else if (layout === 'quote_slide') {
      // Quote: Wide left bar + soft bg
      elements.push(renderDecorDiv(0, 0, 13.33, 7.5, t.offWhite, 't-bg'));
      elements.push(renderDecorDiv(0, 0, 0.8, 7.5, t.primary, 't-leftbar'));
      elements.push(renderDecorDiv(0.8, 2.8, 0.12, 1.8, t.accent, 't-accent'));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom'));
    } else {
      // Default content_slide: Left wide color band + right content area
      elements.push(renderDecorDiv(0, 0, 4.5, 7.15, t.primary, 't-leftband'));
      elements.push(renderDecorDiv(4.5, 0, 0.08, 7.15, t.accent, 't-divider'));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom'));
    }
    return elements;
  };


  return (
    <div className="flex flex-col h-screen w-full overflow-hidden bg-gray-100 text-slate-800 font-sans">
      {/* ===== Header Bar ===== */}
      <header className="h-12 bg-white border-b border-gray-200 flex items-center justify-between px-4 shrink-0 z-40 shadow-sm">
        <div className="flex items-center gap-3">
          <button
            onClick={() => router.push('/')}
            className="flex items-center gap-1.5 px-2.5 py-1 text-gray-500 hover:text-gray-900 hover:bg-gray-100 rounded-md transition text-sm"
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <polyline points="15 18 9 12 15 6"></polyline>
            </svg>
            返回
          </button>
          <div className="h-5 w-px bg-gray-200"></div>
          <div className="flex items-center gap-2">
            <div className="bg-indigo-600 p-1 rounded-md">
              <Icons.FilePresentation className="text-white w-4 h-4" />
            </div>
            <h1 className="text-base font-bold text-slate-800">
              AI PPT <span className="text-gray-400 font-normal text-xs">@小傅哥</span>
            </h1>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <div className="flex items-center gap-1.5 px-2 py-0.5 bg-gray-50 rounded-full border border-gray-200">
            <div className="w-1.5 h-1.5 rounded-full bg-green-500"></div>
            <span className="text-xs text-gray-500">{currentUser || 'Guest'}</span>
          </div>

          <button
            onClick={handleDownloadPptx}
            disabled={!pptData}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition-all ${
              pptData
                ? 'bg-indigo-600 text-white hover:bg-indigo-700 shadow-sm'
                : 'bg-gray-100 text-gray-400 cursor-not-allowed'
            }`}
          >
            <Icons.Download className="w-4 h-4" />
            下载 PPTX
          </button>

          {pptData && (
            <button
              onClick={() => setIsFullscreen(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-gray-100 text-gray-600 hover:bg-gray-200 rounded-md text-sm font-medium transition-all"
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
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition-all ${
              isStylePanelOpen ? 'bg-indigo-100 text-indigo-700' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
            }`}
            title="风格设置"
          >
            🎛️
          </button>

          <button onClick={handleLogout} className="p-1.5 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded-md transition-colors" title="退出登录">
            <Icons.Logout className="w-4 h-4" />
          </button>

          {!isChatOpen && (
            <button onClick={() => setIsChatOpen(true)} className="p-1.5 text-indigo-600 bg-indigo-50 hover:bg-indigo-100 rounded-md transition-colors" title="打开助手">
              <Icons.Chat className="w-4 h-4" />
            </button>
          )}
        </div>
      </header>

      {/* ===== Main 3-Column Layout ===== */}
      <div className="flex flex-1 w-full overflow-hidden">

        {/* ===== Left: Slide Thumbnails / Template Library ===== */}
        <aside className="w-[180px] bg-white border-r border-gray-200 flex flex-col shrink-0 z-30">
          {/* Tab Switcher: Templates / Slides */}
          <div className="px-2 pt-2 pb-1 border-b border-gray-100 flex gap-1">
            <button
              onClick={() => setLeftTab('templates')}
              className={`flex-1 text-[10px] font-medium py-1 rounded transition ${leftTab === 'templates' ? 'bg-indigo-50 text-indigo-600' : 'text-gray-400 hover:bg-gray-50'}`}
            >
              📋 模板
            </button>
            <button
              onClick={() => setLeftTab('slides')}
              className={`flex-1 text-[10px] font-medium py-1 rounded transition ${leftTab === 'slides' ? 'bg-indigo-50 text-indigo-600' : 'text-gray-400 hover:bg-gray-50'}`}
            >
              🖼️ 页面
            </button>
          </div>

          {/* Theme Selector (always visible) */}
          <div className="px-3 pt-2 pb-2 border-b border-gray-100">
            <div className="text-[10px] text-gray-400 font-medium mb-1.5">🎨 主题</div>
            <div className="flex flex-wrap gap-1.5">
              {THEMES.map(t => (
                <button
                  key={t.id}
                  onClick={() => setSelectedThemeId(t.id)}
                  className={`w-6 h-6 rounded-full border-2 transition-all flex items-center justify-center ${
                    selectedThemeId === t.id
                      ? 'border-indigo-500 scale-110 shadow-md'
                      : 'border-gray-200 hover:border-gray-400 hover:scale-105'
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
          <div className="flex-1 overflow-y-auto p-2 space-y-1 scrollbar-thin scrollbar-thumb-gray-200">
            {leftTab === 'templates' ? (
              /* Template Library */
              <div className="space-y-0.5">
                {PPT_TEMPLATES.map((cat) => (
                  <div key={cat.category}>
                    <button
                      onClick={() => setExpandedCategory(expandedCategory === cat.category ? null : cat.category)}
                      className="w-full flex items-center gap-1.5 px-1.5 py-1.5 rounded-md hover:bg-gray-50 transition text-left"
                    >
                      <span className="text-xs">{cat.icon}</span>
                      <span className="text-[11px] font-medium text-gray-600 flex-1">{cat.label}</span>
                      <span className="text-[9px] text-gray-300">{cat.templates.length}</span>
                      <svg
                        className={`w-3 h-3 text-gray-400 transition-transform ${expandedCategory === cat.category ? 'rotate-90' : ''}`}
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
                            <div className="text-[11px] font-medium text-gray-700 group-hover:text-indigo-700">{tpl.name}</div>
                            <div className="text-[9px] text-gray-400 mt-0.5 leading-tight">{tpl.desc}</div>
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
                      : 'ring-1 ring-gray-200 hover:ring-gray-300 hover:shadow-sm'
                  }`}
                >
                  <div className="text-[10px] text-gray-400 mb-1 pl-0.5 font-medium">{idx + 1}</div>
                  <div className="w-full aspect-[16/9] bg-white rounded overflow-hidden relative shadow-sm">
                    {renderMiniSlide(slide, idx)}
                  </div>
                </div>
              ))
            ) : (
              <div className="text-center py-8">
                <Icons.FilePresentation className="w-8 h-8 mx-auto text-gray-200 mb-2" />
                <p className="text-[11px] text-gray-400">暂无幻灯片</p>
                <p className="text-[9px] text-gray-300 mt-1">在右侧输入需求或选择模板</p>
              </div>
            )
          }
          </div>

          {/* Session list */}
          <div className="border-t border-gray-100">
            <div className="h-9 px-3 flex items-center justify-between">
              <span className="text-[10px] font-semibold text-gray-400 uppercase tracking-wider">历史</span>
              <button onClick={handleNewChat} className="p-1 text-gray-400 hover:text-indigo-600 hover:bg-indigo-50 rounded transition" title="新建">
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
                      : 'hover:bg-gray-50 text-gray-500 hover:text-gray-700'
                  }`}
                >
                  <span className="text-[11px] truncate flex-1">{session.title}</span>
                  <button
                    onClick={(e) => handleDeleteSession(e, session.id)}
                    className="p-0.5 rounded opacity-0 group-hover:opacity-100 hover:bg-red-100 text-gray-300 hover:text-red-500 transition"
                  >
                    <Icons.Trash className="w-3 h-3" />
                  </button>
                </div>
              ))}
            </div>
          </div>
        </aside>

        {/* ===== Center: Main Slide Preview ===== */}
        <main className="flex-1 flex flex-col bg-gray-50 h-full overflow-hidden">
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
                <div className="w-24 h-24 mx-auto mb-6 bg-gray-100 rounded-2xl flex items-center justify-center">
                  <Icons.FilePresentation className="w-12 h-12 text-gray-300" />
                </div>
                <p className="text-base text-gray-500 mb-2 font-medium">AI 生成 PPT 后将在此预览</p>
                <p className="text-sm text-gray-400">在右侧对话区描述你的需求</p>
              </div>
            ) : (
              <div
                className="bg-white rounded-lg shadow-xl border border-gray-200 overflow-hidden relative"
                style={{
                  width: '100%',
                  maxWidth: '960px',
                  aspectRatio: '16/9',
                  maxHeight: 'calc(100vh - 160px)',
                }}
              >
                <div className="w-full h-full relative">
                  {renderThemeDecor(pptData.slides[currentSlideIndex], currentSlideIndex)}
                  {renderSlideContent(pptData.slides[currentSlideIndex], currentSlideIndex)}
                </div>
              </div>
            )}
          </div>

          {/* Slide navigation bar */}
          {pptData && pptData.slides.length > 0 && (
            <div className="h-11 px-4 bg-white border-t border-gray-200 flex items-center justify-center gap-4 shrink-0">
              <button
                onClick={() => setCurrentSlideIndex(Math.max(0, currentSlideIndex - 1))}
                disabled={currentSlideIndex === 0}
                className="p-1.5 text-gray-400 hover:text-gray-700 hover:bg-gray-100 rounded-md disabled:opacity-30 disabled:cursor-not-allowed transition"
              >
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="15 18 9 12 15 6"></polyline>
                </svg>
              </button>
              <span className="text-sm text-gray-500 font-mono tabular-nums min-w-[60px] text-center">
                {currentSlideIndex + 1} / {pptData.slides.length}
              </span>
              <button
                onClick={() => setCurrentSlideIndex(Math.min(pptData.slides.length - 1, currentSlideIndex + 1))}
                disabled={currentSlideIndex === pptData.slides.length - 1}
                className="p-1.5 text-gray-400 hover:text-gray-700 hover:bg-gray-100 rounded-md disabled:opacity-30 disabled:cursor-not-allowed transition"
              >
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="9 18 15 12 9 6"></polyline>
                </svg>
              </button>
              <div className="h-4 w-px bg-gray-200 mx-2"></div>
              <span className="text-xs text-gray-400">← → 翻页</span>
            </div>
          )}
        </main>

        {/* ===== Right: Style Panel + Chat Panel ===== */}
        <div className="flex shrink-0 z-20">
          {/* Style Control Panel */}
          <div
            className={`border-l border-gray-200 bg-white flex flex-col transition-all duration-300 ease-[cubic-bezier(0.25,0.1,0.25,1)] ${
              isStylePanelOpen ? 'w-[220px] translate-x-0' : 'w-0 translate-x-full opacity-0 overflow-hidden'
            }`}
          >
            <div className="h-12 px-4 border-b border-gray-100 flex items-center justify-between shrink-0">
              <span className="text-sm font-bold text-slate-800">🎛️ 风格设置</span>
              <button onClick={() => setIsStylePanelOpen(false)} className="p-1.5 text-gray-400 hover:text-gray-600 hover:bg-gray-100 rounded-md transition shrink-0">
                <Icons.Close className="w-4 h-4" />
              </button>
            </div>
            <div className="flex-1 overflow-y-auto p-4 space-y-5 text-sm">
              {/* Style */}
              <div>
                <div className="text-xs text-gray-400 font-medium mb-2">风格</div>
                <div className="grid grid-cols-2 gap-1.5">
                  {STYLE_OPTIONS.map(opt => (
                    <button
                      key={opt.id}
                      onClick={() => setSelectedStyle(opt.id)}
                      className={`px-2 py-1.5 rounded-md text-xs transition ${
                        selectedStyle === opt.id ? 'bg-indigo-100 text-indigo-700 border border-indigo-300' : 'bg-gray-50 text-gray-600 border border-transparent hover:bg-gray-100'
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>
              {/* Structure */}
              <div>
                <div className="text-xs text-gray-400 font-medium mb-2">结构</div>
                <div className="grid grid-cols-2 gap-1.5">
                  {STRUCTURE_OPTIONS.map(opt => (
                    <button
                      key={opt.id}
                      onClick={() => setSelectedStructure(opt.id)}
                      className={`px-2 py-1.5 rounded-md text-xs transition ${
                        selectedStructure === opt.id ? 'bg-indigo-100 text-indigo-700 border border-indigo-300' : 'bg-gray-50 text-gray-600 border border-transparent hover:bg-gray-100'
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>
              {/* Tone */}
              <div>
                <div className="text-xs text-gray-400 font-medium mb-2">色调</div>
                <div className="grid grid-cols-2 gap-1.5">
                  {TONE_OPTIONS.map(opt => (
                    <button
                      key={opt.id}
                      onClick={() => setSelectedTone(opt.id)}
                      className={`px-2 py-1.5 rounded-md text-xs transition ${
                        selectedTone === opt.id ? 'bg-indigo-100 text-indigo-700 border border-indigo-300' : 'bg-gray-50 text-gray-600 border border-transparent hover:bg-gray-100'
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>
              {/* Scene */}
              <div>
                <div className="text-xs text-gray-400 font-medium mb-2">场景</div>
                <div className="grid grid-cols-2 gap-1.5">
                  {SCENE_OPTIONS.map(opt => (
                    <button
                      key={opt.id}
                      onClick={() => setSelectedScene(opt.id)}
                      className={`px-2 py-1.5 rounded-md text-xs transition ${
                        selectedScene === opt.id ? 'bg-indigo-100 text-indigo-700 border border-indigo-300' : 'bg-gray-50 text-gray-600 border border-transparent hover:bg-gray-100'
                      }`}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              </div>
              {/* Layout Preference */}
              <div>
                <div className="text-xs text-gray-400 font-medium mb-2">布局偏好</div>
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
                        selectedLayouts.includes(opt.id) ? 'bg-indigo-100 text-indigo-700 border border-indigo-300' : 'bg-gray-50 text-gray-600 border border-transparent hover:bg-gray-100'
                      }`}
                    >
                      <span>{opt.label}</span>
                      <span className="text-[10px] text-gray-400">{opt.desc}</span>
                    </button>
                  ))}
                </div>
              </div>
            </div>
          </div>

          {/* Chat Panel */}
          <div
            className={`border-l border-gray-200 bg-white flex flex-col transition-all duration-300 ease-[cubic-bezier(0.25,0.1,0.25,1)] ${
              isChatOpen ? 'w-[360px] translate-x-0' : 'w-0 translate-x-full opacity-0 overflow-hidden'
            } shadow-lg`}
          >
            {/* Chat Header */}
            <div className="h-12 px-4 border-b border-gray-100 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2 flex-1 min-w-0">
                <div className="flex items-center justify-center w-8 h-8 rounded-full bg-gradient-to-br from-indigo-500 to-purple-600 text-white shadow-sm shrink-0">
                  <Icons.Sparkles className="w-4 h-4" />
                </div>
                <div className="flex-1 min-w-0">
                  <select
                    value={selectedAgentId}
                    onChange={handleAgentChange}
                    className="w-full bg-transparent text-sm font-bold text-slate-800 focus:outline-none cursor-pointer truncate appearance-none pr-4"
                  >
                    {agents.length === 0 && <option value="">Loading...</option>}
                    {agents.map((agent) => (
                      <option key={agent.agentId} value={agent.agentId}>{agent.agentName}</option>
                    ))}
                  </select>
                  <div className="flex items-center gap-1 mt-0.5">
                    <span className="w-1.5 h-1.5 rounded-full bg-green-500 animate-pulse"></span>
                    <span className="text-[10px] text-gray-400">PPT 助手在线</span>
                  </div>
                </div>
              </div>
              <button onClick={() => setIsChatOpen(false)} className="p-1.5 text-gray-400 hover:text-gray-600 hover:bg-gray-100 rounded-md transition shrink-0">
                <Icons.Close className="w-5 h-5" />
              </button>
            </div>

            {/* Messages */}
            <div className="flex-1 overflow-y-auto p-4 space-y-4 bg-gray-50/50 scrollbar-thin scrollbar-thumb-gray-200">
              {messages.map((msg) => (
                <div key={msg.id} className={`flex gap-2.5 ${msg.role === 'user' ? 'flex-row-reverse' : 'flex-row'}`}>
                  <div className={`shrink-0 w-7 h-7 rounded-full flex items-center justify-center shadow-sm ${
                    msg.role === 'user' ? 'bg-indigo-100 text-indigo-600' : 'bg-white text-indigo-500 border border-gray-100'
                  }`}>
                    {msg.role === 'user' ? <Icons.User className="w-4 h-4" /> : <Icons.Bot className="w-4 h-4" />}
                  </div>
                  <div className="flex flex-col max-w-[85%]">
                    <div className={`p-3 text-sm leading-relaxed whitespace-pre-wrap rounded-xl shadow-sm ${
                      msg.role === 'user'
                        ? 'bg-indigo-600 text-white rounded-tr-sm'
                        : 'bg-white border border-gray-200 text-gray-700 rounded-tl-sm'
                    }`}>
                      {msg.content}
                    </div>
                  </div>
                </div>
              ))}
              <div ref={messagesEndRef} />
            </div>

            {/* Input */}
            <div className="p-3 bg-white border-t border-gray-100 shrink-0">
              {messages.length <= 1 && (
                <div className="flex flex-wrap gap-2 mb-2.5">
                  {quickActions.map((action, idx) => (
                    <button
                      key={idx}
                      onClick={() => setInputValue(action.text)}
                      className="text-xs px-3 py-1.5 bg-indigo-50 text-indigo-600 rounded-full hover:bg-indigo-100 transition border border-indigo-100 font-medium"
                    >
                      {action.label}
                    </button>
                  ))}
                </div>
              )}
              <div className="flex items-end gap-2 bg-gray-50 p-1.5 rounded-xl border border-gray-200 focus-within:border-indigo-300 focus-within:ring-2 focus-within:ring-indigo-50 transition-all">
                <textarea
                  value={inputValue}
                  onChange={(e) => setInputValue(e.target.value)}
                  onKeyDown={handleKeyDown}
                  placeholder={isSending ? 'AI 正在生成...' : '描述你想制作的 PPT...'}
                  disabled={isSending}
                  className="flex-1 px-3 py-2 bg-transparent border-none focus:ring-0 text-sm text-gray-800 placeholder:text-gray-400 resize-none max-h-32 min-h-[44px] scrollbar-thin scrollbar-thumb-gray-200"
                  rows={1}
                  style={{ height: 'auto', minHeight: '44px' }}
                />
                <button
                  onClick={handleSendMessage}
                  disabled={!inputValue.trim() || isSending}
                  className={`p-2 rounded-lg transition-all flex items-center justify-center ${
                    inputValue.trim() && !isSending
                      ? 'bg-indigo-600 text-white shadow-sm hover:bg-indigo-700 active:scale-95'
                      : 'bg-gray-200 text-gray-400 cursor-not-allowed'
                  }`}
                >
                  {isSending ? <Icons.Loader className="w-4 h-4" /> : <Icons.Send className="w-4 h-4" />}
                </button>
              </div>
              <p className="text-center mt-2 text-[10px] text-gray-400">
                {isSending ? 'AI 正在生成 PPT...' : '⌘/Ctrl + Enter 发送'}
              </p>
            </div>
          </div>
        </div>
      </div>

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
            <div className="w-full h-full relative">
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
