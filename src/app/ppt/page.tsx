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
const generatePptx = (data: PptData, theme: PptTheme) => {
  const pres = new pptxgen();
  pres.title = data.title;
  pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5

  data.slides.forEach((slideData, slideIdx) => {
    const slide = pres.addSlide();
    const layout = slideData.layout || '';
    const isTitleSlide = layout === 'title_slide' || slideIdx === 0;
    const isEndSlide = layout === 'end_slide' || slideIdx === data.slides.length - 1;

    // === STEP 1: Professional theme decoration (always rendered) ===
    if (isTitleSlide || isEndSlide) {
      // ---- Cover / End Slide: Top-half navy + bottom white ----
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      slide.addShape((pres as any).shapes.RECTANGLE, {
        x: 0, y: 0, w: 13.33, h: theme.coverNavyHeight,
        fill: { color: theme.primary },
        line: { width: 0 },
      });
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      slide.addShape((pres as any).shapes.RECTANGLE, {
        x: 0, y: theme.coverNavyHeight, w: 13.33, h: 0.12,
        fill: { color: theme.accent },
        line: { width: 0 },
      });
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      slide.addShape((pres as any).shapes.RECTANGLE, {
        x: 0, y: 7.15, w: 13.33, h: 0.35,
        fill: { color: theme.primary },
        line: { width: 0 },
      });
    } else {
      // ---- Content Slide: Full-width navy header + accent stripe ----
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      slide.addShape((pres as any).shapes.RECTANGLE, {
        x: 0, y: 0, w: 13.33, h: theme.contentHeaderHeight,
        fill: { color: theme.primary },
        line: { width: 0 },
      });
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      slide.addShape((pres as any).shapes.RECTANGLE, {
        x: 0, y: theme.contentHeaderHeight, w: 13.33, h: 0.08,
        fill: { color: theme.accent },
        line: { width: 0 },
      });
      // Bottom thin navy bar
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      slide.addShape((pres as any).shapes.RECTANGLE, {
        x: 0, y: 7.15, w: 13.33, h: 0.35,
        fill: { color: theme.primary },
        line: { width: 0 },
      });
      // Subtle left vertical accent line in content area
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      slide.addShape((pres as any).shapes.RECTANGLE, {
        x: 0.7, y: 1.8, w: 0.06, h: 4.8,
        fill: { color: theme.primaryLight },
        line: { width: 0 },
      });
    }

    // === STEP 2: Render AI content elements (skip all shapes) ===
    slideData.elements.forEach((el) => {
      if (el.kind === 'shape') return;

      switch (el.kind) {
        case 'text': {
          let textY = el.y || 0;
          let textColor = el.color || theme.bodyColor;
          let textFontSize = el.fontSize || 18;

          // For content slides, push title text into the header bar
          if (!isTitleSlide && !isEndSlide && textFontSize >= 24) {
            textY = 0.2;
            textColor = theme.white;
          }
          // For cover/end slides, ensure title is white (in navy area)
          if ((isTitleSlide || isEndSlide) && textFontSize >= 24) {
            textColor = theme.white;
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
    if (selectedStyle !== 'professional') styleHints.push(`风格: ${STYLE_OPTIONS.find(o => o.id === selectedStyle)?.label || selectedStyle}`);
    if (selectedStructure !== 'auto') styleHints.push(`结构: ${STRUCTURE_OPTIONS.find(o => o.id === selectedStructure)?.label || selectedStructure}`);
    if (selectedTone !== 'neutral') styleHints.push(`色调: ${TONE_OPTIONS.find(o => o.id === selectedTone)?.label || selectedTone}`);
    if (selectedScene !== 'general') styleHints.push(`场景: ${SCENE_OPTIONS.find(o => o.id === selectedScene)?.label || selectedScene}`);
    if (!selectedLayouts.includes('auto') && selectedLayouts.length > 0) {
      const layoutLabels = selectedLayouts.map(id => LAYOUT_OPTIONS.find(o => o.id === id)?.label || id).join('、');
      styleHints.push(`布局偏好: ${layoutLabels}`);
    }
    const enrichedContent = styleHints.length > 0
      ? `${content}\n\n[用户偏好设置: ${styleHints.join(' | ')}]`
      : content;

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
      console.log('[PPT] Response type:', type, 'content type:', typeof resContent, 'content preview:', typeof resContent === 'string' ? resContent.slice(0, 300) : JSON.stringify(resContent).slice(0, 300));
      const detectedPpt = tryParsePpt(resContent);
      console.log('[PPT] tryParsePpt result:', detectedPpt ? 'DETECTED' : 'null');
      if (detectedPpt) {
        console.log('[PPT] Detected PPT data (type was:', type, ')');
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

  // Render mini slide for thumbnail panel (simplified, very small)
  const renderMiniSlide = (slideData: PptSlide, slideIdx: number) => {
    const layout = slideData.layout || '';
    const isTitleSlide = layout === 'title_slide' || slideIdx === 0;
    const isEndSlide = layout === 'end_slide' || slideIdx === (pptData?.slides.length ?? 1) - 1;
    const t = activeTheme;
    const coverNavyPct = (t.coverNavyHeight / 7.5) * 100;

    return (
      <div className="w-full h-full relative bg-white" style={{ fontSize: '2px' }}>
        {/* Theme decorations (mini) */}
        {isTitleSlide || isEndSlide ? (
          <>
            <div className="absolute inset-x-0 top-0" style={{ height: `${coverNavyPct}%`, backgroundColor: `#${t.primary}` }} />
            <div className="absolute inset-x-0" style={{ top: `${coverNavyPct}%`, height: '2%', backgroundColor: `#${t.accent}` }} />
            <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', backgroundColor: `#${t.primary}` }} />
          </>
        ) : (
          <>
            <div className="absolute inset-x-0 top-0" style={{ height: `${(t.contentHeaderHeight / 7.5) * 100}%`, backgroundColor: `#${t.primary}` }} />
            <div className="absolute inset-x-0" style={{ top: `${(t.contentHeaderHeight / 7.5) * 100}%`, height: '1.5%', backgroundColor: `#${t.accent}` }} />
            <div className="absolute inset-x-0 bottom-0" style={{ height: '5%', backgroundColor: `#${t.primary}` }} />
          </>
        )}
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
    const layout = slideData.layout || '';
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
        let textColor = el.color ? `#${el.color}` : `#${activeTheme.bodyColor}`;
        let fontSize = Math.max(8, (el.fontSize || 18) * 0.7);

        if (!isTitleSlide && !isEndSlide && (el.fontSize || 0) >= 24) {
          yPct = (0.2 / 7.5) * 100;
          textColor = `#${activeTheme.titleColor}`;
        }
        if ((isTitleSlide || isEndSlide) && (el.fontSize || 0) >= 24) {
          textColor = `#${activeTheme.titleColor}`;
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

  // Render theme decorations for preview (matching generatePptx exactly)
  const renderThemeDecor = (slideData: PptSlide, slideIdx: number) => {
    const layout = slideData.layout || '';
    const isTitleSlide = layout === 'title_slide' || slideIdx === 0;
    const isEndSlide = layout === 'end_slide' || slideIdx === (pptData?.slides.length ?? 1) - 1;
    const t = activeTheme;
    const elements: React.ReactNode[] = [];

    if (isTitleSlide || isEndSlide) {
      elements.push(renderDecorDiv(0, 0, 13.33, t.coverNavyHeight, t.primary, 't-top-navy'));
      elements.push(renderDecorDiv(0, t.coverNavyHeight, 13.33, 0.12, t.accent, 't-accent-stripe'));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom'));
    } else {
      elements.push(renderDecorDiv(0, 0, 13.33, t.contentHeaderHeight, t.primary, 't-header'));
      elements.push(renderDecorDiv(0, t.contentHeaderHeight, 13.33, 0.08, t.accent, 't-accent'));
      elements.push(renderDecorDiv(0, 7.15, 13.33, 0.35, t.primary, 't-bottom'));
      elements.push(renderDecorDiv(0.7, 1.8, 0.06, 4.8, t.primaryLight, 't-left-line'));
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

        {/* ===== Left: Slide Thumbnails ===== */}
        <aside className="w-[180px] bg-white border-r border-gray-200 flex flex-col shrink-0 z-30">
          {/* Theme Selector */}
          <div className="px-3 pt-3 pb-2 border-b border-gray-100">
            <div className="text-[11px] text-gray-400 font-medium mb-2">🎨 主题风格</div>
            <div className="flex flex-wrap gap-1.5">
              {THEMES.map(t => (
                <button
                  key={t.id}
                  onClick={() => setSelectedThemeId(t.id)}
                  className={`w-7 h-7 rounded-full border-2 transition-all flex items-center justify-center ${
                    selectedThemeId === t.id
                      ? 'border-indigo-500 scale-110 shadow-md'
                      : 'border-gray-200 hover:border-gray-400 hover:scale-105'
                  }`}
                  title={t.name}
                >
                  <div
                    className="w-5 h-5 rounded-full overflow-hidden"
                    style={{
                      background: `linear-gradient(135deg, #${t.primary} 50%, #${t.accent} 50%)`,
                    }}
                  />
                </button>
              ))}
            </div>
            <div className="text-[10px] text-gray-500 mt-1.5 text-center">{activeTheme.name}</div>
          </div>
          {/* Thumbnails */}
          <div className="flex-1 overflow-y-auto p-3 space-y-3 scrollbar-thin scrollbar-thumb-gray-200">
            {pptData && pptData.slides.length > 0 ? (
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
              <div className="text-center py-10">
                <Icons.FilePresentation className="w-10 h-10 mx-auto text-gray-300 mb-2" />
                <p className="text-[11px] text-gray-400">暂无幻灯片</p>
                <p className="text-[10px] text-gray-300 mt-1">在右侧输入需求生成</p>
              </div>
            )}
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
