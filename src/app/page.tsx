'use client';

import Link from 'next/link';
import { useRouter } from 'next/navigation';
import { useEffect, useState } from 'react';
import {
  ArrowRight,
  FileText,
  LogOut,
  Network,
  PenLine,
  Presentation,
  Sparkles,
  Wand2,
} from 'lucide-react';
import { clearUserInfo, getUserInfo } from '@/utils/cookie';

const products = [
  {
    href: '/ppt',
    agentId: '300001',
    title: 'Presentation Atelier',
    label: 'AI PPT',
    desc: '把主题、受众和结构变成可预览、可导出的演示文稿。',
    cta: '进入 PPT 工坊',
    icon: Presentation,
    accent: 'var(--xerina-rose)',
    meta: ['slides preview', 'pptx export', 'metadata notes'],
  },
  {
    href: '/drawio',
    agentId: '300000',
    title: 'Diagram Studio',
    label: 'Draw.io',
    desc: '用对话生成流程图、架构图和 UML，并直接渲染到 draw.io 画布。',
    cta: '进入图形工作室',
    icon: Network,
    accent: 'var(--xerina-cyan)',
    meta: ['xml rendering', 'canvas context', 'step telemetry'],
  },
  {
    href: '/prompt',
    agentId: '300002',
    title: 'Prompt Forge',
    label: 'Prompt',
    desc: '从零生成、整体改写或局部精修 Prompt，支持流式生成与 Markdown 导出。',
    cta: '进入提示词工坊',
    icon: FileText,
    accent: 'var(--xerina-gold)',
    meta: ['md export', 'stream craft', 'partial edit'],
  },
];

/**
 * description: Renders the product landing hub for PPT and draw.io AI applications.
 * params:
 * - input: No component props.
 * - output: Authenticated landing page with fixed navigation to each Agent workspace.
 */
export default function Home() {
  const router = useRouter();
  const [currentUser, setCurrentUser] = useState('');

  useEffect(() => {
    const userInfo = getUserInfo();

    if (!userInfo || !userInfo.user) {
      router.push('/login');
      return;
    }

    // Cookie access is client-only, so the user label is hydrated after mount.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setCurrentUser(userInfo.user);
  }, [router]);

  const handleLogout = () => {
    clearUserInfo();
    router.push('/login');
  };

  return (
    <main className="xerina-home">
      <header className="xerina-nav">
        <div className="flex items-center gap-3">
          <div className="xerina-brand-mark">
            <Wand2 className="h-5 w-5" />
          </div>
          <div>
            <p className="xerina-kicker">AI Application House</p>
            <h1 className="xerina-wordmark">xerina</h1>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <div className="xerina-user">
            <Sparkles className="h-4 w-4" />
            <span>{currentUser || 'Guest'}</span>
          </div>
          <button className="xerina-icon-button" onClick={handleLogout} title="Logout" type="button">
            <LogOut className="h-4 w-4" />
          </button>
        </div>
      </header>

      <section className="xerina-hero">
        <div className="xerina-hero-copy">
          <p className="xerina-overline">two agents, two surfaces</p>
          <h2>
            让想法先有
            <span>形状</span>
            ，再交给 AI 落笔。
          </h2>
          <p>
            xerina 将演示文稿和 draw.io 图形拆成两个独立工作台：PPT 负责叙事和成果导出，
            Draw.io 负责结构、流程和桌面图形渲染。
          </p>
        </div>

        <div className="xerina-orbit" aria-hidden="true">
          <div className="xerina-orbit-line" />
          <PenLine className="xerina-orbit-pen" />
          <span>agent routed</span>
        </div>
      </section>

      <section className="xerina-product-grid" aria-label="AI application entries">
        {products.map((product, index) => {
          const Icon = product.icon;

          return (
            <Link className={`xerina-product xerina-product-${index + 1}`} href={product.href} key={product.href}>
              <div className="xerina-product-index">0{index + 1}</div>
              <div className="xerina-product-head">
                <div className="xerina-product-icon" style={{ color: product.accent }}>
                  <Icon className="h-7 w-7" />
                </div>
                <div>
                  <p>{product.label}</p>
                  <h3>{product.title}</h3>
                </div>
              </div>
              <p className="xerina-product-desc">{product.desc}</p>
              <div className="xerina-product-meta">
                {product.meta.map((item) => (
                  <span key={item}>{item}</span>
                ))}
              </div>
              <div className="xerina-product-foot">
                <span>Agent {product.agentId}</span>
                <span className="xerina-product-cta">
                  {product.cta}
                  <ArrowRight className="h-4 w-4" />
                </span>
              </div>
            </Link>
          );
        })}
      </section>
    </main>
  );
}


