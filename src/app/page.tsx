'use client';

import { useEffect, useState } from 'react';
import { useRouter } from 'next/navigation';
import { getUserInfo, clearUserInfo } from '@/utils/cookie';

export default function Lobby() {
  const router = useRouter();
  const [currentUser, setCurrentUser] = useState('');

  useEffect(() => {
    const userInfo = getUserInfo();
    if (!userInfo || !userInfo.user) {
      router.push('/login');
      return;
    }
    setCurrentUser(userInfo.user);
  }, [router]);

  const handleLogout = () => {
    clearUserInfo();
    router.push('/login');
  };

  const workspaces = [
    {
      id: 'drawio',
      title: 'Draw.io 绘图',
      desc: 'AI + Draw.io，交互式对话完成流程图、架构图、UML 等绘制',
      icon: (
        <svg xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
          <rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect>
          <line x1="3" y1="9" x2="21" y2="9"></line>
          <line x1="9" y1="21" x2="9" y2="9"></line>
        </svg>
      ),
      gradient: 'from-emerald-500 to-teal-600',
      shadow: 'shadow-emerald-200',
      bgHover: 'hover:border-emerald-300',
      href: '/drawio',
      tags: ['流程图', '架构图', 'UML', '时序图'],
    },
    {
      id: 'ppt',
      title: 'PPT 生成',
      desc: 'AI + PptxGenJS，对话式生成专业 PowerPoint 演示文稿',
      icon: (
        <svg xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
          <rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect>
          <line x1="8" y1="21" x2="16" y2="21"></line>
          <line x1="12" y1="17" x2="12" y2="21"></line>
        </svg>
      ),
      gradient: 'from-indigo-500 to-purple-600',
      shadow: 'shadow-indigo-200',
      bgHover: 'hover:border-indigo-300',
      href: '/ppt',
      tags: ['汇报', '培训', '方案', '总结'],
      badge: 'NEW',
    },
  ];

  return (
    <div className="min-h-screen flex flex-col theme-bg-gradient">
      {/* Header */}
      <header className="h-16 px-8 flex items-center justify-between border-b border-[rgba(255,255,255,0.08)] shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-[12px] grid place-items-center bg-gradient-to-br from-[#62f6c7] to-[#5aa9ff] shadow-[0_8px_20px_rgba(0,0,0,0.3)] text-[rgba(7,10,18,0.92)] font-extrabold text-base">
            AI
          </div>
          <div>
            <h1 className="text-base font-bold text-[rgba(255,255,255,0.92)] tracking-tight">AI 智能体工作台</h1>
            <p className="text-[10px] text-[rgba(255,255,255,0.5)]">选择工具，开始创作</p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-3 py-1.5 bg-[rgba(255,255,255,0.06)] rounded-full border border-[rgba(255,255,255,0.1)]">
            <div className="w-2 h-2 rounded-full bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.4)]"></div>
            <span className="text-xs font-medium text-[rgba(255,255,255,0.7)]">{currentUser}</span>
          </div>
          <button
            onClick={handleLogout}
            className="px-3 py-1.5 text-xs text-[rgba(255,255,255,0.5)] hover:text-[#ff5a7a] hover:bg-[rgba(255,90,122,0.1)] rounded-lg transition-colors border border-transparent hover:border-[rgba(255,90,122,0.2)]"
          >
            退出登录
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 flex flex-col items-center justify-center px-8 py-12">
        <div className="max-w-4xl w-full">
          {/* Title */}
          <div className="text-center mb-12">
            <h2 className="text-3xl font-bold text-[rgba(255,255,255,0.92)] mb-3">选择工作模式</h2>
            <p className="text-[rgba(255,255,255,0.5)] text-sm max-w-md mx-auto">
              选择你想使用的 AI 绘图工具，进入对应的工作区域
            </p>
          </div>

          {/* Workspace Cards */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {workspaces.map((ws) => (
              <button
                key={ws.id}
                onClick={() => router.push(ws.href)}
                className={`
                  group relative text-left p-6 rounded-2xl border border-[rgba(255,255,255,0.1)] 
                  bg-gradient-to-br from-[rgba(255,255,255,0.06)] to-[rgba(255,255,255,0.02)]
                  hover:from-[rgba(255,255,255,0.1)] hover:to-[rgba(255,255,255,0.04)]
                  ${ws.bgHover} hover:shadow-xl ${ws.shadow}
                  transition-all duration-300 cursor-pointer
                `}
              >
                {/* Badge */}
                {ws.badge && (
                  <span className="absolute top-4 right-4 px-2 py-0.5 text-[10px] font-bold bg-gradient-to-r from-amber-400 to-orange-500 text-white rounded-full shadow-sm">
                    {ws.badge}
                  </span>
                )}

                {/* Icon */}
                <div className={`w-14 h-14 rounded-xl bg-gradient-to-br ${ws.gradient} flex items-center justify-center text-white shadow-lg ${ws.shadow} mb-5 group-hover:scale-110 transition-transform duration-300`}>
                  {ws.icon}
                </div>

                {/* Title & Desc */}
                <h3 className="text-xl font-bold text-[rgba(255,255,255,0.92)] mb-2 group-hover:text-white transition-colors">
                  {ws.title}
                </h3>
                <p className="text-sm text-[rgba(255,255,255,0.5)] leading-relaxed mb-4">
                  {ws.desc}
                </p>

                {/* Tags */}
                <div className="flex flex-wrap gap-2">
                  {ws.tags.map((tag) => (
                    <span
                      key={tag}
                      className="px-2 py-0.5 text-[11px] rounded-md bg-[rgba(255,255,255,0.06)] text-[rgba(255,255,255,0.5)] border border-[rgba(255,255,255,0.08)]"
                    >
                      {tag}
                    </span>
                  ))}
                </div>

                {/* Arrow */}
                <div className="absolute bottom-6 right-6 opacity-0 group-hover:opacity-100 transition-all duration-300 translate-x-0 group-hover:translate-x-1">
                  <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-[rgba(255,255,255,0.4)]">
                    <line x1="5" y1="12" x2="19" y2="12"></line>
                    <polyline points="12 5 19 12 12 19"></polyline>
                  </svg>
                </div>
              </button>
            ))}
          </div>

          {/* Footer hint */}
          <p className="text-center text-[11px] text-[rgba(255,255,255,0.3)] mt-10">
            AI Agent Scaffold · @小傅哥 · 更多工具持续接入中…
          </p>
        </div>
      </main>
    </div>
  );
}
