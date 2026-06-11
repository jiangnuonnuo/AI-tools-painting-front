'use client';

import { ArrowRight, Bot, CheckCircle2, Cookie, KeyRound, LogOut, Network, ShieldCheck, UserRound } from 'lucide-react';
import { useEffect, useState } from 'react';
import { useRouter } from 'next/navigation';
import { clearUserInfo, getUserInfo, setUserInfo } from '@/utils/cookie';

const features = [
  {
    title: 'XML 渲染链路',
    desc: 'AI 返回 drawio 类型后直接进入画布，不打断编辑流。',
  },
  {
    title: '画布上下文',
    desc: '需要时导出当前 XML，作为下一轮智能体输入。',
  },
  {
    title: '会话记录',
    desc: '本地保存聊天和画布状态，便于继续设计。',
  },
  {
    title: '人工协作',
    desc: 'AI 生成、人工调整、再携带上下文迭代。',
  },
];

/**
 * description: Provides demo authentication before entering the AI + draw.io workbench.
 * params:
 * - input: No component props.
 * - output: Renders the login form and writes demo auth cookie after validation.
 */
export default function Login() {
  const router = useRouter();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [msg, setMsg] = useState({ text: '', type: '' });
  const [currentUser, setCurrentUser] = useState('');

  const isLoggedIn = Boolean(currentUser);

  useEffect(() => {
    const userInfo = getUserInfo();

    if (!userInfo?.user) {
      return;
    }

    router.push('/');
  }, [router]);

  const handleLogin = (event: React.FormEvent) => {
    event.preventDefault();
    setMsg({ text: '', type: '' });

    if (!username || !password) {
      setMsg({ text: '请输入账号与密码。', type: 'error' });
      return;
    }

    if (username !== 'admin' || password !== 'admin') {
      setMsg({ text: '账号或密码错误（演示账号：admin / admin）。', type: 'error' });
      return;
    }

    setUserInfo(username);
    setCurrentUser(username);
    setMsg({ text: '登录成功，正在跳转...', type: 'info' });

    setTimeout(() => {
      router.push('/');
    }, 500);
  };

  const handleFillDemo = () => {
    setUsername('admin');
    setPassword('admin');
    setMsg({ text: '已填充演示账号。', type: 'info' });
  };

  const handleLogout = () => {
    clearUserInfo();
    setCurrentUser('');
    setMsg({ text: '已退出登录，cookie 已清除。', type: 'info' });
  };

  return (
    <main className="login-shell">
      <div className="login-grid">
        <section className="login-hero">
          <div className="relative z-10 flex h-full flex-col">
            <div className="flex items-center gap-4">
              <div className="brand-mark">
                <Network className="h-5 w-5" />
              </div>
              <div>
                <p className="panel-kicker">AI DRAWING WORKBENCH</p>
                <h1 className="text-[22px] font-semibold text-[var(--workbench-text)]">
                  AI + draw.io 集成工作台
                </h1>
              </div>
            </div>

            <div className="mt-10 max-w-[620px]">
              <p className="text-sm font-semibold uppercase tracking-[0.32em] text-[var(--workbench-accent)]">
                Agent assisted diagramming
              </p>
              <h2 className="mt-4 text-[42px] font-semibold leading-[1.08] tracking-[-0.02em] text-[var(--workbench-text)]">
                用 AI 生成结构图，
                <br />
                在 draw.io 中继续精修。
              </h2>
              <p className="mt-5 max-w-[54ch] text-sm leading-7 text-[var(--workbench-muted)]">
                登录后进入桌面工作台：左侧管理绘图会话，中间保留完整 draw.io 编辑器，右侧与智能体协同生成或修改 XML。
              </p>
            </div>

            <div className="mt-8 grid grid-cols-2 gap-3">
              {features.map((feature) => (
                <div className="feature-card" key={feature.title}>
                  <div className="mb-3 flex items-center gap-2 text-[var(--workbench-accent)]">
                    <CheckCircle2 className="h-4 w-4" />
                    <strong className="text-sm">{feature.title}</strong>
                  </div>
                  <p className="text-xs leading-6 text-[var(--workbench-muted-2)]">{feature.desc}</p>
                </div>
              ))}
            </div>

            <div className="login-visual">
              <span className="login-edge login-edge-1" />
              <span className="login-edge login-edge-2" />
              <span className="login-edge login-edge-3" />
              <span className="login-node login-node-1">Prompt</span>
              <span className="login-node login-node-2">Agent JSON</span>
              <span className="login-node login-node-3">draw.io XML</span>
              <span className="login-node login-node-4">Manual refine</span>
            </div>
          </div>
        </section>

        <section className="login-card">
          <div className="mb-8">
            <div className="assistant-mark mb-5">
              <Bot className="h-4 w-4" />
            </div>
            <p className="panel-kicker">Secure demo access</p>
            <h2 className="mt-2 text-[28px] font-semibold text-[var(--workbench-text)]">登录工作台</h2>
            <p className="mt-3 text-sm leading-6 text-[var(--workbench-muted)]">
              演示账号：admin / admin。登录成功后会在浏览器保存 cookie。
            </p>
          </div>

          {!isLoggedIn ? (
            <form className="space-y-4" onSubmit={handleLogin} autoComplete="on">
              <div>
                <label className="field-label flex items-center gap-2" htmlFor="username">
                  <UserRound className="h-4 w-4" />
                  账号
                </label>
                <input
                  autoComplete="username"
                  className="dark-input mt-2"
                  id="username"
                  onChange={(event) => setUsername(event.target.value)}
                  placeholder="请输入账号"
                  type="text"
                  value={username}
                />
              </div>

              <div>
                <label className="field-label flex items-center gap-2" htmlFor="password">
                  <KeyRound className="h-4 w-4" />
                  密码
                </label>
                <input
                  autoComplete="current-password"
                  className="dark-input mt-2"
                  id="password"
                  onChange={(event) => setPassword(event.target.value)}
                  placeholder="请输入密码"
                  type="password"
                  value={password}
                />
              </div>

              <div className="grid grid-cols-[1fr_auto] gap-3 pt-2">
                <button className="toolbar-button toolbar-button-primary min-h-[46px]" type="submit">
                  登录并保存 Cookie
                  <ArrowRight className="h-4 w-4" />
                </button>
                <button className="toolbar-button min-h-[46px]" onClick={handleFillDemo} type="button">
                  <Cookie className="h-4 w-4" />
                  填充演示账号
                </button>
              </div>
            </form>
          ) : (
            <div className="rounded-2xl border border-[var(--workbench-border)] bg-white/[0.055] p-4">
              <div className="flex items-center justify-between gap-4">
                <div className="min-w-0">
                  <div className="flex items-center gap-2 text-sm font-semibold text-[var(--workbench-text)]">
                    <ShieldCheck className="h-4 w-4 text-[var(--workbench-accent)]" />
                    已登录：{currentUser}
                  </div>
                  <p className="mt-1 text-xs text-[var(--workbench-muted-2)]">欢迎回来</p>
                </div>
                <button className="toolbar-button" onClick={handleLogout} type="button">
                  <LogOut className="h-4 w-4" />
                  退出
                </button>
              </div>
            </div>
          )}

          <div className={`mt-4 min-h-5 text-xs ${msg.type === 'error' ? 'text-[var(--workbench-danger)]' : 'text-[var(--workbench-muted-2)]'}`}>
            {msg.text}
          </div>

          <div className="mt-8 border-t border-[var(--workbench-border)] pt-5 text-xs leading-6 text-[var(--workbench-muted-2)]">
            AI Agent Scaffold · Next.js · draw.io iframe render surface
          </div>
        </section>
      </div>
    </main>
  );
}
