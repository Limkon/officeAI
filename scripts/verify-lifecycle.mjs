import { LifecycleMonitor } from '../packages/gateway/dist/lifecycle.js';
import { createDefaultConfig } from '../packages/shared/dist/index.js';

const settings = {
  ...createDefaultConfig().settings,
  launchMode: 'on-demand',
  idleExitMinutes: 0,
  exitWithOffice: true
};

let now = 0;
let polls = 0;
const exits = [];

const monitor = new LifecycleMonitor(() => settings, {
  launcher: 'office',
  now: () => now,
  pollMs: 5000,
  onExitRequested: (reason) => exits.push(reason),
  detectOffice: async () => {
    polls += 1;
    // 前 2 次轮询模拟 Word 打开，之后模拟 Word 已关闭
    const running = polls <= 2;
    return { checked: true, running, processes: running ? ['winword'] : [] };
  }
});

const tick = async (n) => {
  for (let i = 0; i < n; i += 1) {
    await monitor['check']();
    now += 5000;
  }
};

monitor.watch({ app: 'winword', source: 'addin' }); // Word 打开时侧边栏心跳
await tick(3);
const live = monitor.getStatus();
console.log(
  `[Word 打开中] 退出次数=${exits.length}  活跃会话=${live.sessions.map((s) => s.app).join(',')}  进程=${live.office.processes.join(',')}`
);

now += 200_000; // Word 关闭 + 心跳过期
await tick(3);
console.log(`[Word 关闭后] 退出原因=${JSON.stringify(exits)}`);
