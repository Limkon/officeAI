import test from 'node:test';
import assert from 'node:assert/strict';
import { createDefaultConfig, type AppSettings } from '@officeai/shared';
import { LifecycleMonitor } from './lifecycle.js';
import { isOfficeProcessName } from './process-watch.js';

function settings(patch: Partial<AppSettings> = {}): AppSettings {
  return { ...createDefaultConfig().settings, launchMode: 'on-demand', ...patch };
}

/** Minimal fake clock so the monitor can be tested without real timers */
function createHarness(initial: Partial<AppSettings>, officeStates: Array<boolean>) {
  let now = 0;
  let stateIdx = 0;
  const exits: string[] = [];
  const monitor = new LifecycleMonitor(() => settings(initial), {
    launcher: 'test',
    now: () => now,
    pollMs: 1000,
    onExitRequested: (reason) => exits.push(reason),
    detectOffice: async () => {
      const running = officeStates[Math.min(stateIdx, officeStates.length - 1)];
      stateIdx += 1;
      return { checked: true, running, processes: running ? ['winword'] : [] };
    }
  });

  return {
    monitor,
    exits,
    advance(ms: number) {
      now += ms;
    },
    /** Runs the private poll loop once (start() uses setInterval, so call check indirectly) */
    async tick(times = 1) {
      for (let i = 0; i < times; i += 1) {
        // @ts-expect-error private method accessed from the test harness
        await monitor.check();
        now += 1000;
      }
    }
  };
}

test('isOfficeProcessName matches Office/WPS main binaries only', () => {
  for (const name of ['WINWORD.EXE', 'excel', 'POWERPNT.exe', 'wps.exe', 'et.exe', 'wpp.exe']) {
    assert.equal(isOfficeProcessName(name), true, name);
  }
  for (const name of ['wpscloudsvr.exe', 'wpsupdate.exe', 'notepad.exe', 'explorer.exe', 'code.exe']) {
    assert.equal(isOfficeProcessName(name), false, name);
  }
});

test('LifecycleMonitor does nothing when launch mode is not on-demand', async () => {
  const h = createHarness({ launchMode: 'always' }, [false]);
  h.advance(5 * 60 * 60 * 1000);
  await h.tick(3);
  assert.deepEqual(h.exits, []);
});

test('LifecycleMonitor keeps running while Office is open', async () => {
  const h = createHarness({}, [true, true, true]);
  await h.tick(3);
  assert.deepEqual(h.exits, []);
});

test('LifecycleMonitor exits after Office closes and the grace period passes', async () => {
  const h = createHarness({ exitWithOffice: true, idleExitMinutes: 0 }, [true, false, false]);
  await h.tick(1); // Office seen, keep running
  assert.deepEqual(h.exits, []);
  await h.tick(1); // closed: within 90s grace period
  assert.deepEqual(h.exits, []);
  h.advance(120_000);
  await h.tick(1);
  assert.equal(h.exits.length, 1);
  assert.match(h.exits[0], /Office/);
});

test('LifecycleMonitor keeps running while a task pane heartbeat is fresh', async () => {
  // No office process (e.g. an unexpected process name), but the pane is alive
  const h = createHarness({ exitWithOffice: true, idleExitMinutes: 0, sessionGraceMinutes: 2 }, [false]);
  h.monitor.watch({ app: 'winword', source: 'addin' });
  await h.tick(2);
  assert.deepEqual(h.exits, []);

  // Heartbeat expires → no process, no session → exit
  h.advance(5 * 60_000);
  await h.tick(2);
  assert.equal(h.exits.length, 1);
  assert.match(h.exits[0], /Office/);
});

test('LifecycleMonitor refreshes session grace on each heartbeat', async () => {
  const h = createHarness({ exitWithOffice: true, idleExitMinutes: 0, sessionGraceMinutes: 2 }, [false]);
  for (let i = 0; i < 4; i += 1) {
    h.monitor.watch({ app: 'wps', source: 'addin' });
    h.advance(60_000); // heartbeat every minute, grace is 2 minutes
    await h.tick(1);
  }
  assert.deepEqual(h.exits, []);
  assert.equal(h.monitor.getStatus().sessions.length, 1);
});

test('LifecycleMonitor exits on idle timeout even if Office stays open', async () => {
  const h = createHarness({ idleExitMinutes: 15 }, [true]);
  await h.tick(1);
  h.advance(16 * 60 * 1000);
  await h.tick(1);
  assert.equal(h.exits.length, 1);
  assert.match(h.exits[0], /15 分钟/);
});

test('LifecycleMonitor exits only once even if polled again', async () => {
  const h = createHarness({ idleExitMinutes: 1 }, [false]);
  h.advance(5 * 60 * 1000);
  await h.tick(4);
  assert.equal(h.exits.length, 1);
});
