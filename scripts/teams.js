/*
  Playwright script to open Microsoft Teams and persist login.
  - First run: a Chromium window opens at Teams. Log in once (manual or via env vars).
  - Next runs: login is reused automatically via persistent user data dir.

  Optional auto-fill (may not work with 2FA):
    export MS_EMAIL="you@example.com"
    export MS_PASSWORD="your_password"

  Usage:
    node scripts/teams.js                 # open Teams using persisted profile
    node scripts/teams.js --reset         # clear saved profile and start fresh login
    node scripts/teams.js --list          # print all team names
    node scripts/teams.js --team "Тест"    # open a team by (partial) name
    node scripts/teams.js --team "Тест" --exact   # exact matching
    node scripts/teams.js --watch-join            # watch for "Присоединиться" and click (30s for 10m)
    node scripts/teams.js --watch-join --interval-sec 30 --watch-minutes 10 [--reload-join]
    node scripts/teams.js --prejoin-timeout-sec 120   # wait for pre-join screen and click "Присоединиться сейчас"
    node scripts/teams.js --prejoin                   # explicitly wait for pre-join and click
*/

const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const TEAMS_URL = process.env.TEAMS_URL || 'https://teams.microsoft.com/';
const USER_DATA_DIR = path.resolve(__dirname, '../.pw-user-data/teams');

const rawArgs = process.argv.slice(2);
const args = new Set(rawArgs);

function getArgValue(flag) {
  // --flag=value or --flag value
  const withEq = rawArgs.find(a => a.startsWith(flag + '='));
  if (withEq) return withEq.slice(flag.length + 1);
  const i = rawArgs.indexOf(flag);
  if (i !== -1 && rawArgs[i + 1] && !rawArgs[i + 1].startsWith('--')) return rawArgs[i + 1];
  return null;
}

async function ensureDir(p) {
  await fs.promises.mkdir(p, { recursive: true }).catch(() => {});
}

async function resetProfileIfRequested() {
  if (!args.has('--reset')) return;
  if (!fs.existsSync(USER_DATA_DIR)) return;
  console.log('Resetting saved profile at', USER_DATA_DIR);
  await fs.promises.rm(USER_DATA_DIR, { recursive: true, force: true });
}

function isFirstRun() {
  return !fs.existsSync(USER_DATA_DIR);
}

async function maybeAutoLogin(page) {
  const email = process.env.MS_EMAIL;
  const password = process.env.MS_PASSWORD;
  if (!email || !password) {
    console.log('Tip: Set MS_EMAIL and MS_PASSWORD env vars to attempt auto-login.');
    return;
  }

  try {
    // Microsoft login pages change often; best-effort selectors with generous timeouts.
    // Step 1: Email
    await page.waitForLoadState('domcontentloaded', { timeout: 60_000 });
    // If redirected to login, handle it; otherwise this will time out and we skip.
    const emailInput = await page.waitForSelector('input[type="email"], input[name="loginfmt"]', { timeout: 15_000 });
    if (emailInput) {
      await emailInput.fill('');
      await emailInput.type(email, { delay: 20 });
      await page.keyboard.press('Enter');
    }

    // Step 2: Password
    const pwdInput = await page.waitForSelector('input[type="password"], input[name="passwd"]', { timeout: 30_000 });
    if (pwdInput) {
      await pwdInput.fill('');
      await pwdInput.type(password, { delay: 20 });
      await page.keyboard.press('Enter');
    }

    // Step 3: Stay signed in? Prefer "Yes" to keep session alive.
    const staySignedInYes = await page.waitForSelector('button:has-text("Yes"), input[type="submit"][value="Yes"]', { timeout: 20_000 }).catch(() => null);
    if (staySignedInYes) {
      await staySignedInYes.click();
    }

    console.log('Attempted auto-login with provided credentials.');
  } catch (e) {
    console.log('Auto-login attempt skipped or failed (likely due to 2FA or layout changes). Proceed to login manually in the opened browser window.');
  }
}

async function main() {
  await resetProfileIfRequested();
  await ensureDir(USER_DATA_DIR);

  const firstRun = isFirstRun();
  const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless: false,
    viewport: null,
    args: ['--start-maximized'],
  });

  // Use existing page or open a new one
  const page = context.pages()[0] || await context.newPage();
  await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });

  if (firstRun) {
    console.log('\nПервый запуск: В открывшемся окне войдите в Microsoft Teams.');
    console.log('После входа сессия сохранится. Последующие запуски будут автоматически авторизованы.');
    await maybeAutoLogin(page);
  } else {
    console.log('Открываю Teams с сохранённой сессией...');
  }

  // Keep the window open. If this is used in automation, exit once Teams UI is loaded.
  // Heuristic wait for main app to load (up to 2 minutes), then leave window open.
  try {
    await page.waitForURL(/teams\.microsoft\.com/, { timeout: 120_000 });
  } catch {}

  // Optionally navigate to Teams hub in the app bar (best-effort; labels vary by locale).
  await ensureOnTeamsHub(page).catch(() => {});

  // If user requested teams listing or selection, try to act now.
  const listRequested = args.has('--list') || args.has('--list-teams');
  const teamQuery = getArgValue('--team') || process.env.TEAM_NAME;
  const exact = args.has('--exact');

  if (listRequested || teamQuery) {
    try {
      await waitForTeamsList(page);
    } catch (e) {
      console.log('Не удалось дождаться списка команд. Откройте раздел "Команды" вручную и повторите.');
    }
  }

  if (listRequested) {
    const names = await collectTeamNames(page).catch(() => []);
    if (!names.length) {
      console.log('Список команд пуст или не найден.');
    } else {
      console.log('\nДоступные команды:');
      for (const n of names) console.log('- ' + n);
    }
  }

  if (teamQuery) {
    try {
      const clicked = await clickTeamByName(page, teamQuery, { exact });
      if (clicked) {
        console.log(`Открыта команда: ${clicked}`);
      } else {
        console.log(`Команда не найдена по запросу: "${teamQuery}"${exact ? ' (точное совпадение)' : ''}.`);
      }
    } catch (e) {
      console.log('Ошибка при выборе команды:', e?.message || e);
    }
  }

  // Watch and click "Присоединиться" button if requested
  const watchJoin = args.has('--watch-join') || process.env.WATCH_JOIN === '1';
  let joined = null;
  if (watchJoin) {
    const intervalSec = parseInt(getArgValue('--interval-sec') || process.env.WATCH_INTERVAL_SEC || '30', 10);
    const minutes = parseFloat(getArgValue('--watch-minutes') || process.env.WATCH_MINUTES || '10');
    const reloadEach = args.has('--reload-join') || process.env.WATCH_RELOAD === '1';
    const intervalMs = Math.max(5, intervalSec) * 1000;
    const timeoutMs = Math.max(0.5, minutes) * 60 * 1000;

    console.log(`Наблюдение за кнопкой "Присоединиться": каждые ${Math.round(intervalMs/1000)}с, до ${Math.round(timeoutMs/60000)} мин${reloadEach ? ' (с перезагрузкой страницы)' : ''}.`);
    joined = await waitAndClickJoin(page, { intervalMs, timeoutMs, reloadEach }).catch(() => false);
    if (joined) {
      console.log('Кнопка "Присоединиться" найдена и нажата.');
      const prejoinTimeoutMs = parseInt(getArgValue('--prejoin-timeout-sec') || process.env.PREJOIN_TIMEOUT_SEC || '120', 10) * 1000;
      const prejoined = await waitAndClickPrejoin(context, { timeoutMs: prejoinTimeoutMs }).catch(() => false);
      if (prejoined) {
        console.log('Кнопка "Присоединиться сейчас" нажата. Входим в собрание...');
      } else {
        console.log('Не удалось нажать "Присоединиться сейчас" в отведённое время.');
      }
    } else {
      console.log('За отведённое время кнопка "Присоединиться" не появилась.');
    }
  }

  // Allow explicitly watching for the prejoin screen independently
  const prejoinRequested = args.has('--prejoin') || args.has('--watch-prejoin') || !!getArgValue('--prejoin-timeout-sec') || process.env.WATCH_PREJOIN === '1';
  if (prejoinRequested && !joined) {
    const prejoinTimeoutMs = parseInt(getArgValue('--prejoin-timeout-sec') || process.env.PREJOIN_TIMEOUT_SEC || '120', 10) * 1000;
    console.log(`Ожидание экрана предварительного подключения (до ${Math.round(prejoinTimeoutMs/1000)} сек)...`);
    const prejoined = await waitAndClickPrejoin(context, { timeoutMs: prejoinTimeoutMs }).catch(() => false);
    if (prejoined) {
      console.log('Кнопка "Присоединиться сейчас" нажата. Входим в собрание...');
    } else {
      console.log('Не удалось нажать "Присоединиться сейчас" в отведённое время.');
    }
  }

  // Do not close context so the user can interact; process exits when the window is closed.
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});

// ------- Helpers for Teams UI -------

async function ensureOnTeamsHub(page) {
  // Try known selectors for the left app bar "Teams" button. Best-effort only.
  const selectors = [
    '[data-tid="app-bar-teams"]',
    'button[aria-label="Teams"]',
    'button[aria-label="Команды"]',
    'a[aria-label="Teams"]',
    'a[aria-label="Команды"]'
  ];
  for (const sel of selectors) {
    const el = await page.$(sel);
    if (el) {
      await el.click().catch(() => {});
      break;
    }
  }
}

async function waitForTeamsList(page) {
  // Primary selector from provided HTML snippet
  const primary = 'button[data-testid="team-name"]';
  const fallback = '[role="treeitem"][aria-label], button[aria-label]';
  try {
    await page.waitForSelector(primary, { timeout: 60_000 });
  } catch {
    await page.waitForSelector(fallback, { timeout: 30_000 });
  }
}

async function collectTeamNames(page) {
  const buttons = page.locator('button[data-testid="team-name"]');
  const count = await buttons.count();
  const names = [];
  for (let i = 0; i < count; i++) {
    const btn = buttons.nth(i);
    const aria = (await btn.getAttribute('aria-label')) || '';
    const text = (await btn.innerText().catch(() => '')) || '';
    const name = (aria || text).trim();
    if (name) names.push(name);
  }
  // Deduplicate, preserve order
  return [...new Set(names)];
}

async function clickTeamByName(page, query, { exact = false } = {}) {
  const buttons = page.locator('button[data-testid="team-name"]');
  const count = await buttons.count();
  if (count === 0) return null;

  const norm = s => s.normalize('NFKC').trim();
  const q = norm(query).toLocaleLowerCase();
  let matchIndex = -1;
  let matchName = null;

  for (let i = 0; i < count; i++) {
    const btn = buttons.nth(i);
    const aria = (await btn.getAttribute('aria-label')) || '';
    const text = (await btn.innerText().catch(() => '')) || '';
    const name = norm(aria || text);
    const low = name.toLocaleLowerCase();
    const ok = exact ? (low === q) : low.includes(q);
    if (ok) {
      matchIndex = i;
      matchName = name;
      break;
    }
  }

  if (matchIndex === -1) return null;
  await buttons.nth(matchIndex).click();
  // Wait a bit for the team view to load (best-effort)
  await page.waitForLoadState('domcontentloaded');
  return matchName;
}

async function waitAndClickJoin(page, { intervalMs = 30_000, timeoutMs = 10 * 60 * 1000, reloadEach = false } = {}) {
  const start = Date.now();
  const selectors = [
    'button[data-tid="channel-ongoing-meeting-banner-join-button"]',
    'button[aria-label*="Присоединиться"]',
    'button:has-text("Присоединиться")',
    'button[aria-label*="Join"]',
    'button:has-text("Join")'
  ];

  // Immediate check first
  const foundNow = await findAndClickFirstVisible(page, selectors);
  if (foundNow) return true;

  while (Date.now() - start < timeoutMs) {
    if (reloadEach) {
      try {
        await page.reload({ waitUntil: 'domcontentloaded' });
      } catch {}
    }

    const clicked = await findAndClickFirstVisible(page, selectors);
    if (clicked) return true;

    const leftMs = Math.max(0, timeoutMs - (Date.now() - start));
    console.log(`Кнопка \"Присоединиться\" не найдена. Следующая проверка через ${Math.round(intervalMs/1000)}с. Осталось ~${Math.ceil(leftMs/60000)} мин.`);
    await page.waitForTimeout(intervalMs);
  }
  return false;
}

async function findAndClickFirstVisible(page, selectors) {
  for (const sel of selectors) {
    const loc = page.locator(sel);
    const count = await loc.count();
    if (count === 0) continue;
    for (let i = 0; i < count; i++) {
      const el = loc.nth(i);
      if (await el.isVisible().catch(() => false)) {
        try {
          await el.click({ timeout: 5000 });
          return true;
        } catch {}
      }
    }
    // Fallback: try the first even if not reported visible
    try {
      await loc.first().click({ timeout: 5000 });
      return true;
    } catch {}
  }
  return false;
}

async function waitAndClickPrejoin(context, { timeoutMs = 120_000 } = {}) {
  const endAt = Date.now() + timeoutMs;
  const selectors = [
    '#prejoin-join-button',
    'button#prejoin-join-button',
    'button[data-tid="prejoin-join-button"]',
    'button[aria-label*="Присоединиться сейчас"]',
    'button:has-text("Присоединиться сейчас")',
    'button[aria-label*="Join now"]',
    'button:has-text("Join now")'
  ];

  // Helper to attempt clicking join on a single page
  const tryPage = async (page) => {
    for (const sel of selectors) {
      const loc = page.locator(sel);
      const count = await loc.count();
      if (count === 0) continue;
      const candidate = loc.first();
      // Wait a brief moment for button to become enabled/visible
      try {
        await candidate.waitFor({ state: 'visible', timeout: 5000 });
      } catch {}
      // If disabled, skip
      try {
        if (await candidate.isDisabled()) continue;
      } catch {}
      try {
        await candidate.click({ timeout: 5000 });
        return true;
      } catch {}
    }
    return false;
  };

  // First immediate scan
  for (const page of context.pages()) {
    if (await tryPage(page)) return true;
  }

  // Then poll until timeout, capturing new tabs if any
  while (Date.now() < endAt) {
    // Race: either new page opens, or just wait a bit before scanning again
    const waitMs = 1000;
    const newPagePromise = context.waitForEvent('page', { timeout: waitMs }).catch(() => null);
    const newPage = await newPagePromise;
    if (newPage) {
      try { await newPage.waitForLoadState('domcontentloaded', { timeout: 10_000 }); } catch {}
    }

    for (const page of context.pages()) {
      if (await tryPage(page)) return true;
    }
  }
  return false;
}
