/*
 * XTM login-history capture (Playwright, headed).
 *
 * Opens the XTM Project Manager GUI in a real browser using a PERSISTENT
 * profile so your SSO session is reused on future runs. You log in (SSO/MFA)
 * and navigate to the login/activity-history screen yourself; meanwhile this
 * script records every data response (.serv / .action that returns JSON/text)
 * to a capture directory so we can identify which endpoint carries the login
 * history and what its payload looks like.
 *
 * Usage:
 *   node xtm_login_history_capture.js
 *
 * Close the browser window (or press Ctrl-C) when you've opened the
 * login-history screen. Captured responses land in:
 *   <scratchpad>/xtm_capture/
 */
const PLAYWRIGHT_PATH = process.env.XTM_PLAYWRIGHT_PATH ||
  '/Users/jayjay5032/quarterly-report-q2/node_modules/playwright';
const { chromium } = require(PLAYWRIGHT_PATH);
const fs = require('fs');
const path = require('path');

const GUI = 'https://churchofjesuschrist.xtm-intl.com/project-manager-gui';
const START_URL = process.env.XTM_START_URL || `${GUI}/users-pages.action`;
const PROFILE_DIR = process.env.XTM_PROFILE_DIR ||
  '/Users/jayjay5032/Projects/MartinMonthlyReport/.cache/pw-profile';
const CAPTURE_DIR = process.env.XTM_CAPTURE_DIR ||
  '/private/tmp/claude-502/-Users-jayjay5032-Projects-MartinMonthlyReport/b71f2613-59c3-45c3-9c9d-befda1694377/scratchpad/xtm_capture';
const MAX_MS = parseInt(process.env.XTM_MAX_MS || '900000', 10); // 15 min safety cap

// Noise we never want to save.
const SKIP = [
  'getUserImage', 'getLogo', 'chat/channel', 'sayHelloToServer',
  'pendo', '/scripts/', '/css/', '/themes/', '/node_modules/',
  'main-localization', 'JTemplates', '.woff', '.gif', '.svg', '.ico', '.js?',
];

fs.mkdirSync(CAPTURE_DIR, { recursive: true });
fs.mkdirSync(PROFILE_DIR, { recursive: true });

let n = 0;
const index = [];

(async () => {
  const ctx = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    viewport: null,
    args: ['--start-maximized'],
  });

  const onResponse = async (resp) => {
    const url = resp.url();
    if (!/\.serv|\.action/.test(url)) return;
    if (SKIP.some((s) => url.includes(s))) return;
    let body = '';
    try { body = await resp.text(); } catch (_) { return; }
    if (!body || body.length > 5_000_000) return;
    // Skip obvious HTML login pages (keep JSON/text data).
    const trimmed = body.trimStart();
    const looksHtml = trimmed.startsWith('<!DOCTYPE') || trimmed.startsWith('<html');
    n += 1;
    const safe = url.split('?')[0].split('/').slice(-1)[0].replace(/[^a-z0-9._-]/gi, '_');
    const file = path.join(CAPTURE_DIR, `${String(n).padStart(3, '0')}_${safe}.json`);
    let reqPostData = null;
    try { reqPostData = resp.request().postData(); } catch (_) {}
    const rec = {
      url,
      method: resp.request().method(),
      status: resp.status(),
      contentType: resp.headers()['content-type'] || '',
      requestPostData: reqPostData,
      requestHeaders: resp.request().headers(),
      looksHtml,
      bodyPreview: body.slice(0, 400),
    };
    fs.writeFileSync(file, JSON.stringify({ ...rec, body }, null, 2));
    index.push(rec);
    fs.writeFileSync(path.join(CAPTURE_DIR, '_index.json'), JSON.stringify(index, null, 2));
    process.stdout.write(`captured [${n}] ${resp.status()} ${url}\n`);
  };

  ctx.on('response', onResponse);

  const page = ctx.pages()[0] || (await ctx.newPage());
  await page.goto(START_URL, { waitUntil: 'domcontentloaded' }).catch(() => {});

  // Auto-submit the pre-filled XTM login form (creds saved in the profile).
  await page.waitForTimeout(2500);
  try {
    const pw = page.locator('input[type="password"]');
    if ((await pw.count()) && (await pw.first().inputValue())) {
      const submit = page.locator('button[type="submit"]');
      if (await submit.count()) { await submit.first().click({ timeout: 4000 }); }
    }
  } catch (_) {}

  console.log('\n' + '='.repeat(70));
  console.log('BROWSER OPEN (auto-login attempted). In the window:');
  console.log('  1. Open a PROJECT and go to its Statistics view');
  console.log('     (the one showing "Translation time [hh:mm:ss]").');
  console.log('  2. Let it load, then CLOSE the browser window.');
  console.log(`Captures -> ${CAPTURE_DIR}`);
  console.log('='.repeat(70) + '\n');

  const done = new Promise((resolve) => {
    ctx.on('close', () => resolve('window-closed'));
    setTimeout(() => resolve('timeout'), MAX_MS);
    process.on('SIGINT', () => resolve('sigint'));
  });
  const why = await done;
  console.log(`\nStopping (${why}). Total captured: ${n}. Index: ${path.join(CAPTURE_DIR, '_index.json')}`);
  try { await ctx.close(); } catch (_) {}
  process.exit(0);
})().catch((e) => { console.error('FATAL', e); process.exit(1); });
