/*
 * Fetch XTM user login/logout history via the PM-GUI backend, reusing the
 * persistent (already-logged-in) Playwright profile. Runs headless.
 *
 * The GUI endpoint getUserLoginHistory.serv needs both the session cookies
 * (carried by the persistent profile) AND a `uust` session-token header that
 * the frontend attaches. We load the app in a real page so the app's own
 * fetch machinery is available, then issue the paged POSTs from inside the
 * page context (page.evaluate) so `uust` and cookies are applied exactly as
 * the UI does it.
 *
 * Usage:
 *   node fetch_login_history.js <dateFrom MM-DD-YYYY> <dateTo MM-DD-YYYY> [outfile]
 *   (defaults to the last 35 days)
 *
 * Output: JSON { dateFrom, dateTo, recordsFiltered, count, data:[...] }
 */
const PLAYWRIGHT_PATH = process.env.XTM_PLAYWRIGHT_PATH ||
  '/Users/jayjay5032/quarterly-report-q2/node_modules/playwright';
const { chromium } = require(PLAYWRIGHT_PATH);
const fs = require('fs');

const GUI = 'https://churchofjesuschrist.xtm-intl.com/project-manager-gui';
const PROFILE_DIR = process.env.XTM_PROFILE_DIR ||
  '/Users/jayjay5032/Projects/MartinMonthlyReport/.cache/pw-profile';

// NOTE: getUserLoginHistory.serv expects the request dateFrom/dateTo as
// DD-MM-YYYY, even though it returns the DATE field as MM-DD-YYYY. Verified
// empirically: requesting 06-01-2026..07-02-2026 returned records bounded at
// 2026-01-06..2026-02-07 (i.e. the server read the request as DD-MM).
function fmt(d) {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd}-${mm}-${d.getFullYear()}`;
}
const argFrom = process.argv[2];
const argTo = process.argv[3];
const outFile = process.argv[4] ||
  '/Users/jayjay5032/Projects/MartinMonthlyReport/.cache/login_history/latest.json';
const now = new Date();
const dateTo = argTo || fmt(now);
const dateFrom = argFrom || fmt(new Date(now.getTime() - 35 * 864e5));

const HEADLESS = process.env.XTM_HEADLESS === '1';

// The XTM login page (login.jsp) autofills client/username/password from the
// browser profile's saved credentials. We only submit when the password field
// already has a value — we never type or store credentials ourselves.
async function tryAutoLogin(page) {
  try {
    const pw = page.locator('input[type="password"]');
    if ((await pw.count()) && (await pw.first().inputValue())) {
      const submit = page.locator('button[type="submit"]');
      if (await submit.count()) {
        await submit.first().click({ timeout: 4000 });
        return true;
      }
    }
  } catch (_) {}
  return false;
}

(async () => {
  const ctx = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: HEADLESS,
    viewport: null,
    args: HEADLESS ? [] : ['--start-maximized'],
  });
  const page = ctx.pages()[0] || (await ctx.newPage());
  // The frontend attaches a per-session `uust` token header to every .serv
  // request; getUserLoginHistory.serv rejects (500) without it. Capture it
  // passively from any app request (getAppInitData fires on page load).
  let uust = null;
  page.on('request', (r) => {
    const u = r.headers()['uust'];
    if (u) uust = u;
  });
  await page.goto(`${GUI}/list-users-page.action`, { waitUntil: 'domcontentloaded' });

  // A captured `uust` means the app booted authenticated (it only fires its
  // .serv requests when logged in). Use that as the login signal.
  const waitForUust = async (ms) => {
    const deadline = Date.now() + ms;
    while (Date.now() < deadline && !uust) await page.waitForTimeout(500);
    return !!uust;
  };

  if (!(await waitForUust(6000))) {
    // Not logged in — submit the pre-filled login form (saved creds autofill).
    console.error('Login page detected — submitting saved credentials...');
    await tryAutoLogin(page);
    if (!(await waitForUust(20000))) {
      if (HEADLESS) {
        console.error('Auto-login failed (headless autofill unreliable). Run without XTM_HEADLESS=1.');
        await ctx.close();
        process.exit(2);
      }
      // Headed fallback: keep retrying (covers a slow page or an extra prompt).
      console.error('Still not logged in — retrying / waiting for any extra prompt...');
      const deadline = Date.now() + 300000;
      while (Date.now() < deadline && !uust) {
        await tryAutoLogin(page);
        await page.waitForTimeout(3000);
      }
      if (!uust) {
        console.error('Timed out waiting for login.');
        await ctx.close();
        process.exit(2);
      }
    }
  }
  console.error('Session OK — fetching login history...');

  // Page through all records from inside the app context (uust auto-applied by the app's
  // fetch interceptor; we also pass credentials for cookies).
  const pageSize = 500;
  let start = 0;
  let total = Infinity;
  const all = [];
  while (start < total) {
    const res = await page.evaluate(
      async ({ gui, body, uust }) => {
        const r = await fetch(`${gui}/getUserLoginHistory.serv`, {
          method: 'POST',
          credentials: 'include',
          headers: { 'content-type': 'application/json', accept: 'application/json, text/plain, */*', uust },
          body: JSON.stringify(body),
        });
        const text = await r.text();
        try { return { status: r.status, json: JSON.parse(text) }; }
        catch (_) { return { status: r.status, text: text.slice(0, 300) }; }
      },
      {
        gui: GUI,
        uust,
        body: { length: pageSize, start, orderColumn: 'DATE', orderDir: 'desc',
                searchValue: '', action: null, dateFrom, dateTo },
      }
    );
    if (res.status !== 200 || !res.json) {
      console.error('Bad response at start=' + start, res.status, res.text || '');
      break;
    }
    total = res.json.recordsFiltered ?? 0;
    const batch = res.json.data || [];
    all.push(...batch);
    process.stderr.write(`fetched ${all.length}/${total}\n`);
    if (batch.length === 0) break;
    start += pageSize;
  }

  fs.mkdirSync(require('path').dirname(outFile), { recursive: true });
  fs.writeFileSync(outFile, JSON.stringify(
    { dateFrom, dateTo, recordsFiltered: total, count: all.length, data: all }, null, 2));
  console.log(`OK ${all.length} records (${dateFrom}..${dateTo}) -> ${outFile}`);
  await ctx.close();
  process.exit(0);
})().catch((e) => { console.error('FATAL', e); process.exit(1); });
