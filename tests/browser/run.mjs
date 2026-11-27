#!/usr/bin/env node
// The browser test: the BUILT package, in a real headless Chromium, offline.
//
//   npm run build && npm run test:browser
//
// Each case in tests/browser/cases/*.mjs is bundled on its own with Vite from
// the built `dist/` (the specifiers `document-kit`, `document-kit/docx`, `document-kit/html`,
// `document-kit/layout`, `document-kit/template`, `document-kit/recipe` and `document-kit/ream` resolve
// to dist/), served with `fonts/` from 127.0.0.1 on a free port, and run in a
// fresh page. Every request that is not to that server is aborted, recorded,
// and fails the run.
//
// A case is an ES module whose default export is `async function (ctx)`.
// `ctx.loadFont(file)` returns the bytes of a file in `fonts/` as a Uint8Array
// (the fonts are read into the page before the case starts, so the call is
// synchronous; awaiting it is fine too). The case returns a small
// JSON-serialisable summary and throws on failure. One line is printed per case.
//
// Chromium: env DOCUMENT_KIT_CHROMIUM, else the newest chrome-headless-shell
// Playwright has downloaded, else an installed Google Chrome. This machine may
// be running other browsers: only the one launched here is ever closed.

import { existsSync, readdirSync, readFileSync, rmSync, mkdirSync, writeFileSync, statSync } from 'node:fs';
import { createServer } from 'node:http';
import { homedir, platform } from 'node:os';
import { basename, extname, join, normalize, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = fileURLToPath(new URL('.', import.meta.url));
const repo = resolve(here, '..', '..');
const dist = join(repo, 'dist');
const fontsDir = join(repo, 'fonts');
// DOCUMENT_KIT_BROWSER_CASES points the run at another cases folder (used to
// check the runner itself, e.g. that an outside request fails the run).
const casesDir = process.env.DOCUMENT_KIT_BROWSER_CASES
  ? resolve(process.env.DOCUMENT_KIT_BROWSER_CASES)
  : join(here, 'cases');
const outDir = join(here, '.out');
const CASE_TIMEOUT_MS = 180_000;

function fail(message) {
  console.error(`test:browser: ${message}`);
  process.exit(1);
}

// --- Chromium -----------------------------------------------------------------

function findChromium() {
  const fromEnv = process.env.DOCUMENT_KIT_CHROMIUM;
  if (fromEnv) {
    if (!existsSync(fromEnv)) fail(`DOCUMENT_KIT_CHROMIUM points at ${fromEnv}, which does not exist`);
    return fromEnv;
  }
  const caches = [
    process.env.PLAYWRIGHT_BROWSERS_PATH,
    join(homedir(), 'Library', 'Caches', 'ms-playwright'),
    join(homedir(), '.cache', 'ms-playwright'),
    process.env.LOCALAPPDATA ? join(process.env.LOCALAPPDATA, 'ms-playwright') : undefined,
  ].filter((d) => d && existsSync(d));
  const shells = [];
  for (const cache of caches) {
    for (const dir of readdirSync(cache)) {
      const m = /^chromium_headless_shell-(\d+)$/.exec(dir);
      if (!m) continue;
      const root = join(cache, dir);
      for (const sub of readdirSync(root)) {
        for (const name of ['chrome-headless-shell', 'chrome-headless-shell.exe']) {
          const exe = join(root, sub, name);
          if (existsSync(exe)) shells.push({ revision: Number(m[1]), exe });
        }
      }
    }
  }
  shells.sort((a, b) => b.revision - a.revision);
  if (shells[0]) return shells[0].exe;
  const chromes = {
    darwin: ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'],
    linux: ['/usr/bin/google-chrome', '/usr/bin/google-chrome-stable', '/opt/google/chrome/chrome'],
    win32: [
      join(process.env.PROGRAMFILES ?? 'C:\\Program Files', 'Google', 'Chrome', 'Application', 'chrome.exe'),
    ],
  }[platform()] ?? [];
  const chrome = chromes.find((p) => existsSync(p));
  if (chrome) return chrome;
  fail(
    'no Chromium found. Set DOCUMENT_KIT_CHROMIUM to a Chromium or Chrome executable, ' +
      'or download one with `npx playwright-core install chromium-headless-shell`.',
  );
}

// --- Bundling ----------------------------------------------------------------

const ENTRIES = { '': 'index', '/docx': 'docx', '/html': 'html', '/layout': 'layout', '/pdf-preview': 'pdf-preview', '/template': 'template', '/recipe': 'recipe', '/ream': 'ream' };

async function bundleCase(build, file) {
  const name = basename(file, '.mjs');
  const src = join(outDir, 'src', name);
  const www = join(outDir, 'www', name);
  mkdirSync(src, { recursive: true });
  const entry = join(src, 'entry.mjs');
  writeFileSync(
    entry,
    `import run from ${JSON.stringify(file)};\n` +
      `const files = ${JSON.stringify(fontFiles())};\n` +
      'const fonts = new Map();\n' +
      'let setup;\n' +
      'try {\n' +
      '  await Promise.all(files.map(async (f) => {\n' +
      "    const res = await fetch('/fonts/' + encodeURIComponent(f));\n" +
      "    if (!res.ok) throw new Error('could not read fonts/' + f + ': HTTP ' + res.status);\n" +
      '    fonts.set(f, new Uint8Array(await res.arrayBuffer()));\n' +
      '  }));\n' +
      '} catch (e) { setup = e; }\n' +
      'const ctx = {\n' +
      '  loadFont(file) {\n' +
      '    const bytes = fonts.get(file);\n' +
      "    if (!bytes) throw new Error('no such file in fonts/: ' + file);\n" +
      '    return bytes.slice();\n' +
      '  },\n' +
      '};\n' +
      'window.__runCase = () => (setup ? Promise.reject(setup) : run(ctx));\n' +
      'window.__ready = true;\n',
  );
  await build({
    configFile: false,
    logLevel: 'error',
    root: src,
    resolve: {
      alias: Object.entries(ENTRIES).map(([sub, file]) => ({
        find: new RegExp(`^document-kit${sub.replace('/', '\\/')}$`),
        replacement: join(dist, `${file}.js`),
      })),
    },
    build: {
      outDir: www,
      emptyOutDir: true,
      target: 'es2022',
      minify: false,
      sourcemap: false,
      lib: { entry, formats: ['es'], fileName: () => 'entry.js' },
    },
  });
  writeFileSync(
    join(www, 'index.html'),
    `<!doctype html><meta charset="utf-8"><title>${name}</title><script type="module" src="./entry.js"></script>\n`,
  );
  return name;
}

function fontFiles() {
  return readdirSync(fontsDir).filter((f) => !f.startsWith('.') && statSync(join(fontsDir, f)).isFile());
}

// --- Serving -----------------------------------------------------------------

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.json': 'application/json',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.txt': 'text/plain; charset=utf-8',
  '.md': 'text/markdown; charset=utf-8',
  '.wasm': 'application/wasm',
};

function serve() {
  const roots = { '/fonts/': fontsDir, '/case/': join(outDir, 'www') };
  const server = createServer((req, res) => {
    const path = decodeURIComponent(new URL(req.url ?? '/', 'http://127.0.0.1').pathname);
    for (const [prefix, root] of Object.entries(roots)) {
      if (!path.startsWith(prefix)) continue;
      const target = normalize(join(root, path.slice(prefix.length)));
      if (!target.startsWith(root + sep) || !existsSync(target) || !statSync(target).isFile()) break;
      res.writeHead(200, { 'content-type': TYPES[extname(target)] ?? 'application/octet-stream' });
      res.end(readFileSync(target));
      return;
    }
    res.writeHead(404, { 'content-type': 'text/plain' });
    res.end('not found');
  });
  return new Promise((ok, err) => {
    server.once('error', err);
    server.listen(0, '127.0.0.1', () => ok(server));
  });
}

// --- Running -----------------------------------------------------------------

function oneLine(value, max = 160) {
  const text = typeof value === 'string' ? value : JSON.stringify(value);
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

async function main() {
  for (const file of Object.values(ENTRIES)) {
    if (!existsSync(join(dist, `${file}.js`))) fail(`dist/${file}.js is missing: run \`npm run build\` first`);
  }
  const cases = existsSync(casesDir)
    ? readdirSync(casesDir).filter((f) => f.endsWith('.mjs')).sort()
    : [];
  if (cases.length === 0) fail('no cases in tests/browser/cases');

  const executablePath = findChromium();
  const { build } = await import('vite');
  const { chromium } = await import('playwright-core');

  rmSync(outDir, { recursive: true, force: true });
  const bundled = [];
  for (const file of cases) {
    const name = basename(file, '.mjs');
    try {
      await bundleCase(build, join(casesDir, file));
      bundled.push({ name });
    } catch (error) {
      bundled.push({ name, bundleError: error instanceof Error ? error.message : String(error) });
    }
  }

  const server = await serve();
  const { port } = server.address();
  const origin = `http://127.0.0.1:${port}`;
  const outside = new Map(); // url → the case that asked for it
  let current = '(setup)';
  let browser;
  let failures = 0;
  console.log(`test:browser: ${basename(executablePath)} (${executablePath}), serving ${origin}`);
  try {
    browser = await chromium.launch({
      executablePath,
      headless: true,
      args: ['--disable-background-networking', '--disable-component-update', '--no-first-run'],
    });
    const context = await browser.newContext();
    // Abort and record every request that is not to our own server.
    await context.route('**/*', (route) => {
      const url = route.request().url();
      if (new URL(url).origin === origin) return route.continue();
      if (!outside.has(url)) outside.set(url, current);
      return route.abort('blockedbyclient');
    });
    context.on('request', (request) => {
      const url = request.url();
      if (/^(data|blob|about):/.test(url)) return;
      if (new URL(url).origin !== origin && !outside.has(url)) outside.set(url, current);
    });

    for (const { name, bundleError } of bundled) {
      if (bundleError) {
        failures += 1;
        console.log(`FAIL ${name}  could not bundle: ${oneLine(bundleError, 400)}`);
        continue;
      }
      current = name;
      const page = await context.newPage();
      page.on('websocket', (ws) => outside.set(ws.url(), name));
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      const started = Date.now();
      try {
        await page.goto(`${origin}/case/${name}/index.html`);
        await page.waitForFunction(() => window.__ready === true || undefined, null, {
          timeout: 60_000,
        });
        const summary = await Promise.race([
          page.evaluate(async () => {
            try {
              return { ok: true, value: (await window.__runCase()) ?? null };
            } catch (e) {
              return { ok: false, error: e instanceof Error ? `${e.message}\n${e.stack ?? ''}` : String(e) };
            }
          }),
          new Promise((_, reject) =>
            setTimeout(() => reject(new Error(`timed out after ${CASE_TIMEOUT_MS} ms`)), CASE_TIMEOUT_MS),
          ),
        ]);
        const ms = Date.now() - started;
        if (summary.ok) console.log(`PASS ${name}  ${ms} ms  ${oneLine(summary.value)}`);
        else {
          failures += 1;
          console.log(`FAIL ${name}  ${ms} ms  ${summary.error}`);
        }
      } catch (error) {
        failures += 1;
        const extra = errors.length > 0 ? ` (page errors: ${errors.join('; ')})` : '';
        console.log(`FAIL ${name}  ${error instanceof Error ? error.message : String(error)}${extra}`);
      } finally {
        await page.close();
      }
    }
  } finally {
    if (browser) await browser.close();
    server.closeAllConnections();
    await new Promise((ok) => server.close(ok));
  }

  if (outside.size > 0) {
    failures += 1;
    console.log(`FAIL network  ${outside.size} request(s) outside ${origin} were blocked:`);
    for (const [url, from] of outside) console.log(`  ${from}: ${url}`);
  } else {
    console.log(`PASS network  no request left ${origin}`);
  }
  if (failures > 0) {
    console.log(`test:browser: ${failures} failure(s)`);
    process.exit(1);
  }
  console.log(`test:browser: ${bundled.length} case(s) passed`);
  // Vite's bundler can keep a worker alive after a build; nothing is left to do.
  process.exit(0);
}

main().catch((error) => fail(error instanceof Error ? (error.stack ?? error.message) : String(error)));
