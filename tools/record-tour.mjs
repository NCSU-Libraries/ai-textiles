// Records the captioned walkthrough video, and one short clip per Explore step, as 1080p MP4s.
//
//   tools/quarto.sh render site      # first: the tour shows real pages from site/_site
//   node tools/record-tour.mjs [last-scene]
//
// Writes site/video/tour.mp4 and tour-poster.jpg, site/video/step-<n>.mp4 and .jpg for each step, and
// site/_tour-chapters.md (chapter buttons and the caption transcript, included by index.qmd). Needs Google Chrome and
// ffmpeg (set CHROME and FFMPEG if they aren't at the macOS defaults). [last-scene] stops early, for a quick test, and
// skips the clips.
//
// tools/tour/stage.html holds simplified re-creations of Gemini and Google Sheets, a browser frame, captions and title
// cards. This script serves it over HTTP with the real screenshots (site/images/screens) and the rendered site, drives
// it scene by scene through window.stage (tools/tour/scenes.mjs), and records Chrome's screencast. Numbers come from
// site/_variables.yml, prompts from site/explore.qmd, and the sheet's rows from site/downloads/knitting-practice.csv,
// so the video follows the pages.
import { spawn, spawnSync } from 'node:child_process';
import { createServer } from 'node:http';
import { cpSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { extname, join, resolve } from 'node:path';

const ROOT = resolve(new URL('..', import.meta.url).pathname);
const LAST = process.argv[2] || '';
const W = 1280;
const H = 720;
const SCALE = 1.5;
const BUILD = join(ROOT, 'build', 'tour');
const FRAMES = join(BUILD, 'frames');
const OUT = join(ROOT, 'site', 'video');
const SITE = join(ROOT, 'site', '_site');
if (!existsSync(join(SITE, 'practices.html'))) { console.error('Render the site first: tools/quarto.sh render site'); process.exit(1); }
rmSync(BUILD, { recursive: true, force: true });
mkdirSync(FRAMES, { recursive: true });
mkdirSync(OUT, { recursive: true });
cpSync(join(ROOT, 'tools', 'tour'), join(BUILD, 'www'), { recursive: true });
cpSync(join(ROOT, 'site', 'images', 'lib_logo_white.png'), join(BUILD, 'www', 'logo-white.png'));
cpSync(join(ROOT, 'site', 'images', 'screens'), join(BUILD, 'www', 'shots'), { recursive: true });
cpSync(SITE, join(BUILD, 'www', 'site'), { recursive: true, filter: (p) => !p.includes(join(SITE, 'video')) });

// Facts the scenes use. The variables file is flat enough to read line by line: `section:` then `  key: "value"`.
const vars = {};
let section = '';
for (const line of readFileSync(join(ROOT, 'site', '_variables.yml'), 'utf8').split('\n')) {
  let m = line.match(/^([\w-]+):\s*(?:"([^"]*)"|([^#\s][^#]*?))?\s*(#.*)?$/);
  if (m) { section = m[1]; if (m[2] !== undefined || m[3] !== undefined) vars[m[1]] = (m[2] ?? m[3]).trim(); continue; }
  m = line.match(/^\s+([\w-]+):\s*"([^"]*)"/);
  if (m) vars[section + '.' + m[1]] = m[2];
}
const v = (k) => { if (!(k in vars)) throw new Error('no variable ' + k); return vars[k]; };
const fill = (text) => text.replace(/\{\{< var ([\w.-]+) >\}\}/g, (_, k) => v(k));

// The Explore page's prompts, word for word, in page order (the same ones tools/extract_prompts.py lists).
const explore = readFileSync(join(ROOT, 'site', 'explore.qmd'), 'utf8');
const prompts = [...explore.matchAll(/::: \{\.prompt\}\n```default\n([\s\S]*?)\n```\n:::/g)].map((m) => fill(m[1].trim()));

// The practice data, parsed (quoted fields may hold commas).
function parseCsv(text) {
  const rows = []; let row = [], field = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"' && text[i + 1] === '"') { field += '"'; i++; } else if (c === '"') q = false; else field += c; }
    else if (c === '"') q = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
    else if (c !== '\r') field += c;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}
const csv = parseCsv(readFileSync(join(ROOT, 'site', 'downloads', 'knitting-practice.csv'), 'utf8'));
const data = { rows: csv, cell: (r, c) => csv[r - 1]?.[c] ?? '' };

// A tiny static server for the stage, the screenshots and the site.
const types = { '.html': 'text/html; charset=utf-8', '.png': 'image/png', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml',
  '.json': 'application/json', '.woff2': 'font/woff2', '.woff': 'font/woff', '.jpg': 'image/jpeg' };
const server = createServer((req, res) => {
  let path = join(BUILD, 'www', decodeURIComponent(new URL(req.url, 'http://x').pathname));
  if (existsSync(path) && !extname(path)) path = join(path, 'index.html');
  if (!path.startsWith(join(BUILD, 'www')) || !existsSync(path)) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'content-type': types[extname(path)] || 'application/octet-stream' });
  res.end(readFileSync(path));
});
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const BASE = `http://127.0.0.1:${server.address().port}`;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const chrome = process.env.CHROME || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const port = 9700 + Math.floor(Math.random() * 200);
const proc = spawn(chrome, ['--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', '--no-default-browser-check',
  '--force-color-profile=srgb', `--user-data-dir=${mkdtempSync(join(tmpdir(), 'tour-'))}`, `--remote-debugging-port=${port}`, 'about:blank'],
{ stdio: 'ignore' });

let target;
for (let i = 0; i < 50 && !target; i++) {
  await sleep(200);
  try { target = (await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find((t) => t.type === 'page'); } catch {}
}
const ws = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((r) => ws.addEventListener('open', r, { once: true }));
let id = 0;
const pending = new Map();
const frames = [];
const errors = [];
ws.addEventListener('message', (event) => {
  const msg = JSON.parse(event.data);
  if (msg.id && pending.has(msg.id)) { pending.get(msg.id)(msg); pending.delete(msg.id); }
  if (msg.method === 'Page.screencastFrame') {
    const file = join(FRAMES, String(frames.length).padStart(6, '0') + '.jpg');
    writeFileSync(file, Buffer.from(msg.params.data, 'base64'));
    frames.push({ file, t: msg.params.metadata.timestamp });
    send('Page.screencastFrameAck', { sessionId: msg.params.sessionId });
  }
  if (msg.method === 'Runtime.exceptionThrown') errors.push(msg.params.exceptionDetails.exception?.description);
});
const send = (method, params = {}) => new Promise((res) => { const n = ++id; pending.set(n, res); ws.send(JSON.stringify({ id: n, method, params })); });
async function evaluate(expression) {
  const result = await send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
  if (result.result?.exceptionDetails) throw new Error(result.result.exceptionDetails.exception?.description || 'evaluate failed: ' + expression.slice(0, 80));
  return result.result?.result?.value;
}
const js = (x) => JSON.stringify(x);
const S = (code) => evaluate(`(async () => { ${code} })()`);

// Cursor: drawn into the stage and moved with real mouse events.
let mouse = { x: W * 0.62, y: H * 0.72 };
async function moveTo(x, y, ms = 700) {
  const steps = Math.max(10, Math.round(ms / 16));
  const from = { ...mouse };
  for (let i = 1; i <= steps; i++) {
    const p = i / steps;
    const e = p < 0.5 ? 4 * p * p * p : 1 - Math.pow(-2 * p + 2, 3) / 2;
    await send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: from.x + (x - from.x) * e, y: from.y + (y - from.y) * e });
    await sleep(ms / steps);
  }
  mouse = { x, y };
}
async function toCell(ref, ms = 600) {
  const p = await evaluate(`stage.sheet.where(${js(ref)})`);
  if (p) await moveTo(p.x, p.y, ms);
  return !!p;
}
const CURSOR = `(() => {
  const style = document.createElement('style');
  style.textContent = '#tour-cursor { position: fixed; left: 0; top: 0; z-index: 2147483647; pointer-events: none; transform: translate(-200px, -200px); } #tour-cursor svg { display: block; filter: drop-shadow(0 1px 2px rgba(0,0,0,.4)); } body.carded #tour-cursor { opacity: 0; }';
  document.head.append(style);
  const c = document.createElement('div'); c.id = 'tour-cursor';
  c.innerHTML = '<svg width="26" height="26" viewBox="0 0 24 24"><path d="M5 2.6 19.2 12.6l-6.4 1.2 3.8 7.2-2.7 1.4-3.8-7.3-4.9 4.3Z" fill="#111" stroke="#fff" stroke-width="1.5" stroke-linejoin="round"/></svg>';
  document.body.append(c);
  document.addEventListener('mousemove', (e) => { c.style.transform = 'translate(' + (e.clientX - 6) + 'px, ' + (e.clientY - 3) + 'px)'; }, true);
})()`;

// Captions, collected for the transcript, with the chapter each belongs to.
const transcript = [];
let chapter = '';
const onCaption = (text) => { if (text && transcript.at(-1)?.text !== text) transcript.push({ chapter, text }); };

const scenes = (await import(join(ROOT, 'tools', 'tour', 'scenes.mjs'))).default({ S, js, sleep, moveTo, toCell, evaluate, v, prompts, data, onCaption });

await send('Runtime.enable');
await send('Page.enable');
await send('Emulation.setDeviceMetricsOverride', { width: W, height: H, deviceScaleFactor: SCALE, mobile: false });
await send('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value: 'light' }] });
await send('Page.navigate', { url: BASE + '/stage.html' });
for (let i = 0; i < 100; i++) { await sleep(100); if ((await evaluate('document.readyState')) === 'complete') break; }
await evaluate(CURSOR);
await send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: mouse.x, y: mouse.y });
await sleep(400);
await send('Page.startScreencast', { format: 'jpeg', quality: 90, maxWidth: W * SCALE, maxHeight: H * SCALE, everyNthFrame: 1 });
while (!frames.length) await sleep(20);

// Each scene's start, in seconds from the first frame (screencast timestamps are seconds since the epoch).
const marks = [];
let failed = null;
for (const [name, run, meta = {}] of scenes) {
  const started = Date.now();
  if (meta.chapter) chapter = meta.chapter;
  marks.push({ name, ...meta, at: started / 1000 - frames[0].t });
  try { await run(); } catch (error) { failed = name + ': ' + error.message; break; }
  console.log('scene', name.padEnd(12), ((Date.now() - started) / 1000).toFixed(1) + 's');
  if (name === LAST) break;
}
await sleep(300);
await send('Page.stopScreencast');
await sleep(500);
ws.close();
proc.kill();
server.close();

if (failed) console.log('FAILED in scene', failed);
if (errors.length) console.log('page errors:', errors.join(' | '));
console.log('frames', frames.length);
if (failed) process.exit(1);

// Constant 30 fps from timestamped frames: each frame lasts until the next one arrives.
const lines = ['ffconcat version 1.0'];
frames.forEach((frame, i) => {
  const next = frames[i + 1];
  lines.push(`file '${frame.file}'`, `duration ${(next ? next.t - frame.t : 1.0).toFixed(4)}`);
});
lines.push(`file '${frames[frames.length - 1].file}'`);
writeFileSync(join(BUILD, 'frames.ffconcat'), lines.join('\n') + '\n');
const ffmpeg = process.env.FFMPEG || '/opt/homebrew/bin/ffmpeg';
const ENCODE = ['-c:v', 'libx264', '-preset', 'slow', '-crf', '24', '-pix_fmt', 'yuv420p', '-colorspace', 'bt709', '-color_primaries', 'bt709',
  '-color_trc', 'bt709', '-movflags', '+faststart'];
const run = (args) => { const r = spawnSync(ffmpeg, ['-y', '-loglevel', 'error', ...args], { stdio: 'inherit' }); if (r.status !== 0) throw new Error('ffmpeg failed'); };
const video = join(OUT, 'tour.mp4');
run(['-f', 'concat', '-safe', '0', '-i', join(BUILD, 'frames.ffconcat'), '-vf', 'fps=30,scale=1920:1080:flags=lanczos', ...ENCODE, video]);
// Poster: the title card, a moment after it fades in.
run(['-ss', '1.2', '-i', video, '-frames:v', '1', '-q:v', '3', join(OUT, 'tour-poster.jpg')]);
const total = Number(spawnSync(ffmpeg.replace(/ffmpeg$/, 'ffprobe'), ['-v', 'error', '-show_entries', 'format=duration', '-of', 'csv=p=0', video]).stdout.toString().trim());
console.log('video', video, total.toFixed(1) + 's');
if (LAST) process.exit(0);

// One clip per Explore step: from the step's first scene to the next step's (or the outro's) first scene.
const fmt = (s) => `${Math.floor(s / 60)}:${String(Math.floor(s % 60)).padStart(2, '0')}`;
const chapters = marks.filter((m) => m.chapter);
const clips = [];
chapters.forEach((m, i) => {
  const end = (chapters[i + 1] || marks.find((x) => x.name === 'outro') || { at: total }).at;
  if (m.clip) {
    const file = `step-${m.clip}.mp4`;
    run(['-ss', m.at.toFixed(2), '-i', video, '-t', (end - m.at).toFixed(2), '-an', ...ENCODE, join(OUT, file)]);
    run(['-ss', '1.0', '-i', join(OUT, file), '-frames:v', '1', '-q:v', '3', join(OUT, `step-${m.clip}.jpg`)]);
    clips.push({ ...m, file, length: end - m.at });
  }
});

// Chapter buttons and the caption transcript for the home page.
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const md = ['<!-- Generated by tools/record-tour.mjs; don\'t edit. -->', '', '<div class="chapters" aria-label="Chapters">'];
for (const m of chapters) md.push(`<button type="button" data-t="${m.at.toFixed(1)}"><span>${fmt(m.at)}</span> ${esc(m.chapter)}</button>`);
md.push('</div>', '', '::: {.callout-note collapse="true" appearance="simple"}', '## Read the captions', '');
let last = '';
for (const t of transcript) {
  if (t.chapter && t.chapter !== last) { md.push('', `**${t.chapter}**`, ''); last = t.chapter; }
  md.push(`- ${t.text}`);
}
md.push(':::', '');
writeFileSync(join(ROOT, 'site', '_tour-chapters.md'), md.join('\n'));
for (const c of clips) console.log('clip', c.file.padEnd(16), c.length.toFixed(1) + 's');
console.log('wrote site/_tour-chapters.md');
process.exit(0);
