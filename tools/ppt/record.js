// Registra le animazioni di index.html come video (senza testi) ed estrae i testi
// con posizione e stile, per costruire una versione PowerPoint con testi modificabili.
//
// Uso: node tools/ppt/record.js <index.html> <cartella-uscita> [numeri slide...]
// Richiede: Playwright con Chromium, ffmpeg con libx264.
'use strict';
const { chromium } = require('playwright');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const FPS = 25;
const FADE = 1; // secondi di dissolvenza per chiudere il loop senza scatti
const H = ['.head .eyebrow', '.head .title'];

// mode "loop": video ripetuto all'infinito, registrato da `start` per `L` secondi.
// mode "once": video che parte da 0 e resta fermo sull'ultimo fotogramma.
// mode "still": immagine fissa.
// blocks: elementi che diventano caselle di testo in PowerPoint.
const SPECS = {
  1: { mode: 'loop', start: 1, L: 16, blocks: ['.cover-text .eyebrow', '.cover-pre', '.cover-title', { sel: '.team li', exclude: '.ini' }, '.team .ini', '.hint'] },
  2: { mode: 'loop', start: 3, L: 12, blocks: [...H, '.id-top span', '.id-photo p', '.id-grid dt', '.id-grid dd', { sel: '.stamp', stamp: true }, '.side .kicker', '.side .fact'] },
  3: { mode: 'loop', start: 2.5, L: 12, blocks: [...H, '.scale-note', '.plist li b', { sel: '.plist li', exclude: 'b' }, '.belt-label', '.cols3 h3', '.cols3 p'] },
  4: { mode: 'loop', start: 1.5, L: 16, blocks: [...H, '.foot-left', '.right-col .lead', '.meter-head .kicker', '.step span'] },
  5: { mode: 'once', dur: 18, blocks: [...H, '.dim-col .lead', '.stat p'], hide: ['[data-k="replay"]'] },
  6: { mode: 'once', dur: 25, blocks: [...H, '.layers .lead', '.layer h3', '.layer p'] },
  7: { mode: 'loop', start: 1.5, L: 12, blocks: [...H, '.mag-col .lead', '.stat .big', '.stat p', '.mag-col .small'] },
  8: { mode: 'loop', start: 4, L: 14, blocks: [...H, '.tag', '.legend li span', '.atmo-col .body', 'svg text'] },
  9: { mode: 'once', dur: 25, blocks: [...H, '.tag', '.grs-col .lead', '.facts li', '.grs-col .small'] },
  10: { mode: 'loop', start: 2, L: 16, blocks: [...H, '.sub', '.mcard h3', '.mcard p'] },
  11: { mode: 'loop', start: 2, L: 16, blocks: [...H, '.ring-col .lead', '.facts li', '.cmp p'] },
  12: { mode: 'loop', start: 2, L: 20, blocks: [...H, '.juno-col .lead', '.juno-row .kicker', '.chips li', '.when b', '.when span', '.trip span', '.box .kicker', '.box .small'] },
  13: { mode: 'loop', start: 1.5, L: 14, blocks: [...H, '.quote', '.end-col .body'] },
  14: { mode: 'still', at: 2, blocks: [...H, '.qf .kicker', '.qf .lead', '.qf .small', { sel: '.qb', box: true, click: true }] },
  15: { mode: 'loop', start: 3, L: 14, blocks: ['.thanks h2', '.thanks .lead', { sel: '.thanks .team li', exclude: '.ini' }, '.thanks .team .ini', '.credits'] },
};

// Orologio virtuale: performance.now, requestAnimationFrame e animazioni CSS
// avanzano solo quando chiamiamo __step(ms), quindi ogni fotogramma è preciso.
const SHIM = `(() => {
  let vt = 0; const q = [];
  performance.now = () => vt;
  window.requestAnimationFrame = cb => { q.push(cb); return q.length; };
  window.cancelAnimationFrame = () => {};
  window.__step = dt => {
    vt += dt;
    for (const a of document.getAnimations()) {
      if (a.__v0 === undefined) a.__v0 = vt - dt;
      a.pause(); a.currentTime = Math.max(0, vt - a.__v0);
    }
    const cbs = q.splice(0, q.length);
    for (const cb of cbs) { try { cb(vt); } catch (e) { console.error(e && e.stack || e); } }
  };
})();`;

// Eseguita nella pagina: segna gli elementi di testo e ne legge testo, stile e posizione.
function extract(specs) {
  const slide = document.querySelector('.slide.active');
  const stageR = document.getElementById('stage').getBoundingClientRect();
  const px = v => parseFloat(v) || 0;
  const rel = r => ({ x: r.left - stageR.left, y: r.top - stageR.top, w: r.width, h: r.height });
  const isSvg = e => e instanceof SVGElement;
  function delayOf(el) {
    for (let e = el; e && e !== slide; e = e.parentElement) {
      const cs = getComputedStyle(e);
      if (cs.animationName && cs.animationName !== 'none') return px(cs.animationDelay);
    }
    for (const e of el.querySelectorAll('*')) {
      const cs = getComputedStyle(e);
      if (cs.animationName && cs.animationName !== 'none') return px(cs.animationDelay);
    }
    return 0;
  }
  function runOf(e, t) {
    const cs = getComputedStyle(e);
    return {
      t, size: px(cs.fontSize), weight: parseInt(cs.fontWeight, 10) || 400, italic: cs.fontStyle === 'italic',
      color: isSvg(e) ? cs.fill : cs.color, fam: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(),
      ls: cs.letterSpacing === 'normal' ? 0 : px(cs.letterSpacing), upper: cs.textTransform === 'uppercase',
      banded: !!(e.closest && e.closest('.banded')), sup: cs.verticalAlign === 'super', node: null,
    };
  }
  function paragraphs(root, exclude) {
    const P = []; let cur = null;
    const blockish = e => { const d = getComputedStyle(e).display; return !/^(inline|inline-block|inline-flex|inline-grid|contents)$/.test(d); };
    (function walk(node, pel) {
      const ncs = node.nodeType === 1 ? getComputedStyle(node) : null;
      const rowFlex = !!ncs && /flex/.test(ncs.display) && ncs.flexDirection.indexOf('row') === 0;
      for (const ch of node.childNodes) {
        if (ch.nodeType === 3) {
          const t = ch.textContent.replace(/\s+/g, ' ');
          if (!t) continue;
          if (!cur) { cur = { runs: [], el: pel, nodes: [] }; P.push(cur); }
          const r = runOf(ch.parentElement, t); cur.runs.push(r); cur.nodes.push(ch);
        } else if (ch.nodeType === 1) {
          if (exclude && ch.matches(exclude)) continue;
          if (ch.tagName === 'BR') { if (cur) cur.runs.push({ br: true }); continue; }
          if (getComputedStyle(ch).display === 'none') continue;
          if (!isSvg(ch) && blockish(ch) && !rowFlex) { cur = null; walk(ch, ch); cur = null; } else walk(ch, pel);
        }
      }
    })(root, root);
    const out = [];
    for (const p of P) {
      const runs = p.runs;
      while (runs.length && !runs[0].br && !runs[0].t.trim()) runs.shift();
      while (runs.length && !runs[runs.length - 1].br && !runs[runs.length - 1].t.trim()) runs.pop();
      if (!runs.length) continue;
      if (!runs[0].br) runs[0].t = runs[0].t.replace(/^\s+/, '');
      const last = runs[runs.length - 1]; if (!last.br) last.t = last.t.replace(/\s+$/, '');
      let L = 1e9, T = 1e9, R = -1e9, B = -1e9; const tops = new Set();
      for (const n of p.nodes) {
        const rg = document.createRange(); rg.selectNodeContents(n);
        for (const r of rg.getClientRects()) {
          if (r.width < 0.5) continue;
          L = Math.min(L, r.left); T = Math.min(T, r.top); R = Math.max(R, r.right); B = Math.max(B, r.bottom);
          tops.add(Math.round(r.top / 4));
        }
      }
      const cs = getComputedStyle(p.el);
      const size = runs.find(r => !r.br).size;
      let lh = cs.lineHeight === 'normal' ? size * 1.2 : px(cs.lineHeight);
      let align = cs.textAlign;
      if (isSvg(p.el)) { lh = size * 1.2; align = cs.textAnchor === 'middle' ? 'center' : 'left'; }
      if (align === 'start') align = 'left';
      if (align === 'end') align = 'right';
      out.push({
        runs: runs.map(r => { const c = Object.assign({}, r); delete c.node; return c; }),
        align, lh, lines: Math.max(1, tops.size),
        text: { x: L - stageR.left, y: T - stageR.top, r: R - stageR.left, b: B - stageR.top },
      });
    }
    return out;
  }
  const blocks = [];
  for (const spec0 of specs) {
    const spec = typeof spec0 === 'string' ? { sel: spec0 } : spec0;
    slide.querySelectorAll(spec.sel).forEach(el => {
      const cs = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      const paras = paragraphs(el, spec.exclude);
      if (!paras.length) return;
      el.setAttribute('data-ppt', spec.exclude ? 'x' : '1');
      if (spec.exclude) el.querySelectorAll(spec.exclude).forEach(c => c.setAttribute('data-ppt-keep', '1'));
      const b = {
        sel: spec.sel, delay: delayOf(el), click: !!spec.click, paras,
        rect: rel(r), rot: 0, box: null, stamp: !!spec.stamp,
        pad: [px(cs.paddingTop), px(cs.paddingRight), px(cs.paddingBottom), px(cs.paddingLeft)],
        content: null,
      };
      if (!isSvg(el)) {
        const w = el.offsetWidth, h = el.offsetHeight, cx = r.left + r.width / 2 - stageR.left, cy = r.top + r.height / 2 - stageR.top;
        b.rect = { x: cx - w / 2, y: cy - h / 2, w, h };
        const m = cs.transform && cs.transform !== 'none' ? cs.transform.match(/matrix\(([^)]+)\)/) : null;
        if (m) { const v = m[1].split(',').map(Number); b.rot = Math.atan2(v[1], v[0]) * 180 / Math.PI; }
        b.content = { x: b.rect.x + b.pad[3] + px(cs.borderLeftWidth), r: b.rect.x + w - b.pad[1] - px(cs.borderRightWidth) };
      }
      if (spec.box) {
        b.box = { fill: cs.backgroundColor, radius: px(cs.borderTopLeftRadius), line: cs.borderTopColor, lineW: px(cs.borderTopWidth) };
      }
      blocks.push(b);
    });
  }
  const notes = slide.querySelector('.notes');
  return { blocks, notes: notes ? notes.textContent.trim() : '' };
}

const GHOST_CSS = `
[data-ppt="1"], [data-ppt="1"] *, [data-ppt="x"] { color: transparent !important; -webkit-text-fill-color: transparent !important; text-shadow: none !important; }
[data-ppt="x"] > :not([data-ppt-keep]) { color: transparent !important; -webkit-text-fill-color: transparent !important; }
[data-ppt="1"].banded, [data-ppt="1"] .banded { background: none !important; }
svg text[data-ppt] { fill: transparent !important; }
body { padding: 0 !important; }
#ui, #progress { display: none !important; }`;

async function openSlide(browser, file, n) {
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 1 });
  page.on('pageerror', e => console.error('slide ' + n + ' pageerror: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') console.error('slide ' + n + ' console: ' + m.text()); });
  await page.addInitScript(SHIM);
  await page.goto(file + '#slide-' + n);
  await page.addStyleTag({ content: 'body{padding:0!important}#ui,#progress{display:none!important}' });
  await page.evaluate(() => window.dispatchEvent(new Event('resize')));
  await page.evaluate(() => window.__step(0));
  return page;
}
async function advance(page, ms, chunk) {
  while (ms > 0) { const d = Math.min(ms, chunk); await page.evaluate(x => window.__step(x), d); ms -= d; }
}
function encoder(out, crf) {
  const p = spawn('ffmpeg', ['-hide_banner', '-loglevel', 'error', '-y', '-f', 'image2pipe', '-framerate', String(FPS), '-i', '-',
    '-c:v', 'libx264', '-preset', 'slow', '-crf', String(crf), '-pix_fmt', 'yuv420p', '-movflags', '+faststart', out], { stdio: ['pipe', 'inherit', 'inherit'] });
  const done = new Promise((res, rej) => p.on('close', c => c === 0 ? res() : rej(new Error('ffmpeg exit ' + c))));
  return { write: buf => new Promise(r => p.stdin.write(buf) ? r() : p.stdin.once('drain', r)), end: () => { p.stdin.end(); return done; } };
}

async function doSlide(browser, file, out, n) {
  const spec = SPECS[n];
  const t0 = Date.now();
  // 1) estrazione dei testi a animazioni d'ingresso concluse
  let page = await openSlide(browser, file, n);
  await advance(page, 6000, 500);
  const data = await page.evaluate(extract, spec.blocks);
  await page.close();
  if (process.env.EXTRACT_ONLY) { // aggiorna solo i testi, tiene il video già registrato
    const f = path.join(out, 's' + String(n).padStart(2, '0') + '.json');
    const meta = JSON.parse(fs.readFileSync(f, 'utf8'));
    meta.blocks = data.blocks; meta.notes = data.notes;
    fs.writeFileSync(f, JSON.stringify(meta, null, 1));
    console.log('slide ' + n + ': testi aggiornati');
    return;
  }
  // 2) registrazione senza testi
  page = await openSlide(browser, file, n);
  await page.evaluate(([specs, css, hide]) => {
    const slide = document.querySelector('.slide.active');
    for (const s0 of specs) {
      const s = typeof s0 === 'string' ? { sel: s0 } : s0;
      slide.querySelectorAll(s.sel).forEach(el => {
        el.setAttribute('data-ppt', s.exclude ? 'x' : '1');
        if (s.exclude) el.querySelectorAll(s.exclude).forEach(c => c.setAttribute('data-ppt-keep', '1'));
      });
    }
    (hide || []).forEach(sel => slide.querySelectorAll(sel).forEach(el => { el.style.visibility = 'hidden'; }));
    const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
  }, [spec.blocks, GHOST_CSS, spec.hide]);
  const shot = () => page.screenshot({ type: 'jpeg', quality: 93 });
  const base = path.join(out, 's' + String(n).padStart(2, '0'));
  const frameMs = 1000 / FPS;
  const meta = { n, mode: spec.mode, notes: data.notes, blocks: data.blocks };
  if (spec.mode === 'still') {
    await advance(page, spec.at * 1000, 250);
    fs.writeFileSync(base + '.jpg', await shot());
    meta.image = path.basename(base + '.jpg');
  } else {
    const startMs = spec.mode === 'loop' ? spec.start * 1000 : 0;
    const durS = spec.mode === 'loop' ? spec.L + FADE : spec.dur;
    await advance(page, startMs, 250);
    const raw = spec.mode === 'loop' ? base + '.raw.mp4' : base + '.mp4';
    const enc = encoder(raw, spec.mode === 'loop' ? 14 : 23);
    const frames = Math.round(durS * FPS);
    for (let f = 0; f < frames; f++) {
      if (f) await page.evaluate(x => window.__step(x), frameMs);
      await enc.write(await shot());
    }
    await enc.end();
    if (spec.mode === 'loop') {
      const L = spec.L, F = FADE;
      await new Promise((res, rej) => {
        const p = spawn('ffmpeg', ['-hide_banner', '-loglevel', 'error', '-y', '-i', raw, '-filter_complex',
          `[0]split=3[a][b][c];[a]trim=start=${L}:end=${L + F},setpts=PTS-STARTPTS[e];[b]trim=start=0:end=${F},setpts=PTS-STARTPTS[s];` +
          `[e][s]xfade=transition=fade:duration=${F}:offset=0[x];[c]trim=start=${F}:end=${L},setpts=PTS-STARTPTS[m];[x][m]concat=n=2:v=1:a=0[o]`,
          '-map', '[o]', '-c:v', 'libx264', '-preset', 'slow', '-crf', '23', '-pix_fmt', 'yuv420p', '-movflags', '+faststart', base + '.mp4'], { stdio: 'inherit' });
        p.on('close', c => c === 0 ? res() : rej(new Error('xfade ' + c)));
      });
      fs.unlinkSync(raw);
    }
    meta.video = path.basename(base + '.mp4');
    meta.durMs = Math.round((spec.mode === 'loop' ? spec.L : spec.dur) * 1000);
    meta.loop = spec.mode === 'loop';
  }
  await page.close();
  fs.writeFileSync(base + '.json', JSON.stringify(meta, null, 1));
  console.log('slide ' + n + ' fatta in ' + Math.round((Date.now() - t0) / 1000) + ' s');
}

(async () => {
  const [file, out, ...which] = process.argv.slice(2);
  fs.mkdirSync(out, { recursive: true });
  const list = which.length ? which.map(Number) : Object.keys(SPECS).map(Number);
  const workers = +(process.env.WORKERS || 2);
  const browsers = await Promise.all(Array.from({ length: workers }, () => chromium.launch({
    args: ['--use-gl=angle', '--use-angle=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist'] })));
  const url = 'file://' + path.resolve(file);
  let i = 0;
  await Promise.all(browsers.map(async b => {
    while (i < list.length) { const n = list[i++]; await doSlide(b, url, out, n); }
    await b.close();
  }));
})().catch(e => { console.error(e); process.exit(1); });
