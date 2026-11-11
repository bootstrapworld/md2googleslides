// Times calculateFontSize() over the example decks and writes the resulting sizes as JSON.
// usage: npx tsx bench/font-fit.mts out.json   (run on two branches, then diff the JSON)
import fs from 'fs';
import path from 'path';
const root = path.resolve(import.meta.dirname, '..');
const t0 = performance.now();
const {calculateFontSize} = await import(root + '/src/layout/presentation_helpers.ts');
const {default: extractSlides} = await import(root + '/src/parser/extract_slides.ts');
const tImport = performance.now() - t0;

const files = ['2-tables.md', 'autoFittingText.md', 'customLayout.md', 'debug.md', 'example.md', 'lotsOfImages.md', 'order-of-operations/slides.mkd'].map(f => path.join(root, 'examples', f));
const css = fs.readFileSync(path.join(root, 'examples/shared.css'), 'utf8');
const texts: {rawText: string; src: string}[] = [];
for (const f of files) {
  try {
    for (const s of extractSlides(fs.readFileSync(f, 'utf8'), css) as any[]) {
      for (const t of [s.title, s.subtitle, ...(s.bodies ?? []).map((b: any) => b.text)]) if (t?.rawText?.trim()) texts.push({rawText: t.rawText, src: path.basename(f)});
    }
  } catch (e) { console.error('skip', f, String(e).slice(0, 80)); }
}
// plus a few synthetic long blocks, like a real autofit stress test
const long = 'The quick brown fox jumps over the lazy dog. '.repeat(40);
texts.push({rawText: long, src: 'synthetic'}, {rawText: Array.from({length: 25}, (_, i) => `bullet number ${i} with some words`).join('\n'), src: 'synthetic'});

const EMU = 12700;
const sizes = [{w: 640, h: 300}, {w: 640, h: 80}, {w: 300, h: 300}];
const fonts = ['Arial', 'Georgia', 'Courier New'];
const constraintsList = ['both', 'horizontal'];
const mkEl = (id: string, w: number, h: number, font: string, size: number) => ({
  objectId: id,
  size: {width: {magnitude: w * EMU, unit: 'EMU'}, height: {magnitude: h * EMU, unit: 'EMU'}},
  transform: {scaleX: 1, scaleY: 1},
  shape: {text: {textElements: [
    {paragraphMarker: {style: {lineSpacing: 115, spaceBelow: {magnitude: 4, unit: 'PT'}, spaceAbove: {magnitude: 0, unit: 'PT'}}}},
    {textRun: {content: 'x', style: {fontSize: {magnitude: size, unit: 'PT'}, weightedFontFamily: {fontFamily: font, weight: 400}}}},
  ]}},
});
function run(tag: string) {
  const out: Record<string, number> = {};
  let i = 0;
  const t = performance.now();
  for (const tx of texts) for (const sz of sizes) for (const f of fonts) for (const c of constraintsList) {
    const el = mkEl(`${tag}-${i++}`, sz.w, sz.h, f, 32);
    out[`${tx.src}|${tx.rawText.slice(0, 24).replace(/\n/g, '/')}|${sz.w}x${sz.h}|${f}|${c}|${i}`] = calculateFontSize([el], {rawText: tx.rawText, textRuns: [], listMarkers: [], big: false}, c);
  }
  return {out, ms: performance.now() - t, calls: i};
}
const cold = run('cold');                 // includes font-cache / first-use costs
const warm = [run('w1'), run('w2'), run('w3')];
const best = Math.min(...warm.map(w => w.ms));
const result = {import_ms: +tImport.toFixed(1), calls: cold.calls, texts: texts.length, cold_ms: +cold.ms.toFixed(1), warm_ms_best: +best.toFixed(1), per_call_us: +(best / cold.calls * 1000).toFixed(1), sizes: cold.out};
fs.writeFileSync(process.argv[2], JSON.stringify(result, null, 1));
console.log(`texts=${texts.length} calls=${cold.calls} import=${result.import_ms}ms cold=${result.cold_ms}ms warm(best of 3)=${result.warm_ms_best}ms (${result.per_call_us}µs/call)`);
