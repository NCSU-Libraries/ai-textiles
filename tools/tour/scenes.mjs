// Scenes for tools/record-tour.mjs: the Explore page's eight steps, in order.
//
// Prompts come from site/explore.qmd (prompts[0] is step 3's), numbers from site/_variables.yml, and the sheet's rows
// from the practice CSV. Gemini's answers follow the spike in instructor/dry-run.md, shortened to fit; the screenshots
// are the real ones from site/images/screens. Each scene is [name, run, { chapter, clip }]: a chapter starts a button
// on the home page, and a clip becomes site/video/step-<clip>.mp4 for that step on the Explore page.
export default ({ S, js, sleep, moveTo, toCell, v, prompts, data, onCaption }) => {
  const n = (k) => v(k).replace(/,/g, '');
  const last = n('data.last-row-n');
  const cap = (t) => { onCaption(t); return S(`stage.caption(${js(t)})`); };
  const step = (t) => S(`stage.step(${js(t)})`);
  const layout = (l) => S(`stage.layout(${js(l)})`);
  const cursor = (on) => S(`document.body.classList.toggle('carded', ${!on})`);
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const f = (formula) => `<span class="f">${esc(formula)}</span>`;

  // Gemini
  const gem = {
    type: async (text, delay = 12) => { await S(`await stage.gem.type(${js(text)}, ${delay})`); await sleep(400); await S('await stage.gem.send()'); },
    paste: async (text) => { await S(`await stage.gem.paste(${js(text)})`); await sleep(900); await S('await stage.gem.send()'); },
    answer: (blocks, gap = 380, think = 1100) => S(`await stage.gem.answer(${js(blocks)}, ${gap}, ${think})`),
    list: (items, gap = 520) => S(`await stage.gem.list(${js(items)}, ${gap})`),
    mark: (i, on = true) => S(`stage.gem.mark(${i}, ${on})`),
  };

  // Sheets
  const L = 'ABCDEFGHIJKLM';
  const NUMERIC = new Set([3, 4, 8, 9, 10, 11, 12]);
  const date = (s) => { const [y, m, d] = s.split('-'); return `${+m}/${+d}/${y}`; };
  // Cells for sheet rows first..first+count-1 of the Data tab (row 1 is the header).
  function dataCells(first, count) {
    const cells = {};
    for (let r = first; r < first + count; r++) {
      const row = data.rows[r - 1]; if (!row) continue;
      row.forEach((val, c) => {
        const isNum = r > 1 && NUMERIC.has(c) && val !== '' && !isNaN(Number(val));
        cells[L[c] + r] = { v: r > 1 && c === 0 ? date(val) : val, cls: isNum ? 'n' : '' };
      });
    }
    return cells;
  }
  const sheet = {
    tab: (name) => S(`stage.sheet.tab(${js(name)})`),
    setup: (name, cfg) => S(`stage.sheet.setup(${js(name)}, ${js(cfg)})`),
    set: (ref, val, cls = '', tab) => S(`stage.sheet.set(${js(ref)}, ${js(val)}, ${js(cls)}${tab ? ', ' + js(tab) : ''})`),
    select: (ref, formula) => S(`stage.sheet.select(${js(ref)}${formula !== undefined ? ', ' + js(formula) : ''})`),
    hl: (refs) => S(`stage.sheet.hl(${js(refs)})`),
    formula: async (ref, formula, result, cls = 'n', delay = 6) => {
      await toCell(ref, 450);
      await S(`await stage.sheet.formula(${js(ref)}, ${js(formula)}, ${js(result)}, ${js(cls)}, ${delay})`);
    },
  };
  const DATA_WIDTHS = [78, 118, 70, 46, 46, 74, 150, 78, 46, 46, 46, 66, 150];
  const showData = (first) => sheet.setup('Data', { cells: { ...dataCells(1, 1), ...dataCells(first, 22) }, widths: DATA_WIDTHS, first, rows: 22,
    headerRow: first === 1 ? 1 : 0 });

  // Numbers the sheet shows, computed from the data the way Sheets would.
  const general = (x) => String(Number(x.toPrecision(10)));
  const rows = data.rows.slice(1);
  const groups = {};
  for (const r of rows) if (r[5]) (groups[r[5]] ||= []).push(Number(r[10]));
  const summary = Object.keys(groups).sort().map((k) => [k, groups[k].length, general(groups[k].reduce((a, b) => a + b, 0) / groups[k].length)]);
  const sj = summary.find(([k]) => k === 'S/J')[2];
  const rejects = rows.map((r) => r[12]).filter((x) => x !== '' && !isNaN(Number(x))).map(Number);
  const rejectSum = rejects.reduce((a, b) => a + b, 0);
  const rejectMean = general(rejectSum / rejects.length);
  const rejectZero = general(rejectSum / rows.length);
  // Lots that repeat within the first screen of rows, to point at.
  const seen = {}, repeats = [];
  for (let r = 2; r <= 23; r++) { const lot = data.cell(r, 1); (seen[lot] ||= []).push('B' + r); }
  for (const refs of Object.values(seen)) if (refs.length > 1) repeats.push(...refs);

  return [
    ['intro', async () => {
      await cursor(false);
      await showData(1);
      await S(`stage.card('AI-Assisted Data Exploration for Research', 'Describe your data to Gemini, run its formulas in Google Sheets, and check every number.', { points: ['Look', 'Describe', 'Ask', 'Check'], instant: true })`);
      await sleep(3800);
      await S('stage.card(null)');
      await sleep(700);
    }],

    ['look', async () => {
      await step('1 · Look first');
      await cursor(true);
      await cap('Your own copy of the practice sheet. No AI yet.');
      await sleep(2200);
      await toCell('A2');
      await sheet.select('A2');
      await cap('Click A2, then press Cmd + ↓ to jump to the last row');
      await sleep(1800);
      await showData(Number(last) - 20);
      await sheet.select('A' + last);
      await toCell('A' + last, 500);
      await S(`stage.answer('Rows of data', ${js(v('data.rows'))})`);
      await cap(`Row ${v('data.last-row')}, minus the header: ${v('data.rows')} rows`);
      await sleep(3000);
      await S('stage.answer(null)');
      await showData(1);
      await sheet.hl(repeats);
      await cap('One row is one lot knitted on one day, so lots repeat');
      await sleep(3200);
      await sheet.hl([]);
    }, { chapter: 'Look first', clip: '1' }],

    ['security', async () => {
      await step('2 · Data security');
      await S('stage.gem.wide(true)');
      await layout('gem');
      await sleep(800);
      await S(`stage.gem.highlight('acct', true)`);
      await cap('In Gemini, look for Edu under your name, at the bottom left');
      await sleep(3000);
      await S(`stage.gem.highlight('acct', false)`);
      await S(`stage.gem.highlight('disc', true)`);
      await cap('With your NC State login, chats aren’t used to improve the models');
      await sleep(3200);
      await S(`stage.gem.highlight('disc', false)`);
      await layout('browser');
      await S(`await stage.browser.load('site/data-security.html', ${js(v('site-short') + '/data-security.html')}, 0.85)`);
      await S(`await stage.browser.scrollTo('#guide', 1400)`);
      await cap('For your own data, the decision guide says what you can share');
      await sleep(3400);
      await S(`stage.note('<span class="from">Today</span>You describe the data. Gemini never sees a real row.')`);
      await cap('Today you won’t upload anything');
      await sleep(3000);
      await S('stage.note(null)');
    }, { chapter: 'Data security', clip: '2' }],

    ['describe', async () => {
      await step('3 · Describe the data');
      await S('stage.gem.wide(false)');
      await layout('split');
      await cap('Start a new chat and paste the prompt from the Explore page');
      await sleep(900);
      await gem.paste(prompts[0]);
      await cap('The layout, the header row, 5 made-up rows, and notes');
      await sleep(2600);
      await gem.answer(["<p>Understood. I'm ready for your first formula request.</p>"]);
      await cap('Gemini knows the layout, but not one real row');
      await sleep(2800);
    }, { chapter: 'Describe the data', clip: '3' }],

    ['problems', async () => {
      await step('4 · Find the problems');
      await sheet.setup('Work', { cells: { A1: { v: 'Put your formulas on this tab, starting in A3, so their results never spill onto the Data tab.', cls: 'note spill' } },
        widths: [96, 150, 20, 100, 80, 104], first: 1, rows: 22 });
      await sheet.tab('Work');
      await cap('Ask for checks, each with a formula you can run');
      await gem.type(prompts[1]);
      await gem.answer([]);
      await gem.list([
        `<b>Non-numeric text in numeric fields</b> (e.g., <code>"any"</code> in Fabric Diameter)${f(`=SUMPRODUCT((Data!J2:J${last}<>"")*NOT(ISNUMBER(Data!J2:J${last})))`)}`,
        `<b>Missing values in numeric fields</b> (e.g., blank Reject Qty)${f(`=COUNTBLANK(Data!M2:M${last})`)}`,
        `<b>Dates stored as text</b>${f(`=SUMPRODUCT((Data!A2:A${last}<>"")*NOT(ISNUMBER(Data!A2:A${last})))`)}`,
        `<b>Duplicate Lot IDs</b>${f(`=ROWS(Data!B2:B${last})-COUNTA(UNIQUE(Data!B2:B${last}))`)}`,
        `<b>Malformed text formatting</b> (e.g., trailing <code>% %</code> in Composition)${f(`=COUNTIF(Data!G2:G${last}, "*% %")`)}`,
        `<b>Invalid or zero production quantities</b>${f(`=COUNTIF(Data!L2:L${last}, "<=0")`)}`,
      ], 480);
      await cap('Paste each formula into the Work tab, from A3 down');
      const checks = [
        [`=SUMPRODUCT((Data!J2:J${last}<>"")*NOT(ISNUMBER(Data!J2:J${last})))`, n('key.any-rows'), 'text in Fabric Diameter'],
        [`=COUNTBLANK(Data!M2:M${last})`, n('key.reject-blank'), 'blank Reject Qty'],
        [`=SUMPRODUCT((Data!A2:A${last}<>"")*NOT(ISNUMBER(Data!A2:A${last})))`, '0', 'dates as text'],
        [`=ROWS(Data!B2:B${last})-COUNTA(UNIQUE(Data!B2:B${last}))`, n('key.lots-repeat'), 'repeated Lot IDs'],
        [`=COUNTIF(Data!G2:G${last}, "*% %")`, '13', 'Composition ending "% %"'],
        [`=COUNTIF(Data!L2:L${last}, "<=0")`, n('key.qty-zero'), 'Production Qty of 0 or less'],
      ];
      for (const [i, [formula, result, label]] of checks.entries()) {
        await sheet.formula('A' + (3 + i), formula, result);
        await sheet.set('B' + (3 + i), label, 'note');
        await sleep(250);
      }
      await gem.mark(3);
      await sheet.hl(['A6']);
      await cap(`${v('key.lots-repeat')} “duplicate” lots? Lots run for days, so that’s expected`);
      await sleep(3600);
      await gem.mark(3, false);
      await sheet.hl([]);
      await cap(`The Explore page lists ${v('key.problems')} problems. These formulas find 4 it missed.`);
      const missed = [
        [`=ROWS(Data!A2:M${last})-ROWS(UNIQUE(Data!A2:M${last}))`, n('key.repeated-rows'), 'exact copies of a row'],
        [`=COUNTIF(Data!B2:B${last},"nan")`, n('key.lot-nan'), 'Lot says "nan"'],
        [`=MAX(Data!E2:E${last})`, n('key.gauge-max'), 'a gauge of 2,468?'],
        [`=SUMPRODUCT((Data!M2:M${last}<>"")*NOT(ISNUMBER(Data!M2:M${last})))`, '1', 'text in Reject Qty'],
      ];
      for (const [i, [formula, result, label]] of missed.entries()) {
        await sheet.formula('A' + (10 + i), formula, result);
        await sheet.set('B' + (10 + i), label, 'note');
        await sleep(450);
      }
      await sleep(1200);
      const row = Number(n('key.reject-note-row'));
      await showData(row - 8);
      await sheet.tab('Data');
      await layout('sheet');
      await sleep(700);
      await sheet.select('M' + row);
      await sheet.hl(['M' + row]);
      await toCell('M' + row, 600);
      await cap(`Row ${v('key.reject-note-row')} holds a note where a number belongs`);
      await sleep(3400);
      await sheet.hl([]);
      await layout('split');
    }, { chapter: 'Find the problems', clip: '4' }],

    ['summary', async () => {
      await step('5 · Summarize, with a chart');
      await sheet.tab('Work');
      await cap('Ask for a summary table, as a formula');
      await gem.type(prompts[2]);
      const query = `=QUERY(Data!F2:K${last}, "SELECT F, COUNT(F), AVG(K) WHERE F IS NOT NULL GROUP BY F LABEL F 'Construction', COUNT(F) 'Lot Count', AVG(K) 'Mean GSM'")`;
      await gem.answer([f(query), '<p><b>How to Create the Chart</b></p>'], 500);
      await gem.list(['Select the output table, including the headers.', 'Click <b>Insert &gt; Chart</b>.', 'Chart type: <b>Column chart</b>.',
        'Series: remove <b>Lot Count</b>, leaving only Mean GSM.'], 420);
      await sheet.formula('D3', query, 'Construction', '', 4);
      for (const [ref, val] of [['E3', 'Lot Count'], ['F3', 'Mean GSM']]) await sheet.set(ref, val);
      for (const [i, [k, count, mean]] of summary.entries()) {
        await sheet.set('D' + (4 + i), k);
        await sheet.set('E' + (4 + i), String(count), 'n');
        await sheet.set('F' + (4 + i), mean, 'n');
      }
      await cap('One formula writes the whole table');
      await sleep(2200);
      await sheet.hl(['E3']);
      await cap('“Lot Count” counts rows, not lots');
      await sleep(2800);
      await sheet.hl([]);
      await cursor(false);
      await S(`await stage.shot('From our rehearsal', 'shots/sheets-summary.png', 540)`);
      await cap('Insert → Chart, then remove the count series');
      await sleep(3800);
      await S('stage.shot(null)');
      await cursor(true);
      await sheet.setup('Compare', { cells: {
        A1: { v: 'Check each number two ways', cls: 'title' },
        A3: { v: 'Statistic' }, B3: { v: 'First way' }, C3: { v: 'Second way' }, D3: { v: 'Compare to (decimals)' }, E3: { v: 'Match?' },
        A4: { v: 'Data rows, not counting the header' }, D4: { v: '0', cls: 'n' },
        A5: { v: 'Mean GSM for single jersey (S/J)' }, D5: { v: '1', cls: 'n' },
        A6: { v: 'Mean Reject Qty' }, D6: { v: '2', cls: 'n' },
        A7: { v: 'Your own:' }, D7: { v: '2', cls: 'n' }, A8: { v: 'Your own:' }, D8: { v: '2', cls: 'n' },
      }, widths: [214, 96, 96, 130, 64], first: 1, rows: 22, headerRow: 3 });
      await sheet.tab('Compare');
      await cap('Write the S/J mean in Compare, under First way');
      await toCell('B5');
      await sheet.select('B5', '');
      await sleep(300);
      await sheet.set('B5', sj, 'n');
      await sleep(1800);
    }, { chapter: 'Summarize, with a chart', clip: '5' }],

    ['check', async () => {
      await step('6 · Check it a second way');
      await cap('Ask for a different formula for the same number');
      await gem.type(prompts[3]);
      await gem.answer([f(`=AVERAGEIF(Data!F2:F${last}, "S/J", Data!K2:K${last})`), '<p>Alternatively, using FILTER:</p>',
        f(`=AVERAGE(FILTER(Data!K2:K${last}, Data!F2:F${last} = "S/J"))`)]);
      await sheet.formula('C5', `=AVERAGEIF(Data!F2:F${last}, "S/J", Data!K2:K${last})`, sj);
      await sheet.set('E5', 'match', 'match');
      await cap('Both formulas give the same number, so Match? says match');
      await sleep(2600);
      await cap('Do the same for the row count from step 1');
      await toCell('B4');
      await sheet.select('B4', '');
      await sheet.set('B4', n('data.rows'), 'n');
      await sheet.formula('C4', '=COUNTA(Data!A2:A)', n('data.rows'));
      await sheet.set('E4', 'match', 'match');
      await sleep(1800);
      await cursor(false);
      await S(`await stage.shot('From our rehearsal: a pivot table, a third way', 'shots/sheets-pivot.png', 1000)`);
      await cap('Making a pivot table? Untick Show totals before you chart it.');
      await sleep(4400);
      await S('stage.shot(null)');
      await cursor(true);
    }, { chapter: 'Check it a second way', clip: '6' }],

    ['blanks', async () => {
      await step('7 · What does a blank mean?');
      await cap('Reject Qty has blanks. Ask what to do with them.');
      await gem.type(prompts[4], 16);
      await gem.answer([
        '<p><b>No, blank cells should generally not count as zero.</b></p>',
        '<p>Since 0 is explicitly recorded in some of your rows, a blank cell strongly suggests missing or unrecorded data, not zero defects.</p>',
        '<p>Ignores blanks:</p>', f(`=AVERAGE(Data!M2:M${last})`),
        '<p>If you confirm blanks mean 0 rejects:</p>', f(`=SUM(Data!M2:M${last}) / ROWS(Data!M2:M${last})`),
      ], 450);
      await sheet.formula('B6', `=AVERAGE(Data!M2:M${last})`, rejectMean);
      await sheet.formula('C6', `=SUM(Data!M2:M${last}) / ROWS(Data!M2:M${last})`, rejectZero);
      await sheet.set('E6', 'differs', 'differs');
      await sheet.hl(['B6', 'C6']);
      await cap('They differ on purpose. One skips blanks, one counts them as zero.');
      await sleep(3400);
      await cap('Which would you report? Ask whoever recorded the data.');
      await sleep(3200);
      await sheet.hl([]);
    }, { chapter: 'What a blank means', clip: '7' }],

    ['own', async () => {
      await step('8 · Your own data');
      await cursor(false);
      await layout('browser');
      await S(`await stage.browser.load('site/practices.html', ${js(v('site-short') + '/practices.html')}, 0.85)`);
      await cap('Recommended practices and prompts, on one page');
      await sleep(2600);
      await S(`await stage.browser.scrollTo('#own-data', 1400)`);
      await cap('Then plan the same steps for your own data');
      await sleep(3800);
    }, { chapter: 'Your own data', clip: '8' }],

    ['outro', async () => {
      await S('stage.caption(null); stage.step(null)');
      await S(`stage.card('What you’ll leave with', null, { list: ['Your own copy of the practice sheet', 'The data’s problems, found with formulas Gemini wrote', 'A summary table and a chart', 'One number worked out two ways, and a reason for any difference', 'Knowing what’s safe to share, and why your account matters', 'A plan for your own data'], url: ${js(v('site-short'))} })`);
      await sleep(5600);
    }],
  ];
};
