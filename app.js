(() => {
  'use strict';

  const APP_VERSION = '1.0.0';
  const BUILTIN_TEMPLATE_NAME = '水务公司运行日报表（2026年11月）.xls';
  const BUILTIN_TEMPLATE_URL = 'templates/水务公司运行日报表（2026年11月）.xls';
  const DB_NAME = 'water-report-mobile';
  const DB_STORE = 'settings';
  const TEMPLATE_KEY = 'savedTemplate';
  const DRAFT_KEY = 'waterReportRawDraft';

  const state = {
    parsed: null,
    savedTemplate: null,
    excelBlob: null,
    pngBlob: null,
    text: '',
    pngObjectUrl: '',
    deferredInstallPrompt: null,
  };

  const $ = (id) => document.getElementById(id);
  const rawTextEl = $('rawText');
  const messageEl = $('message');
  const previewCard = $('previewCard');
  const resultCard = $('resultCard');

  const escapeHtml = (s) => String(s)
    .replaceAll('&', '&amp;').replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;').replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');

  function showMessage(text, type='ok') {
    messageEl.textContent = text;
    messageEl.className = `message ${type}`;
  }
  function hideMessage() {
    messageEl.className = 'message hidden';
    messageEl.textContent = '';
  }

  function normalizeRaw(raw) {
    return String(raw || '')
      .replace(/\u3000/g, ' ')
      .replace(/\u00a0/g, ' ')
      .replace(/\u202f/g, ' ')
      .replace(/[\u200b-\u200f\u2060\ufeff]/g, '')
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/[ \t]+/g, ' ')
      .replace(/ *\n */g, '\n')
      .trim();
  }

  function findPlantHeader(text, plant) {
    const escaped = plant.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const patterns = [
      new RegExp(`${escaped}(?:污水处理厂|再生水厂)?\\s*[:：-]?\\s*(?:(?:20\\d{2})年\\s*)?(\\d{1,2})\\s*月\\s*(\\d{1,2})\\s*日?\\s*(?:交接班(?:情况)?|交班情况|交接情况)\\s*[:：]?`, 'i'),
      new RegExp(`${escaped}[^\\n]{0,24}?(\\d{1,2})\\s*月\\s*(\\d{1,2})\\s*日?[^\\n]{0,16}?(?:交接班(?:情况)?|交班情况|交接情况)\\s*[:：]?`, 'i'),
    ];
    for (const re of patterns) {
      const m = re.exec(text);
      if (m) return m;
    }
    throw new Error(`未识别到【${plant}交接班日期】。请直接整段复制微信原文，无需删姓名或时间。`);
  }

  function must(re, text, label) {
    const m = re.exec(text);
    if (!m) throw new Error(`未识别到【${label}】。请检查微信原文是否完整复制。`);
    return m;
  }

  function normalizeTimeLabel(label) {
    return String(label)
      .replace(/\s+/g, '')
      .replace(/(\d{1,2})时/g, (_, h) => `${String(Number(h)).padStart(2, '0')}时`);
  }

  function standardizeQualityItem(item, plantName, idx) {
    const compact = String(item).replace(/\s+/g, ' ').trim();
    const metricM = /进水在线监测\s*(COD|总磷|总氮|氨氮)\s*[:：]?/i.exec(compact);
    if (!metricM) throw new Error(`${plantName}第${idx}项水质数据未识别到指标名称。`);
    const metric = metricM[1].toUpperCase() === 'COD' ? 'COD' : metricM[1];

    const avgM = /平均\s*值\s*([0-9]+(?:\.[0-9]+)?)\s*mg\/L/i.exec(compact);
    if (!avgM) throw new Error(`${plantName}${metric}未识别到平均值。`);

    const timePat = '(\\d{1,2}\\s*时(?:\\s*[-—~至]\\s*\\d{1,2}\\s*时)?)';
    const maxM = new RegExp(`最大\\s*(?:值)?\\s*${timePat}\\s*([0-9]+(?:\\.[0-9]+)?)\\s*mg/L`, 'i').exec(compact);
    const minM = new RegExp(`最小\\s*(?:值)?\\s*${timePat}\\s*([0-9]+(?:\\.[0-9]+)?)\\s*mg/L`, 'i').exec(compact);
    if (!maxM) throw new Error(`${plantName}${metric}未识别到最大值。`);
    if (!minM) throw new Error(`${plantName}${metric}未识别到最小值。`);

    const designM = /设计\s*值\s*[≤<]\s*=?\s*([0-9]+(?:\.[0-9]+)?)\s*mg(?:\/L)?/i.exec(compact);
    if (!designM) throw new Error(`${plantName}${metric}未识别到设计值。`);

    const metricEsc = metric.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const outM = new RegExp(`在线监测出水\\s*${metricEsc}\\s*日平均\\s*值(?:为)?\\s*([0-9]+(?:\\.[0-9]+)?)\\s*mg/L\\s*[，,]?\\s*(达标|不达标|超标)?`, 'i').exec(compact);
    if (!outM) throw new Error(`${plantName}${metric}未识别到出水日平均值。`);

    const avg = avgM[1];
    const maxTime = normalizeTimeLabel(maxM[1]);
    const maxValue = maxM[2];
    const minTime = normalizeTimeLabel(minM[1]);
    const minValue = minM[2];
    const design = designM[1];
    const outAvg = outM[1];
    const outStatus = outM[2] || '达标';

    const exceedM = /(?:超标数据为|超标数据|超标数值|数据超标值|超标值|超标为)\s*[:：]?\s*(.*?)(?=[（(]?\s*设计值)/i.exec(compact);
    let condition = '正常';
    let exceedText = '';
    if (exceedM) {
      const pairRe = new RegExp(`${timePat}\\s*([0-9]+(?:\\.[0-9]+)?)\\s*mg/L`, 'ig');
      const pairs = [];
      let p;
      while ((p = pairRe.exec(exceedM[1])) !== null) {
        pairs.push(`${normalizeTimeLabel(p[1])}${p[2]}mg/L`);
      }
      if (!pairs.length) throw new Error(`${plantName}${metric}识别到“超标数据”字样，但未识别到超标时刻/数值。`);
      exceedText = pairs.join('、');
      condition = `超标数据为：${exceedText}`;
    } else if (!compact.includes('正常')) {
      throw new Error(`${plantName}${metric}未识别到“正常”或“超标数据”状态。`);
    }

    const text = `${idx}、进水在线监测${metric}：平均值${avg}mg/L，最大值${maxTime}${maxValue}mg/L，最小值${minTime}${minValue}mg/L，${condition}（设计值≤${design}mg/L）。在线监测出水${metric}日平均值为${outAvg}mg/L，${outStatus}。`;
    return { text, metric, exceedText, status: exceedText ? '超标' : '正常' };
  }

  function extractQualityItems(block, plantName) {
    const text = normalizeRaw(block);
    const circled = ['①', '②', '③', '④'];
    const starts = [];
    for (let i = 0; i < 4; i++) {
      let pos = text.indexOf(circled[i]);
      if (pos < 0) {
        const re = new RegExp(`(?:^|\\n)\\s*${i + 1}[、.．]\\s*(?=(?:\\d{1,2}\\s*月\\s*\\d{1,2}\\s*日?\\s*)?进水在线监测)`, 'm');
        const m = re.exec(text);
        if (!m) throw new Error(`${plantName}缺少第${i + 1}项水质数据。`);
        pos = m.index;
      }
      starts.push(pos);
    }

    const items = [];
    for (let i = 0; i < 4; i++) {
      const start = starts[i];
      const end = i < 3 ? starts[i + 1] : text.length;
      let item = text.slice(start, end).trim();
      item = item.replace(/^[①②③④]\s*/, '').replace(new RegExp(`^${i + 1}[、.．]\\s*`), '');
      item = item.replace(/^\d{1,2}\s*月\s*\d{1,2}\s*日?\s*/, '').replace(/\s+/g, ' ').trim();
      const done = [...item.matchAll(/(?:不达标|达标|超标)[。.]?/g)];
      if (done.length) {
        const last = done[done.length - 1];
        item = item.slice(0, last.index + last[0].length);
      }
      items.push(standardizeQualityItem(item, plantName, i + 1));
    }
    return items;
  }

  function parseWechatText(raw) {
    const text = normalizeRaw(raw);
    if (!text) throw new Error('微信原始数据为空。');

    let southHeader = null;
    let northHeader = null;
    let southError = null;
    try { southHeader = findPlantHeader(text, '南园'); } catch (e) { southError = e; }
    try { northHeader = findPlantHeader(text, '北园'); } catch (_) {}

    let reportMonth;
    let reportDay;
    if (!southHeader && !northHeader) {
      const matches = [...text.matchAll(/(20\d{2})年\s*(\d{1,2})月\s*(\d{1,2})日/g)];
      if (!matches.length) throw southError || new Error('未识别到南园/北园交接班日期。');
      const last = matches[matches.length - 1];
      reportMonth = Number(last[2]);
      reportDay = Number(last[3]);
    } else {
      const h = southHeader || northHeader;
      reportMonth = Number(h[1]);
      reportDay = Number(h[2]);
    }

    if (southHeader && northHeader) {
      if (Number(southHeader[1]) !== Number(northHeader[1]) || Number(southHeader[2]) !== Number(northHeader[2])) {
        throw new Error('数据校验失败：南园与北园标题日期不一致。');
      }
    }

    const yearM = /(20\d{2})年\s*\d{1,2}月\s*\d{1,2}日/.exec(text);
    const year = yearM ? Number(yearM[1]) : new Date().getFullYear();

    const one = must(/一水厂[\s\S]*?(?:昨日|昨天|咋日)?\s*供水\s*([0-9]+)\s*吨[\s\S]*?压力\s*(?:全天)?\s*([0-9.]+)\s*MPa/i, text, '一水厂供水量/压力');
    const two = must(/二水厂[\s\S]*?(?:昨日|昨天|咋日)?\s*供水\s*([0-9]+)\s*吨[\s\S]*?压力\s*(?:全天)?\s*([0-9.]+)\s*MPa/i, text, '二水厂供水量/压力');
    const oneSupply = Number(one[1]);
    const onePressure = one[2];
    const twoSupply = Number(two[1]);
    const twoPressure = two[2];

    let twoDetail = null;
    const twoDetailM = /二水厂[\s\S]*?供水\s*([0-9]+)\s*吨[\s\S]*?低压供水\s*([0-9]+)\s*吨[\s\S]*?压力\s*([0-9.]+)\s*MPa[\s\S]*?高压供水\s*([0-9]+)\s*吨[\s\S]*?压力\s*([0-9.]+)\s*MPa/i.exec(text);
    if (twoDetailM) {
      const [, total, lowSupply, lowPressure, highSupply, highPressure] = twoDetailM;
      if (Number(total) !== twoSupply) throw new Error('二水厂总供水识别不一致。');
      if (Number(lowSupply) + Number(highSupply) !== twoSupply) throw new Error('二水厂低压+高压与总供水不一致。');
      twoDetail = {
        lowSupply: Number(lowSupply),
        lowPressure,
        highSupply: Number(highSupply),
        highPressure,
      };
    }

    if (!southHeader) southHeader = findPlantHeader(text, '南园');
    if (!northHeader) northHeader = findPlantHeader(text, '北园');
    if (northHeader.index <= southHeader.index + southHeader[0].length) throw new Error('识别到的北园交接班标题位置异常。');

    const southStart = southHeader.index + southHeader[0].length;
    const northStart = northHeader.index;
    const northBodyStart = northHeader.index + northHeader[0].length;
    const southBlock = text.slice(southStart, northStart);
    const northBlock = text.slice(northBodyStart);

    const treatment = must(/南园污水处理厂(?:处量|处理)\s*([0-9]+)\s*吨[\s\S]*?一期\s*([0-9]+)\s*吨[，,]\s*二期\s*([0-9]+)\s*吨[\s\S]*?北园再生水厂处理\s*([0-9]+)\s*吨[\s\S]*?合计处理(?:污水)?\s*([0-9]+)\s*吨[\s\S]*?北盛镇日处理量\s*([0-9]+)\s*吨[\s\S]*?往北园泵站送量\s*([0-9]+)\s*吨[，,]\s*景观桥泵站\s*([0-9]+)\s*吨[\s\S]*?永安污水处理量\s*([0-9]+)\s*吨/i, southBlock, '南园/北园/永安处理量');
    const nums = treatment.slice(1).map(Number);
    const [southTotal, phase1, phase2, northTotal, statedMainTotal, beishengTotal, beishengToNorth, beishengToScenic, yonganTotal] = nums;

    const warnings = [];
    if (phase1 + phase2 !== southTotal) {
      throw new Error(`南园一期${phase1}+二期${phase2}=${phase1 + phase2}，但南园总量为${southTotal}吨。`);
    }
    if (southTotal + northTotal !== statedMainTotal) {
      throw new Error(`南园${southTotal}+北园${northTotal}=${southTotal + northTotal}，但原文合计为${statedMainTotal}吨。`);
    }
    if (beishengToNorth + beishengToScenic !== beishengTotal) {
      warnings.push(`北盛镇总量${beishengTotal}吨与分项${beishengToNorth}+${beishengToScenic}=${beishengToNorth + beishengToScenic}吨不一致；正式日报保留原始上报总量。`);
    }

    const northTreatmentM = /北园再生水厂\s*([0-9]+)\s*吨/.exec(northBlock);
    if (northTreatmentM && Number(northTreatmentM[1]) !== northTotal) {
      throw new Error('南园总报文与北园交接班报文中的北园处理量不一致。');
    }

    const southQuality = extractQualityItems(southBlock, '南园');
    const northQuality = extractQualityItems(northBlock, '北园');

    const reportDate = new Date(year, reportMonth - 1, reportDay);
    const dataDate = new Date(reportDate);
    dataDate.setDate(dataDate.getDate() - 1);

    return {
      year, reportMonth, reportDay, reportDate, dataDate,
      oneSupply, onePressure, twoSupply, twoPressure, twoDetail,
      southTotal, phase1, phase2, northTotal, statedMainTotal,
      beishengTotal, beishengToNorth, beishengToScenic, yonganTotal,
      southQuality, northQuality, warnings,
      supplyTotal: oneSupply + twoSupply,
      mainSewageTotal: southTotal + northTotal,
      fullSewageTotal: southTotal + northTotal + yonganTotal,
    };
  }

  function buildSupplyText(d) {
    const second = d.twoDetail
      ? `2、金阳二水厂：${d.twoSupply}吨（低压供水${d.twoDetail.lowSupply}吨，压力${d.twoDetail.lowPressure}MPa；高压供水${d.twoDetail.highSupply}吨，压力${d.twoDetail.highPressure}MPa）。`
      : `2、金阳二水厂：${d.twoSupply}吨，压力${d.twoPressure}MPa。`;
    return `供水总量${d.supplyTotal}吨，具体如下：\n1、金阳一水厂：${d.oneSupply}吨，压力${d.onePressure}MPa；\n${second}`;
  }

  function buildBodyText(d) {
    const southDesc = `南园污水处理厂处理${d.southTotal}吨（一期${d.phase1}吨，二期${d.phase2}吨）（设计能力为5.5万吨/天）`;
    return [
      `一、处理量：污水处理合计${d.mainSewageTotal}吨，其中：${southDesc}，北园再生水厂处理${d.northTotal}吨，北盛镇日处理量${d.beishengTotal}吨（往北园泵站送量${d.beishengToNorth}吨，景观桥泵站${d.beishengToScenic}吨）。`,
      '二、水质情况：',
      '（一）南园污水处理厂',
      ...d.southQuality.map(x => x.text),
      '',
      '（二）北园再生水厂',
      ...d.northQuality.map(x => x.text),
    ].join('\n');
  }

  function buildTextSummary(d) {
    const dt = d.dataDate;
    return `${dt.getFullYear()}年${dt.getMonth() + 1}月${dt.getDate()}日供排水情况：\n1、供水合计：${d.supplyTotal}吨\n金阳一水厂：${d.oneSupply}吨\n金阳二水厂：${d.twoSupply}吨\n2、污水处理合计：${d.fullSewageTotal}吨\n南园污水处理厂：${d.southTotal}吨\n北园再生水厂：${d.northTotal}吨\n永安污水处理厂：${d.yonganTotal}吨`;
  }

  function formatDateLine(d) {
    const dd = d.dataDate;
    const rd = d.reportDate;
    return `数据日期：${dd.getFullYear()}年${dd.getMonth() + 1}月${dd.getDate()}日　　　　　　　　　　　　　　　　　　　　　　　　　汇编日期：${rd.getFullYear()}年${rd.getMonth() + 1}月${rd.getDate()}日`;
  }

  function buildFooter(d) {
    const dt = d.dataDate;
    return `${dt.getMonth() + 1}月${dt.getDate()}日，金阳新区污水处理量合计${d.fullSewageTotal}吨，其中永安污水处理厂污水处理量为${d.yonganTotal}吨。`;
  }

  function exceedRegex() {
    return /(?:超标数据为：|超标数据为|超标数据|超标数值|数据超标值|超标值|超标为)\s*[:：]?\s*[^\n（）()]*?mg\/L(?=\s*(?:[（(]|[。.;；]|设计值|$))/ig;
  }

  function renderRichText(text) {
    const re = exceedRegex();
    let html = '';
    let last = 0;
    let m;
    while ((m = re.exec(text)) !== null) {
      html += escapeHtml(text.slice(last, m.index));
      html += `<span class="exceed">${escapeHtml(m[0])}</span>`;
      last = m.index + m[0].length;
    }
    html += escapeHtml(text.slice(last));
    return html.replace(/\n/g, '<br>');
  }

  function renderPreview(d) {
    previewCard.classList.remove('hidden');
    $('dataPreview').innerHTML = [
      ['数据日期', `${d.dataDate.getFullYear()}-${String(d.dataDate.getMonth() + 1).padStart(2, '0')}-${String(d.dataDate.getDate()).padStart(2, '0')}`],
      ['供水合计', `${d.supplyTotal}吨`],
      ['污水处理合计', `${d.fullSewageTotal}吨`],
      ['南园+北园', `${d.mainSewageTotal}吨`],
      ['一水厂', `${d.oneSupply}吨`],
      ['二水厂', `${d.twoSupply}吨`],
      ['南园', `${d.southTotal}吨`],
      ['北园', `${d.northTotal}吨`],
    ].map(([k, v]) => `<div class="metric"><span>${escapeHtml(k)}</span><strong>${escapeHtml(v)}</strong></div>`).join('');

    const quality = [
      ...d.southQuality.map(x => ({...x, plant: '南园'})),
      ...d.northQuality.map(x => ({...x, plant: '北园'})),
    ];
    $('qualityPreview').innerHTML = quality.map(x => {
      const cls = x.exceedText ? 'quality-tag exceed-tag' : 'quality-tag';
      const suffix = x.exceedText ? `：${x.exceedText}` : '：正常';
      return `<span class="${cls}">${escapeHtml(`${x.plant}${x.metric}${suffix}`)}</span>`;
    }).join('');

    $('warningBox').innerHTML = d.warnings.map(w => `<div class="warning">⚠️ ${escapeHtml(w)}</div>`).join('');
  }

  function renderReport(d) {
    $('dataDateText').textContent = `数据日期：${d.dataDate.getFullYear()}年${d.dataDate.getMonth() + 1}月${d.dataDate.getDate()}日`;
    $('compileDateText').textContent = `汇编日期：${d.reportDate.getFullYear()}年${d.reportDate.getMonth() + 1}月${d.reportDate.getDate()}日`;
    $('supplyText').textContent = buildSupplyText(d);
    $('wastewaterText').innerHTML = renderRichText(buildBodyText(d));
    $('footerText').textContent = buildFooter(d);
    state.text = buildTextSummary(d);
    $('textOutput').textContent = state.text;
  }

  function cloneSheet(ws) {
    if (typeof structuredClone === 'function') return structuredClone(ws);
    return JSON.parse(JSON.stringify(ws));
  }

  function sheetDate(name) {
    const m = /^\s*(\d{1,2})月\s*(\d{1,2})日\s*$/.exec(String(name || '').trim());
    return m ? [Number(m[1]), Number(m[2])] : null;
  }

  function chooseTemplateSheet(wb, month, day) {
    const target = month * 100 + day;
    const items = wb.SheetNames
      .map(name => ({name, md: sheetDate(name)}))
      .filter(x => x.md && x.md[0] === month);

    if (!items.length) throw new Error(`当前Excel模板中没有${month}月日报工作表，请先选择${month}月模板。`);
    const before = items.filter(x => x.md[0] * 100 + x.md[1] < target);
    if (before.length) {
      return before.sort((a, b) => (b.md[0] * 100 + b.md[1]) - (a.md[0] * 100 + a.md[1]))[0].name;
    }
    const same = items.find(x => x.md[1] === day);
    if (same) return same.name;
    return items.sort((a, b) => (b.md[0] * 100 + b.md[1]) - (a.md[0] * 100 + a.md[1]))[0].name;
  }

  function setCell(ws, addr, value) {
    const old = ws[addr] || {};
    ws[addr] = {...old, t: typeof value === 'number' ? 'n' : 's', v: value};
    delete ws[addr].w;
    delete ws[addr].h;
    delete ws[addr].r;
    delete ws[addr].R;
    delete ws[addr].f;
  }

  function patchAlignmentNoWrap(cell) {
    if (!cell) return;
    cell.s = cell.s || {};
    cell.s.alignment = {...(cell.s.alignment || {}), wrapText: false, vertical: 'center'};
  }

  function makeRichRuns(text) {
    const re = exceedRegex();
    const runs = [];
    let last = 0;
    let m;
    while ((m = re.exec(text)) !== null) {
      if (m.index > last) runs.push({t: 's', v: text.slice(last, m.index)});
      runs.push({t: 's', v: m[0], s: {color: {rgb: 'FF0000'}, font: {color: {rgb: 'FFFF0000'}}}});
      last = m.index + m[0].length;
    }
    if (last < text.length) runs.push({t: 's', v: text.slice(last)});
    return runs;
  }

  function estimateBodyHeight(bodyText, baseRows) {
    const explicitLines = bodyText.split('\n');
    let visualLines = 0;
    for (const line of explicitLines) {
      const weighted = [...line].reduce((n, ch) => n + (/[^\x00-\xff]/.test(ch) ? 1 : 0.58), 0);
      visualLines += Math.max(1, Math.ceil(weighted / 78));
    }
    const estimatedPt = visualLines * 19.5 + 16;
    const current = (baseRows?.[4]?.hpt || 252) + (baseRows?.[5]?.hpt || 117);
    return Math.max(current, estimatedPt);
  }

  function openDb() {
    return new Promise((resolve, reject) => {
      if (!('indexedDB' in window)) return reject(new Error('当前浏览器不支持本机模板保存。'));
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE);
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error || new Error('本机存储打开失败。'));
    });
  }

  async function idbGet(key) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, 'readonly');
      const req = tx.objectStore(DB_STORE).get(key);
      req.onsuccess = () => resolve(req.result ?? null);
      req.onerror = () => reject(req.error);
      tx.oncomplete = () => db.close();
    });
  }

  async function idbSet(key, value) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, 'readwrite');
      tx.objectStore(DB_STORE).put(value, key);
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onerror = () => { db.close(); reject(tx.error); };
    });
  }

  async function idbDelete(key) {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, 'readwrite');
      tx.objectStore(DB_STORE).delete(key);
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onerror = () => { db.close(); reject(tx.error); };
    });
  }

  function updateTemplateUi() {
    if (state.savedTemplate) {
      $('templateName').textContent = state.savedTemplate.name;
      $('templateStatus').textContent = '本机已保存';
    } else {
      $('templateName').textContent = BUILTIN_TEMPLATE_NAME;
      $('templateStatus').textContent = '内置模板';
    }
  }

  async function restoreSavedTemplate() {
    try {
      const saved = await idbGet(TEMPLATE_KEY);
      if (saved && saved.name && saved.buffer) state.savedTemplate = saved;
    } catch (_) {}
    updateTemplateUi();
  }

  async function loadTemplateArrayBuffer() {
    if (state.savedTemplate?.buffer) return state.savedTemplate.buffer.slice(0);
    const resp = await fetch(BUILTIN_TEMPLATE_URL, {cache: 'no-store'});
    if (!resp.ok) throw new Error('内置Excel模板读取失败，请选择本月Excel模板。');
    return await resp.arrayBuffer();
  }

  function verifyGeneratedWorkbook(bytes, d) {
    const check = XLSX.read(bytes, {type: 'array', cellStyles: true});
    const name = `${d.reportMonth}月${d.reportDay}日`;
    const ws = check.Sheets[name];
    if (!ws) throw new Error(`Excel生成校验失败：未找到新工作表“${name}”。`);
    const checks = [
      ['C2', d.oneSupply], ['D2', d.twoSupply], ['E2', d.supplyTotal],
      ['C3', d.southTotal], ['D3', d.northTotal], ['E3', d.yonganTotal], ['F3', d.fullSewageTotal],
    ];
    for (const [addr, expected] of checks) {
      const actual = Number(ws[addr]?.v);
      if (actual !== Number(expected)) throw new Error(`Excel生成校验失败：${addr}应为${expected}，实际为${ws[addr]?.v ?? '空'}。`);
    }
    if (!String(ws['B4']?.v || '').includes(String(d.supplyTotal))) throw new Error('Excel生成校验失败：供水正文未正确写入。');
    if (!String(ws['B5']?.v || '').includes(String(d.mainSewageTotal))) throw new Error('Excel生成校验失败：污水正文未正确写入。');
  }

  async function buildExcel(d) {
    if (typeof XLSX === 'undefined') throw new Error('Excel组件未加载。请保持网络连接并刷新页面后重试。');
    const arrayBuffer = await loadTemplateArrayBuffer();
    const wb = XLSX.read(arrayBuffer, {type: 'array', cellStyles: true, cellNF: true, cellHTML: true, bookVBA: true});
    const newName = `${d.reportMonth}月${d.reportDay}日`;
    const templateName = chooseTemplateSheet(wb, d.reportMonth, d.reportDay);
    const templateWs = wb.Sheets[templateName];
    const ws = cloneSheet(templateWs);
    const bodyText = buildBodyText(d);

    setCell(ws, 'A2', formatDateLine(d));
    setCell(ws, 'C2', d.oneSupply);
    setCell(ws, 'D2', d.twoSupply);
    setCell(ws, 'E2', d.supplyTotal);
    setCell(ws, 'C3', d.southTotal);
    setCell(ws, 'D3', d.northTotal);
    setCell(ws, 'E3', d.yonganTotal);
    setCell(ws, 'F3', d.fullSewageTotal);
    setCell(ws, 'B4', buildSupplyText(d));
    setCell(ws, 'B5', bodyText);
    setCell(ws, 'A7', buildFooter(d));
    patchAlignmentNoWrap(ws['A2']);

    const runs = makeRichRuns(bodyText);
    if (runs.some(r => r.s)) ws['B5'].R = runs;

    if (Array.isArray(ws['!rows'])) {
      const rows = ws['!rows'];
      while (rows.length < 7) rows.push({});
      const need = estimateBodyHeight(bodyText, rows);
      const base5 = rows[4]?.hpt || 252;
      const base6 = rows[5]?.hpt || 117;
      const total = base5 + base6;
      if (need > total) rows[5] = {...(rows[5] || {}), hpt: base6 + (need - total)};
      rows[1] = {...(rows[1] || {}), hpt: rows[1]?.hpt || 27};
    }

    // 第一张汇总表完全不修改。只新增/替换目标日报工作表。
    if (wb.Sheets[newName]) {
      wb.Sheets[newName] = ws;
    } else {
      wb.SheetNames.push(newName);
      wb.Sheets[newName] = ws;
    }

    const out = XLSX.write(wb, {bookType: 'xls', type: 'array', cellStyles: true, bookSST: true});
    verifyGeneratedWorkbook(out, d);
    return new Blob([out], {type: 'application/vnd.ms-excel'});
  }

  function canvasLooksValid(canvas) {
    if (!canvas || canvas.width < 500 || canvas.height < 250) return false;
    const ctx = canvas.getContext('2d', {willReadFrequently: true});
    const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
    const step = Math.max(4, Math.floor((canvas.width * canvas.height) / 12000) * 4);
    let nonWhite = 0;
    let samples = 0;
    for (let i = 0; i < data.length; i += step) {
      const r = data[i], g = data[i + 1], b = data[i + 2], a = data[i + 3];
      if (a > 10 && (r < 245 || g < 245 || b < 245)) nonWhite++;
      samples++;
    }
    return samples > 0 && nonWhite / samples > 0.01;
  }

  async function makePngBlob() {
    if (typeof html2canvas === 'undefined') throw new Error('图片组件未加载。请保持网络连接并刷新页面后重试。');
    if (document.fonts?.ready) await document.fonts.ready;
    await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
    const node = $('reportCapture');
    const canvas = await html2canvas(node, {
      backgroundColor: '#ffffff',
      scale: 1,
      useCORS: true,
      logging: false,
      width: node.scrollWidth,
      height: node.scrollHeight,
      windowWidth: node.scrollWidth,
      windowHeight: node.scrollHeight,
    });
    if (!canvasLooksValid(canvas)) throw new Error('日报图片生成异常：检测到图片可能为空白，请刷新页面后重试。');
    return await new Promise((resolve, reject) => {
      canvas.toBlob(b => b ? resolve(b) : reject(new Error('PNG生成失败。')), 'image/png');
    });
  }

  function blobDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 6000);
  }

  async function shareBlob(blob, filename, title) {
    if (!blob) throw new Error('请先生成日报。');
    if (!navigator.share) throw new Error('当前浏览器不支持直接分享，请使用下载/保存按钮。');
    const file = new File([blob], filename, {type: blob.type || 'application/octet-stream'});
    if (navigator.canShare && !navigator.canShare({files: [file]})) {
      throw new Error('当前浏览器不支持分享此文件类型，请使用下载按钮。');
    }
    await navigator.share({title, files: [file]});
  }

  function setPngPreview(blob) {
    if (state.pngObjectUrl) URL.revokeObjectURL(state.pngObjectUrl);
    state.pngObjectUrl = URL.createObjectURL(blob);
    $('pngPreview').src = state.pngObjectUrl;
  }

  function parseAndRender() {
    hideMessage();
    const d = parseWechatText(rawTextEl.value);
    state.parsed = d;
    renderPreview(d);
    renderReport(d);
    showMessage('数据解析与校验通过。', 'ok');
    return d;
  }

  async function generateAll() {
    const d = parseAndRender();
    renderReport(d);
    state.excelBlob = await buildExcel(d);
    state.pngBlob = await makePngBlob();
    state.text = buildTextSummary(d);
    $('textOutput').textContent = state.text;
    setPngPreview(state.pngBlob);
    resultCard.classList.remove('hidden');
    resultCard.scrollIntoView({behavior: 'smooth', block: 'start'});
    showMessage('Excel、日报图片、文字版均已生成并完成内容校验。', 'ok');
  }

  function excelFilename(d) {
    return `水务公司运行日报表（${d.reportDate.getFullYear()}年${d.reportMonth}月）_更新至${d.reportMonth}月${d.reportDay}日.xls`;
  }
  function pngFilename(d) {
    return `${d.reportDate.getFullYear()}年${d.reportMonth}月${d.reportDay}日_水务运行日报.png`;
  }
  function txtFilename(d) {
    return `${d.dataDate.getFullYear()}年${d.dataDate.getMonth() + 1}月${d.dataDate.getDate()}日_供排水情况.txt`;
  }

  function updateNetworkBadge() {
    const el = $('networkBadge');
    if (navigator.onLine) {
      el.textContent = '在线';
      el.classList.remove('offline');
    } else {
      el.textContent = '离线';
      el.classList.add('offline');
    }
  }

  function isStandalone() {
    return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
  }
  function isIOS() {
    return /iPhone|iPad|iPod/i.test(navigator.userAgent);
  }

  function configureInstallUi() {
    if (isStandalone()) return;
    $('installCard').classList.remove('hidden');
    if (isIOS()) $('installHint').textContent = 'Safari 点“分享”→“添加到主屏幕”，以后像 App 一样打开。';
  }

  function openInstallHelp() {
    const steps = isIOS()
      ? '1. 请使用 Safari 打开本页面。<br>2. 点底部“分享”按钮。<br>3. 向下找到“添加到主屏幕”。<br>4. 名称保留“水务日报”，点“添加”。'
      : '如果浏览器支持安装，会直接弹出安装提示；若没有提示，可在浏览器菜单中选择“安装应用”或“添加到主屏幕”。';
    $('installSteps').innerHTML = steps;
    $('installModal').classList.remove('hidden');
  }

  async function init() {
    const draft = localStorage.getItem(DRAFT_KEY);
    if (draft) rawTextEl.value = draft;
    await restoreSavedTemplate();
    updateNetworkBadge();
    configureInstallUi();
  }

  $('templateFile').addEventListener('change', async () => {
    const file = $('templateFile').files[0];
    if (!file) return;
    try {
      if (!/\.(xls|xlsx)$/i.test(file.name)) throw new Error('请选择 .xls 或 .xlsx Excel 文件。');
      const buffer = await file.arrayBuffer();
      state.savedTemplate = {name: file.name, buffer};
      await idbSet(TEMPLATE_KEY, state.savedTemplate);
      updateTemplateUi();
      showMessage(`已保存本机模板：${file.name}。以后打开本页面会自动使用。`, 'ok');
    } catch (e) {
      showMessage(e.message || String(e), 'error');
    } finally {
      $('templateFile').value = '';
    }
  });

  $('clearTemplateBtn').addEventListener('click', async () => {
    try { await idbDelete(TEMPLATE_KEY); } catch (_) {}
    state.savedTemplate = null;
    updateTemplateUi();
    showMessage('已恢复使用内置11月模板。', 'ok');
  });

  rawTextEl.addEventListener('input', () => {
    try { localStorage.setItem(DRAFT_KEY, rawTextEl.value); } catch (_) {}
  });

  $('clearTextBtn').addEventListener('click', () => {
    rawTextEl.value = '';
    try { localStorage.removeItem(DRAFT_KEY); } catch (_) {}
    previewCard.classList.add('hidden');
    resultCard.classList.add('hidden');
    hideMessage();
    rawTextEl.focus();
  });

  $('pasteBtn').addEventListener('click', async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (!text) throw new Error('剪贴板中没有文字。');
      rawTextEl.value = text;
      localStorage.setItem(DRAFT_KEY, text);
      showMessage('已从剪贴板粘贴。', 'ok');
    } catch (_) {
      showMessage('浏览器未允许直接读取剪贴板。请长按文本框后选择“粘贴”。', 'error');
      rawTextEl.focus();
    }
  });

  $('parseBtn').addEventListener('click', () => {
    try { parseAndRender(); }
    catch (e) { showMessage(e.message || String(e), 'error'); }
  });

  $('generateBtn').addEventListener('click', async () => {
    const btn = $('generateBtn');
    btn.disabled = true;
    btn.textContent = '正在生成…';
    try {
      await generateAll();
    } catch (e) {
      console.error(e);
      showMessage(e.message || String(e), 'error');
    } finally {
      btn.disabled = false;
      btn.textContent = '生成日报';
    }
  });

  $('downloadExcel').addEventListener('click', () => {
    if (!state.excelBlob || !state.parsed) return showMessage('请先点击“生成日报”。', 'error');
    blobDownload(state.excelBlob, excelFilename(state.parsed));
  });
  $('downloadPng').addEventListener('click', () => {
    if (!state.pngBlob || !state.parsed) return showMessage('请先点击“生成日报”。', 'error');
    blobDownload(state.pngBlob, pngFilename(state.parsed));
  });
  $('downloadTxt').addEventListener('click', () => {
    if (!state.text || !state.parsed) return showMessage('请先点击“生成日报”。', 'error');
    blobDownload(new Blob([state.text], {type: 'text/plain;charset=utf-8'}), txtFilename(state.parsed));
  });

  $('shareExcel').addEventListener('click', async () => {
    try { await shareBlob(state.excelBlob, excelFilename(state.parsed), '水务运行日报 Excel'); }
    catch (e) { if (e.name !== 'AbortError') showMessage(e.message || String(e), 'error'); }
  });
  $('sharePng').addEventListener('click', async () => {
    try { await shareBlob(state.pngBlob, pngFilename(state.parsed), '水务运行日报图片'); }
    catch (e) { if (e.name !== 'AbortError') showMessage(e.message || String(e), 'error'); }
  });

  $('copyText').addEventListener('click', async () => {
    if (!state.text) return showMessage('请先点击“生成日报”。', 'error');
    try {
      await navigator.clipboard.writeText(state.text);
      showMessage('文字版已复制，可直接粘贴到微信。', 'ok');
    } catch (_) {
      const ta = document.createElement('textarea');
      ta.value = state.text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
      showMessage('文字版已复制，可直接粘贴到微信。', 'ok');
    }
  });

  window.addEventListener('online', updateNetworkBadge);
  window.addEventListener('offline', updateNetworkBadge);
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    state.deferredInstallPrompt = e;
    if (!isStandalone()) $('installCard').classList.remove('hidden');
  });

  $('installBtn').addEventListener('click', async () => {
    if (state.deferredInstallPrompt) {
      state.deferredInstallPrompt.prompt();
      await state.deferredInstallPrompt.userChoice.catch(() => {});
      state.deferredInstallPrompt = null;
    } else {
      openInstallHelp();
    }
  });
  $('closeInstallModal').addEventListener('click', () => $('installModal').classList.add('hidden'));
  $('installModal').addEventListener('click', (e) => {
    if (e.target === $('installModal')) $('installModal').classList.add('hidden');
  });

  window.WaterReportMobile = {
    APP_VERSION,
    parseWechatText,
    buildBodyText,
    buildSupplyText,
    buildTextSummary,
    buildFooter,
    formatDateLine,
  };

  if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }

  init().catch(() => {});
})();
