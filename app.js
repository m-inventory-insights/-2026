/* Dashboard logic: the selected workbooks are parsed in the browser only. */
(() => {
  'use strict';
  const CLASS_KEY = 'sales-analysis-product-classes-v1';
  const state = {
    files: { invoices: null, deliveries: null, returns: null },
    data: null,
    classifications: loadClassifications(),
    names: { invoices: '', deliveries: '', returns: '' }
  };
  const $ = (selector) => document.querySelector(selector);
  const $$ = (selector) => Array.from(document.querySelectorAll(selector));

  function loadClassifications() {
    try {
      const raw = localStorage.getItem(CLASS_KEY);
      const parsed = raw ? JSON.parse(raw) : {};
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (_) { return {}; }
  }
  function saveClassifications() {
    try { localStorage.setItem(CLASS_KEY, JSON.stringify(state.classifications)); } catch (_) { /* storage may be unavailable */ }
  }
  function clean(value) {
    if (value === undefined || value === null) return '';
    return String(value).replace(/\u00a0/g, ' ').replace(/[\u200e\u200f\u202a-\u202e]/g, '').replace(/\s+/g, ' ').trim();
  }
  function norm(value) {
    return clean(value).toLocaleLowerCase().normalize('NFKC').replace(/\s+/g, ' ');
  }
  function normHeader(value) {
    return clean(value).toLocaleLowerCase().normalize('NFKD').replace(/[^\p{L}\p{N}]/gu, '');
  }
  function normId(value) { return norm(value).replace(/\s+/g, ''); }
  function productKey(value) { return norm(value); }
  function esc(value) {
    return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function isBlank(value) { return value === null || value === undefined || clean(value) === ''; }
  function numberOrNull(value) {
    if (value === null || value === undefined || clean(value) === '') return null;
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    const stripped = clean(value).replace(/[٬,\s]/g, '').replace(/[٫]/g, '.').replace(/[^\d.\-]/g, '');
    if (!stripped || stripped === '-' || stripped === '.') return null;
    const n = Number(stripped);
    return Number.isFinite(n) ? n : null;
  }
  function sum(values) { return values.reduce((acc, v) => acc + (Number.isFinite(v) ? v : 0), 0); }
  function fmtNumber(value, digits = 2) {
    if (value === null || value === undefined || !Number.isFinite(Number(value))) return '—';
    return new Intl.NumberFormat('en-US', { maximumFractionDigits: digits, minimumFractionDigits: 0 }).format(Number(value));
  }
  function fmtMoney(value) { return fmtNumber(value, 2); }
  function fmtPercent(value) { return `${fmtNumber(value, 1)}%`; }
  function excelDate(value) {
    if (value === null || value === undefined || value === '') return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === 'number' && Number.isFinite(value)) {
      // Excel's date epoch (including its historical 1900 compatibility offset).
      return new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    }
    const d = new Date(value);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  function fmtDate(value, includeTime = false) {
    const d = excelDate(value);
    if (!d) return clean(value) || '—';
    const datePart = `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
    if (!includeTime) return datePart;
    const mins = d.getUTCHours() * 60 + d.getUTCMinutes();
    if (mins === 0) return datePart;
    return `${datePart} ${String(d.getUTCHours()).padStart(2, '0')}:${String(d.getUTCMinutes()).padStart(2, '0')}`;
  }
  function categoryDefault(product) {
    const s = clean(product);
    // Names that strongly indicate an assembled / finished trailer or vehicle body.
    return /(تريلا|قلاب|الطيارة|كساحة|تنك\s+مثبت|شاسية|شاسيه|مقطورة|قلاب خلفي|قلاب قطعتين)/i.test(s) ? 'finished' : 'raw';
  }
  function categoryOf(product) {
    const key = productKey(product);
    const saved = state.classifications[key];
    return saved === 'finished' || saved === 'raw' || saved === 'unclassified' ? saved : categoryDefault(product);
  }
  function categoryLabel(cat) {
    if (cat === 'finished') return 'منتجات تامة';
    if (cat === 'raw') return 'مواد خام / مبيعات غير مباشرة';
    return 'غير مصنف';
  }
  function categoryTag(cat) {
    const klass = cat === 'finished' ? 'tag-finished' : cat === 'raw' ? 'tag-raw' : 'tag-unknown';
    return `<span class="tag ${klass}">${esc(categoryLabel(cat))}</span>`;
  }
  function setStatus(text, kind = 'idle') {
    const el = $('#data-status');
    el.className = `status-pill ${kind === 'ready' ? 'status-ready' : kind === 'error' ? 'status-error' : 'status-idle'}`;
    el.innerHTML = `<span class="status-dot"></span>${esc(text)}`;
  }
  function showImportMessage(text, type = 'success') {
    const el = $('#import-message');
    el.hidden = false;
    el.className = `import-message ${type}`;
    el.textContent = text;
  }
  function findHeader(headers, candidates, required = true) {
    const normalized = headers.map(normHeader);
    for (const candidate of candidates) {
      const idx = normalized.indexOf(normHeader(candidate));
      if (idx >= 0) return idx;
    }
    if (required) throw new Error(`لم يتم العثور على العمود المطلوب: ${candidates[0]}`);
    return -1;
  }
  async function readWorkbook(file) {
    if (!file) throw new Error('من فضلك اختر الملفات الثلاثة أولًا.');
    if (!window.XLSX) throw new Error('تعذر تحميل مكتبة قراءة Excel. افتح الموقع مع اتصال بالإنترنت ثم أعد المحاولة.');
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: false, raw: true });
    const firstName = workbook.SheetNames && workbook.SheetNames[0];
    if (!firstName) throw new Error(`الملف ${file.name} لا يحتوي على ورقة بيانات.`);
    const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[firstName], { header: 1, defval: null, raw: true, blankrows: true });
    const headerIndex = matrix.findIndex(row => Array.isArray(row) && row.some(v => !isBlank(v)));
    if (headerIndex < 0) throw new Error(`الملف ${file.name} فارغ.`);
    return { rows: matrix, headerIndex, headers: matrix[headerIndex].map(v => clean(v)), sheetName: firstName };
  }
  function parseInvoices(sheet) {
    const h = sheet.headers;
    const ix = {
      date: findHeader(h, ['Date']), salesperson: findHeader(h, ['Salesperson']), product: findHeader(h, ['Invoice lines/Product', 'Product']),
      customer: findHeader(h, ['Partner', 'Customer']), origin: findHeader(h, ['Origin']), quantity: findHeader(h, ['Total Quantity:', 'Total Quantity']), amount: findHeader(h, ['Total Signed'])
    };
    const carry = { date: null, salesperson: '', customer: '', origin: '' };
    const out = [];
    for (let i = sheet.headerIndex + 1; i < sheet.rows.length; i++) {
      const row = sheet.rows[i] || [];
      const prod = clean(row[ix.product]);
      if (!prod) continue;
      const current = { date: row[ix.date], salesperson: clean(row[ix.salesperson]), customer: clean(row[ix.customer]), origin: clean(row[ix.origin]) };
      const startsNewGroup = [ix.date, ix.salesperson, ix.customer, ix.origin].some(j => !isBlank(row[j]));
      if (startsNewGroup) {
        carry.date = current.date;
        carry.salesperson = current.salesperson;
        carry.customer = current.customer;
        carry.origin = current.origin;
      }
      out.push({
        sourceRow: i + 1,
        date: carry.date,
        salesperson: carry.salesperson || 'غير محدد بملف الفواتير',
        product: prod,
        customer: carry.customer || '',
        origin: carry.origin || '',
        quantity: numberOrNull(row[ix.quantity]),
        amount: numberOrNull(row[ix.amount])
      });
    }
    if (!out.length) throw new Error('لم يتم العثور على صفوف منتجات في ملف الفواتير. راجع عنوان عمود Invoice lines/Product.');
    return out;
  }
  function parseDeliveries(sheet) {
    const h = sheet.headers;
    const ix = {
      date: findHeader(h, ['Date of Transfer']), source: findHeader(h, ['Source Document']), customer: findHeader(h, ['Customer Name']),
      product: findHeader(h, ['Product']), quantity: findHeader(h, ['Stock Moves/Quantity Done', 'Quantity Done']), chassis: findHeader(h, ['Chassis No.', 'Chassis No']), reference: findHeader(h, ['Reference'])
    };
    const out = [];
    for (let i = sheet.headerIndex + 1; i < sheet.rows.length; i++) {
      const row = sheet.rows[i] || [];
      if (row.every(isBlank)) continue;
      const source = clean(row[ix.source]);
      const product = clean(row[ix.product]);
      if (!source && !product) continue;
      out.push({ sourceRow: i + 1, date: row[ix.date], sourceDocument: source, customer: clean(row[ix.customer]), product: product || 'منتج غير محدد', quantity: numberOrNull(row[ix.quantity]) ?? 0, chassis: clean(row[ix.chassis]), reference: clean(row[ix.reference]) });
    }
    if (!out.length) throw new Error('لم يتم العثور على صفوف تسليمات في الملف.');
    return out;
  }
  function parseReturns(sheet) {
    const h = sheet.headers;
    const ix = {
      reference: findHeader(h, ['Reference']), customer: findHeader(h, ['Customer Name']), date: findHeader(h, ['Scheduled Date']),
      sourceDocument: findHeader(h, ['Source Document']), status: findHeader(h, ['Status'], false), priority: findHeader(h, ['Priority'], false)
    };
    const out = [];
    for (let i = sheet.headerIndex + 1; i < sheet.rows.length; i++) {
      const row = sheet.rows[i] || [];
      if (row.every(isBlank)) continue;
      const sourceDocument = clean(row[ix.sourceDocument]);
      const reference = clean(row[ix.reference]);
      if (!sourceDocument && !reference) continue;
      out.push({ sourceRow: i + 1, reference, customer: clean(row[ix.customer]), date: row[ix.date], sourceDocument, status: ix.status >= 0 ? clean(row[ix.status]) : '', priority: ix.priority >= 0 ? clean(row[ix.priority]) : '' });
    }
    return out;
  }
  async function analyzeSelectedFiles() {
    const button = $('#analyze-files');
    button.disabled = true;
    button.innerHTML = 'جارٍ تحليل الملفات…';
    setStatus('جارٍ قراءة الملفات');
    try {
      const [invSheet, delSheet, retSheet] = await Promise.all([
        readWorkbook(state.files.invoices), readWorkbook(state.files.deliveries), readWorkbook(state.files.returns)
      ]);
      const invoices = parseInvoices(invSheet);
      const deliveries = parseDeliveries(delSheet);
      const returns = parseReturns(retSheet);
      state.data = { invoices, deliveries, returns, fileNames: { ...state.names } };
      renderAll();
      $('#workspace').hidden = false;
      $('#export-workbook').disabled = false;
      setStatus('تم تحميل البيانات بنجاح', 'ready');
      showImportMessage(`اكتمل التحليل: ${invoices.length} سطر فاتورة، و${deliveries.length} سطر تسليم، و${returns.length} سجل مرتجع. يمكنك مراجعة التصنيف وتصدير التقارير.`, 'success');
      $('#workspace').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (error) {
      console.error(error);
      setStatus('تعذر التحليل', 'error');
      showImportMessage(error && error.message ? error.message : 'حدث خطأ أثناء قراءة الملفات.', 'error');
    } finally {
      button.disabled = !Object.values(state.files).every(Boolean);
      button.innerHTML = 'تحليل الملفات <span aria-hidden="true">←</span>';
    }
  }
  function dataModel() {
    const d = state.data;
    if (!d) return null;
    const invoiceOrders = new Set(d.invoices.map(r => normId(r.origin)).filter(Boolean));
    const sellerByOrigin = new Map();
    for (const r of d.invoices) {
      const k = normId(r.origin);
      if (!k) continue;
      if (!sellerByOrigin.has(k)) sellerByOrigin.set(k, new Set());
      if (r.salesperson && r.salesperson !== 'غير محدد بملف الفواتير') sellerByOrigin.get(k).add(r.salesperson);
    }
    const deliveryByReference = new Map();
    for (const row of d.deliveries) {
      const key = normId(row.reference);
      if (key && !deliveryByReference.has(key)) deliveryByReference.set(key, row);
    }
    const deliveries = d.deliveries.map(r => {
      const matched = !!r.sourceDocument && invoiceOrders.has(normId(r.sourceDocument));
      const sellers = r.sourceDocument && sellerByOrigin.has(normId(r.sourceDocument)) ? Array.from(sellerByOrigin.get(normId(r.sourceDocument))) : [];
      return { ...r, hasInvoice: matched, salesperson: sellers.length ? sellers.join('، ') : 'غير معروف من ملف التسليمات' };
    });
    const returns = d.returns.map(r => {
      const sourceReference = clean(r.sourceDocument).replace(/^return\s+of\s+/i, '').trim();
      const match = deliveryByReference.get(normId(sourceReference)) || null;
      return {
        ...r, sourceReference, matchedDelivery: match,
        product: match ? match.product : 'غير معروف — لا يوجد تسليم مطابق',
        deliveryQuantity: match ? match.quantity : null,
        quantityInferred: !!match,
        chassis: match ? match.chassis : '',
        deliveryReference: match ? match.reference : sourceReference,
        deliverySourceDocument: match ? match.sourceDocument : '',
        matchStatus: match ? 'مرتجع مربوط بالتسليم' : 'لم يُعثر على تسليم مطابق'
      };
    });
    const productNames = new Map();
    for (const row of d.invoices) if (row.product) productNames.set(productKey(row.product), clean(row.product));
    for (const row of d.deliveries) if (row.product) productNames.set(productKey(row.product), clean(row.product));
    for (const row of returns) if (row.product && !row.product.startsWith('غير معروف')) productNames.set(productKey(row.product), clean(row.product));
    const products = Array.from(productNames.values()).sort((a,b) => a.localeCompare(b, 'ar'));
    const sellers = Array.from(new Set(d.invoices.map(r => clean(r.salesperson)).filter(s => s && s !== 'غير محدد بملف الفواتير'))).sort((a,b) => a.localeCompare(b, 'en'));
    const uninvoiced = deliveries.filter(r => !r.hasInvoice);
    return { ...d, invoiceOrders, deliveries, returns, products, sellers, uninvoiced, unInvoicedOrderCount: new Set(uninvoiced.map(r => normId(r.sourceDocument) || `row-${r.sourceRow}`)).size };
  }
  function productMetrics(model) {
    const totalInvoiceAmount = sum(model.invoices.map(r => r.amount).filter(v => v !== null));
    const totalDeliveryQty = sum(model.deliveries.map(r => r.quantity));
    const returnedByProduct = new Map();
    for (const r of model.returns) {
      if (!r.matchedDelivery || r.deliveryQuantity === null) continue;
      const k = productKey(r.product);
      returnedByProduct.set(k, (returnedByProduct.get(k) || 0) + r.deliveryQuantity);
    }
    return model.products.map(product => {
      const key = productKey(product);
      const invRows = model.invoices.filter(r => productKey(r.product) === key);
      const delRows = model.deliveries.filter(r => productKey(r.product) === key);
      const qtyKnown = sum(invRows.map(r => r.quantity).filter(v => v !== null));
      const amountKnown = sum(invRows.map(r => r.amount).filter(v => v !== null));
      const deliveryQty = sum(delRows.map(r => r.quantity));
      const returnedQty = returnedByProduct.get(key) || 0;
      const missingQty = invRows.filter(r => r.quantity === null).length;
      const missingAmount = invRows.filter(r => r.amount === null).length;
      return {
        product, category: categoryOf(product), invoiceLines: invRows.length,
        invoiceQty: qtyKnown, invoiceQtyMissing: missingQty, invoiceAmount: amountKnown, invoiceAmountMissing: missingAmount,
        deliveryLines: delRows.length, deliveryQty, returnedQty, netDeliveryQty: deliveryQty - returnedQty,
        amountShare: totalInvoiceAmount ? amountKnown / totalInvoiceAmount * 100 : 0,
        deliveryShare: totalDeliveryQty ? deliveryQty / totalDeliveryQty * 100 : 0
      };
    });
  }
  function sellerMetrics(model) {
    return model.sellers.map(seller => {
      const rows = model.invoices.filter(r => clean(r.salesperson) === seller);
      const getCat = cat => rows.filter(r => categoryOf(r.product) === cat);
      const finished = getCat('finished');
      const raw = getCat('raw');
      const unknown = getCat('unclassified');
      const finishedQty = sum(finished.map(r => r.quantity).filter(v => v !== null));
      const rawQty = sum(raw.map(r => r.quantity).filter(v => v !== null));
      const unclassifiedQty = sum(unknown.map(r => r.quantity).filter(v => v !== null));
      return {
        seller,
        finishedQty,
        finishedAmount: sum(finished.map(r => r.amount).filter(v => v !== null)),
        rawQty,
        rawAmount: sum(raw.map(r => r.amount).filter(v => v !== null)),
        unclassifiedQty,
        unclassifiedAmount: sum(unknown.map(r => r.amount).filter(v => v !== null)),
        totalQty: finishedQty + rawQty + unclassifiedQty,
        totalAmount: sum(rows.map(r => r.amount).filter(v => v !== null))
      };
    }).sort((a,b) => b.totalAmount - a.totalAmount);
  }
  function sellerProductMetrics(model) {
    const products = model.products;
    const pairs = [];
    for (const seller of model.sellers) {
      for (const product of products) {
        const matched = model.invoices.filter(r => clean(r.salesperson) === seller && productKey(r.product) === productKey(product));
        pairs.push({
          seller, product, category: categoryOf(product),
          quantity: sum(matched.map(r => r.quantity).filter(v => v !== null)),
          amount: sum(matched.map(r => r.amount).filter(v => v !== null))
        });
      }
    }
    return pairs;
  }
  function renderAll() {
    const model = dataModel();
    if (!model) return;
    model.productMetrics = productMetrics(model);
    model.sellerMetrics = sellerMetrics(model);
    model.sellerProductMetrics = sellerProductMetrics(model);
    state.model = model;
    renderDashboard(model);
    renderProducts(model);
    renderSalespeople(model);
    renderUninvoiced(model);
    renderReturns(model);
    renderClassification(model);
    renderDetails(model);
    $('#uninvoiced-nav-count').textContent = String(model.unInvoicedOrderCount);
    $('#period-label').textContent = derivePeriod(model);
    refreshFileCardLabels();
  }
  function derivePeriod(model) {
    const allDates = [...model.invoices, ...model.deliveries, ...model.returns].map(r => excelDate(r.date)).filter(Boolean).sort((a,b) => a.getTime()-b.getTime());
    if (!allDates.length) return 'الفترة حسب الملفات';
    const first = allDates[0]; const last = allDates[allDates.length - 1];
    const start = `${String(first.getUTCMonth()+1).padStart(2,'0')}/${first.getUTCFullYear()}`;
    const end = `${String(last.getUTCMonth()+1).padStart(2,'0')}/${last.getUTCFullYear()}`;
    return start === end ? `الفترة: ${start}` : `من ${fmtDate(first)} إلى ${fmtDate(last)}`;
  }
  function makeKpi(label, value, helper, icon) {
    return `<article class="kpi-card"><div class="kpi-top"><div class="kpi-label">${esc(label)}</div><div class="kpi-icon">${icon}</div></div><div class="kpi-value">${esc(value)}</div><div class="kpi-helper">${esc(helper)}</div></article>`;
  }
  function renderDashboard(model) {
    const invAmount = sum(model.invoices.map(r => r.amount).filter(v => v !== null));
    const delQty = sum(model.deliveries.map(r => r.quantity));
    const missingAmt = model.invoices.filter(r => r.amount === null).length;
    const uniqueInvoiceOrders = new Set(model.invoices.map(r => normId(r.origin)).filter(Boolean)).size;
    $('#kpi-grid').innerHTML = [
      makeKpi('القيمة المسجلة بالفواتير', fmtMoney(invAmount), `${model.invoices.filter(r=>r.amount!==null).length} سطرًا له قيمة مسجلة`, '＄'),
      makeKpi('كمية التسليم الإجمالية', fmtNumber(delQty), `${model.deliveries.length} سطر تسليم`, '⇢'),
      makeKpi('الأوامر غير المفوترة', fmtNumber(model.unInvoicedOrderCount, 0), `${model.uninvoiced.length} سطر تسليم غير مفوتر`, '⚑'),
      makeKpi('سجلات المرتجعات', fmtNumber(model.returns.length, 0), `${model.returns.filter(r=>r.matchedDelivery).length} مرتبط بتسليم`, '↶')
    ].join('');
    const classAmounts = ['finished','raw','unclassified'].map(cat => ({ label: categoryLabel(cat), value: sum(model.invoices.filter(r => categoryOf(r.product) === cat).map(r=>r.amount).filter(v=>v!==null)), category:cat })).filter(r=>r.value>0);
    renderBarChart('category-chart', classAmounts, invAmount, true);
    const topProducts = [...model.productMetrics].sort((a,b)=>b.invoiceAmount-a.invoiceAmount).filter(r=>r.invoiceAmount>0).slice(0,6).map(r=>({label:r.product,value:r.invoiceAmount,category:r.category}));
    renderBarChart('top-products-chart', topProducts, Math.max(0,...topProducts.map(r=>r.value)), false);
    const notes = [
      { tone: model.unInvoicedOrderCount ? 'warn' : 'good', title: `${model.unInvoicedOrderCount} أمر تسليم غير مفوتر`, detail: model.unInvoicedOrderCount ? `تتوزع على ${model.unInvoiced.length} سطر تسليم؛ راجع تقرير الأوامر غير المفوترة.` : 'كل أوامر التسليم لها Origin مطابق بالفواتير.' },
      { tone: missingAmt ? 'warn' : 'good', title: `${missingAmt} سطر فاتورة بلا قيمة مسجلة`, detail: 'لا يتم تقدير القيمة الناقصة تلقائيًا أو توزيعها على أصناف أخرى.' },
      { tone: model.invoices.some(r=>r.quantity===null) ? 'warn' : 'good', title: `${model.invoices.filter(r=>r.quantity===null).length} سطر فاتورة بلا كمية مسجلة`, detail: 'تظهر الكميات المعروفة منفصلة عن عدد السطور التي تنقصها كمية.' },
      { tone: 'warn', title: 'بيانات الإنتاج غير متاحة في الملفات الحالية', detail: 'مساهمة المنتجات تُحسب من التسليمات والقيمة المسجلة، لا من تقرير إنتاج مستقل.' }
    ];
    $('#quality-notes').innerHTML = notes.map(n => `<div class="quality-item ${n.tone==='good'?'good':''}"><div class="quality-bullet">${n.tone==='good'?'✓':'!'}</div><div class="quality-text"><strong>${esc(n.title)}</strong><span>${esc(n.detail)}</span></div></div>`).join('');
    const sellers = model.sellerMetrics.slice(0,6);
    $('#seller-mini-table').innerHTML = `<table class="mini-table"><thead><tr><th>البائع</th><th>كمية المنتجات التامة</th><th>كمية المواد الخام</th><th>إجمالي الكمية</th><th>قيمة الفواتير</th></tr></thead><tbody>${sellers.map(s=>`<tr><td>${esc(s.seller)}</td><td class="num">${fmtNumber(s.finishedQty)}</td><td class="num">${fmtNumber(s.rawQty)}</td><td class="num">${fmtNumber(s.totalQty)}</td><td class="num">${fmtMoney(s.totalAmount)}</td></tr>`).join('') || '<tr><td colspan="5">لا توجد بيانات بائعين</td></tr>'}</tbody></table>`;
  }
  function renderBarChart(targetId, rows, maxValue, isCategory) {
    const el = document.getElementById(targetId);
    if (!rows.length) { el.innerHTML = '<div class="empty-chart">لا توجد قيم مسجلة لعرضها</div>'; return; }
    const max = Math.max(Number(maxValue)||0, ...rows.map(r=>r.value), 1);
    el.innerHTML = rows.map(r => {
      const pct = Math.min(100, Math.max(0, r.value / max * 100));
      return `<div class="bar-row"><div class="bar-label" title="${esc(r.label)}">${esc(r.label)}</div><div class="bar-track"><div class="bar-fill" style="width:${pct.toFixed(2)}%;${r.category==='raw'?'background:linear-gradient(90deg,#4c84c4,#9bbce6)':r.category==='unclassified'?'background:linear-gradient(90deg,#c88b30,#f1d09a)':''}"></div></div><div class="bar-value">${fmtMoney(r.value)}</div></div>`;
    }).join('');
  }
  function cellHtml(value, cls = '') {
    if (value === null || value === undefined || value === '') return `<td class="dim ${cls}">—</td>`;
    return `<td class="${cls}">${esc(value)}</td>`;
  }
  function renderTable(id, columns, rows, emptyMessage = 'لا توجد بيانات لعرضها') {
    const table = document.getElementById(id);
    const headers = `<thead><tr>${columns.map(c=>`<th>${esc(c.label)}</th>`).join('')}</tr></thead>`;
    const body = rows.length ? `<tbody>${rows.map(row=>`<tr>${columns.map(col=>{
      if (col.render) return `<td class="${col.cls||''}">${col.render(row)}</td>`;
      let value = typeof col.value === 'function' ? col.value(row) : row[col.key];
      if (col.format) value = col.format(value, row);
      if (col.html) return `<td class="${col.cls||''}">${value ?? ''}</td>`;
      return cellHtml(value, col.cls || '');
    }).join('')}</tr>`).join('')}</tbody>` : `<tbody><tr><td class="empty-cell" colspan="${columns.length}">${esc(emptyMessage)}</td></tr></tbody>`;
    table.innerHTML = headers + body;
  }
  function renderProducts(model) {
    const q = norm($('#product-search').value || '');
    const cat = $('#product-category-filter').value || 'all';
    const rows = model.productMetrics.filter(r => (!q || norm(r.product).includes(q)) && (cat === 'all' || r.category === cat));
    $('#products-count').textContent = `${rows.length} منتج`;
    renderTable('products-table', [
      {label:'المنتج',key:'product',cls:'product-name'},
      {label:'التصنيف',key:'category',html:true,render:r=>categoryTag(r.category)},
      {label:'كمية الفاتورة المسجلة',value:r=>r.invoiceQty,cls:'num',format:v=>fmtNumber(v)},
      {label:'سطور فاتورة بلا كمية',key:'invoiceQtyMissing',cls:'num',format:v=>fmtNumber(v,0)},
      {label:'القيمة المسجلة',key:'invoiceAmount',cls:'num',format:v=>fmtMoney(v)},
      {label:'سطور بلا قيمة',key:'invoiceAmountMissing',cls:'num',format:v=>fmtNumber(v,0)},
      {label:'كمية التسليم',key:'deliveryQty',cls:'num',format:v=>fmtNumber(v)},
      {label:'كمية مرتجع مرجعية*',key:'returnedQty',cls:'num',format:v=>fmtNumber(v)},
      {label:'صافي كمية التسليم*',key:'netDeliveryQty',cls:'num',format:v=>fmtNumber(v)},
      {label:'% من القيمة المسجلة',key:'amountShare',cls:'num',format:v=>fmtPercent(v)},
      {label:'% من كمية التسليم',key:'deliveryShare',cls:'num',format:v=>fmtPercent(v)}
    ], rows);
  }
  function renderSalespeople(model) {
    const summary = model.sellerMetrics;
    renderTable('sellers-table', [
      {label:'البائع',key:'seller',cls:'product-name'},
      {label:'إجمالي كمية المنتجات التامة',key:'finishedQty',cls:'num',format:v=>fmtNumber(v)},
      {label:'قيمة المنتجات التامة',key:'finishedAmount',cls:'num',format:v=>fmtMoney(v)},
      {label:'إجمالي كمية المواد الخام / المبيعات غير المباشرة',key:'rawQty',cls:'num',format:v=>fmtNumber(v)},
      {label:'قيمة المواد الخام',key:'rawAmount',cls:'num',format:v=>fmtMoney(v)},
      {label:'كمية المنتجات غير المصنفة',key:'unclassifiedQty',cls:'num',format:v=>fmtNumber(v)},
      {label:'قيمة المنتجات غير المصنفة',key:'unclassifiedAmount',cls:'num',format:v=>fmtMoney(v)},
      {label:'إجمالي الكمية المسجلة',key:'totalQty',cls:'num',format:v=>fmtNumber(v)},
      {label:'إجمالي قيمة الفواتير',key:'totalAmount',cls:'num',format:v=>fmtMoney(v)}
    ], summary, 'لم يتم العثور على أسماء بائعين في الفواتير.');
    const sellerSelect = $('#seller-filter');
    const existing = sellerSelect.value || 'all';
    sellerSelect.innerHTML = '<option value="all">كل البائعين</option>' + model.sellers.map(s=>`<option value="${esc(s)}">${esc(s)}</option>`).join('');
    sellerSelect.value = model.sellers.includes(existing) ? existing : 'all';
    renderSellerProducts(model);
  }
  function renderSellerProducts(model) {
    const q = norm($('#seller-product-search').value || '');
    const seller = $('#seller-filter').value || 'all';
    const cat = $('#seller-category-filter').value || 'all';
    const categoryOrder = { finished: 0, raw: 1, unclassified: 2 };
    const rows = model.sellerProductMetrics.filter(r => (!q || norm(r.product).includes(q) || norm(r.seller).includes(q)) && (seller==='all'||r.seller===seller) && (cat==='all'||r.category===cat))
      .sort((a,b) => a.seller.localeCompare(b.seller, 'ar') || categoryOrder[a.category] - categoryOrder[b.category] || a.product.localeCompare(b.product, 'ar'));
    $('#seller-product-count').textContent = `${new Set(rows.map(r=>r.seller)).size} بائع · ${new Set(rows.map(r=>r.product)).size} منتج`;
    // Each seller/product pair is already aggregated: show the sum of known quantity and value, never the number of invoice lines.
    renderTable('seller-products-table', [
      {label:'البائع',key:'seller'}, {label:'المنتج',key:'product',cls:'product-name'},
      {label:'التصنيف',html:true,render:r=>categoryTag(r.category)},
      {label:'إجمالي الكمية المباعة حسب الفواتير',key:'quantity',cls:'num',format:v=>fmtNumber(v)},
      {label:'إجمالي القيمة المسجلة',key:'amount',cls:'num',format:v=>fmtMoney(v)}
    ], rows);
  }
  function renderUninvoiced(model) {
    const lines = model.uninvoiced;
    const docs = new Map();
    for (const r of lines) {
      const key = normId(r.sourceDocument) || `row-${r.sourceRow}`;
      if (!docs.has(key)) docs.set(key, { sourceDocument:r.sourceDocument || 'غير محدد', customer:r.customer, lines:0, quantity:0, products:new Set() });
      const item = docs.get(key); item.lines += 1; item.quantity += r.quantity; item.products.add(r.product);
    }
    $('#uninvoiced-summary').innerHTML = [
      {label:'أوامر غير مفوترة',value:docs.size},{label:'سطور تسليم غير مفوترة',value:lines.length},{label:'إجمالي الكمية',value:fmtNumber(sum(lines.map(r=>r.quantity)))}
    ].map(s=>`<div class="inline-stat"><label>${esc(s.label)}</label><strong>${esc(s.value)}</strong></div>`).join('');
    renderTable('uninvoiced-table', [
      {label:'تاريخ التسليم',value:r=>fmtDate(r.date,true)}, {label:'أمر البيع / Source Document',key:'sourceDocument'},
      {label:'مرجع التسليم',key:'reference'}, {label:'العميل',key:'customer'}, {label:'المنتج',key:'product',cls:'product-name'},
      {label:'الكمية',key:'quantity',cls:'num',format:v=>fmtNumber(v)}, {label:'رقم الشاسيه',key:'chassis'},
      {label:'البائع',key:'salesperson'}, {label:'الحالة',html:true,render:()=>'<span class="tag tag-alert">غير مفوتر</span>'}
    ], lines, 'ممتاز — لم تظهر تسليمات غير مفوترة وفق قاعدة الربط الحالية.');
  }
  function renderReturns(model) {
    renderTable('returns-table', [
      {label:'مرجع المرتجع',key:'reference'}, {label:'تاريخ المرتجع',value:r=>fmtDate(r.date,true)},
      {label:'Source Document للمرتجع',key:'sourceDocument'}, {label:'مرجع التسليم المطابق',key:'deliveryReference'},
      {label:'أمر البيع',key:'deliverySourceDocument'}, {label:'العميل',value:r=>r.customer || r.matchedDelivery?.customer || '—'},
      {label:'المنتج المرتبط',key:'product',cls:'product-name'}, {label:'كمية التسليم المرجعية',key:'deliveryQuantity',cls:'num',format:v=>v===null?'غير متاحة':fmtNumber(v)},
      {label:'رقم الشاسيه',key:'chassis'}, {label:'الحالة',key:'status'}, {label:'نتيجة المطابقة',html:true,render:r=>`<span class="tag ${r.matchedDelivery?'tag-ok':'tag-alert'}">${esc(r.matchStatus)}</span>`}
    ], model.returns);
  }
  function renderClassification(model) {
    const q = norm($('#classification-search').value || '');
    const rows = model.products.filter(p=>!q||norm(p).includes(q));
    $('#classification-count').textContent = `${rows.length} منتج`;
    const usage = new Map();
    for (const r of model.invoices) usage.set(productKey(r.product),(usage.get(productKey(r.product))||0)+1);
    for (const r of model.deliveries) usage.set(productKey(r.product),(usage.get(productKey(r.product))||0));
    renderTable('classification-table', [
      {label:'المنتج',value:r=>r,cls:'product-name'},
      {label:'عدد سطور الفواتير',value:r=>usage.get(productKey(r))||0,cls:'num',format:v=>fmtNumber(v,0)},
      {label:'التصنيف الحالي',value:r=>categoryLabel(categoryOf(r))},
      {label:'تغيير التصنيف',html:true,render:r=>`<select class="classification-select" aria-label="تصنيف ${esc(r)}" data-product="${esc(r)}"><option value="finished" ${categoryOf(r)==='finished'?'selected':''}>منتجات تامة</option><option value="raw" ${categoryOf(r)==='raw'?'selected':''}>مواد خام / مبيعات غير مباشرة</option><option value="unclassified" ${categoryOf(r)==='unclassified'?'selected':''}>غير مصنف</option></select>`}
    ], rows);
    $$('#classification-table select[data-product]').forEach(select => select.addEventListener('change', e => {
      const product = e.currentTarget.dataset.product;
      state.classifications[productKey(product)] = e.currentTarget.value;
      saveClassifications();
      renderAll();
    }));
  }
  function renderDetails(model) {
    const source = $('#details-source').value || 'invoices';
    const q = norm($('#details-search').value || '');
    let rows, columns;
    if (source === 'invoices') {
      rows = model.invoices.filter(r=>!q||[r.product,r.salesperson,r.customer,r.origin].some(v=>norm(v).includes(q)));
      columns = [
        {label:'صف المصدر',key:'sourceRow',cls:'num'}, {label:'التاريخ',value:r=>fmtDate(r.date,true)}, {label:'البائع',key:'salesperson'},
        {label:'المنتج',key:'product',cls:'product-name'}, {label:'العميل',key:'customer'}, {label:'Origin',key:'origin'},
        {label:'الكمية المسجلة',key:'quantity',cls:'num',format:v=>v===null?'غير مسجلة':fmtNumber(v)},
        {label:'القيمة المسجلة',key:'amount',cls:'num',format:v=>v===null?'غير مسجلة':fmtMoney(v)},
        {label:'التصنيف',html:true,render:r=>categoryTag(categoryOf(r.product))}
      ];
    } else if (source === 'deliveries') {
      rows = model.deliveries.filter(r=>!q||[r.product,r.sourceDocument,r.reference,r.customer,r.chassis].some(v=>norm(v).includes(q)));
      columns = [
        {label:'صف المصدر',key:'sourceRow',cls:'num'}, {label:'تاريخ التسليم',value:r=>fmtDate(r.date,true)}, {label:'Source Document',key:'sourceDocument'},
        {label:'مرجع التسليم',key:'reference'}, {label:'العميل',key:'customer'}, {label:'المنتج',key:'product',cls:'product-name'},
        {label:'الكمية',key:'quantity',cls:'num',format:v=>fmtNumber(v)}, {label:'رقم الشاسيه',key:'chassis'},
        {label:'مطابقة فاتورة',html:true,render:r=>`<span class="tag ${r.hasInvoice?'tag-ok':'tag-alert'}">${r.hasInvoice?'مفوتر':'غير مفوتر'}</span>`}
      ];
    } else {
      rows = model.returns.filter(r=>!q||[r.reference,r.sourceDocument,r.customer,r.product,r.deliveryReference].some(v=>norm(v).includes(q)));
      columns = [
        {label:'صف المصدر',key:'sourceRow',cls:'num'}, {label:'مرجع المرتجع',key:'reference'}, {label:'التاريخ',value:r=>fmtDate(r.date,true)},
        {label:'العميل',key:'customer'}, {label:'Source Document',key:'sourceDocument'}, {label:'المنتج المرتبط',key:'product'},
        {label:'كمية التسليم المرجعية',key:'deliveryQuantity',cls:'num',format:v=>v===null?'غير متاحة':fmtNumber(v)}, {label:'الحالة',key:'status'}
      ];
    }
    $('#details-count').textContent = `${rows.length} سجل`;
    renderTable('details-table', columns, rows);
  }
  function refreshFileCardLabels() {
    for (const type of ['invoices','deliveries','returns']) {
      const file = state.files[type];
      const labelId = type === 'invoices' ? 'invoice-file-label' : type === 'deliveries' ? 'delivery-file-label' : 'return-file-label';
      const el = document.getElementById(labelId);
      if (file) el.textContent = file.name;
    }
  }
  function csvEscape(value) {
    let s = value === null || value === undefined ? '' : String(value);
    if (typeof value === 'string' && /^[=+\-@]/.test(s)) s = `'${s}`; // avoid formula execution in spreadsheet apps
    return `"${s.replace(/"/g,'""')}"`;
  }
  function downloadCSV(filename, headers, rows) {
    const content = '\ufeff' + [headers, ...rows].map(row => row.map(csvEscape).join(',')).join('\r\n');
    const blob = new Blob([content], {type:'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a'); link.href = url; link.download = filename; link.click();
    setTimeout(()=>URL.revokeObjectURL(url), 800);
  }
  function exportProductsCSV() {
    const m = state.model; if(!m) return;
    const q = norm($('#product-search').value || ''); const cat = $('#product-category-filter').value || 'all';
    const rows = m.productMetrics.filter(r=>(!q||norm(r.product).includes(q))&&(cat==='all'||r.category===cat));
    downloadCSV('تقرير_المنتجات.csv', ['المنتج','التصنيف','كمية الفاتورة المسجلة','سطور بلا كمية','القيمة المسجلة','سطور بلا قيمة','كمية التسليم','كمية مرتجع مرجعية','صافي كمية التسليم','نسبة القيمة %','نسبة التسليم %'], rows.map(r=>[r.product,categoryLabel(r.category),r.invoiceQty,r.invoiceQtyMissing,r.invoiceAmount,r.invoiceAmountMissing,r.deliveryQty,r.returnedQty,r.netDeliveryQty,r.amountShare,r.deliveryShare]));
  }
  function exportSellersCSV() {
    const m = state.model; if(!m) return;
    downloadCSV('ملخص_البائعين.csv',['البائع','إجمالي كمية المنتجات التامة','قيمة المنتجات التامة','إجمالي كمية المواد الخام / المبيعات غير المباشرة','قيمة المواد الخام','كمية المنتجات غير المصنفة','قيمة المنتجات غير المصنفة','إجمالي الكمية المسجلة','إجمالي القيمة المسجلة'],m.sellerMetrics.map(r=>[r.seller,r.finishedQty,r.finishedAmount,r.rawQty,r.rawAmount,r.unclassifiedQty,r.unclassifiedAmount,r.totalQty,r.totalAmount]));
  }
  function exportSellerProductsCSV() {
    const m = state.model; if(!m) return;
    const q = norm($('#seller-product-search').value || ''); const seller = $('#seller-filter').value || 'all'; const cat=$('#seller-category-filter').value||'all';
    const rows=m.sellerProductMetrics.filter(r=>(!q||norm(r.product).includes(q)||norm(r.seller).includes(q))&&(seller==='all'||r.seller===seller)&&(cat==='all'||r.category===cat))
      .sort((a,b) => a.seller.localeCompare(b.seller, 'ar') || ({finished:0,raw:1,unclassified:2}[a.category] - {finished:0,raw:1,unclassified:2}[b.category]) || a.product.localeCompare(b.product, 'ar'));
    downloadCSV('تفاصيل_البائع_والمنتج.csv',['البائع','المنتج','التصنيف','إجمالي الكمية المسجلة','إجمالي القيمة المسجلة'],rows.map(r=>[r.seller,r.product,categoryLabel(r.category),r.quantity,r.amount]));
  }
  function exportUninvoicedCSV() {
    const m=state.model;if(!m)return;
    downloadCSV('الأوامر_غير_المفوترة.csv',['تاريخ التسليم','Source Document','Reference','العميل','المنتج','الكمية','رقم الشاسيه','البائع'],m.uninvoiced.map(r=>[fmtDate(r.date,true),r.sourceDocument,r.reference,r.customer,r.product,r.quantity,r.chassis,r.salesperson]));
  }
  function exportReturnsCSV() {
    const m=state.model;if(!m)return;
    downloadCSV('تقرير_المرتجعات.csv',['مرجع المرتجع','التاريخ','Source Document','مرجع التسليم المطابق','العميل','المنتج المرتبط','كمية التسليم المرجعية التقديرية','رقم الشاسيه','الحالة','نتيجة المطابقة'],m.returns.map(r=>[r.reference,fmtDate(r.date,true),r.sourceDocument,r.deliveryReference,r.customer||r.matchedDelivery?.customer,r.product,r.deliveryQuantity,r.chassis,r.status,r.matchStatus]));
  }
  function exportClassificationCSV() {
    const m=state.model;if(!m)return;
    downloadCSV('تصنيف_المنتجات.csv',['المنتج','التصنيف'],m.products.map(p=>[p,categoryLabel(categoryOf(p))]));
  }
  function exportDetailsCSV() {
    const m=state.model;if(!m)return;
    const source=$('#details-source').value||'invoices'; const q=norm($('#details-search').value||'');
    if(source==='invoices'){
      const rows=m.invoices.filter(r=>!q||[r.product,r.salesperson,r.customer,r.origin].some(v=>norm(v).includes(q)));
      downloadCSV('تفاصيل_الفواتير.csv',['صف المصدر','التاريخ','البائع','المنتج','العميل','Origin','الكمية','القيمة','التصنيف'],rows.map(r=>[r.sourceRow,fmtDate(r.date,true),r.salesperson,r.product,r.customer,r.origin,r.quantity,r.amount,categoryLabel(categoryOf(r.product))]));
    }else if(source==='deliveries'){
      const rows=m.deliveries.filter(r=>!q||[r.product,r.sourceDocument,r.reference,r.customer,r.chassis].some(v=>norm(v).includes(q)));
      downloadCSV('تفاصيل_التسليمات.csv',['صف المصدر','التاريخ','Source Document','Reference','العميل','المنتج','الكمية','رقم الشاسيه','الحالة'],rows.map(r=>[r.sourceRow,fmtDate(r.date,true),r.sourceDocument,r.reference,r.customer,r.product,r.quantity,r.chassis,r.hasInvoice?'مفوتر':'غير مفوتر']));
    }else{
      const rows=m.returns.filter(r=>!q||[r.reference,r.sourceDocument,r.customer,r.product,r.deliveryReference].some(v=>norm(v).includes(q)));
      downloadCSV('تفاصيل_المرتجعات.csv',['صف المصدر','Reference','التاريخ','العميل','Source Document','المنتج المرتبط','كمية التسليم المرجعية','الحالة'],rows.map(r=>[r.sourceRow,r.reference,fmtDate(r.date,true),r.customer,r.sourceDocument,r.product,r.deliveryQuantity,r.status]));
    }
  }
  function exportFullWorkbook() {
    const m=state.model;if(!m)return;
    if(!window.XLSX){showImportMessage('مكتبة Excel غير متاحة. تأكد من اتصال الإنترنت.','error');return;}
    const wb=XLSX.utils.book_new();
    const add=(name, headers, rows)=>{const ws=XLSX.utils.aoa_to_sheet([headers,...rows]); ws['!views']=[{rightToLeft:true}]; XLSX.utils.book_append_sheet(wb,ws,name);};
    const invoiceTotal=sum(m.invoices.map(r=>r.amount).filter(v=>v!==null));
    const metrics=[
      ['المؤشر','القيمة'],['قيمة الفواتير المسجلة',invoiceTotal],['عدد سطور الفواتير',m.invoices.length],['عدد أوامر الفواتير الفريدة',new Set(m.invoices.map(r=>normId(r.origin)).filter(Boolean)).size],['كمية التسليم الإجمالية',sum(m.deliveries.map(r=>r.quantity))],['عدد سطور التسليم',m.deliveries.length],['الأوامر غير المفوترة',m.unInvoicedOrderCount],['سطور التسليم غير المفوترة',m.uninvoiced.length],['عدد سجلات المرتجعات',m.returns.length],['سطور فاتورة بلا كمية',m.invoices.filter(r=>r.quantity===null).length],['سطور فاتورة بلا قيمة',m.invoices.filter(r=>r.amount===null).length],['ملاحظة','لا يوجد ملف إنتاج مستقل؛ المساهمة مبنية على التسليمات والقيمة المسجلة.'],['ملاحظة المرتجع','كمية المرتجع غير مسجلة؛ كمية التسليم المطابقة معروضة كمرجع تقديري.']
    ];
    add('الملخص',metrics[0],metrics.slice(1));
    add('المنتجات',['المنتج','التصنيف','كمية الفاتورة المسجلة','سطور بلا كمية','القيمة المسجلة','سطور بلا قيمة','عدد سطور الفاتورة','كمية التسليم','كمية مرتجع مرجعية تقديرية','صافي كمية التسليم','نسبة القيمة المسجلة %','نسبة كمية التسليم %'],m.productMetrics.map(r=>[r.product,categoryLabel(r.category),r.invoiceQty,r.invoiceQtyMissing,r.invoiceAmount,r.invoiceAmountMissing,r.invoiceLines,r.deliveryQty,r.returnedQty,r.netDeliveryQty,r.amountShare,r.deliveryShare]));
    add('ملخص البائعين',['البائع','إجمالي كمية المنتجات التامة','قيمة المنتجات التامة','إجمالي كمية المواد الخام / المبيعات غير المباشرة','قيمة المواد الخام','كمية غير مصنفة','قيمة غير مصنفة','إجمالي الكمية المسجلة','إجمالي القيمة المسجلة'],m.sellerMetrics.map(r=>[r.seller,r.finishedQty,r.finishedAmount,r.rawQty,r.rawAmount,r.unclassifiedQty,r.unclassifiedAmount,r.totalQty,r.totalAmount]));
    add('البائع والمنتج',['البائع','المنتج','التصنيف','إجمالي الكمية المسجلة','إجمالي القيمة المسجلة'],m.sellerProductMetrics.map(r=>[r.seller,r.product,categoryLabel(r.category),r.quantity,r.amount]));
    add('غير مفوتر',['صف المصدر','التاريخ','Source Document','Reference','العميل','المنتج','الكمية','رقم الشاسيه','البائع'],m.uninvoiced.map(r=>[r.sourceRow,fmtDate(r.date,true),r.sourceDocument,r.reference,r.customer,r.product,r.quantity,r.chassis,r.salesperson]));
    add('المرتجعات',['صف المصدر','مرجع المرتجع','التاريخ','Source Document','مرجع التسليم المطابق','العميل','المنتج المرتبط','كمية التسليم المرجعية التقديرية','رقم الشاسيه','الحالة','نتيجة المطابقة'],m.returns.map(r=>[r.sourceRow,r.reference,fmtDate(r.date,true),r.sourceDocument,r.deliveryReference,r.customer||r.matchedDelivery?.customer,r.product,r.deliveryQuantity,r.chassis,r.status,r.matchStatus]));
    add('الفواتير - المصدر',['صف المصدر','التاريخ','البائع','المنتج','العميل','Origin','الكمية المسجلة','القيمة المسجلة','التصنيف'],m.invoices.map(r=>[r.sourceRow,fmtDate(r.date,true),r.salesperson,r.product,r.customer,r.origin,r.quantity,r.amount,categoryLabel(categoryOf(r.product))]));
    add('التسليمات - المصدر',['صف المصدر','التاريخ','Source Document','Reference','العميل','المنتج','الكمية','رقم الشاسيه','حالة الفوترة','البائع المستدل'],m.deliveries.map(r=>[r.sourceRow,fmtDate(r.date,true),r.sourceDocument,r.reference,r.customer,r.product,r.quantity,r.chassis,r.hasInvoice?'مفوتر':'غير مفوتر',r.salesperson]));
    add('المرتجعات - المصدر',['صف المصدر','Reference','التاريخ','العميل','Source Document','الحالة','المنتج المرتبط','كمية التسليم المرجعية التقديرية'],m.returns.map(r=>[r.sourceRow,r.reference,fmtDate(r.date,true),r.customer,r.sourceDocument,r.status,r.product,r.deliveryQuantity]));
    add('تصنيف المنتجات',['المنتج','التصنيف'],m.products.map(p=>[p,categoryLabel(categoryOf(p))]));
    XLSX.writeFile(wb,`تقرير_تحليل_المبيعات_${new Date().toISOString().slice(0,10)}.xlsx`);
  }

  // Navigation and filters
  function activateView(name) {
    $$('.nav-item').forEach(btn=>btn.classList.toggle('active',btn.dataset.view===name));
    $$('.view').forEach(view=>view.classList.toggle('active',view.id===`view-${name}`));
  }
  $$('.nav-item').forEach(btn=>btn.addEventListener('click',()=>activateView(btn.dataset.view)));
  $$('[data-go-view]').forEach(btn=>btn.addEventListener('click',()=>activateView(btn.dataset.goView)));
  for (const [inputId, type] of [['invoice-file','invoices'],['delivery-file','deliveries'],['return-file','returns']]) {
    document.getElementById(inputId).addEventListener('change', e=>{
      state.files[type] = e.target.files && e.target.files[0] ? e.target.files[0] : null;
      state.names[type] = state.files[type]?.name || '';
      refreshFileCardLabels();
      $('#analyze-files').disabled = !Object.values(state.files).every(Boolean);
      if (state.files[type]) setStatus('تم اختيار ملف — أكمل الملفات الثلاثة');
    });
  }
  $('#analyze-files').addEventListener('click', analyzeSelectedFiles);
  $('#export-workbook').addEventListener('click', exportFullWorkbook);
  $('#export-products').addEventListener('click', exportProductsCSV);
  $('#export-sellers').addEventListener('click', exportSellersCSV);
  $('#export-seller-products').addEventListener('click', exportSellerProductsCSV);
  $('#export-uninvoiced').addEventListener('click', exportUninvoicedCSV);
  $('#export-returns').addEventListener('click', exportReturnsCSV);
  $('#export-classification').addEventListener('click', exportClassificationCSV);
  $('#export-details').addEventListener('click', exportDetailsCSV);
  $('#product-search').addEventListener('input',()=>state.data&&renderProducts(state.model));
  $('#product-category-filter').addEventListener('change',()=>state.data&&renderProducts(state.model));
  $('#seller-product-search').addEventListener('input',()=>state.data&&renderSellerProducts(state.model));
  $('#seller-filter').addEventListener('change',()=>state.data&&renderSellerProducts(state.model));
  $('#seller-category-filter').addEventListener('change',()=>state.data&&renderSellerProducts(state.model));
  $('#classification-search').addEventListener('input',()=>state.data&&renderClassification(state.model));
  $('#details-source').addEventListener('change',()=>state.data&&renderDetails(state.model));
  $('#details-search').addEventListener('input',()=>state.data&&renderDetails(state.model));
  $('#reset-classification').addEventListener('click',()=>{
    state.classifications={}; saveClassifications(); if(state.data)renderAll();
    showImportMessage('تمت إعادة التصنيف المبدئي وفق أسماء المنتجات.','success');
  });
})();
