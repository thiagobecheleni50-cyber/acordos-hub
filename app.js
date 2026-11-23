/* ==========================================================================
   AcordosHub — Application Logic, XLSX/CSV Import, Live Sheets & Reports
   ========================================================================== */

(() => {
  'use strict';

  const DEFAULT_SHEET_ID = '1FXKNHPd_YwR2C-zSDaXlFiH3EMEjM5GvDcv8ykW8gK8';
  const DEFAULT_SHEET_GID = '2121317474';
  const STORAGE_KEYS = {
    CONFIG: 'acordoshub_config_v2',
    DATA: 'acordoshub_local_data_v2'
  };

  const DEFAULT_HEADERS = [
    'Código',
    'Parte / Envolvido',
    'Objeto do Acordo',
    'Categoria',
    'Responsável',
    'Valor (R$)',
    'Data Início',
    'Prazo / Vencimento',
    'Data Pagamento',
    'Situação Pagamento',
    'Progresso (%)',
    'Status'
  ];

  const DEMO_ROWS = [
    {
      _id: 'demo-1', _rowNumber: 2,
      'Código': 'ACD-2026-001',
      'Parte / Envolvido': 'Construtora Horizonte Ltda',
      'Objeto do Acordo': 'Readequação de cronograma de entrega de obras civis Fase II',
      'Categoria': 'Contratual',
      'Responsável': 'Mariana Silva',
      'Valor (R$)': 245000,
      'Data Início': '2026-06-15',
      'Prazo / Vencimento': '2026-12-10',
      'Data Pagamento': '2026-11-18',
      'Situação Pagamento': 'Pago em Dia',
      'Progresso (%)': 100,
      'Status': 'Concluído'
    },
    {
      _id: 'demo-2', _rowNumber: 3,
      'Código': 'ACD-2026-002',
      'Parte / Envolvido': 'Sindicato dos Técnicos Industriais',
      'Objeto do Acordo': 'Acordo coletivo de banco de horas e adicional operacional',
      'Categoria': 'Trabalhista',
      'Responsável': 'Carlos Eduardo',
      'Valor (R$)': 88500,
      'Data Início': '2026-07-02',
      'Prazo / Vencimento': '2026-11-10',
      'Data Pagamento': '2026-11-08',
      'Situação Pagamento': 'Pago em Dia',
      'Progresso (%)': 100,
      'Status': 'Concluído'
    },
    {
      _id: 'demo-3', _rowNumber: 4,
      'Código': 'ACD-2026-003',
      'Parte / Envolvido': 'TechCloud Sistemas S.A.',
      'Objeto do Acordo': 'Compensação de SLA e migração sem custos de ambiente',
      'Categoria': 'Comercial',
      'Responsável': 'Fernanda Costa',
      'Valor (R$)': 134000,
      'Data Início': '2026-08-10',
      'Prazo / Vencimento': '2026-11-12',
      'Data Pagamento': '',
      'Situação Pagamento': 'Vencido',
      'Progresso (%)': 42,
      'Status': 'Atrasado'
    },
    {
      _id: 'demo-4', _rowNumber: 5,
      'Código': 'ACD-2026-004',
      'Parte / Envolvido': 'Secretaria Municipal de Infraestrutura',
      'Objeto do Acordo': 'Termo de cooperação técnica para pavimentação e drenagem',
      'Categoria': 'Institucional',
      'Responsável': 'Rafael Mendes',
      'Valor (R$)': 420000,
      'Data Início': '2026-08-24',
      'Prazo / Vencimento': '2027-02-28',
      'Data Pagamento': '2026-10-28',
      'Situação Pagamento': 'Pago em Dia',
      'Progresso (%)': 75,
      'Status': 'Em Andamento'
    },
    {
      _id: 'demo-5', _rowNumber: 6,
      'Código': 'ACD-2026-005',
      'Parte / Envolvido': 'Distribuidora Atlas Logística',
      'Objeto do Acordo': 'Renegociação de débitos em aberto em 6 parcelas fixas',
      'Categoria': 'Financeiro',
      'Responsável': 'Mariana Silva',
      'Valor (R$)': 67200,
      'Data Início': '2026-09-05',
      'Prazo / Vencimento': '2026-12-05',
      'Data Pagamento': '2026-11-20',
      'Situação Pagamento': 'Pago em Dia',
      'Progresso (%)': 85,
      'Status': 'Atenção'
    },
    {
      _id: 'demo-6', _rowNumber: 7,
      'Código': 'ACD-2026-006',
      'Parte / Envolvido': 'Instituto Verde Ambiental',
      'Objeto do Acordo': 'Compromisso de compensação florestal e reflorestamento',
      'Categoria': 'Regulatório',
      'Responsável': 'Juliana Rocha',
      'Valor (R$)': 195000,
      'Data Início': '2026-09-18',
      'Prazo / Vencimento': '2027-01-20',
      'Data Pagamento': '',
      'Situação Pagamento': 'Em Aberto',
      'Progresso (%)': 25,
      'Status': 'Aguardando Assinatura'
    },
    {
      _id: 'demo-7', _rowNumber: 8,
      'Código': 'ACD-2026-007',
      'Parte / Envolvido': 'Grupo Energia Sul',
      'Objeto do Acordo': 'Revisão tarifária e crédito de energia incentivada',
      'Categoria': 'Comercial',
      'Responsável': 'Carlos Eduardo',
      'Valor (R$)': 312000,
      'Data Início': '2026-10-03',
      'Prazo / Vencimento': '2026-11-01',
      'Data Pagamento': '2026-10-30',
      'Situação Pagamento': 'Pago em Dia',
      'Progresso (%)': 100,
      'Status': 'Concluído'
    },
    {
      _id: 'demo-8', _rowNumber: 9,
      'Código': 'ACD-2026-008',
      'Parte / Envolvido': 'Consultoria Nexus Jurídico',
      'Objeto do Acordo': 'Encerramento amigável de demanda cível nº 08412-90',
      'Categoria': 'Judicial',
      'Responsável': 'Fernanda Costa',
      'Valor (R$)': 54000,
      'Data Início': '2026-10-19',
      'Prazo / Vencimento': '2026-11-05',
      'Data Pagamento': '2026-11-11',
      'Situação Pagamento': 'Pago com Atraso',
      'Progresso (%)': 100,
      'Status': 'Concluído'
    },
    {
      _id: 'demo-9', _rowNumber: 10,
      'Código': 'ACD-2026-009',
      'Parte / Envolvido': 'Alpha Equipamentos Médicos',
      'Objeto do Acordo': 'Substituição em garantia de lote de monitores multiparâmetro',
      'Categoria': 'Contratual',
      'Responsável': 'Rafael Mendes',
      'Valor (R$)': 118400,
      'Data Início': '2026-11-04',
      'Prazo / Vencimento': '2026-11-18',
      'Data Pagamento': '',
      'Situação Pagamento': 'Vencido',
      'Progresso (%)': 35,
      'Status': 'Atrasado'
    },
    {
      _id: 'demo-10', _rowNumber: 11,
      'Código': 'ACD-2026-010',
      'Parte / Envolvido': 'Banco Fomento Regional',
      'Objeto do Acordo': 'Reperfilamento de linha de capital de giro com carência',
      'Categoria': 'Financeiro',
      'Responsável': 'Juliana Rocha',
      'Valor (R$)': 510000,
      'Data Início': '2026-11-11',
      'Prazo / Vencimento': '2027-03-15',
      'Data Pagamento': '2026-11-21',
      'Situação Pagamento': 'Pago em Dia',
      'Progresso (%)': 95,
      'Status': 'Concluído'
    }
  ];

  const state = {
    config: {
      sheetId: DEFAULT_SHEET_ID,
      sheetGid: DEFAULT_SHEET_GID,
      appsScriptUrl: '',
      mode: 'demo',
      syncInterval: 30
    },
    headers: [...DEFAULT_HEADERS],
    rows: [],
    schema: {},
    xlsxWorkbook: null, // cached multi-sheet workbook
    customReports: {
      total: null,  // { fileName, headers, rows, schema }
      onTime: null  // { fileName, headers, rows, schema }
    },
    activeReportKind: 'total',
    filters: {
      search: '',
      status: '',
      payment: '',
      category: '',
      owner: '',
      dateFrom: '',
      dateTo: '',
      quick: 'all'
    },
    sort: { column: null, dir: 'asc' },
    selectedIds: new Set(),
    pendingAction: null,
    syncTimer: null,
    lastSyncAt: new Date()
  };

  /* -------------------- Schema Detection & Parsing -------------------- */
  function normalizeKey(str) {
    return String(str || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  }

  function detectSchema(headers) {
    const find = (patterns, exclude = []) => headers.find(h => {
      const n = normalizeKey(h);
      if (exclude.some(ex => n.includes(ex))) return false;
      return patterns.some(p => n.includes(p));
    }) || null;

    const idCol = find(['codigo', 'cod', 'id', 'numero', 'num', 'protocolo', 'processo']) || headers[0] || null;
    const partyCol = find(['parte', 'envolvido', 'empresa', 'fornecedor', 'cliente', 'orgao', 'contraparte', 'nome']) || headers[1] || headers[0] || null;
    const descCol = find(['objeto', 'descricao', 'assunto', 'titulo', 'acordo', 'resumo']) || headers[2] || null;
    const categoryCol = find(['categoria', 'tipo', 'area', 'natureza', 'setor', 'modalidade']) || null;
    const ownerCol = find(['responsavel', 'gestor', 'dono', 'analista', 'coordenador']) || null;
    const valueCol = find(['valor', 'montante', 'custo', 'preco', 'orcamento', 'r$', 'total', 'financeiro']) || null;
    const startCol = find(['inicio', 'abertura', 'assinatura', 'criacao', 'emissao']) || null;
    const dueCol = find(['prazo', 'vencimento', 'termino', 'fim', 'limite', 'vigencia']) || null;
    const payDateCol = find(['data pagamento', 'data de pagamento', 'pago em', 'dt pagamento', 'liquidacao', 'quitacao']) || null;
    const payStatusCol = find(['situacao pagamento', 'pagamento', 'adimplencia', 'pontualidade', 'em dia'], ['data', 'dt', 'valor']) || null;
    const progressCol = find(['progresso', 'avanco', 'cumprimento', '%', 'percentual']) || null;
    const statusCol = find(['status', 'situacao', 'etapa', 'fase', 'estado'], ['pagamento']) || null;

    return { idCol, partyCol, descCol, categoryCol, ownerCol, valueCol, startCol, dueCol, payDateCol, payStatusCol, progressCol, statusCol };
  }

  function parseNumber(val) {
    if (typeof val === 'number' && !Number.isNaN(val)) return val;
    if (val === null || val === undefined) return 0;
    let s = String(val).replace(/[R$\s%]/g, '').trim();
    if (!s) return 0;
    if (s.includes(',') && s.includes('.')) {
      s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    } else if (s.includes(',')) {
      s = s.replace(',', '.');
    }
    const num = parseFloat(s);
    return Number.isNaN(num) ? 0 : num;
  }

  function formatBRL(val) {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(parseNumber(val));
  }

  function parseDateValue(val) {
    if (!val) return null;
    const str = String(val).trim();
    if (!str) return null;
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
      const [y, m, d] = str.slice(0, 10).split('-').map(Number);
      return new Date(y, m - 1, d);
    }
    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) {
      const [d, m, y] = str.split('/').map(Number);
      return new Date(y, m - 1, d);
    }
    const parsed = new Date(str);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function formatDateBR(val) {
    const dt = parseDateValue(val);
    if (!dt) return val ? String(val) : '—';
    return dt.toLocaleDateString('pt-BR');
  }

  function escapeHtml(str) {
    return String(str === null || str === undefined ? '' : str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function getStatusClass(statusStr) {
    const s = normalizeKey(statusStr);
    if (s.includes('em dia') || s.includes('conclu') || s.includes('cumpr') || s.includes('finaliz') || s.includes('quitad')) return 'status-concluido';
    if (s.includes('atras') || s.includes('vencid') || s.includes('crit') || s.includes('cancel') || s.includes('inadimp')) return 'status-atrasado';
    if (s.includes('atenc') || s.includes('alert') || s.includes('pend') || s.includes('aberto') || s.includes('risco')) return 'status-atencao';
    if (s.includes('aguard') || s.includes('analise') || s.includes('assin')) return 'status-aguardando';
    return 'status-andamento';
  }

  function getStatusColor(statusStr) {
    const map = {
      'status-concluido': '#10b981',
      'status-andamento': '#2563eb',
      'status-aguardando': '#06b6d4',
      'status-atencao': '#f59e0b',
      'status-atrasado': '#dc2626'
    };
    return map[getStatusClass(statusStr)] || '#ea580c';
  }

  function isRowCompleted(row, schema = state.schema) {
    const st = schema.statusCol ? normalizeKey(row[schema.statusCol]) : '';
    return st.includes('conclu') || st.includes('cumpr') || st.includes('finaliz') || st.includes('quitad');
  }

  function isRowPaidOnTime(row, schema = state.schema) {
    const payStatus = schema.payStatusCol ? normalizeKey(row[schema.payStatusCol]) : '';
    const status = schema.statusCol ? normalizeKey(row[schema.statusCol]) : '';

    if (payStatus) {
      if (payStatus.includes('em dia') || payStatus.includes('pontual') || payStatus.includes('no prazo')) return true;
      if (payStatus.includes('atras') || payStatus.includes('vencid') || payStatus.includes('aberto') || payStatus.includes('pendent')) return false;
      if (payStatus.includes('pago') || payStatus.includes('quitad')) {
        // Verify dates if available
        const paidDt = schema.payDateCol ? parseDateValue(row[schema.payDateCol]) : null;
        const dueDt = schema.dueCol ? parseDateValue(row[schema.dueCol]) : null;
        if (paidDt && dueDt) return paidDt <= dueDt;
        return true;
      }
    }

    // Fallback: compare payment date vs due date
    const paidDt = schema.payDateCol ? parseDateValue(row[schema.payDateCol]) : null;
    const dueDt = schema.dueCol ? parseDateValue(row[schema.dueCol]) : null;
    if (paidDt && dueDt) return paidDt <= dueDt;

    if (status.includes('pago em dia') || status.includes('em dia')) return true;
    return false;
  }

  function getDaysUntilDue(row, schema = state.schema) {
    if (!schema.dueCol) return null;
    const due = parseDateValue(row[schema.dueCol]);
    if (!due) return null;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return Math.round((due - today) / 86400000);
  }

  function isRowOverdue(row, schema = state.schema) {
    if (isRowCompleted(row, schema) || isRowPaidOnTime(row, schema)) return false;
    const st = schema.statusCol ? normalizeKey(row[schema.statusCol]) : '';
    const pst = schema.payStatusCol ? normalizeKey(row[schema.payStatusCol]) : '';
    if (st.includes('atras') || st.includes('vencid') || pst.includes('vencid')) return true;
    const days = getDaysUntilDue(row, schema);
    return days !== null && days < 0;
  }

  /* -------------------- CSV Parser -------------------- */
  function parseCSV(text) {
    const clean = text.replace(/^\uFEFF/, '');
    const firstLine = clean.split(/\r?\n/)[0] || '';
    const commaCount = (firstLine.match(/,/g) || []).length;
    const semiCount = (firstLine.match(/;/g) || []).length;
    const tabCount = (firstLine.match(/\t/g) || []).length;
    let delim = ',';
    if (tabCount > commaCount && tabCount > semiCount) delim = '\t';
    else if (semiCount > commaCount) delim = ';';

    const rows = [];
    let row = [], field = '', inQuotes = false;
    for (let i = 0; i < clean.length; i++) {
      const ch = clean[i], next = clean[i + 1];
      if (inQuotes) {
        if (ch === '"' && next === '"') { field += '"'; i++; }
        else if (ch === '"') inQuotes = false;
        else field += ch;
      } else {
        if (ch === '"') inQuotes = true;
        else if (ch === delim) { row.push(field.trim()); field = ''; }
        else if (ch === '\r' && next === '\n') { row.push(field.trim()); rows.push(row); row = []; field = ''; i++; }
        else if (ch === '\n') { row.push(field.trim()); rows.push(row); row = []; field = ''; }
        else field += ch;
      }
    }
    if (field.length > 0 || row.length > 0) { row.push(field.trim()); rows.push(row); }
    const nonEmpty = rows.filter(r => r.some(c => c !== ''));
    if (!nonEmpty.length) return { headers: [], rows: [] };

    const headers = nonEmpty[0].map((h, i) => h || `Coluna ${i + 1}`);
    const dataRows = nonEmpty.slice(1).map((r, idx) => {
      const obj = { _id: `csv-${idx + 2}`, _rowNumber: idx + 2 };
      headers.forEach((h, ci) => obj[h] = r[ci] !== undefined ? r[ci] : '');
      return obj;
    });
    return { headers, rows: dataRows };
  }

  async function parseUploadedFile(file) {
    const lower = file.name.toLowerCase();
    if (lower.endsWith('.xlsx')) {
      const buf = await file.arrayBuffer();
      const wb = await window.XLSXEngine.readXLSX(buf);
      const firstNonEmpty = wb.sheets.find(s => s.headers.length > 0) || wb.sheets[0];
      return { type: 'xlsx', workbook: wb, headers: firstNonEmpty.headers, rows: firstNonEmpty.rows, sheetName: firstNonEmpty.name };
    }
    const text = await file.text();
    const parsed = parseCSV(text);
    return { type: 'csv', workbook: null, headers: parsed.headers, rows: parsed.rows, sheetName: file.name };
  }

  /* -------------------- Persistence & Live Google Sheets Sync -------------------- */
  function loadConfig() {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEYS.CONFIG) || 'null');
      if (saved) state.config = { ...state.config, ...saved };
    } catch (_) {}
  }

  function saveConfig() { localStorage.setItem(STORAGE_KEYS.CONFIG, JSON.stringify(state.config)); }

  function saveLocalData() {
    try {
      localStorage.setItem(STORAGE_KEYS.DATA, JSON.stringify({ headers: state.headers, rows: state.rows }));
    } catch (_) {}
  }

  function loadLocalOrDemoData() {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEYS.DATA) || 'null');
      if (saved && Array.isArray(saved.headers) && Array.isArray(saved.rows) && saved.headers.length > 0) {
        state.headers = saved.headers;
        state.rows = saved.rows;
        state.schema = detectSchema(state.headers);
        return;
      }
    } catch (_) {}
    state.headers = [...DEFAULT_HEADERS];
    state.rows = DEMO_ROWS.map(r => ({ ...r }));
    state.schema = detectSchema(state.headers);
  }

  async function syncFromSheet(silent = false) {
    const { mode, appsScriptUrl, sheetId, sheetGid } = state.config;
    if (mode === 'demo') {
      updateConnectionUI('demo', 'Modo Demonstração');
      state.lastSyncAt = new Date();
      renderAll();
      return;
    }

    document.getElementById('syncPillText').textContent = 'Sincronizando...';

    try {
      if (mode === 'apps_script' && appsScriptUrl) {
        const res = await fetch(appsScriptUrl, { method: 'GET', cache: 'no-store' });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const json = await res.json();
        if (!json.ok) throw new Error(json.error || 'Erro no Apps Script');

        state.headers = json.headers.length ? json.headers : [...DEFAULT_HEADERS];
        state.rows = (json.rows || []).map((r, i) => ({ _id: `gs-${r._rowNumber || i + 2}`, ...r }));
        state.schema = detectSchema(state.headers);
        state.lastSyncAt = new Date();
        saveLocalData();

        const title = json.spreadsheetTitle ? `${json.spreadsheetTitle} — ${json.sheetName}` : `Planilha Conectada (${json.sheetName || 'Aba ativa'})`;
        document.getElementById('sheetSubtitle').textContent = title;
        updateConnectionUI('live', 'Planilha Ao Vivo (Bidirecional)');
        if (!silent) showToast('Planilha sincronizada em tempo real!', 'success');
        renderAll();
        return;
      }

      if (mode === 'public_csv' && sheetId) {
        const url = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/export?format=csv&gid=${encodeURIComponent(sheetGid || '0')}&_t=${Date.now()}`;
        const res = await fetch(url, { cache: 'no-store' });
        if (!res.ok) throw new Error(`Acesso negado (HTTP ${res.status}). Verifique o compartilhamento ou use Apps Script.`);
        const text = await res.text();
        if (text.trim().startsWith('<!DOCTYPE') || text.includes('ServiceLogin')) {
          throw new Error('Planilha privada: ative "Qualquer pessoa com o link (Leitor)" ou conecte via Apps Script.');
        }
        const parsed = parseCSV(text);
        if (!parsed.headers.length) throw new Error('A planilha retornou vazia.');
        state.headers = parsed.headers;
        state.rows = parsed.rows;
        state.schema = detectSchema(state.headers);
        state.lastSyncAt = new Date();
        saveLocalData();
        updateConnectionUI('live', 'Planilha Ao Vivo (Leitura)');
        if (!silent) showToast('Dados atualizados direto da planilha Google!', 'success');
        renderAll();
      }
    } catch (err) {
      updateConnectionUI('error', 'Erro na sincronização');
      if (!silent) showToast(err.message || 'Falha ao conectar à planilha.', 'error');
    }
  }

  function setupAutoSync() {
    if (state.syncTimer) { clearInterval(state.syncTimer); state.syncTimer = null; }
    const secs = Number(state.config.syncInterval || 0);
    if (secs > 0 && state.config.mode !== 'demo') {
      state.syncTimer = setInterval(() => syncFromSheet(true), secs * 1000);
    }
  }

  function updateConnectionUI(status, label) {
    const pill = document.getElementById('syncPill');
    const pillText = document.getElementById('syncPillText');
    pill.classList.remove('live', 'error');
    pillText.textContent = label;
    if (status === 'live') pill.classList.add('live');
    else if (status === 'error') pill.classList.add('error');
  }

  /* -------------------- Filtering & Sorting -------------------- */
  function getFilteredRows() {
    const { search, status, payment, category, owner, dateFrom, dateTo, quick } = state.filters;
    const { statusCol, payStatusCol, categoryCol, ownerCol, valueCol, dueCol } = state.schema;
    const q = normalizeKey(search);

    let list = state.rows.filter(row => {
      if (q) {
        const hay = state.headers.map(h => normalizeKey(row[h])).join(' ');
        if (!hay.includes(q)) return false;
      }
      if (status && statusCol && String(row[statusCol] || '') !== status) return false;
      if (payment && payStatusCol && String(row[payStatusCol] || '') !== payment) return false;
      if (category && categoryCol && String(row[categoryCol] || '') !== category) return false;
      if (owner && ownerCol && String(row[ownerCol] || '') !== owner) return false;

      if ((dateFrom || dateTo) && dueCol) {
        const dt = parseDateValue(row[dueCol]);
        if (!dt) return false;
        if (dateFrom && dt < parseDateValue(dateFrom)) return false;
        if (dateTo && dt > parseDateValue(dateTo)) return false;
      }

      if (quick === 'ontime') {
        if (!isRowPaidOnTime(row)) return false;
      } else if (quick === 'due30') {
        const d = getDaysUntilDue(row);
        if (isRowCompleted(row) || d === null || d < 0 || d > 30) return false;
      } else if (quick === 'overdue') {
        if (!isRowOverdue(row)) return false;
      } else if (quick === 'highvalue') {
        if (!valueCol || parseNumber(row[valueCol]) < 100000) return false;
      } else if (quick === 'completed') {
        if (!isRowCompleted(row)) return false;
      }

      return true;
    });

    if (state.sort.column) {
      const col = state.sort.column;
      const mult = state.sort.dir === 'asc' ? 1 : -1;
      list = [...list].sort((a, b) => {
        if (col === state.schema.valueCol || col === state.schema.progressCol) {
          return (parseNumber(a[col]) - parseNumber(b[col])) * mult;
        }
        if (col === state.schema.startCol || col === state.schema.dueCol || col === state.schema.payDateCol) {
          return ((parseDateValue(a[col])?.getTime() || 0) - (parseDateValue(b[col])?.getTime() || 0)) * mult;
        }
        return String(a[col] || '').localeCompare(String(b[col] || ''), 'pt-BR', { numeric: true }) * mult;
      });
    }

    return list;
  }

  /* -------------------- Report Datasets (Totais & Pagos em Dia) -------------------- */
  function getTotalReportData() {
    if (state.customReports.total) return state.customReports.total;
    return {
      fileName: null,
      headers: state.headers,
      rows: state.rows,
      schema: state.schema
    };
  }

  function getOnTimeReportData() {
    if (state.customReports.onTime) return state.customReports.onTime;
    const onTimeRows = state.rows.filter(r => isRowPaidOnTime(r, state.schema));
    return {
      fileName: null,
      headers: state.headers,
      rows: onTimeRows,
      schema: state.schema
    };
  }

  /* -------------------- Render Hero, Reports Hub & KPIs -------------------- */
  function renderHeroAndReports() {
    const totalData = getTotalReportData();
    const onTimeData = getOnTimeReportData();

    const totalCount = totalData.rows.length;
    const totalVal = totalData.schema.valueCol
      ? totalData.rows.reduce((acc, r) => acc + parseNumber(r[totalData.schema.valueCol]), 0)
      : 0;
    const avgTicket = totalCount > 0 ? totalVal / totalCount : 0;

    const onTimeCount = onTimeData.rows.length;
    const onTimeVal = onTimeData.schema.valueCol
      ? onTimeData.rows.reduce((acc, r) => acc + parseNumber(r[onTimeData.schema.valueCol]), 0)
      : 0;
    const punctualityPct = totalCount > 0 ? Math.min(100, Math.round((onTimeCount / totalCount) * 100)) : 0;

    // Hero Presentation 4-card stats
    document.getElementById('heroMetricsPanel').innerHTML = `
      <div class="hero-stat">
        <div class="hero-stat-lbl">Acordos Totais</div>
        <div class="hero-stat-val">${totalCount}</div>
        <div class="hero-stat-sub">Base consolidada ativa</div>
      </div>
      <div class="hero-stat">
        <div class="hero-stat-lbl">Volume Financeiro</div>
        <div class="hero-stat-val">${formatBRL(totalVal)}</div>
        <div class="hero-stat-sub">Ticket médio ${formatBRL(avgTicket)}</div>
      </div>
      <div class="hero-stat">
        <div class="hero-stat-lbl">Pagos em Dia</div>
        <div class="hero-stat-val" style="color:#6ee7b7;">${onTimeCount}</div>
        <div class="hero-stat-sub">${formatBRL(onTimeVal)} quitados no prazo</div>
      </div>
      <div class="hero-stat">
        <div class="hero-stat-lbl">Taxa de Pontualidade</div>
        <div class="hero-stat-val" style="color:#fdba74;">${punctualityPct}%</div>
        <div class="hero-stat-sub">Adimplência em dia</div>
      </div>
    `;

    // Relatório de Acordos Totais Card
    document.getElementById('reportTotalSource').textContent = totalData.fileName
      ? `Arquivo dedicado: ${totalData.fileName}`
      : 'Fonte: Base principal consolidada';
    document.getElementById('btnClearReportTotalFile').style.display = totalData.fileName ? 'inline-flex' : 'none';
    document.getElementById('reportTotalMetrics').innerHTML = `
      <div class="report-mini-kpi">
        <div class="report-mini-lbl">Total Acordos</div>
        <div class="report-mini-val">${totalCount}</div>
      </div>
      <div class="report-mini-kpi">
        <div class="report-mini-lbl">Valor Total</div>
        <div class="report-mini-val">${formatBRL(totalVal)}</div>
      </div>
      <div class="report-mini-kpi">
        <div class="report-mini-lbl">Ticket Médio</div>
        <div class="report-mini-val">${formatBRL(avgTicket)}</div>
      </div>
    `;

    // Relatório de Acordos Pagos em Dia Card
    document.getElementById('reportOnTimeSource').textContent = onTimeData.fileName
      ? `Arquivo dedicado: ${onTimeData.fileName}`
      : 'Fonte: Filtrado automaticamente da base principal';
    document.getElementById('btnClearReportOnTimeFile').style.display = onTimeData.fileName ? 'inline-flex' : 'none';
    document.getElementById('reportOnTimeMetrics').innerHTML = `
      <div class="report-mini-kpi">
        <div class="report-mini-lbl">Pagos em Dia</div>
        <div class="report-mini-val" style="color:#059669;">${onTimeCount}</div>
      </div>
      <div class="report-mini-kpi">
        <div class="report-mini-lbl">Valor Pago em Dia</div>
        <div class="report-mini-val">${formatBRL(onTimeVal)}</div>
      </div>
      <div class="report-mini-kpi">
        <div class="report-mini-lbl">Pontualidade</div>
        <div class="report-mini-val" style="color:#059669;">${punctualityPct}%</div>
      </div>
    `;
    document.getElementById('onTimeProgressBar').style.width = `${punctualityPct}%`;
  }

  function renderKPIs(filtered) {
    const totalAll = state.rows.length;
    const total = filtered.length;
    const { valueCol } = state.schema;

    const totalValue = valueCol ? filtered.reduce((acc, r) => acc + parseNumber(r[valueCol]), 0) : 0;
    const avgValue = total > 0 ? totalValue / total : 0;
    const onTimeCount = filtered.filter(r => isRowPaidOnTime(r)).length;
    const onTimePct = total > 0 ? Math.round((onTimeCount / total) * 100) : 0;
    const overdueCount = filtered.filter(r => isRowOverdue(r)).length;
    const dueSoonCount = filtered.filter(r => {
      if (isRowCompleted(r)) return false;
      const d = getDaysUntilDue(r);
      return d !== null && d >= 0 && d <= 30;
    }).length;

    const kpis = [
      {
        label: 'Acordos Filtrados',
        value: String(total),
        sub: total === totalAll ? 'Base completa monitorada' : `de ${totalAll} registros totais`,
        accent: '#ea580c',
        iconBg: '#ffedd5',
        icon: `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#ea580c" stroke-width="2.2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>`
      },
      {
        label: 'Valor Consolidado',
        value: formatBRL(totalValue),
        sub: `Ticket médio: ${formatBRL(avgValue)}`,
        accent: '#2563eb',
        iconBg: '#dbeafe',
        icon: `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.2"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>`
      },
      {
        label: 'Pagos em Dia',
        value: `${onTimeCount} (${onTimePct}%)`,
        sub: 'Adimplência pontual no filtro',
        accent: '#10b981',
        iconBg: '#d1fae5',
        progress: onTimePct,
        icon: `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>`
      },
      {
        label: 'Vencem em até 30d',
        value: String(dueSoonCount),
        sub: 'Acompanhamento prioritário',
        accent: '#f59e0b',
        iconBg: '#fef3c7',
        icon: `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>`
      },
      {
        label: 'Atrasados / Críticos',
        value: String(overdueCount),
        sub: overdueCount > 0 ? 'Requer ação imediata' : 'Nenhum atraso identificado',
        accent: '#dc2626',
        iconBg: '#fee2e2',
        icon: `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#dc2626" stroke-width="2.2"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`
      }
    ];

    document.getElementById('kpiGrid').innerHTML = kpis.map(k => `
      <article class="kpi-card">
        <div class="kpi-bar" style="background:${k.accent};"></div>
        <div class="kpi-top">
          <span class="kpi-label">${k.label}</span>
          <div class="kpi-icon" style="background:${k.iconBg};">${k.icon}</div>
        </div>
        <div class="kpi-value">${k.value}</div>
        <div class="kpi-sub">${k.sub}</div>
        ${k.progress !== undefined ? `
          <div class="kpi-progress-track">
            <div class="kpi-progress-fill" style="width:${k.progress}%; background:${k.accent};"></div>
          </div>
        ` : ''}
      </article>
    `).join('');

    const timeStr = (state.lastSyncAt || new Date()).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    document.getElementById('lastUpdatedLabel').textContent = `Última atualização às ${timeStr}`;
  }

  /* -------------------- Render Charts -------------------- */
  function renderCharts(filtered) {
    renderStatusDonut(filtered);
    renderCategoryBars(filtered);
    renderMonthlyChart(filtered);
  }

  function renderStatusDonut(filtered) {
    const container = document.getElementById('chartStatusDonut');
    document.getElementById('tagStatusCount').textContent = `${filtered.length} acordo(s)`;

    const col = state.schema.statusCol || state.schema.payStatusCol;
    if (!col || filtered.length === 0) {
      container.innerHTML = `<div class="empty-state" style="padding:24px 10px;">Sem dados para exibir.</div>`;
      return;
    }

    const counts = {};
    filtered.forEach(r => {
      const key = String(r[col] || 'Não informado').trim() || 'Não informado';
      counts[key] = (counts[key] || 0) + 1;
    });

    const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    const total = filtered.length;
    const radius = 56;
    const circumference = 2 * Math.PI * radius;
    let offset = 0;

    const segments = entries.map(([label, count]) => {
      const dash = (count / total) * circumference;
      const seg = `<circle cx="76" cy="76" r="${radius}" fill="transparent" stroke="${getStatusColor(label)}" stroke-width="20" stroke-dasharray="${dash} ${circumference - dash}" stroke-dashoffset="${-offset}" transform="rotate(-90 76 76)"></circle>`;
      offset += dash;
      return seg;
    }).join('');

    const legend = entries.map(([label, count]) => {
      const pct = Math.round((count / total) * 100);
      const activeStyle = state.filters.status === label ? 'background:var(--ember-50); outline:1px solid var(--ember-500);' : '';
      return `
        <div class="legend-item" data-filter-status="${escapeHtml(label)}" style="${activeStyle}">
          <span class="legend-left">
            <span class="legend-dot" style="background:${getStatusColor(label)};"></span>
            <span>${escapeHtml(label)}</span>
          </span>
          <span class="legend-val">${count} <span style="color:var(--text-muted);font-weight:500;">(${pct}%)</span></span>
        </div>
      `;
    }).join('');

    container.innerHTML = `
      <div class="donut-wrap">
        <div class="donut-svg-box">
          <svg width="152" height="152" viewBox="0 0 152 152" aria-hidden="true">
            <circle cx="76" cy="76" r="${radius}" fill="transparent" stroke="#f1f5f9" stroke-width="20"></circle>
            ${segments}
          </svg>
          <div class="donut-center">
            <span class="donut-center-num">${total}</span>
            <span class="donut-center-lbl">acordos</span>
          </div>
        </div>
        <div class="donut-legend">${legend}</div>
      </div>
    `;

    container.querySelectorAll('[data-filter-status]').forEach(el => {
      el.addEventListener('click', () => {
        const val = el.getAttribute('data-filter-status');
        state.filters.status = state.filters.status === val ? '' : val;
        document.getElementById('filterStatus').value = state.filters.status;
        renderAll();
      });
    });
  }

  function renderCategoryBars(filtered) {
    const container = document.getElementById('chartCategoryBars');
    const catCol = state.schema.categoryCol || state.schema.ownerCol;
    const valCol = state.schema.valueCol;

    if (!catCol || filtered.length === 0) {
      container.innerHTML = `<div class="empty-state" style="padding:24px 10px;">Sem dados de categoria.</div>`;
      return;
    }

    const totals = {};
    filtered.forEach(r => {
      const k = String(r[catCol] || 'Outros').trim() || 'Outros';
      totals[k] = (totals[k] || 0) + (valCol ? parseNumber(r[valCol]) : 1);
    });

    const entries = Object.entries(totals).sort((a, b) => b[1] - a[1]).slice(0, 6);
    const maxVal = Math.max(...entries.map(e => e[1]), 1);
    const palette = ['#dc2626', '#ea580c', '#f59e0b', '#10b981', '#2563eb', '#06b6d4'];

    container.innerHTML = `<div class="hbar-list">${entries.map(([name, val], idx) => `
      <div class="hbar-item">
        <div class="hbar-meta">
          <span class="hbar-name">${escapeHtml(name)}</span>
          <span class="hbar-val">${valCol ? formatBRL(val) : `${val} acordo(s)`}</span>
        </div>
        <div class="hbar-track">
          <div class="hbar-fill" style="width:${Math.max(6, Math.round((val / maxVal) * 100))}%; background:${palette[idx % palette.length]};"></div>
        </div>
      </div>
    `).join('')}</div>`;
  }

  function renderMonthlyChart(filtered) {
    const container = document.getElementById('chartMonthly');
    const dateCol = state.schema.startCol || state.schema.dueCol;

    if (!dateCol || filtered.length === 0) {
      container.innerHTML = `<div class="empty-state" style="padding:24px 10px;">Sem datas identificadas.</div>`;
      return;
    }

    const byMonth = {};
    filtered.forEach(r => {
      const dt = parseDateValue(r[dateCol]);
      if (!dt) return;
      const key = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}`;
      byMonth[key] = (byMonth[key] || 0) + 1;
    });

    const entries = Object.entries(byMonth).sort((a, b) => a[0].localeCompare(b[0])).slice(-7);
    if (!entries.length) {
      container.innerHTML = `<div class="empty-state" style="padding:24px 10px;">Nenhuma data válida.</div>`;
      return;
    }

    const maxCount = Math.max(...entries.map(e => e[1]), 1);
    const monthsPt = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

    container.innerHTML = `
      <div class="vchart-wrap">
        ${entries.map(([ym, count]) => `
          <div class="vbar-col" title="${ym}: ${count} acordo(s)">
            <span class="vbar-num">${count}</span>
            <div class="vbar-bar" style="height:${Math.max(12, Math.round((count / maxCount) * 100))}%;"></div>
          </div>
        `).join('')}
      </div>
      <div class="vbar-labels">
        ${entries.map(([ym]) => {
          const [y, m] = ym.split('-').map(Number);
          return `<div class="vbar-lbl">${monthsPt[m - 1]}/${String(y).slice(-2)}</div>`;
        }).join('')}
      </div>
    `;
  }

  /* -------------------- Filters & Table -------------------- */
  function populateFilterOptions() {
    const fillSelect = (id, col, defaultLabel, currentVal) => {
      const sel = document.getElementById(id);
      const unique = col
        ? [...new Set(state.rows.map(r => String(r[col] || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pt-BR'))
        : [];
      sel.innerHTML = `<option value="">${defaultLabel}</option>` + unique.map(v =>
        `<option value="${escapeHtml(v)}" ${v === currentVal ? 'selected' : ''}>${escapeHtml(v)}</option>`
      ).join('');
    };

    fillSelect('filterStatus', state.schema.statusCol, 'Todos os status', state.filters.status);
    fillSelect('filterPayment', state.schema.payStatusCol, 'Todas as situações', state.filters.payment);
    fillSelect('filterCategory', state.schema.categoryCol, 'Todas as categorias', state.filters.category);
    fillSelect('filterOwner', state.schema.ownerCol, 'Todos os responsáveis', state.filters.owner);
  }

  function renderTable(filtered) {
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');
    document.getElementById('tableCount').innerHTML = `Exibindo <strong>${filtered.length}</strong> de <strong>${state.rows.length}</strong> acordos`;

    const allChecked = filtered.length > 0 && filtered.every(r => state.selectedIds.has(r._id));
    const displayHeaders = state.headers.filter(h => !(state.headers.length > 10 && normalizeKey(h).includes('observac')));

    thead.innerHTML = `
      <tr>
        <th style="width:40px;"><input type="checkbox" id="chkSelectAll" ${allChecked ? 'checked' : ''} aria-label="Selecionar todos" /></th>
        ${displayHeaders.map(h => {
          const arrow = state.sort.column === h ? (state.sort.dir === 'asc' ? ' ▲' : ' ▼') : '';
          return `<th class="sortable" data-sort-col="${escapeHtml(h)}">${escapeHtml(h)}${arrow}</th>`;
        }).join('')}
        <th style="text-align:right; width:90px;">Ações</th>
      </tr>
    `;

    thead.querySelector('#chkSelectAll')?.addEventListener('change', (e) => {
      filtered.forEach(r => e.target.checked ? state.selectedIds.add(r._id) : state.selectedIds.delete(r._id));
      updateSelectionToolbar();
      renderTable(filtered);
    });

    thead.querySelectorAll('[data-sort-col]').forEach(th => {
      th.addEventListener('click', () => {
        const col = th.getAttribute('data-sort-col');
        state.sort.dir = (state.sort.column === col && state.sort.dir === 'asc') ? 'desc' : 'asc';
        state.sort.column = col;
        renderAll();
      });
    });

    if (filtered.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="${displayHeaders.length + 2}">
            <div class="empty-state">
              <h3>Nenhum acordo encontrado</h3>
              <p>Ajuste os filtros aplicados, importe uma planilha (.xlsx / .csv) ou clique em <strong>+ Incluir Acordo</strong>.</p>
            </div>
          </td>
        </tr>
      `;
      updateSelectionToolbar();
      return;
    }

    tbody.innerHTML = filtered.map(row => `
      <tr>
        <td><input type="checkbox" class="row-check" data-row-id="${escapeHtml(row._id)}" ${state.selectedIds.has(row._id) ? 'checked' : ''} aria-label="Selecionar registro" /></td>
        ${displayHeaders.map(h => renderTableCell(row, h)).join('')}
        <td style="text-align:right; white-space:nowrap;">
          <button type="button" class="btn btn-ghost-danger" data-delete-id="${escapeHtml(row._id)}" title="Excluir acordo">
            Excluir
          </button>
        </td>
      </tr>
    `).join('');

    tbody.querySelectorAll('.row-check').forEach(chk => {
      chk.addEventListener('change', (e) => {
        const id = e.target.getAttribute('data-row-id');
        e.target.checked ? state.selectedIds.add(id) : state.selectedIds.delete(id);
        updateSelectionToolbar();
      });
    });

    tbody.querySelectorAll('[data-delete-id]').forEach(btn => {
      btn.addEventListener('click', () => {
        const target = state.rows.find(r => r._id === btn.getAttribute('data-delete-id'));
        if (target) requestDeleteConfirmation([target]);
      });
    });

    updateSelectionToolbar();
  }

  function renderTableCell(row, header, schema = state.schema) {
    const val = row[header];
    const { idCol, partyCol, valueCol, startCol, dueCol, payDateCol, payStatusCol, progressCol, statusCol } = schema;

    if (header === idCol) return `<td><span class="cell-id">${escapeHtml(val || '—')}</span></td>`;
    if (header === partyCol) return `<td><div class="cell-primary">${escapeHtml(val || '—')}</div></td>`;
    if (header === valueCol) return `<td class="cell-money">${formatBRL(val)}</td>`;
    if (header === startCol || header === payDateCol) return `<td style="white-space:nowrap;">${formatDateBR(val)}</td>`;
    if (header === dueCol) {
      const days = getDaysUntilDue(row, schema);
      let badge = '';
      if (isRowPaidOnTime(row, schema)) {
        badge = `<div><span class="deadline-tag deadline-ontime">Em Dia</span></div>`;
      } else if (!isRowCompleted(row, schema) && days !== null) {
        if (days < 0) badge = `<div><span class="deadline-tag deadline-late">Atrasado ${Math.abs(days)}d</span></div>`;
        else if (days <= 30) badge = `<div><span class="deadline-tag deadline-soon">Vence em ${days}d</span></div>`;
      }
      return `<td style="white-space:nowrap;">${formatDateBR(val)}${badge}</td>`;
    }
    if (header === progressCol) {
      const pct = Math.min(100, Math.max(0, Math.round(parseNumber(val))));
      const color = pct >= 90 ? '#10b981' : pct >= 50 ? '#ea580c' : '#f59e0b';
      return `
        <td>
          <div class="mini-progress">
            <div class="mini-progress-bar"><div class="mini-progress-fill" style="width:${pct}%; background:${color};"></div></div>
            <span class="mini-progress-txt">${pct}%</span>
          </div>
        </td>
      `;
    }
    if (header === statusCol || header === payStatusCol) {
      return `<td><span class="status-badge ${getStatusClass(val)}"><span class="status-dot"></span>${escapeHtml(val || '—')}</span></td>`;
    }
    return `<td>${escapeHtml(val || '—')}</td>`;
  }

  function updateSelectionToolbar() {
    const btn = document.getElementById('btnDeleteSelected');
    const validSelected = [...state.selectedIds].filter(id => state.rows.some(r => r._id === id));
    document.getElementById('selectedCount').textContent = String(validSelected.length);
    btn.style.display = validSelected.length > 0 ? 'inline-flex' : 'none';
  }

  function renderAll() {
    renderHeroAndReports();
    populateFilterOptions();
    const filtered = getFilteredRows();
    renderKPIs(filtered);
    renderCharts(filtered);
    renderTable(filtered);
  }

  /* -------------------- Executive Report Modal -------------------- */
  function openReportModal(kind) {
    state.activeReportKind = kind;
    const isTotal = kind === 'total';
    const data = isTotal ? getTotalReportData() : getOnTimeReportData();
    const allTotalCount = getTotalReportData().rows.length;
    const totalVal = data.schema.valueCol
      ? data.rows.reduce((acc, r) => acc + parseNumber(r[data.schema.valueCol]), 0)
      : 0;
    const avgVal = data.rows.length > 0 ? totalVal / data.rows.length : 0;
    const pctShare = allTotalCount > 0 ? Math.round((data.rows.length / allTotalCount) * 100) : 0;

    document.getElementById('reportModalTitle').textContent = isTotal
      ? 'Relatório Gerencial — Acordos Totais'
      : 'Relatório Gerencial — Acordos Pagos em Dia';

    const headersToShow = data.headers.slice(0, 9);
    const bodyHtml = `
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; padding-bottom:10px; border-bottom:1px solid var(--border-subtle);">
        <div style="font-size:12.5px; color:var(--text-secondary);">
          <strong>Fonte:</strong> ${escapeHtml(data.fileName || (isTotal ? 'Base Principal Consolidada' : 'Filtro Automático de Acordos Pagos em Dia'))}
        </div>
        <div style="font-size:12px; color:var(--text-muted);">
          Gerado em ${new Date().toLocaleString('pt-BR')}
        </div>
      </div>

      <div class="report-metrics-row">
        <div class="report-mini-kpi">
          <div class="report-mini-lbl">${isTotal ? 'Total de Acordos' : 'Acordos Pagos em Dia'}</div>
          <div class="report-mini-val" style="color:${isTotal ? '#dc2626' : '#059669'};">${data.rows.length}</div>
        </div>
        <div class="report-mini-kpi">
          <div class="report-mini-lbl">${isTotal ? 'Valor Total Consolidado' : 'Valor Pago em Dia'}</div>
          <div class="report-mini-val">${formatBRL(totalVal)}</div>
        </div>
        <div class="report-mini-kpi">
          <div class="report-mini-lbl">${isTotal ? 'Ticket Médio' : 'Taxa de Pontualidade'}</div>
          <div class="report-mini-val">${isTotal ? formatBRL(avgVal) : `${pctShare}%`}</div>
        </div>
      </div>

      <div class="table-scroll" style="border:1px solid var(--border-subtle); border-radius:12px;">
        <table class="data-table" style="min-width:760px;">
          <thead>
            <tr>${headersToShow.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr>
          </thead>
          <tbody>
            ${data.rows.length ? data.rows.map(r => `
              <tr>${headersToShow.map(h => renderTableCell(r, h, data.schema)).join('')}</tr>
            `).join('') : `<tr><td colspan="${headersToShow.length}" style="text-align:center; padding:32px; color:var(--text-muted);">Nenhum registro neste relatório.</td></tr>`}
          </tbody>
        </table>
      </div>
    `;

    document.getElementById('reportModalBody').innerHTML = bodyHtml;
    openModal('modalReport');
  }

  /* -------------------- Export CSV & XLSX -------------------- */
  function downloadCSV(headers, rows, filename) {
    const esc = (v) => {
      const s = String(v === null || v === undefined ? '' : v);
      return /[",;\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const lines = [
      headers.map(esc).join(';'),
      ...rows.map(r => headers.map(h => esc(r[h])).join(';'))
    ];
    const blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    triggerDownload(blob, filename);
    showToast(`Arquivo ${filename} baixado!`, 'info');
  }

  function downloadXLSX(headers, rows, filename, sheetName) {
    const bytes = window.XLSXEngine.writeXLSX(headers, rows, sheetName);
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    triggerDownload(blob, filename);
    showToast(`Planilha ${filename} baixada!`, 'success');
  }

  function triggerDownload(blob, filename) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  /* -------------------- Add & Confirm Modals (Passwordless) -------------------- */
  function openAddModal() {
    const container = document.getElementById('dynamicFormFields');
    const { idCol, valueCol, startCol, dueCol, payDateCol, payStatusCol, progressCol, statusCol } = state.schema;
    const nextNum = String(state.rows.length + 1).padStart(3, '0');
    const todayIso = new Date().toISOString().slice(0, 10);

    const fieldsHtml = state.headers.map((h, idx) => {
      const inputId = `add_field_${idx}`;
      const norm = normalizeKey(h);

      if (h === statusCol) {
        return `
          <div class="field-group">
            <label class="field-label" for="${inputId}">${escapeHtml(h)}</label>
            <select id="${inputId}" name="${escapeHtml(h)}" class="select" required>
              ${['Em Andamento', 'Concluído', 'Aguardando Assinatura', 'Atenção', 'Atrasado'].map(s => `<option value="${s}">${s}</option>`).join('')}
            </select>
          </div>
        `;
      }
      if (h === payStatusCol) {
        return `
          <div class="field-group">
            <label class="field-label" for="${inputId}">${escapeHtml(h)}</label>
            <select id="${inputId}" name="${escapeHtml(h)}" class="select">
              ${['Pago em Dia', 'Em Aberto', 'Pago com Atraso', 'Vencido'].map(s => `<option value="${s}">${s}</option>`).join('')}
            </select>
          </div>
        `;
      }
      if (h === startCol || h === dueCol || h === payDateCol) {
        return `
          <div class="field-group">
            <label class="field-label" for="${inputId}">${escapeHtml(h)}</label>
            <input type="date" id="${inputId}" name="${escapeHtml(h)}" class="input" value="${h === startCol ? todayIso : ''}" />
          </div>
        `;
      }
      if (h === valueCol) {
        return `
          <div class="field-group">
            <label class="field-label" for="${inputId}">${escapeHtml(h)}</label>
            <input type="number" step="0.01" min="0" id="${inputId}" name="${escapeHtml(h)}" class="input" placeholder="Ex: 50000" required />
          </div>
        `;
      }
      if (h === progressCol) {
        return `
          <div class="field-group">
            <label class="field-label" for="${inputId}">${escapeHtml(h)}</label>
            <input type="number" min="0" max="100" id="${inputId}" name="${escapeHtml(h)}" class="input" value="0" />
          </div>
        `;
      }
      if (norm.includes('observac') || norm.includes('objeto') || norm.includes('descric')) {
        return `
          <div class="field-group" style="grid-column: 1 / -1;">
            <label class="field-label" for="${inputId}">${escapeHtml(h)}</label>
            <textarea id="${inputId}" name="${escapeHtml(h)}" class="textarea" rows="2"></textarea>
          </div>
        `;
      }
      return `
        <div class="field-group">
          <label class="field-label" for="${inputId}">${escapeHtml(h)}</label>
          <input type="text" id="${inputId}" name="${escapeHtml(h)}" class="input" value="${h === idCol ? `ACD-2026-${nextNum}` : ''}" ${idx < 2 ? 'required' : ''} />
        </div>
      `;
    }).join('');

    container.innerHTML = `<div class="form-grid-2">${fieldsHtml}</div>`;
    openModal('modalAdd');
  }

  function openActionConfirmation({ type, title, summaryHtml, buttonText, buttonClass, payload }) {
    state.pendingAction = { type, payload };
    document.getElementById('confirmModalTitle').textContent = title;
    const box = document.getElementById('confirmSummaryBox');
    box.className = type === 'delete' ? 'confirm-summary danger' : 'confirm-summary';
    box.innerHTML = summaryHtml;
    const btn = document.getElementById('btnConfirmAction');
    btn.textContent = buttonText;
    btn.className = `btn ${buttonClass}`;
    document.getElementById('confirmActionError').classList.remove('visible');
    openModal('modalConfirmAction');
    setTimeout(() => btn.focus(), 60);
  }

  function requestDeleteConfirmation(rowsToDelete) {
    const { idCol, partyCol, valueCol } = state.schema;
    const listHtml = rowsToDelete.slice(0, 4).map(r => {
      const code = idCol ? r[idCol] : r._id;
      const party = partyCol ? r[partyCol] : '';
      const val = valueCol ? ` — ${formatBRL(r[valueCol])}` : '';
      return `<div><strong>${escapeHtml(code)}</strong> ${escapeHtml(party)}${val}</div>`;
    }).join('');
    const extra = rowsToDelete.length > 4 ? `<div><em>+ ${rowsToDelete.length - 4} outro(s) registro(s)</em></div>` : '';

    openActionConfirmation({
      type: 'delete',
      title: 'Confirmar Exclusão',
      summaryHtml: `<div style="color:#9f1239;font-weight:700;">Atenção: você está prestes a excluir ${rowsToDelete.length} acordo(s):</div>${listHtml}${extra}`,
      buttonText: 'Confirmar Exclusão',
      buttonClass: 'btn-danger',
      payload: rowsToDelete
    });
  }

  async function executePendingAction() {
    const action = state.pendingAction;
    if (!action) return;
    const btn = document.getElementById('btnConfirmAction');
    btn.disabled = true;

    try {
      const isLiveAppsScript = state.config.mode === 'apps_script' && state.config.appsScriptUrl;

      if (action.type === 'add') {
        if (isLiveAppsScript) {
          const res = await fetch(state.config.appsScriptUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'add', data: action.payload })
          });
          const json = await res.json();
          if (!json.ok) throw new Error(json.error || 'Erro ao gravar na planilha.');
          closeModal('modalConfirmAction');
          showToast('Acordo incluído na planilha Google!', 'success');
          await syncFromSheet(true);
        } else {
          state.rows.unshift({ _id: 'local-' + Date.now(), _rowNumber: state.rows.length + 2, ...action.payload });
          saveLocalData();
          closeModal('modalConfirmAction');
          showToast('Acordo incluído com sucesso!', 'success');
          renderAll();
        }
      }

      if (action.type === 'delete') {
        const rowsToDelete = action.payload;
        const idsSet = new Set(rowsToDelete.map(r => r._id));
        if (isLiveAppsScript) {
          const rowNumbers = rowsToDelete.map(r => r._rowNumber).filter(Boolean);
          const res = await fetch(state.config.appsScriptUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'delete', rowNumbers })
          });
          const json = await res.json();
          if (!json.ok) throw new Error(json.error || 'Erro ao excluir da planilha.');
          rowsToDelete.forEach(r => state.selectedIds.delete(r._id));
          closeModal('modalConfirmAction');
          showToast(`${rowsToDelete.length} acordo(s) excluído(s) da planilha!`, 'success');
          await syncFromSheet(true);
        } else {
          state.rows = state.rows.filter(r => !idsSet.has(r._id));
          rowsToDelete.forEach(r => state.selectedIds.delete(r._id));
          saveLocalData();
          closeModal('modalConfirmAction');
          showToast(`${rowsToDelete.length} acordo(s) excluído(s)!`, 'success');
          renderAll();
        }
      }
    } catch (err) {
      const errEl = document.getElementById('confirmActionError');
      errEl.textContent = err.message || 'Erro ao executar operação.';
      errEl.classList.add('visible');
    } finally {
      btn.disabled = false;
    }
  }

  /* -------------------- File Dropzones & Handlers -------------------- */
  function setupDropzone(zoneId, onFile) {
    const zone = document.getElementById(zoneId);
    ['dragenter', 'dragover'].forEach(ev => zone.addEventListener(ev, (e) => {
      e.preventDefault();
      zone.classList.add('dragover');
    }));
    ['dragleave', 'drop'].forEach(ev => zone.addEventListener(ev, (e) => {
      e.preventDefault();
      zone.classList.remove('dragover');
    }));
    zone.addEventListener('drop', (e) => {
      const file = e.dataTransfer?.files?.[0];
      if (file) onFile(file);
    });
  }

  async function handleMainXlsxFile(file) {
    try {
      const parsed = await parseUploadedFile(file);
      if (!parsed.headers.length) throw new Error('Planilha .xlsx vazia.');
      state.xlsxWorkbook = parsed.workbook;
      state.headers = parsed.headers;
      state.rows = parsed.rows;
      state.schema = detectSchema(state.headers);
      state.selectedIds.clear();
      state.lastSyncAt = new Date();
      saveLocalData();

      document.getElementById('xlsxFileLabel').innerHTML = `<strong>${escapeHtml(file.name)}</strong> (${parsed.rows.length} linhas)`;
      document.getElementById('sheetSubtitle').textContent = `Excel carregado: ${file.name} (${parsed.sheetName})`;
      updateConnectionUI('live', `Excel: ${file.name.slice(0, 18)}`);

      const sel = document.getElementById('xlsxSheetSelector');
      if (parsed.workbook && parsed.workbook.sheets.length > 1) {
        sel.innerHTML = parsed.workbook.sheets.map((s, i) =>
          `<option value="${i}">${escapeHtml(s.name)} (${s.rows.length})</option>`
        ).join('');
        sel.style.display = 'inline-block';
      } else {
        sel.style.display = 'none';
      }

      showToast(`${parsed.rows.length} acordos importados de ${file.name}!`, 'success');
      renderAll();
    } catch (err) {
      showToast(err.message || 'Erro ao ler arquivo .xlsx.', 'error');
    }
  }

  async function handleMainCsvFile(file) {
    try {
      const parsed = await parseUploadedFile(file);
      if (!parsed.headers.length) throw new Error('Arquivo .csv vazio ou inválido.');
      state.headers = parsed.headers;
      state.rows = parsed.rows;
      state.schema = detectSchema(state.headers);
      state.selectedIds.clear();
      state.lastSyncAt = new Date();
      saveLocalData();

      document.getElementById('csvFileLabel').innerHTML = `<strong>${escapeHtml(file.name)}</strong> (${parsed.rows.length} linhas)`;
      document.getElementById('sheetSubtitle').textContent = `CSV carregado: ${file.name}`;
      updateConnectionUI('live', `CSV: ${file.name.slice(0, 18)}`);
      showToast(`${parsed.rows.length} acordos importados de ${file.name}!`, 'success');
      renderAll();
    } catch (err) {
      showToast(err.message || 'Erro ao ler arquivo .csv.', 'error');
    }
  }

  async function handleDedicatedReportFile(kind, file) {
    try {
      const parsed = await parseUploadedFile(file);
      if (!parsed.headers.length) throw new Error('Arquivo de relatório vazio.');
      state.customReports[kind] = {
        fileName: file.name,
        headers: parsed.headers,
        rows: parsed.rows,
        schema: detectSchema(parsed.headers)
      };
      renderHeroAndReports();
      showToast(`Relatório "${kind === 'total' ? 'Acordos Totais' : 'Pagos em Dia'}" atualizado com ${file.name} (${parsed.rows.length} linhas)!`, 'success');
    } catch (err) {
      showToast(err.message || 'Erro ao carregar arquivo do relatório.', 'error');
    }
  }

  /* -------------------- Modals & Toasts -------------------- */
  function openModal(id) { document.getElementById(id).classList.add('open'); }
  function closeModal(id) { document.getElementById(id).classList.remove('open'); }
  function showToast(message, type = 'info') {
    const stack = document.getElementById('toastStack');
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    el.textContent = message;
    stack.appendChild(el);
    setTimeout(() => el.remove(), 3800);
  }

  /* -------------------- Event Wiring -------------------- */
  function bindEvents() {
    document.querySelectorAll('[data-close-modal]').forEach(btn => {
      btn.addEventListener('click', () => closeModal(btn.getAttribute('data-close-modal')));
    });
    document.querySelectorAll('.modal-backdrop').forEach(bg => {
      bg.addEventListener('click', (e) => { if (e.target === bg) bg.classList.remove('open'); });
    });

    // Primary & Hero buttons
    document.getElementById('btnNewAgreement').addEventListener('click', openAddModal);
    document.getElementById('btnHeroNew').addEventListener('click', openAddModal);
    document.getElementById('btnTableNew').addEventListener('click', openAddModal);
    document.getElementById('btnSyncNow').addEventListener('click', () => syncFromSheet(false));

    // Reports View & Export buttons
    document.getElementById('btnHeroReportTotal').addEventListener('click', () => openReportModal('total'));
    document.getElementById('btnViewReportTotal').addEventListener('click', () => openReportModal('total'));
    document.getElementById('btnHeroReportOnTime').addEventListener('click', () => openReportModal('onTime'));
    document.getElementById('btnViewReportOnTime').addEventListener('click', () => openReportModal('onTime'));

    const todayTag = () => new Date().toISOString().slice(0, 10);
    document.getElementById('btnDownloadTotalXlsx').addEventListener('click', () => {
      const d = getTotalReportData();
      downloadXLSX(d.headers, d.rows, `relatorio-acordos-totais-${todayTag()}.xlsx`, 'Acordos Totais');
    });
    document.getElementById('btnDownloadTotalCsv').addEventListener('click', () => {
      const d = getTotalReportData();
      downloadCSV(d.headers, d.rows, `relatorio-acordos-totais-${todayTag()}.csv`);
    });
    document.getElementById('btnDownloadOnTimeXlsx').addEventListener('click', () => {
      const d = getOnTimeReportData();
      downloadXLSX(d.headers, d.rows, `relatorio-acordos-pagos-em-dia-${todayTag()}.xlsx`, 'Pagos em Dia');
    });
    document.getElementById('btnDownloadOnTimeCsv').addEventListener('click', () => {
      const d = getOnTimeReportData();
      downloadCSV(d.headers, d.rows, `relatorio-acordos-pagos-em-dia-${todayTag()}.csv`);
    });

    document.getElementById('btnPrintReport').addEventListener('click', () => window.print());
    document.getElementById('btnModalReportXlsx').addEventListener('click', () => {
      const isT = state.activeReportKind === 'total';
      const d = isT ? getTotalReportData() : getOnTimeReportData();
      downloadXLSX(d.headers, d.rows, `relatorio-${isT ? 'acordos-totais' : 'pagos-em-dia'}-${todayTag()}.xlsx`, isT ? 'Acordos Totais' : 'Pagos em Dia');
    });
    document.getElementById('btnModalReportCsv').addEventListener('click', () => {
      const isT = state.activeReportKind === 'total';
      const d = isT ? getTotalReportData() : getOnTimeReportData();
      downloadCSV(d.headers, d.rows, `relatorio-${isT ? 'acordos-totais' : 'pagos-em-dia'}-${todayTag()}.csv`);
    });

    // Dedicated Report File Inputs
    document.getElementById('reportTotalFileInput').addEventListener('change', (e) => {
      const f = e.target.files?.[0];
      if (f) handleDedicatedReportFile('total', f);
      e.target.value = '';
    });
    document.getElementById('btnClearReportTotalFile').addEventListener('click', () => {
      state.customReports.total = null;
      renderHeroAndReports();
      showToast('Relatório de Acordos Totais voltou a usar a base geral.', 'info');
    });

    document.getElementById('reportOnTimeFileInput').addEventListener('change', (e) => {
      const f = e.target.files?.[0];
      if (f) handleDedicatedReportFile('onTime', f);
      e.target.value = '';
    });
    document.getElementById('btnClearReportOnTimeFile').addEventListener('click', () => {
      state.customReports.onTime = null;
      renderHeroAndReports();
      showToast('Relatório Pagos em Dia voltou ao filtro automático.', 'info');
    });

    // Main XLSX & CSV Uploads + Drag & Drop
    document.getElementById('xlsxFileInput').addEventListener('change', (e) => {
      const f = e.target.files?.[0];
      if (f) handleMainXlsxFile(f);
      e.target.value = '';
    });
    setupDropzone('dropzoneXlsx', handleMainXlsxFile);

    document.getElementById('xlsxSheetSelector').addEventListener('change', (e) => {
      const idx = Number(e.target.value);
      const sheet = state.xlsxWorkbook?.sheets?.[idx];
      if (!sheet) return;
      state.headers = sheet.headers;
      state.rows = sheet.rows;
      state.schema = detectSchema(state.headers);
      state.selectedIds.clear();
      saveLocalData();
      showToast(`Aba "${sheet.name}" carregada (${sheet.rows.length} linhas).`, 'info');
      renderAll();
    });

    document.getElementById('csvFileInput').addEventListener('change', (e) => {
      const f = e.target.files?.[0];
      if (f) handleMainCsvFile(f);
      e.target.value = '';
    });
    setupDropzone('dropzoneCsv', handleMainCsvFile);

    // Filtered Table Exports
    document.getElementById('btnExportXlsx').addEventListener('click', () => {
      downloadXLSX(state.headers, getFilteredRows(), `acordos-filtrado-${todayTag()}.xlsx`, 'Acordos Filtrados');
    });
    document.getElementById('btnExportCsv').addEventListener('click', () => {
      downloadCSV(state.headers, getFilteredRows(), `acordos-filtrado-${todayTag()}.csv`);
    });

    // Batch Delete
    document.getElementById('btnDeleteSelected').addEventListener('click', () => {
      const selected = state.rows.filter(r => state.selectedIds.has(r._id));
      if (selected.length) requestDeleteConfirmation(selected);
    });

    // Connect Google Sheets Modal
    const openConnect = () => {
      document.getElementById('cfgSheetId').value = state.config.sheetId;
      document.getElementById('cfgSheetGid').value = state.config.sheetGid;
      document.getElementById('cfgAppsScriptUrl').value = state.config.appsScriptUrl;
      document.getElementById('cfgSyncInterval').value = String(state.config.syncInterval);
      document.getElementById('cfgConnectionMode').value = state.config.mode;
      openModal('modalConnect');
    };
    document.getElementById('btnOpenConnect').addEventListener('click', openConnect);
    document.getElementById('btnBannerConnect').addEventListener('click', openConnect);

    document.getElementById('btnSaveConnection').addEventListener('click', async () => {
      state.config.sheetId = document.getElementById('cfgSheetId').value.trim() || DEFAULT_SHEET_ID;
      state.config.sheetGid = document.getElementById('cfgSheetGid').value.trim() || DEFAULT_SHEET_GID;
      state.config.appsScriptUrl = document.getElementById('cfgAppsScriptUrl').value.trim();
      state.config.syncInterval = Number(document.getElementById('cfgSyncInterval').value || 30);
      state.config.mode = document.getElementById('cfgConnectionMode').value;
      if (state.config.appsScriptUrl && state.config.mode === 'demo') state.config.mode = 'apps_script';
      saveConfig();
      setupAutoSync();
      closeModal('modalConnect');
      await syncFromSheet(false);
    });

    const resetDemo = () => {
      state.config.mode = 'demo';
      saveConfig();
      setupAutoSync();
      state.headers = [...DEFAULT_HEADERS];
      state.rows = DEMO_ROWS.map(r => ({ ...r }));
      state.schema = detectSchema(state.headers);
      state.customReports.total = null;
      state.customReports.onTime = null;
      document.getElementById('xlsxFileLabel').textContent = 'Nenhum arquivo .xlsx inserido';
      document.getElementById('csvFileLabel').textContent = 'Nenhum arquivo .csv inserido';
      document.getElementById('xlsxSheetSelector').style.display = 'none';
      saveLocalData();
      updateConnectionUI('demo', 'Modo Demonstração');
      closeModal('modalConnect');
      showToast('Base de exemplo restaurada.', 'info');
      renderAll();
    };
    document.getElementById('btnResetDemo').addEventListener('click', resetDemo);
    document.getElementById('btnResetDemoQuick').addEventListener('click', resetDemo);

    document.getElementById('btnCopyAppsScript').addEventListener('click', async () => {
      try {
        const res = await fetch('apps-script/Code.gs');
        const code = res.ok ? await res.text() : 'Abra o arquivo apps-script/Code.gs no repositório.';
        await navigator.clipboard.writeText(code);
        showToast('Código Code.gs copiado!', 'success');
      } catch (_) {
        showToast('Copie o arquivo apps-script/Code.gs no repositório.', 'info');
      }
    });

    // Filters
    document.getElementById('filterSearch').addEventListener('input', e => { state.filters.search = e.target.value; renderAll(); });
    document.getElementById('filterStatus').addEventListener('change', e => { state.filters.status = e.target.value; renderAll(); });
    document.getElementById('filterPayment').addEventListener('change', e => { state.filters.payment = e.target.value; renderAll(); });
    document.getElementById('filterCategory').addEventListener('change', e => { state.filters.category = e.target.value; renderAll(); });
    document.getElementById('filterOwner').addEventListener('change', e => { state.filters.owner = e.target.value; renderAll(); });
    document.getElementById('filterDateFrom').addEventListener('change', e => { state.filters.dateFrom = e.target.value; renderAll(); });
    document.getElementById('filterDateTo').addEventListener('change', e => { state.filters.dateTo = e.target.value; renderAll(); });

    document.querySelectorAll('#quickChips .chip').forEach(chip => {
      chip.addEventListener('click', () => {
        document.querySelectorAll('#quickChips .chip').forEach(c => c.classList.remove('active'));
        chip.classList.add('active');
        state.filters.quick = chip.getAttribute('data-quick');
        renderAll();
      });
    });

    document.getElementById('btnClearFilters').addEventListener('click', () => {
      state.filters = { search: '', status: '', payment: '', category: '', owner: '', dateFrom: '', dateTo: '', quick: 'all' };
      ['filterSearch', 'filterStatus', 'filterPayment', 'filterCategory', 'filterOwner', 'filterDateFrom', 'filterDateTo'].forEach(id => {
        document.getElementById(id).value = '';
      });
      document.querySelectorAll('#quickChips .chip').forEach(c => c.classList.toggle('active', c.getAttribute('data-quick') === 'all'));
      renderAll();
    });

    // Add Form -> Confirmation Modal
    document.getElementById('formAddAgreement').addEventListener('submit', (e) => {
      e.preventDefault();
      const formData = new FormData(e.target);
      const record = {};
      state.headers.forEach(h => {
        const val = formData.get(h);
        record[h] = (h === state.schema.valueCol || h === state.schema.progressCol) ? parseNumber(val) : String(val || '').trim();
      });

      closeModal('modalAdd');
      const { idCol, partyCol, valueCol, statusCol } = state.schema;
      openActionConfirmation({
        type: 'add',
        title: 'Confirmar Inclusão',
        summaryHtml: `
          <div style="color:#9a1b0c;font-weight:700;">Resumo do novo acordo a ser incluído:</div>
          <div><strong>${escapeHtml(idCol || 'Registro')}:</strong> ${escapeHtml(record[idCol] || '—')}</div>
          ${partyCol ? `<div><strong>${escapeHtml(partyCol)}:</strong> ${escapeHtml(record[partyCol] || '—')}</div>` : ''}
          ${valueCol ? `<div><strong>${escapeHtml(valueCol)}:</strong> ${formatBRL(record[valueCol])}</div>` : ''}
          ${statusCol ? `<div><strong>${escapeHtml(statusCol)}:</strong> ${escapeHtml(record[statusCol] || '—')}</div>` : ''}
        `,
        buttonText: 'Confirmar Inclusão',
        buttonClass: 'btn-primary',
        payload: record
      });
    });

    document.getElementById('formConfirmAction').addEventListener('submit', async (e) => {
      e.preventDefault();
      document.getElementById('confirmActionError').classList.remove('visible');
      await executePendingAction();
    });
  }

  async function init() {
    loadConfig();
    loadLocalOrDemoData();
    bindEvents();
    setupAutoSync();
    if (state.config.mode !== 'demo') await syncFromSheet(true);
    else {
      updateConnectionUI('demo', 'Modo Demonstração');
      renderAll();
    }
  }

  document.addEventListener('DOMContentLoaded', init);
})();
