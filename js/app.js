(() => {
  'use strict';

  const STORAGE_KEY = 'control-hub:v1';
  const PREFS_KEY = 'control-hub:prefs';
  const STATUS_TIMEOUT_MS = 8000;
  const RECENT_LIMIT = 12;
  const DB_COLLECTION = 'tools';

  // Rodando como artefato do claude.ai? Lá a lista fica num banco compartilhado entre as contas.
  const HOSTED = typeof window.claude?.use === 'function';

  // Ferramentas iniciais do modo local. No artefato, a lista vem do banco compartilhado.
  const DEFAULT_TOOLS = [
    { id: 'gestta', name: 'Gestta', url: 'https://app.gestta.com.br', icon: '📈', category: 'Acompanhamento', color: '#6b7280', description: 'Andamento das empresas.', tags: ['empresas'] },
    { id: 'portal-cliente', name: 'Portal do Cliente', url: '', icon: '🧑‍💼', category: 'Atalhos', color: '#6b7280', description: 'Acesso rápido pelo app.', tags: ['cliente'] },
    { id: 'planilha-dp', name: 'Planilha DP', url: '', icon: '👥', category: 'Atalhos', color: '#6b7280', description: 'Acesso rápido pelo app.', tags: ['planilha', 'dp'] },
    { id: 'planilha-fiscal', name: 'Planilha Fiscal', url: '', icon: '🧾', category: 'Atalhos', color: '#6b7280', description: 'Acesso rápido pelo app.', tags: ['planilha', 'fiscal'] },
    { id: 'planilha-contabil', name: 'Planilha Contábil', url: '', icon: '📒', category: 'Atalhos', color: '#6b7280', description: 'Acesso rápido pelo app.', tags: ['planilha', 'contábil'] },
    { id: 'zappy', name: 'Zappy', url: '', icon: '💬', category: 'Comunicação', color: '#22a55b', description: 'Só os contatos.', tags: ['whatsapp', 'contatos'] },
    { id: 'zimbra', name: 'Zimbra', url: '', icon: '✉️', category: 'Comunicação', color: '#22a55b', description: 'Agendamento das salas e monitoramento das comunicações.', tags: ['e-mail', 'salas'], notes: 'Monitorar ao máximo as comunicações — tanto no e-mail quanto no WhatsApp.' },
    { id: 'pasta-servidor', name: 'Pasta Servidor', url: '', icon: '🗄️', category: 'Monitoramento', color: '#22a55b', description: 'Arquivos fora do padrão ou salvos no local errado.', tags: ['servidor', 'arquivos'], notes: 'Monitorar arquivos fora do padrão e salvos em local errado. Como diagnosticar quando um arquivo não for salvo?', monitor: false },
    { id: 'cardapio', name: 'Cardápio', url: '', icon: '🍽️', category: 'Outros', color: '#4f7cff', description: '', tags: [] },
  ];

  // ---------- Estado ----------
  const state = {
    tools: [],
    loaded: false,
    status: {},            // id -> 'online' | 'offline' | 'checking'
    view: 'all',           // all | favorites | recent | offline | cat:<nome>
    query: '',
    readOnly: false,
    // O artefato bloqueia requisições a outros sites, então a checagem de status só existe no modo local.
    canCheck: !HOSTED,
    prefs: { theme: null, layout: 'grid', sort: 'manual' },
  };

  let db = null;          // namespace do banco compartilhado (modo artefato)
  let downloads = null;   // namespace de downloads (modo artefato)

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

  const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));

  const escapeHtml = (s = '') => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function safeUrl(url) {
    try {
      const u = new URL(url);
      return ['http:', 'https:'].includes(u.protocol) ? u.href : null;
    } catch {
      return null;
    }
  }

  function hostOf(url) {
    if (!url) return 'Link não configurado';
    try { return new URL(url).host; } catch { return url; }
  }

  const validId = (id) => typeof id === 'string' && /^[A-Za-z0-9_\-.~:@+]{1,120}$/.test(id) && id !== '.' && id !== '..';

  function normalizeTool(t, i = 0) {
    return {
      id: validId(t.id) ? t.id : uid(),
      name: String(t.name || 'Sem nome').slice(0, 60),
      url: safeUrl(t.url) || '',
      icon: t.icon || '🔧',
      category: String(t.category || 'Geral').trim() || 'Geral',
      color: /^#[0-9a-f]{6}$/i.test(t.color || '') ? t.color : '#4f7cff',
      description: t.description || '',
      tags: Array.isArray(t.tags) ? t.tags.filter(Boolean).map(String) : [],
      notes: t.notes || '',
      favorite: !!t.favorite,
      monitor: t.monitor !== false,
      newTab: t.newTab !== false,
      order: Number.isFinite(t.order) ? t.order : i,
      uses: Number(t.uses) || 0,
      lastUsed: t.lastUsed || null,
      createdAt: t.createdAt || Date.now(),
    };
  }

  // ---------- Persistência ----------
  // Duas implementações com a mesma interface: banco compartilhado (artefato) e localStorage (local).
  const localStore = {
    shared: false,
    load() {
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        return raw ? JSON.parse(raw).map(normalizeTool) : DEFAULT_TOOLS.map(normalizeTool);
      } catch {
        return DEFAULT_TOOLS.map(normalizeTool);
      }
    },
    persistAll() {
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state.tools)); } catch { toast('Não foi possível salvar no navegador.'); }
    },
    async saveTool() { this.persistAll(); },
    async saveTools() { this.persistAll(); },
    async deleteTool() { this.persistAll(); },
  };

  const dbStore = {
    shared: true,
    async saveTool(t) {
      const { id, ...body } = t;
      await db.doc(`${DB_COLLECTION}/${id}`).set(body);
    },
    async saveTools(list) {
      // Uma escrita por vez: o banco pede que escritas no mesmo documento não se sobreponham.
      for (const t of list) await this.saveTool(t);
    },
    async deleteTool(id) {
      await db.doc(`${DB_COLLECTION}/${id}`).delete();
    },
  };

  let store = localStore;

  // Aplica a mudança na tela na hora e grava em seguida; avisa se a gravação falhar.
  async function persist(fn) {
    try {
      await fn();
    } catch (e) {
      if (e?.code === 'invalid_argument' && store.shared) {
        setReadOnly(true);
        toast('Você não tem permissão para editar esta lista.');
      } else if (e?.code === 'quota_exceeded') {
        toast('Limite de ferramentas atingido. Exclua alguma para adicionar novas.');
      } else {
        toast('Não foi possível salvar. Tente de novo em instantes.');
      }
    }
  }

  function loadPrefs() {
    try {
      Object.assign(state.prefs, JSON.parse(localStorage.getItem(PREFS_KEY) || '{}'));
    } catch { /* usa padrões */ }
  }

  function savePrefs() {
    try { localStorage.setItem(PREFS_KEY, JSON.stringify(state.prefs)); } catch { /* ignora */ }
  }

  // ---------- Consultas ----------
  const categories = () => [...new Set(state.tools.map((t) => t.category))].sort((a, b) => a.localeCompare(b, 'pt-BR'));

  function matchesQuery(t, q) {
    if (!q) return true;
    const hay = [t.name, t.description, t.category, t.url, t.notes, ...t.tags].join(' ').toLowerCase();
    return q.toLowerCase().split(/\s+/).every((w) => hay.includes(w));
  }

  function inView(t, view) {
    if (view === 'favorites') return t.favorite;
    if (view === 'recent') return !!t.lastUsed;
    if (view === 'offline') return state.status[t.id] === 'offline';
    if (view.startsWith('cat:')) return t.category === view.slice(4);
    return true;
  }

  function visibleTools() {
    let list = state.tools.filter((t) => inView(t, state.view) && matchesQuery(t, state.query));
    const sort = state.view === 'recent' ? 'recent' : state.prefs.sort;
    const by = {
      manual: (a, b) => a.order - b.order,
      name: (a, b) => a.name.localeCompare(b.name, 'pt-BR'),
      usage: (a, b) => b.uses - a.uses || a.order - b.order,
      recent: (a, b) => (b.lastUsed || 0) - (a.lastUsed || 0),
    }[sort];
    list.sort(by);
    if (state.view === 'recent') list = list.slice(0, RECENT_LIMIT);
    // favoritas primeiro na visão geral com ordem manual
    if (state.view === 'all' && sort === 'manual') list.sort((a, b) => b.favorite - a.favorite);
    return list;
  }

  // ---------- Renderização ----------
  function render() {
    renderChrome();
    renderSidebar();
    renderStats();
    renderTools();
  }

  function renderChrome() {
    $('#btn-check').hidden = !state.canCheck;
    $('[data-view="offline"]').parentElement.hidden = !state.canCheck;
    $('#btn-reset').hidden = store.shared;
    $('#btn-add').hidden = state.readOnly;
    $('#btn-add-empty').hidden = state.readOnly;
    $('#input-import').closest('label').hidden = state.readOnly;
  }

  function showBanner(text, kind = 'warn') {
    const el = $('#banner');
    el.textContent = text;
    el.className = `banner ${kind}`;
    el.hidden = !text;
  }

  function renderSidebar() {
    const counts = {
      all: state.tools.length,
      favorites: state.tools.filter((t) => t.favorite).length,
      recent: Math.min(state.tools.filter((t) => t.lastUsed).length, RECENT_LIMIT),
      offline: state.tools.filter((t) => state.status[t.id] === 'offline').length,
    };
    $$('[data-count]').forEach((el) => { el.textContent = counts[el.dataset.count] || ''; });

    const catList = $('#category-list');
    catList.innerHTML = categories().map((c) => {
      const tools = state.tools.filter((t) => t.category === c);
      const color = tools[0]?.color || 'var(--primary)';
      return `<li><button class="nav-item" data-view="cat:${escapeHtml(c)}">
        <span class="nav-dot" style="background:${color}"></span>${escapeHtml(c)}
        <span class="count">${tools.length}</span></button></li>`;
    }).join('');

    $$('.nav-item').forEach((b) => b.classList.toggle('active', b.dataset.view === state.view));
    $('#category-options').innerHTML = categories().map((c) => `<option value="${escapeHtml(c)}">`).join('');
  }

  function renderStats() {
    const stats = [
      ['Ferramentas', state.tools.length],
      ['Categorias', categories().length],
    ];
    if (state.canCheck) {
      const monitored = state.tools.filter((t) => t.monitor && t.url);
      const online = monitored.filter((t) => state.status[t.id] === 'online').length;
      const offline = monitored.filter((t) => state.status[t.id] === 'offline').length;
      const checked = online + offline;
      stats.push(['Online', checked ? `${online}/${checked}` : '—'], ['Com problema', checked ? offline : '—']);
    }
    stats.push(
      ['Sem link', state.tools.filter((t) => !t.url).length],
      ['Acessos', state.tools.reduce((s, t) => s + t.uses, 0)],
    );
    $('#stats').innerHTML = stats.map(([label, value]) =>
      `<div class="stat"><div class="stat-label">${label}</div><div class="stat-value">${value}</div></div>`).join('');
  }

  const STATUS_LABEL = { online: 'Online', offline: 'Fora do ar', checking: 'Verificando…' };

  function renderTools() {
    const list = visibleTools();
    const container = $('#tools');
    container.className = `tools ${state.prefs.layout}`;
    $('#sort').value = state.prefs.sort;
    $('#btn-layout').textContent = state.prefs.layout === 'grid' ? '☰' : '▦';

    const titles = { all: 'Todas as ferramentas', favorites: 'Favoritas', recent: 'Usadas recentemente', offline: 'Com problema' };
    $('#view-title').textContent = titles[state.view] || state.view.slice(4);

    if (!state.loaded) {
      container.innerHTML = '';
      $('#empty').hidden = true;
      showBanner('Carregando ferramentas…', 'info');
      return;
    }
    $('#empty').hidden = list.length > 0;

    const draggable = !state.readOnly && state.prefs.sort === 'manual' && state.view !== 'recent' && !state.query;
    const ro = state.readOnly;

    container.innerHTML = list.map((t) => {
      const st = state.canCheck && t.monitor && t.url ? state.status[t.id] : null;
      const used = t.lastUsed ? `Usada ${timeAgo(t.lastUsed)}` : 'Nunca usada';
      // Links reais (<a>): dentro do artefato, window.open não funciona para todos.
      const target = t.newTab || HOSTED ? ' target="_blank" rel="noopener noreferrer"' : '';
      return `
      <article class="card${t.url ? '' : ' no-link'}" data-id="${escapeHtml(t.id)}" style="--accent:${t.color}" ${draggable ? 'draggable="true"' : ''}>
        <div class="card-head">
          <div class="card-icon">${escapeHtml(t.icon)}</div>
          <div class="card-titles">
            <h3 class="card-title" title="${escapeHtml(t.name)}">${escapeHtml(t.name)}</h3>
            <div class="card-host">${escapeHtml(hostOf(t.url))}</div>
          </div>
          <button class="icon-btn fav-btn ${t.favorite ? 'on' : ''}" data-action="fav" title="Favoritar" aria-label="Favoritar" ${ro ? 'disabled' : ''}>⭐</button>
        </div>
        <p class="card-desc">${escapeHtml(t.description)}</p>
        ${t.tags.length ? `<div class="card-tags">${t.tags.map((g) => `<span class="tag">${escapeHtml(g)}</span>`).join('')}</div>` : ''}
        <div class="card-foot">
          ${st ? `<span class="status ${st}"><span class="status-dot"></span>${STATUS_LABEL[st]}</span><span class="card-meta"></span>` : `<span class="card-meta">${used}</span>`}
          ${state.canCheck && t.monitor && t.url ? `<button class="icon-btn" data-action="check" title="Verificar status" aria-label="Verificar status">📡</button>` : ''}
          ${t.url ? `<button class="icon-btn" data-action="copy" title="Copiar link" aria-label="Copiar link">🔗</button>` : ''}
          ${ro ? '' : `<button class="icon-btn" data-action="edit" title="Editar" aria-label="Editar">✏️</button>`}
          ${t.url
            ? `<a class="btn btn-primary btn-sm" data-action="open" href="${escapeHtml(t.url)}"${target}>Abrir ↗</a>`
            : (ro ? '' : `<button class="btn btn-ghost btn-sm" data-action="edit">Definir link</button>`)}
        </div>
      </article>`;
    }).join('');
  }

  function timeAgo(ts) {
    const s = Math.round((Date.now() - ts) / 1000);
    if (s < 60) return 'agora';
    const m = Math.round(s / 60);
    if (m < 60) return `há ${m} min`;
    const h = Math.round(m / 60);
    if (h < 24) return `há ${h} h`;
    const d = Math.round(h / 24);
    return `há ${d} dia${d > 1 ? 's' : ''}`;
  }

  // ---------- Ações ----------
  const findTool = (id) => state.tools.find((t) => t.id === id);

  // Chamado quando o link "Abrir" é seguido: só registra o uso, a navegação é do próprio <a>.
  function recordUse(t) {
    t.uses += 1;
    t.lastUsed = Date.now();
    if (!state.readOnly) persist(() => store.saveTool(t));
    setTimeout(render, 0);
  }

  function openTool(t) {
    if (!t.url) {
      if (!state.readOnly) {
        openDialog(t);
        toast('Defina o link desta ferramenta.');
      }
      return;
    }
    const link = $(`.card[data-id="${CSS.escape(t.id)}"] a[data-action="open"]`);
    if (link) link.click();
  }

  async function checkStatus(t) {
    if (!t.url || !state.canCheck) return;
    state.status[t.id] = 'checking';
    renderTools();
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), STATUS_TIMEOUT_MS);
    try {
      // no-cors: a resposta é opaca, mas a promessa só resolve se o host respondeu.
      await fetch(t.url, { mode: 'no-cors', cache: 'no-store', signal: ctrl.signal });
      state.status[t.id] = 'online';
    } catch {
      state.status[t.id] = 'offline';
    } finally {
      clearTimeout(timer);
    }
    render();
  }

  async function checkAll() {
    const btn = $('#btn-check');
    btn.disabled = true;
    const monitored = state.tools.filter((t) => t.monitor && t.url);
    await Promise.all(monitored.map(checkStatus));
    btn.disabled = false;
    const off = monitored.filter((t) => state.status[t.id] === 'offline').length;
    toast(off ? `${off} ferramenta(s) com problema.` : 'Tudo online ✅');
  }

  async function copyLink(t) {
    if (!t.url) return toast('Link não configurado.');
    try {
      await navigator.clipboard.writeText(t.url);
      toast('Link copiado.');
    } catch {
      toast(t.url);
    }
  }

  function setReadOnly(value) {
    state.readOnly = value;
    if (value && store.shared) showBanner('Você está vendo a lista em modo leitura. Peça acesso de Colaborador (Contributor) para editar.');
    render();
  }

  // ---------- Confirmação na página ----------
  // confirm() não aparece dentro de artefatos, então a página usa um diálogo próprio.
  function ask(message, buttons) {
    const dlg = $('#ask-dialog');
    $('#ask-message').textContent = message;
    const actions = $('#ask-actions');
    actions.innerHTML = '<span class="spacer"></span>' + buttons.map((b, i) =>
      `<button type="button" class="btn ${b.kind || 'btn-ghost'}" data-i="${i}">${escapeHtml(b.label)}</button>`).join('');
    return new Promise((resolve) => {
      const done = (value) => {
        dlg.removeEventListener('close', onClose);
        actions.onclick = null;
        if (dlg.open) dlg.close();
        resolve(value);
      };
      const onClose = () => done(null);
      actions.onclick = (e) => {
        const btn = e.target.closest('[data-i]');
        if (btn) done(buttons[btn.dataset.i].value);
      };
      dlg.addEventListener('close', onClose);
      dlg.showModal();
    });
  }

  // ---------- Modal de ferramenta ----------
  const dialog = $('#tool-dialog');
  const form = $('#tool-form');

  function openDialog(tool) {
    if (state.readOnly) return;
    form.reset();
    $('#dialog-title').textContent = tool ? 'Editar ferramenta' : 'Nova ferramenta';
    $('#btn-delete').hidden = !tool;
    const t = tool || {
      icon: '🔧', color: '#4f7cff', monitor: true, newTab: true,
      category: state.view.startsWith('cat:') ? state.view.slice(4) : '',
    };
    form.elements.id.value = t.id || '';
    form.elements.name.value = t.name || '';
    form.elements.url.value = t.url || '';
    form.elements.icon.value = t.icon || '';
    form.elements.description.value = t.description || '';
    form.elements.category.value = t.category || '';
    form.elements.color.value = t.color || '#4f7cff';
    form.elements.tags.value = (t.tags || []).join(', ');
    form.elements.notes.value = t.notes || '';
    form.elements.favorite.checked = !!t.favorite;
    form.elements.monitor.checked = t.monitor !== false;
    form.elements.newTab.checked = t.newTab !== false;
    dialog.showModal();
    form.elements.name.focus();
  }

  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const f = form.elements;
    const rawUrl = f.url.value.trim();
    const url = rawUrl ? safeUrl(rawUrl) : '';
    if (url === null) {
      toast('Informe uma URL http(s) válida.');
      f.url.focus();
      return;
    }
    const data = {
      name: f.name.value.trim(),
      url,
      icon: f.icon.value.trim() || '🔧',
      description: f.description.value.trim(),
      category: f.category.value.trim() || 'Geral',
      color: f.color.value,
      tags: f.tags.value.split(',').map((s) => s.trim()).filter(Boolean),
      notes: f.notes.value.trim(),
      favorite: f.favorite.checked,
      monitor: f.monitor.checked,
      newTab: f.newTab.checked,
    };
    let tool = findTool(f.id.value);
    if (tool) {
      Object.assign(tool, data);
      toast('Ferramenta atualizada.');
    } else {
      const order = Math.max(-1, ...state.tools.map((t) => t.order)) + 1;
      tool = normalizeTool({ ...data, order });
      state.tools.push(tool);
      toast('Ferramenta adicionada.');
    }
    persist(() => store.saveTool(tool));
    dialog.close();
    render();
  });

  $('#btn-cancel').addEventListener('click', () => dialog.close());

  $('#btn-delete').addEventListener('click', async () => {
    const t = findTool(form.elements.id.value);
    if (!t) return;
    dialog.close();
    const ok = await ask(`Excluir "${t.name}"?${store.shared ? '\nA ferramenta some para todos que usam esta lista.' : ''}`, [
      { label: 'Cancelar', value: false },
      { label: 'Excluir', value: true, kind: 'btn-danger' },
    ]);
    if (!ok) return;
    state.tools = state.tools.filter((x) => x.id !== t.id);
    delete state.status[t.id];
    persist(() => store.deleteTool(t.id));
    if (state.view.startsWith('cat:') && !categories().includes(state.view.slice(4))) state.view = 'all';
    render();
    toast('Ferramenta excluída.');
  });

  // ---------- Import / Export ----------
  async function exportData() {
    const filename = `control-hub-${new Date().toISOString().slice(0, 10)}.json`;
    const json = JSON.stringify({ app: 'control-hub', version: 1, exportedAt: new Date().toISOString(), tools: state.tools }, null, 2);
    if (downloads) {
      try {
        await downloads.save({ filename, data: json });
      } catch (e) {
        if (e?.code !== 'declined') toast('Não foi possível gerar o arquivo aqui.');
      }
      return;
    }
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([json], { type: 'application/json' }));
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  async function importData(file) {
    let tools;
    try {
      const json = JSON.parse(await file.text());
      tools = Array.isArray(json) ? json : json.tools;
      if (!Array.isArray(tools)) throw new Error('formato');
    } catch {
      toast('Arquivo inválido: use um JSON exportado pelo Control Hub.');
      return;
    }
    const mode = await ask(`Importar ${tools.length} ferramenta(s).\nSubstituir a lista atual ou juntar com ela?`, [
      { label: 'Cancelar', value: null },
      { label: 'Juntar', value: 'merge' },
      { label: 'Substituir', value: 'replace', kind: 'btn-danger' },
    ]);
    if (!mode) return;
    const incoming = tools.map(normalizeTool);
    let removed = [];
    if (mode === 'replace') {
      const keep = new Set(incoming.map((t) => t.id));
      removed = state.tools.filter((t) => !keep.has(t.id)).map((t) => t.id);
      state.tools = incoming;
    } else {
      const ids = new Set(state.tools.map((t) => t.id));
      const base = Math.max(-1, ...state.tools.map((t) => t.order)) + 1;
      incoming.forEach((t, i) => {
        if (ids.has(t.id)) t.id = uid();
        t.order = base + i;
        state.tools.push(t);
      });
    }
    state.status = {};
    render();
    await persist(async () => {
      await store.saveTools(incoming);
      for (const id of removed) await store.deleteTool(id);
    });
    toast(`${incoming.length} ferramenta(s) importada(s).`);
  }

  // ---------- Arrastar e soltar (ordem manual) ----------
  let dragId = null;
  const tools = $('#tools');

  tools.addEventListener('dragstart', (e) => {
    const card = e.target.closest('.card');
    if (!card) return;
    dragId = card.dataset.id;
    card.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
  });
  tools.addEventListener('dragend', (e) => {
    e.target.closest('.card')?.classList.remove('dragging');
    $$('.drag-over').forEach((el) => el.classList.remove('drag-over'));
  });
  tools.addEventListener('dragover', (e) => {
    const card = e.target.closest('.card');
    if (!card || !dragId) return;
    e.preventDefault();
    $$('.drag-over').forEach((el) => el.classList.remove('drag-over'));
    if (card.dataset.id !== dragId) card.classList.add('drag-over');
  });
  tools.addEventListener('drop', (e) => {
    const card = e.target.closest('.card');
    if (!card || !dragId || card.dataset.id === dragId) return;
    e.preventDefault();
    const ordered = [...state.tools].sort((a, b) => a.order - b.order);
    const from = ordered.findIndex((t) => t.id === dragId);
    const to = ordered.findIndex((t) => t.id === card.dataset.id);
    const [moved] = ordered.splice(from, 1);
    ordered.splice(to, 0, moved);
    const changed = ordered.filter((t, i) => t.order !== i);
    ordered.forEach((t, i) => { t.order = i; });
    dragId = null;
    persist(() => store.saveTools(changed));
    renderTools();
  });

  // ---------- Eventos ----------
  tools.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-action]');
    const card = e.target.closest('.card');
    if (!card || !btn) return;
    const t = findTool(card.dataset.id);
    if (!t) return;
    const action = btn.dataset.action;
    if (action === 'open') recordUse(t);
    else if (action === 'edit') openDialog(t);
    else if (action === 'copy') copyLink(t);
    else if (action === 'check') checkStatus(t);
    else if (action === 'fav' && !state.readOnly) {
      t.favorite = !t.favorite;
      persist(() => store.saveTool(t));
      render();
    }
  });

  tools.addEventListener('dblclick', (e) => {
    const card = e.target.closest('.card');
    if (card && !e.target.closest('button, a')) openTool(findTool(card.dataset.id));
  });

  $('#sidebar').addEventListener('click', (e) => {
    const item = e.target.closest('.nav-item');
    if (!item) return;
    state.view = item.dataset.view;
    $('#sidebar').classList.remove('open');
    render();
  });

  $('#search').addEventListener('input', (e) => { state.query = e.target.value.trim(); renderTools(); });
  $('#search').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      const first = visibleTools()[0];
      if (first) openTool(first);
    } else if (e.key === 'Escape') {
      e.target.value = '';
      state.query = '';
      renderTools();
    }
  });

  $('#sort').addEventListener('change', (e) => { state.prefs.sort = e.target.value; savePrefs(); renderTools(); });
  $('#btn-layout').addEventListener('click', () => {
    state.prefs.layout = state.prefs.layout === 'grid' ? 'list' : 'grid';
    savePrefs();
    renderTools();
  });
  $('#btn-theme').addEventListener('click', () => {
    state.prefs.theme = currentTheme() === 'dark' ? 'light' : 'dark';
    applyTheme();
    savePrefs();
  });
  $('#btn-add').addEventListener('click', () => openDialog());
  $('#btn-add-empty').addEventListener('click', () => openDialog());
  $('#btn-check').addEventListener('click', checkAll);
  $('#btn-menu').addEventListener('click', () => $('#sidebar').classList.toggle('open'));
  $('#btn-export').addEventListener('click', exportData);
  $('#input-import').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) importData(file);
    e.target.value = '';
  });
  $('#btn-reset').addEventListener('click', async () => {
    const ok = await ask('Restaurar as ferramentas iniciais? As ferramentas atuais serão substituídas.', [
      { label: 'Cancelar', value: false },
      { label: 'Restaurar', value: true, kind: 'btn-danger' },
    ]);
    if (!ok) return;
    state.tools = DEFAULT_TOOLS.map(normalizeTool);
    state.status = {};
    state.view = 'all';
    store.persistAll();
    render();
  });

  document.addEventListener('keydown', (e) => {
    const typing = /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement?.tagName);
    const modal = dialog.open || $('#ask-dialog').open;
    if ((e.key === '/' && !typing) || (e.key.toLowerCase() === 'k' && (e.ctrlKey || e.metaKey))) {
      e.preventDefault();
      $('#search').focus();
    } else if (e.key.toLowerCase() === 'n' && !typing && !modal && !e.ctrlKey && !e.metaKey && !state.readOnly) {
      e.preventDefault();
      openDialog();
    }
  });

  // ---------- Utilidades ----------
  let toastTimer;
  function toast(msg) {
    const el = $('#toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove('show'), 2600);
  }

  const systemDark = () => window.matchMedia('(prefers-color-scheme: dark)').matches;

  function currentTheme() {
    return document.documentElement.dataset.theme || state.prefs.theme || (systemDark() ? 'dark' : 'light');
  }

  // Sem escolha explícita, não marca data-theme: o CSS segue o sistema (e o tema do visualizador no artefato).
  function applyTheme() {
    if (state.prefs.theme) document.documentElement.dataset.theme = state.prefs.theme;
  }

  // ---------- Início ----------
  function startLocal(banner) {
    store = localStore;
    state.tools = localStore.load();
    state.loaded = true;
    showBanner(banner || '');
    render();
  }

  async function startShared() {
    render(); // estado "carregando"
    const [dbNs, dl, user] = await Promise.all([
      window.claude.use('db'),
      window.claude.use('downloads'),
      window.claude.use('user'),
    ]);
    downloads = dl;
    if (!dbNs) {
      startLocal('A lista compartilhada não está disponível nesta visualização. As alterações ficam só neste navegador.');
      return;
    }
    db = dbNs;
    store = dbStore;
    // can() === null: a plataforma não informou; mantém a edição e decide pela primeira escrita recusada.
    if (user && (await user.can('data.write').catch(() => null)) === false) state.readOnly = true;

    db.collection(DB_COLLECTION).onSnapshot((snap) => {
      state.tools = snap.docs.map((d, i) => normalizeTool({ ...d.data(), id: d.id }, i));
      if (!state.loaded) {
        state.loaded = true;
        showBanner(state.readOnly ? 'Você está vendo a lista em modo leitura. Peça acesso de Colaborador (Contributor) para editar.' : '');
      }
      render();
    }, () => {
      showBanner('A conexão com a lista compartilhada caiu. Recarregue a página para continuar.');
    });
  }

  loadPrefs();
  applyTheme();
  if (HOSTED) startShared();
  else startLocal();
})();
