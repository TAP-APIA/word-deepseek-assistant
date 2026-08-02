/* DeepSeek Word 助手 - 任务窗格逻辑 */
/* global Office, Word, localStorage */

(() => {
  'use strict';

  // 直连 DeepSeek 官方 API（已实测支持浏览器 CORS 跨域，无需本地服务）
  const DEEPSEEK_API = 'https://api.deepseek.com/chat/completions';
  const MODELS_API = 'https://api.deepseek.com/models';
  const SYSTEM_PROMPT =
    '你是运行在 Microsoft Word 侧边栏中的写作助手。用户消息末尾可能附带【当前选中内容】和【文档内容】作为上下文。\n' +
    '当你的回复中包含需要写入 Word 文档的正文时，请严格遵守以下规则：\n' +
    '1. 只把要插入文档的正文放在 Markdown 代码块（以```开头、以```结尾）中；\n' +
    '2. 代码块外只写修改说明、建议或摘要，不要包含正文；\n' +
    '3. 代码块内只放正文本身，不要使用任何 Markdown 标记；\n' +
    '4. 如果回复只是回答问题、不需要插入文档，则不要使用代码块。';
  const MAX_HISTORY = 20; // 发送给模型的对话轮数上限
  const HISTORY_KEY = 'ds_conversations';
  const MAX_HISTORY_CONV = 20;
  // deepseek-chat / deepseek-reasoner 已于 2026-07-24 弃用
  const LEGACY_MODELS = ['deepseek-chat', 'deepseek-reasoner'];
  const FALLBACK_MODELS = ['deepseek-v4-flash', 'deepseek-v4-pro'];

  let apiKey = localStorage.getItem('ds_api_key') || '';
  let model = localStorage.getItem('ds_model') || '';
  let conversation = [];
  let streaming = false;
  let lastAssistantText = '';
  let currentConvId = null;
  let conversations = loadHistory();
  let cachedModels = null;

  const $ = (id) => document.getElementById(id);

  const els = {
    hostBadge: $('hostBadge'),
    messages: $('messages'),
    userInput: $('userInput'),
    sendBtn: $('sendBtn'),
    actionBar: $('actionBar'),
    actionNote: $('actionNote'),
    newChatBtn: $('newChatBtn'),
    settingsBtn: $('settingsBtn'),
    historyBtn: $('historyBtn'),
    historyBackdrop: $('historyBackdrop'),
    historyPanel: $('historyPanel'),
    historyList: $('historyList'),
    historyCloseBtn: $('historyCloseBtn'),
    ctxInfo: $('ctxInfo'),
    includeSelection: $('includeSelection'),
    includeDocStart: $('includeDocStart'),
    refreshCtxBtn: $('refreshCtxBtn'),
    statusLine: $('statusLine'),
    toast: $('toast'),
    settingsModal: $('settingsModal'),
    apiKeyInput: $('apiKeyInput'),
    modelSelect: $('modelSelect'),
    customModelInput: $('customModelInput'),
    fetchModelsBtn: $('fetchModelsBtn'),
    settingsStatus: $('settingsStatus'),
    saveSettingsBtn: $('saveSettingsBtn'),
    cancelSettingsBtn: $('cancelSettingsBtn'),
  };

  /* ---------- 通用 UI ---------- */

  let toastTimer = null;
  function toast(text, isError = false) {
    els.toast.textContent = text;
    els.toast.style.background = isError ? '#b91c1c' : '#111827';
    els.toast.classList.remove('hidden');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => els.toast.classList.add('hidden'), 2600);
  }

  function setStatus(text, isWarn = false) {
    els.statusLine.textContent = text;
    els.statusLine.classList.toggle('warn', isWarn);
  }

  function escapeHtml(s) {
    return s
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function formatMessage(text) {
    const blocks = text.split(/```/);
    let html = '';
    blocks.forEach((block, i) => {
      if (i % 2 === 1) {
        const lines = block.split('\n');
        lines.shift(); // 语言标识
        html += `<span class="code">${escapeHtml(lines.join('\n'))}</span>`;
      } else {
        html += escapeHtml(block)
          .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
          .replace(/`([^`]+)`/g, '<code>$1</code>')
          .replace(/\n/g, '<br>');
      }
    });
    return html;
  }

  function scrollToBottom() {
    els.messages.scrollTop = els.messages.scrollHeight;
  }

  function renderWelcome() {
    if (els.messages.childElementCount > 0) return;
    const div = document.createElement('div');
    div.className = 'welcome';
    div.innerHTML =
      '在下方输入指令，让 DeepSeek 帮你写作、润色或翻译。<br>' +
      '勾选“附带选中内容”后，模型就能看到你当前选中的文字。';
    els.messages.appendChild(div);
  }

  function addMessage(role, text, isStreaming = false) {
    const div = document.createElement('div');
    div.className = `msg ${role}`;
    if (isStreaming) {
      div.innerHTML = '<span class="typing"></span>';
    } else {
      div.innerHTML = formatMessage(text) || '（空）';
    }
    els.messages.appendChild(div);
    scrollToBottom();
    return div;
  }

  function showActions() {
    els.actionBar.classList.remove('hidden');
  }

  /* ---------- 历史对话 ---------- */

  function loadHistory() {
    try {
      const raw = localStorage.getItem(HISTORY_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch {
      return [];
    }
  }

  function saveHistory() {
    try {
      localStorage.setItem(HISTORY_KEY, JSON.stringify(conversations.slice(0, MAX_HISTORY_CONV)));
    } catch (err) {
      console.warn('保存历史失败：', err);
    }
  }

  function stripContext(content) {
    const marker = '\n\n以下是当前 Word 文档中读取到的上下文';
    const i = content.indexOf(marker);
    return i >= 0 ? content.slice(0, i) : content;
  }

  function getTitleFromText(text) {
    const t = stripContext(text).trim().replace(/\s+/g, ' ');
    return (t.length > 24 ? t.slice(0, 24) + '…' : t) || '（无标题）';
  }

  function saveCurrentConversation() {
    if (!conversation.length) return;
    if (!currentConvId) {
      currentConvId = 'conv_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    }
    const firstUser = conversation.find((m) => m.role === 'user');
    const rec = {
      id: currentConvId,
      title: getTitleFromText(firstUser ? firstUser.content : ''),
      messages: conversation.slice(),
      updatedAt: Date.now(),
    };
    const idx = conversations.findIndex((c) => c.id === currentConvId);
    if (idx >= 0) {
      conversations[idx] = rec;
    } else {
      conversations.unshift(rec);
    }
    if (conversations.length > MAX_HISTORY_CONV) {
      conversations = conversations.slice(0, MAX_HISTORY_CONV);
    }
    saveHistory();
  }

  function formatTime(ts) {
    const min = Math.floor((Date.now() - ts) / 60000);
    if (min < 1) return '刚刚';
    if (min < 60) return min + ' 分钟前';
    const hr = Math.floor(min / 60);
    if (hr < 24) return hr + ' 小时前';
    const day = Math.floor(hr / 24);
    if (day < 7) return day + ' 天前';
    return new Date(ts).toLocaleDateString();
  }

  function renderHistoryList() {
    els.historyList.innerHTML = '';
    const sorted = conversations.slice().sort((a, b) => b.updatedAt - a.updatedAt);
    if (!sorted.length) {
      els.historyList.innerHTML = '<div class="history-empty">暂无历史对话</div>';
      return;
    }
    for (const conv of sorted) {
      const li = document.createElement('li');
      li.className = 'history-item' + (conv.id === currentConvId ? ' active' : '');

      const title = document.createElement('span');
      title.className = 'history-title';
      title.textContent = conv.title;

      const time = document.createElement('span');
      time.className = 'history-time';
      time.textContent = formatTime(conv.updatedAt);

      const del = document.createElement('button');
      del.className = 'history-del';
      del.textContent = '✕';
      del.title = '删除';
      del.addEventListener('click', (e) => {
        e.stopPropagation();
        deleteConversation(conv.id);
      });

      li.appendChild(title);
      li.appendChild(time);
      li.appendChild(del);
      li.addEventListener('click', () => loadConversation(conv.id));
      els.historyList.appendChild(li);
    }
  }

  function openHistory() {
    renderHistoryList();
    els.historyPanel.classList.add('open');
    els.historyBackdrop.classList.remove('hidden');
  }

  function closeHistory() {
    els.historyPanel.classList.remove('open');
    els.historyBackdrop.classList.add('hidden');
  }

  function deleteConversation(id) {
    conversations = conversations.filter((c) => c.id !== id);
    if (currentConvId === id) currentConvId = null;
    saveHistory();
    renderHistoryList();
  }

  function loadConversation(id) {
    saveCurrentConversation();
    const conv = conversations.find((c) => c.id === id);
    if (!conv) return;

    currentConvId = id;
    conversation = conv.messages.slice();
    lastAssistantText = '';
    els.messages.innerHTML = '';
    els.actionBar.classList.add('hidden');

    for (const m of conversation) {
      if (m.role === 'user') {
        addMessage('user', stripContext(m.content));
      } else if (m.role === 'assistant') {
        lastAssistantText = m.content;
        addMessage('assistant', m.content);
      }
    }
    if (lastAssistantText) {
      showActions();
      updateActionNote();
    }
    renderWelcome();
    closeHistory();
    updateSelectionInfo();
  }

  /* ---------- Word 集成 ---------- */

  function inOffice() {
    return typeof Office !== 'undefined' && typeof Word !== 'undefined' && Office.context;
  }

  async function updateSelectionInfo() {
    if (!inOffice()) {
      els.ctxInfo.textContent = '';
      return;
    }
    try {
      await Word.run(async (context) => {
        const sel = context.document.getSelection().getRange();
        context.load(sel, 'text');
        await context.sync();
        const len = (sel.text || '').trim().length;
        els.ctxInfo.textContent = len > 0 ? `已选 ${len} 字` : '当前无选中内容';
      });
    } catch {
      els.ctxInfo.textContent = '';
    }
  }

  async function refreshDocInfo() {
    if (typeof Word === 'undefined') {
      els.ctxInfo.textContent = 'Word API 未加载';
      return;
    }
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        context.load(body, 'text');
        await context.sync();
        const len = (body.text || '').trim().length;
        els.ctxInfo.textContent = len > 0 ? `文档共 ${len} 字` : '文档为空';
      });
    } catch (err) {
      els.ctxInfo.textContent = '读取文档失败：' + (err && err.message ? err.message : String(err));
    }
  }

  /**
   * 按需读取文档上下文，附加到用户消息里。
   */
  async function gatherContext() {
    const wantSelection = els.includeSelection.checked;
    const wantDoc = els.includeDocStart.checked;
    if (!wantSelection && !wantDoc) {
      els.ctxInfo.textContent = '未勾选“附带选中内容/附带文档内容”';
      return '';
    }
    if (typeof Word === 'undefined') {
      els.ctxInfo.textContent = 'Word API 未加载，无法读取文档';
      return '';
    }
    try {
      return await Word.run(async (context) => {
        let selRef = null;
        if (wantSelection) {
          selRef = context.document.getSelection().getRange();
          context.load(selRef, 'text');
        }
        let bodyRef = null;
        if (wantDoc) {
          bodyRef = context.document.body;
          context.load(bodyRef, 'text');
        }
        await context.sync();

        const parts = [];
        if (wantSelection) {
          const selText = (selRef.text || '').trim();
          els.ctxInfo.textContent = `已选 ${selText.length} 字`;
          if (selText.length > 0) parts.push(`【当前选中内容】（${selText.length} 字）：\n${selText.slice(0, 2000)}`);
        }
        if (wantDoc) {
          const docText = (bodyRef.text || '').trim();
          els.ctxInfo.textContent = (els.ctxInfo.textContent ? els.ctxInfo.textContent + '；' : '') + `已读文档 ${docText.length} 字`;
          if (docText.length > 0) parts.push(`【文档内容】（全文 ${docText.length} 字，附前 16000 字）：\n${docText.slice(0, 16000)}`);
        }
        if (parts.length === 0) {
          els.ctxInfo.textContent = (els.ctxInfo.textContent || '') + '；无可附带内容';
          return '';
        }
        return (
          '以下是当前 Word 文档中读取到的上下文（来自“附带选中内容/附带文档内容”选项），请基于这些内容回答，不要编造文档里不存在的细节：\n' +
          parts.join('\n\n')
        );
      });
    } catch (err) {
      els.ctxInfo.textContent = '读取文档失败：' + (err && err.message ? err.message : String(err));
      return '';
    }
  }

  function normalizeText(t) {
    return t.replace(/\r\n/g, '\r').replace(/\n/g, '\r');
  }

  function extractInsertText(text) {
    const blocks = [];
    const re = /```([\s\S]*?)```/g;
    let m = null;
    while ((m = re.exec(text)) !== null) {
      let content = m[1];
      const lines = content.split('\n');
      if (lines.length > 1 && /^[A-Za-z0-9_+-]+\s*$/.test(lines[0].trim())) {
        lines.shift();
        content = lines.join('\n');
      }
      blocks.push(content.trim());
    }
    return blocks.length ? blocks.join('\n\n') : null;
  }

  function hasCodeBlock(text) {
    return /```/.test(text);
  }

  function updateActionNote() {
    els.actionNote.textContent = hasCodeBlock(lastAssistantText)
      ? '检测到代码块：点击下方按钮将只插入代码块内的正文（修改说明不会写入）'
      : '将插入整条回复';
  }

  async function applyAction(kind) {
    const text = lastAssistantText;
    if (!text) return;
    if (!inOffice()) {
      toast('当前不在 Word 中运行，无法写入文档', true);
      return;
    }
    const extracted = extractInsertText(text);
    const finalText = extracted !== null ? extracted : text;
    try {
      await Word.run(async (context) => {
        const selRange = context.document.getSelection().getRange();
        if (kind === 'replace') {
          selRange.insertText(normalizeText(finalText), Word.InsertLocation.replace);
        } else if (kind === 'after') {
          selRange.insertText(normalizeText(finalText), Word.InsertLocation.after);
        } else if (kind === 'start') {
          context.document.body.insertText(normalizeText(finalText), Word.InsertLocation.start);
        } else if (kind === 'end') {
          context.document.body.insertText(normalizeText(finalText), Word.InsertLocation.end);
        }
        await context.sync();
      });
      const labels = { replace: '已替换选中内容', after: '已插入到光标处', start: '已插入到文首', end: '已插入到文末' };
      toast((labels[kind] || '已应用到文档') + (extracted !== null ? '（仅代码块正文）' : ''));
    } catch (err) {
      toast('应用失败：' + err.message, true);
    }
  }

  /* ---------- 对话逻辑 ---------- */

  async function streamChat(bubble) {
    lastAssistantText = '';
    let res;
    try {
      res = await fetch(DEEPSEEK_API, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model,
          messages: [{ role: 'system', content: SYSTEM_PROMPT }, ...conversation],
          stream: true,
          max_tokens: 4096,
        }),
      });
    } catch (err) {
      throw new Error('无法连接 DeepSeek，请检查网络（' + err.message + '）');
    }

    if (!res.ok) {
      let detail = '';
      try {
        const j = await res.json();
        detail = (j.error && (j.error.message || j.error)) || JSON.stringify(j);
      } catch {
        detail = await res.text();
      }
      if (res.status === 401) throw new Error('API Key 无效或未填写（401），请在设置中检查');
      if (res.status === 402) throw new Error('DeepSeek 账户余额不足（402），请前往平台充值');
      if (res.status === 429) throw new Error('请求过于频繁（429），请稍后再试');
      throw new Error(detail || `请求失败（HTTP ${res.status}）`);
    }

    if (!res.body) {
      const j = await res.json();
      lastAssistantText = (j.choices && j.choices[0] && j.choices[0].message.content) || '';
      bubble.innerHTML = formatMessage(lastAssistantText) || '（空回复）';
      return;
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder('utf-8');
    let buf = '';

    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });

      let sep;
      while ((sep = buf.indexOf('\n\n')) >= 0) {
        const raw = buf.slice(0, sep);
        buf = buf.slice(sep + 2);
        for (const line of raw.split('\n')) {
          if (!line.startsWith('data:')) continue;
          const data = line.slice(5).trim();
          if (data === '[DONE]') continue;
          try {
            const j = JSON.parse(data);
            const delta = j.choices && j.choices[0] && j.choices[0].delta
              ? j.choices[0].delta.content || ''
              : '';
            if (delta) {
              lastAssistantText += delta;
              bubble.innerHTML = formatMessage(lastAssistantText);
              scrollToBottom();
            }
          } catch {
            /* 忽略无法解析的片段 */
          }
        }
      }
    }

    if (!lastAssistantText) {
      throw new Error('未收到模型回复');
    }
    bubble.innerHTML = formatMessage(lastAssistantText);
  }

  async function handleSend() {
    const text = els.userInput.value.trim();
    if (!text || streaming) return;
    if (!apiKey) {
      toast('请先在设置中填写 DeepSeek API Key', true);
      openSettings();
      return;
    }
    if (!model || LEGACY_MODELS.includes(model)) {
      toast('当前模型已弃用或未选择，请先在设置中重新选择模型', true);
      openSettings();
      return;
    }

    streaming = true;
    els.sendBtn.disabled = true;
    els.sendBtn.textContent = '生成中';
    els.userInput.value = '';

    addMessage('user', text);

    // 附加文档上下文
    let contextNote = '';
    try {
      contextNote = await gatherContext();
    } catch (err) {
      console.warn('读取文档上下文失败：', err);
      els.ctxInfo.textContent = '读取上下文失败：' + (err && err.message ? err.message : String(err));
    }
    const userContent = contextNote ? `${text}\n\n${contextNote}` : text;
    conversation.push({ role: 'user', content: userContent });
    setStatus(
      contextNote
        ? '已附带文档上下文，正在生成回复…'
        : '未附带文档内容（请勾选“附带文档内容”或“附带选中内容”）',
      !contextNote
    );

    const bubble = addMessage('assistant', '', true);
    try {
      await streamChat(bubble);
      conversation.push({ role: 'assistant', content: lastAssistantText });
      // 控制历史长度
      if (conversation.length > MAX_HISTORY * 2) {
        conversation = conversation.slice(conversation.length - MAX_HISTORY * 2);
      }
      saveCurrentConversation();
      showActions();
      updateActionNote();
    } catch (err) {
      bubble.classList.add('error');
      bubble.innerHTML = '请求失败：' + escapeHtml(err.message || String(err));
      conversation.pop();
    }

    streaming = false;
    els.sendBtn.disabled = false;
    els.sendBtn.textContent = '发送';
    els.userInput.focus();
  }

  function newChat() {
    saveCurrentConversation();
    conversation = [];
    lastAssistantText = '';
    currentConvId = null;
    els.messages.innerHTML = '';
    els.actionBar.classList.add('hidden');
    renderWelcome();
    updateSelectionInfo();
  }

  /* ---------- 设置 ---------- */

  function openSettings() {
    els.apiKeyInput.value = apiKey;
    els.customModelInput.value = LEGACY_MODELS.includes(model) ? '' : model;
    els.settingsModal.classList.remove('hidden');
    els.saveSettingsBtn.disabled = true;
    if (apiKey) {
      els.modelSelect.innerHTML = '<option value="">正在获取模型列表…</option>';
      fetchModels();
    } else {
      els.modelSelect.innerHTML = '<option value="">填写 API Key 后点击“获取模型列表”</option>';
      setSettingsStatus('请先填写 DeepSeek API Key，再获取模型列表');
    }
    els.apiKeyInput.focus();
  }

  function closeSettings() {
    els.settingsModal.classList.add('hidden');
  }

  function setSettingsStatus(msg, isError = false) {
    els.settingsStatus.textContent = msg;
    els.settingsStatus.style.color = isError ? '#b91c1c' : '#6b7280';
  }

  function populateModelSelect(list) {
    els.modelSelect.innerHTML = '';
    const opts = list.slice();
    if (model && !LEGACY_MODELS.includes(model) && !opts.includes(model)) {
      opts.unshift(model);
    }
    let selected = false;
    for (const id of opts) {
      const opt = document.createElement('option');
      opt.value = id;
      opt.textContent = id;
      if (id === model && !LEGACY_MODELS.includes(model)) {
        opt.selected = true;
        selected = true;
      }
      els.modelSelect.appendChild(opt);
    }
    if (!selected && opts.length) {
      els.modelSelect.selectedIndex = 0;
    }
    els.modelSelect.disabled = false;
    els.saveSettingsBtn.disabled = false;
  }

  async function fetchModels() {
    const key = els.apiKeyInput.value.trim();
    if (!key) {
      setSettingsStatus('请先填写 DeepSeek API Key', true);
      return;
    }
    apiKey = key;
    localStorage.setItem('ds_api_key', key);
    refreshStatus();
    setSettingsStatus('正在获取模型列表…');
    try {
      const res = await fetch(MODELS_API, {
        headers: { Authorization: `Bearer ${key}` },
      });
      if (!res.ok) {
        let detail = '';
        try {
          const j = await res.json();
          detail = (j.error && (j.error.message || j.error)) || JSON.stringify(j);
        } catch {
          detail = await res.text();
        }
        if (res.status === 401) throw new Error('API Key 无效（401）');
        if (res.status === 429) throw new Error('请求过于频繁（429），请稍后再试');
        throw new Error(detail || `请求失败（HTTP ${res.status}）`);
      }
      const j = await res.json();
      const list = (j.data || []).map((m) => m.id).filter(Boolean);
      if (!list.length) throw new Error('模型列表为空');
      cachedModels = list;
      populateModelSelect(list);
      setSettingsStatus('模型列表已更新（共 ' + list.length + ' 个），选择后点击“保存设置”');
    } catch (err) {
      populateModelSelect(FALLBACK_MODELS);
      setSettingsStatus('获取模型失败：' + err.message + '。已显示当前官方模型（deepseek-v4-flash / deepseek-v4-pro），也可在"自定义模型"中手动填写。', true);
    }
  }

  function saveSettings() {
    apiKey = els.apiKeyInput.value.trim();
    const custom = els.customModelInput.value.trim();
    const selected = els.modelSelect.value;
    if (!custom && !selected) {
      toast('请先获取模型列表并选择模型', true);
      return;
    }
    model = custom || selected;
    if (LEGACY_MODELS.includes(model)) {
      toast('警告：' + model + ' 已弃用（2026-07-24），请选择 deepseek-v4-flash 或 deepseek-v4-pro', true);
      return;
    }
    localStorage.setItem('ds_api_key', apiKey);
    localStorage.setItem('ds_model', model);
    closeSettings();
    refreshStatus();
    toast('设置已保存：' + model);
  }

  /* ---------- 初始化 ---------- */

  function bind() {
    els.sendBtn.addEventListener('click', handleSend);
    els.userInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey && !e.isComposing && e.keyCode !== 229) {
        e.preventDefault();
        handleSend();
      }
    });
    els.newChatBtn.addEventListener('click', newChat);
    els.settingsBtn.addEventListener('click', openSettings);
    els.fetchModelsBtn.addEventListener('click', fetchModels);
    els.customModelInput.addEventListener('input', () => {
      els.saveSettingsBtn.disabled = !els.customModelInput.value.trim() && !els.modelSelect.value;
    });
    els.modelSelect.addEventListener('change', () => {
      els.saveSettingsBtn.disabled = false;
    });
    els.historyBtn.addEventListener('click', openHistory);
    els.historyCloseBtn.addEventListener('click', closeHistory);
    els.historyBackdrop.addEventListener('click', closeHistory);
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeHistory();
    });
    els.cancelSettingsBtn.addEventListener('click', closeSettings);
    els.saveSettingsBtn.addEventListener('click', saveSettings);
    els.settingsModal.addEventListener('click', (e) => {
      if (e.target === els.settingsModal) closeSettings();
    });
    els.refreshCtxBtn.addEventListener('click', updateSelectionInfo);
    els.actionBar.querySelectorAll('.action-btn').forEach((btn) => {
      btn.addEventListener('click', () => applyAction(btn.dataset.action));
    });
  }

  function refreshStatus() {
    setStatus(apiKey ? '直连 DeepSeek（无需本地服务）' : '请先在设置中填写 DeepSeek API Key', !apiKey);
  }

  document.addEventListener('DOMContentLoaded', () => {
    bind();
    renderWelcome();
    refreshStatus();
    if (!apiKey) openSettings();
  });

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady((info) => {
      const host = info && info.host;
      if (!host) {
        els.hostBadge.textContent = '未运行在 Word 中';
        els.hostBadge.classList.add('warn');
        setStatus('当前页面未作为 Word 加载项运行，聊天可用，但文档读写不可用。', true);
        return;
      }
      els.hostBadge.textContent = `已连接 ${host}`;
      updateSelectionInfo();
      refreshDocInfo();
    });
  } else {
    els.hostBadge.textContent = 'Office.js 未加载';
    els.hostBadge.classList.add('warn');
    setStatus('Office.js 加载失败，请检查网络后重新打开任务窗格。', true);
  }
})();
