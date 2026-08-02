/* DeepSeek Word 助手 - 任务窗格逻辑 */
/* global Office, Word, localStorage */

(() => {
  'use strict';

  // 直连 DeepSeek 官方 API（已实测支持浏览器 CORS 跨域，无需本地服务）
  const DEEPSEEK_API = 'https://api.deepseek.com/chat/completions';
  const MAX_HISTORY = 20; // 发送给模型的对话轮数上限
  const HISTORY_KEY = 'ds_conversations';
  const MAX_HISTORY_CONV = 20;

  let apiKey = localStorage.getItem('ds_api_key') || '';
  let model = localStorage.getItem('ds_model') || 'deepseek-chat';
  let conversation = [];
  let streaming = false;
  let lastAssistantText = '';
  let currentConvId = null;
  let conversations = loadHistory();

  const $ = (id) => document.getElementById(id);

  const els = {
    hostBadge: $('hostBadge'),
    messages: $('messages'),
    userInput: $('userInput'),
    sendBtn: $('sendBtn'),
    actionBar: $('actionBar'),
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
    if (lastAssistantText) showActions();
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
        const sel = context.document.getSelection();
        const t = sel.getText();
        await context.sync();
        const len = (t.value || '').trim().length;
        els.ctxInfo.textContent = len > 0 ? `已选 ${len} 字` : '当前无选中内容';
      });
    } catch {
      els.ctxInfo.textContent = '';
    }
  }

  /**
   * 按需读取文档上下文，附加到用户消息里。
   */
  async function gatherContext() {
    const wantSelection = els.includeSelection.checked;
    const wantDoc = els.includeDocStart.checked;
    if (!inOffice() || (!wantSelection && !wantDoc)) {
      els.ctxInfo.textContent = '';
      return '';
    }

    return Word.run(async (context) => {
      const jobs = [];
      if (wantSelection) {
        const sel = context.document.getSelection();
        jobs.push({ kind: 'sel', ref: sel.getText() });
      }
      if (wantDoc) {
        jobs.push({ kind: 'doc', ref: context.document.body.getText() });
      }
      await context.sync();

      const parts = [];
      for (const job of jobs) {
        const text = (job.ref.value || '').trim();
        if (job.kind === 'sel') {
          els.ctxInfo.textContent = `已选 ${text.length} 字`;
          if (text.length > 0) parts.push(`【当前选中内容】（${text.length} 字）：\n${text.slice(0, 2000)}`);
        } else {
          if (text.length > 0) parts.push(`【文档开头】（前 4000 字）：\n${text.slice(0, 4000)}`);
        }
      }
      if (parts.length === 0) return '';
      return (
        '以下是当前 Word 文档中读取到的上下文（来自“附带选中内容/文档开头”选项），请基于这些内容回答，不要编造文档里不存在的细节：\n' +
        parts.join('\n\n')
      );
    });
  }

  function normalizeText(t) {
    return t.replace(/\r\n/g, '\r').replace(/\n/g, '\r');
  }

  async function applyAction(kind) {
    const text = lastAssistantText;
    if (!text) return;
    if (!inOffice()) {
      toast('当前不在 Word 中运行，无法写入文档', true);
      return;
    }
    try {
      await Word.run(async (context) => {
        if (kind === 'replace') {
          context.document.getSelection().insertText(normalizeText(text), Word.InsertLocation.replace);
        } else if (kind === 'after') {
          context.document.getSelection().insertText(normalizeText(text), Word.InsertLocation.after);
        } else if (kind === 'start') {
          context.document.body.insertText(normalizeText(text), Word.InsertLocation.start);
        } else if (kind === 'end') {
          context.document.body.insertText(normalizeText(text), Word.InsertLocation.end);
        }
        await context.sync();
      });
      const labels = { replace: '已替换选中内容', after: '已插入到光标处', start: '已插入到文首', end: '已插入到文末' };
      toast(labels[kind] || '已应用到文档');
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
          messages: conversation,
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
      els.ctxInfo.textContent = '读取上下文失败';
    }
    const userContent = contextNote ? `${text}\n\n${contextNote}` : text;
    conversation.push({ role: 'user', content: userContent });

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
    els.modelSelect.value = ['deepseek-chat', 'deepseek-reasoner'].includes(model) ? model : 'deepseek-chat';
    els.customModelInput.value = ['deepseek-chat', 'deepseek-reasoner'].includes(model) ? '' : model;
    els.settingsModal.classList.remove('hidden');
    els.apiKeyInput.focus();
  }

  function closeSettings() {
    els.settingsModal.classList.add('hidden');
  }

  function saveSettings() {
    apiKey = els.apiKeyInput.value.trim();
    const custom = els.customModelInput.value.trim();
    model = custom || els.modelSelect.value;
    localStorage.setItem('ds_api_key', apiKey);
    localStorage.setItem('ds_model', model);
    closeSettings();
    refreshStatus();
    toast('设置已保存' + (apiKey ? '' : '（未填写 API Key）'));
  }

  /* ---------- 初始化 ---------- */

  function bind() {
    els.sendBtn.addEventListener('click', handleSend);
    els.userInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        handleSend();
      }
    });
    els.newChatBtn.addEventListener('click', newChat);
    els.settingsBtn.addEventListener('click', openSettings);
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
    });
  } else {
    els.hostBadge.textContent = 'Office.js 未加载';
    els.hostBadge.classList.add('warn');
    setStatus('Office.js 加载失败，请检查网络后重新打开任务窗格。', true);
  }
})();
