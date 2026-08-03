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
  let reasoningEffort = localStorage.getItem('ds_reasoning_effort') || 'high';
  let conversation = [];
  let streaming = false;
  let lastAssistantText = '';
  let lastReasoningText = '';
  let latestEdit = null; // 最新回复的修改记录 { insertedText, replacedText, applied, bubbleEl, isLatest }
  let editLog = []; // 更早回复的已应用修改记录（按应用顺序）
  let pendingRollback = null; // 待回退的记录（确认弹窗）
  let editSeq = 0; // 用于生成唯一的书签名/定位标记
  const MAX_EDITS = 20; // 内存中保留的存档点数上限
  let currentConvId = null;
  let conversations = loadHistory();
  let cachedModels = null;

  const $ = (id) => document.getElementById(id);

  const els = {
    hostBadge: $('hostBadge'),
    messages: $('messages'),
    userInput: $('userInput'),
    sendBtn: $('sendBtn'),
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
    effortSelect: $('effortSelect'),
    customModelInput: $('customModelInput'),
    fetchModelsBtn: $('fetchModelsBtn'),
    settingsStatus: $('settingsStatus'),
    saveSettingsBtn: $('saveSettingsBtn'),
    cancelSettingsBtn: $('cancelSettingsBtn'),
    confirmModal: $('confirmModal'),
    confirmMsg: $('confirmMsg'),
    confirmOkBtn: $('confirmOkBtn'),
    confirmCancelBtn: $('confirmCancelBtn'),
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

  // 把 Markdown 风格的回复转换成 Word insertHtml 支持的 HTML 子集，
  // 让标题/加粗/列表/代码块等格式能真实落到文档里
  function markdownToWordHtml(text) {
    const lines = String(text || '').split(/\r\n|\r|\n/);
    const out = [];
    let i = 0;
    let inCode = false;
    let codeBuf = [];
    let para = [];

    function inline(s) {
      return escapeHtml(s)
        .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
        .replace(/(^|[^*])\*([^*]+)\*/g, '$1<em>$2</em>')
        .replace(/~~([^~]+)~~/g, '<s>$1</s>')
        .replace(/`([^`]+)`/g, '<span style="font-family:Consolas,monospace">$1</span>')
        .replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, '<a href="$2">$1</a>');
    }

    function flushPara() {
      if (para.length) {
        out.push('<p>' + inline(para.join(' ')) + '</p>');
        para = [];
      }
    }

    while (i < lines.length) {
      const line = lines[i];
      const trimmed = line.trim();

      if (trimmed.startsWith('```')) {
        if (!inCode) {
          flushPara();
          inCode = true;
          codeBuf = [];
        } else {
          out.push('<p style="font-family:Consolas,monospace">' + escapeHtml(codeBuf.join('\n')).replace(/\n/g, '<br>') + '</p>');
          inCode = false;
          codeBuf = [];
        }
        i++;
        continue;
      }

      if (inCode) {
        codeBuf.push(line);
        i++;
        continue;
      }

      if (!trimmed) {
        flushPara();
        i++;
        continue;
      }

      const h = /^(#{1,6})\s+(.*)$/.exec(trimmed);
      if (h) {
        flushPara();
        const lv = h[1].length;
        out.push('<h' + lv + '>' + inline(h[2]) + '</h' + lv + '>');
        i++;
        continue;
      }

      const ul = /^[-*]\s+(.*)$/.exec(trimmed);
      if (ul) {
        flushPara();
        const items = [];
        while (i < lines.length && /^[-*]\s+/.test(lines[i].trim())) {
          items.push('<li>' + inline(lines[i].trim().replace(/^[-*]\s+/, '')) + '</li>');
          i++;
        }
        out.push('<ul>' + items.join('') + '</ul>');
        continue;
      }

      const ol = /^\d+[.)]\s+(.*)$/.exec(trimmed);
      if (ol) {
        flushPara();
        const items = [];
        while (i < lines.length && /^\d+[.)]\s+/.test(lines[i].trim())) {
          items.push('<li>' + inline(lines[i].trim().replace(/^\d+[.)]\s+/, '')) + '</li>');
          i++;
        }
        out.push('<ol>' + items.join('') + '</ol>');
        continue;
      }

      const q = /^>\s?(.*)$/.exec(trimmed);
      if (q) {
        flushPara();
        const quotes = [];
        while (i < lines.length && /^>\s?/.test(lines[i].trim())) {
          quotes.push(inline(lines[i].trim().replace(/^>\s?/, '')));
          i++;
        }
        out.push('<p style="margin-left:24px;color:#6b7280">' + quotes.join('<br>') + '</p>');
        continue;
      }

      para.push(trimmed);
      i++;
    }

    flushPara();
    if (inCode && codeBuf.length) {
      out.push('<p style="font-family:Consolas,monospace">' + escapeHtml(codeBuf.join('\n')).replace(/\n/g, '<br>') + '</p>');
    }
    return out.join('');
  }

  function renderAssistantHtml(reasoning, content) {
    let html = '';
    if (reasoning) {
      html += '<details class="reasoning" open><summary>思考过程</summary><div class="reasoning-body">' +
        formatMessage(reasoning) + '</div></details>';
    }
    const body = formatMessage(content);
    if (body) html += body;
    if (!html) html = '（空）';
    return html;
  }

  // 仅在贴近底部时跟随最新内容；用户已上滑阅读时不打断其位置
  function smartScrollToBottom(el) {
    if (!el) return;
    if (el.scrollHeight - el.scrollTop - el.clientHeight > 64) return;
    el.scrollTop = el.scrollHeight;
  }

  function scrollToBottom() {
    smartScrollToBottom(els.messages);
  }

  // 流式更新气泡内容：原地更新“思考过程”与正文，不重建 DOM，
  // 这样思考过程中的滚动条可以正常拖拽，且折叠/展开状态不会被重置
  function renderStreamingBubble(bubble, reasoning, content) {
    const typing = bubble.querySelector('.typing');
    if (typing) typing.remove();

    let details = bubble.querySelector('.reasoning');
    if (reasoning) {
      if (!details) {
        details = document.createElement('details');
        details.className = 'reasoning';
        details.open = true;
        const summary = document.createElement('summary');
        summary.textContent = '思考过程';
        const body = document.createElement('div');
        body.className = 'reasoning-body';
        details.appendChild(summary);
        details.appendChild(body);
        bubble.insertBefore(details, bubble.firstChild);
      }
      const bodyEl = details.querySelector('.reasoning-body');
      bodyEl.innerHTML = formatMessage(reasoning);
      smartScrollToBottom(bodyEl);
    } else if (details) {
      details.remove();
    }

    let contentEl = bubble.querySelector('.msg-content');
    if (!contentEl) {
      contentEl = document.createElement('div');
      contentEl.className = 'msg-content';
      bubble.appendChild(contentEl);
    }
    const formatted = formatMessage(content);
    contentEl.innerHTML = formatted || '';
    if (!formatted && !reasoning) contentEl.innerHTML = '（空）';
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

  function addMessage(role, text, isStreaming = false, reasoning = '') {
    const div = document.createElement('div');
    div.className = `msg ${role}`;
    if (isStreaming) {
      div.innerHTML = '<span class="typing"></span>';
    } else {
      div.innerHTML = role === 'assistant'
        ? renderAssistantHtml(reasoning, text)
        : formatMessage(text) || '（空）';
    }
    els.messages.appendChild(div);
    scrollToBottom();
    return div;
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
    latestEdit = null;
    editLog = [];
    pendingRollback = null;

    for (const m of conversation) {
      if (m.role === 'user') {
        addMessage('user', stripContext(m.content));
      } else if (m.role === 'assistant') {
        lastAssistantText = m.content;
        lastReasoningText = m.reasoning || '';
        addMessage('assistant', m.content, false, lastReasoningText);
      }
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

  function bookmarksSupported(context) {
    return (
      typeof context.document.bookmarks !== 'undefined' &&
      typeof context.document.bookmarks.add === 'function' &&
      typeof context.document.bookmarks.getItemOrNullObject === 'function'
    );
  }

  const MARKER_SEARCH_OPTS = {
    matchCase: true,
    matchWholeWord: false,
    ignorePunct: false,
    ignoreSpace: false,
  };

  // 生成只在本次写入中存在的唯一定位标记（PUA 字符 + 随机串，几乎不可能与正文冲突）
  function makeMarker() {
    return (
      '\uE000DSINS' + Date.now().toString(36) + '_' + (++editSeq) + '_' +
      Math.random().toString(36).slice(2, 8) + '\uE001'
    );
  }

  async function performApply(rec) {
    await Word.run(async (context) => {
      const selRange = context.document.getSelection().getRange();
      context.load(selRange, 'text');
      await context.sync();
      const replacedText = (selRange.text || '').trim();
      rec.replacedText = replacedText;

      if (!bookmarksSupported(context)) {
        // 旧版 Word 不支持书签：退化为整文快照（文档很长时会较慢，但保证兼容）
        const before = context.document.body.getOoxml();
        await context.sync();
        rec.beforeOoxml = before.value;
        const html = rec.insertedHtml || markdownToWordHtml(rec.insertedText);
        const canHtml = typeof selRange.insertHtml === 'function';
        const payload = canHtml ? html : rec.insertedText;
        const loc = replacedText.length > 0 ? Word.InsertLocation.replace : Word.InsertLocation.after;
        if (canHtml) {
          selRange.insertHtml(payload, loc);
        } else {
          selRange.insertText(payload, loc);
        }
        await context.sync();
        const after = context.document.body.getOoxml();
        await context.sync();
        rec.afterOoxml = after.value;
        return;
      }

      // 1) 只保存被替换内容（选区级 OOXML，体积与文档长度无关），并把选区清空成插入点
      if (replacedText.length > 0) {
        const before = selRange.getOoxml();
        selRange.delete();
        await context.sync();
        rec.beforeOoxml = before.value;
      } else {
        rec.beforeOoxml = null;
      }

      // 2) 在插入点放唯一标记，用书签圈住标记，再在书签末尾写入真实文本
      //    （Word 的“end”插入发生在书签范围内，书签会保留并扩展覆盖新内容）
      const marker = makeMarker();
      selRange.insertText(marker, Word.InsertLocation.after);
      await context.sync();
      const hits = context.document.body.search(marker, MARKER_SEARCH_OPTS);
      context.load(hits, 'text');
      await context.sync();
      if (!hits.items || !hits.items.length) {
        throw new Error('无法定位插入位置，请重试');
      }
      const hitRange = hits.items[0];
      const bmName = 'DS_EDIT_' + Date.now().toString(36) + '_' + editSeq;
      const newBm = context.document.bookmarks.add(bmName, hitRange);
      const html = rec.insertedHtml || markdownToWordHtml(rec.insertedText);
      const range = newBm.getRange();
      if (typeof range.insertHtml === 'function') {
        range.insertHtml(html, Word.InsertLocation.end);
      } else {
        range.insertText(rec.insertedText, Word.InsertLocation.end);
      }
      await context.sync();

      // 3) 删除标记（书签已扩展覆盖真实文本，删除标记不会影响书签）
      const hits2 = context.document.body.search(marker, MARKER_SEARCH_OPTS);
      context.load(hits2, 'text');
      await context.sync();
      if (hits2.items && hits2.items.length) {
        hits2.items[0].delete();
      }
      await context.sync();
      rec.bookmarkName = bmName;
    });
  }

  // 兜底方案：整文 OOXML 快照（仅在不支持书签的旧版 Word 使用）
  async function getBodyOoxml() {
    return Word.run(async (context) => {
      const result = context.document.body.getOoxml();
      await context.sync();
      return result.value;
    });
  }

  async function restoreBodyOoxml(ooxml) {
    await Word.run(async (context) => {
      context.document.body.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();
    });
  }

  // 将一次修改恢复到应用前的状态（只操作书签标记的局部范围，与文档长度无关）
  async function restoreEdit(rec, keepBookmark) {
    if (!rec.bookmarkName) {
      if (!rec.beforeOoxml) throw new Error('没有存档点，无法撤销');
      await restoreBodyOoxml(rec.beforeOoxml);
      return;
    }
    await Word.run(async (context) => {
      const bm = context.document.bookmarks.getItemOrNullObject(rec.bookmarkName);
      await context.sync();
      if (bm.isNullObject) throw new Error('修改标记已丢失（文档可能已被改动），无法撤销');
      const bmRange = bm.getRange();
      if (keepBookmark) {
        // 撤销后仍要能“接受”：内容操作可能连带删除书签，若被删则在原区域重建
        if (rec.beforeOoxml) {
          bmRange.insertOoxml(rec.beforeOoxml, Word.InsertLocation.replace);
        } else {
          bmRange.delete();
        }
        await context.sync();
        const bm2 = context.document.bookmarks.getItemOrNullObject(rec.bookmarkName);
        await context.sync();
        if (bm2.isNullObject) {
          context.document.bookmarks.add(rec.bookmarkName, bmRange);
          await context.sync();
        }
      } else {
        // 记录被丢弃（回退等）：先删书签，再恢复内容
        bm.delete();
        if (rec.beforeOoxml) {
          bmRange.insertOoxml(rec.beforeOoxml, Word.InsertLocation.replace);
        } else {
          bmRange.delete();
        }
        await context.sync();
      }
    });
  }

  // 重新应用：清空书签当前位置的内容，再按与首次写入相同的流程写入真实文本
  // （不改变 beforeOoxml 存档，因此之后仍可撤销）
  async function applyTextToBookmark(rec) {
    await Word.run(async (context) => {
      const bm = context.document.bookmarks.getItemOrNullObject(rec.bookmarkName);
      await context.sync();
      if (bm.isNullObject) throw new Error('修改标记已丢失（文档可能已被改动），无法应用');
      const bmRange = bm.getRange();
      // 先删书签再清空内容，避免内容操作连带删除书签
      bm.delete();
      bmRange.delete();
      await context.sync();

      const marker = makeMarker();
      bmRange.insertText(marker, Word.InsertLocation.after);
      await context.sync();
      const hits = context.document.body.search(marker, MARKER_SEARCH_OPTS);
      context.load(hits, 'text');
      await context.sync();
      if (!hits.items || !hits.items.length) {
        throw new Error('无法定位插入位置，请重试');
      }
      const hitRange = hits.items[0];
      const newBm = context.document.bookmarks.add(rec.bookmarkName, hitRange);
      const html = rec.insertedHtml || markdownToWordHtml(rec.insertedText);
      const range = newBm.getRange();
      if (typeof range.insertHtml === 'function') {
        range.insertHtml(html, Word.InsertLocation.end);
      } else {
        range.insertText(rec.insertedText, Word.InsertLocation.end);
      }
      await context.sync();

      const hits2 = context.document.body.search(marker, MARKER_SEARCH_OPTS);
      context.load(hits2, 'text');
      await context.sync();
      if (hits2.items && hits2.items.length) {
        hits2.items[0].delete();
      }
      await context.sync();
    });
  }

  // 记录被丢弃时清理文档中的书签（防止长期累积）
  async function deleteBookmark(rec) {
    if (!rec || !rec.bookmarkName || !inOffice()) return;
    try {
      await Word.run(async (context) => {
        const bm = context.document.bookmarks.getItemOrNullObject(rec.bookmarkName);
        await context.sync();
        if (!bm.isNullObject) bm.delete();
        await context.sync();
      });
    } catch (err) {
      console.warn('清理书签失败：', err);
    }
  }

  function appendToggleButton(bubble, record) {
    const wrap = document.createElement('div');
    wrap.className = 'reply-actions';
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'toggle-btn';
    btn.addEventListener('click', () => onRecordBtn(record));
    wrap.appendChild(btn);
    bubble.appendChild(wrap);
    record.btnEl = btn;
    updateButtonFor(record);
  }

  function updateButtonFor(record) {
    const btn = record.btnEl;
    if (!btn) return;
    btn.disabled = false;
    if (record.isLatest) {
      // 最新一条回复：沿用原逻辑，撤销/接受切换
      btn.textContent = record.applied ? '撤销' : '接受';
      btn.title = record.applied ? '撤销上次写入' : '将回复（含格式）应用到文档';
    } else {
      btn.textContent = record.applied ? '回退' : '接受';
      btn.title = record.applied ? '回退到此次修改之前' : '将回复应用到文档';
    }
  }

  function onRecordBtn(record) {
    if (record.isLatest) {
      if (record.applied) {
        undoLatest();
      } else {
        applyRecord(record);
      }
    } else if (record.applied) {
      showConfirmRollback(record);
    } else {
      applyRecord(record);
    }
  }

  async function applyRecord(record) {
    if (!inOffice()) {
      toast('当前不在 Word 中运行，无法应用', true);
      return;
    }
    try {
      const isReapply = record.hadApplied && !record.applied;
      if (record.bookmarkName) {
        // 书签标记了该次修改影响的区域：只在该区域重新写入，速度快且不动文档其他部分
        await applyTextToBookmark(record);
      } else if (record.afterOoxml) {
        // 旧版兼容：整文恢复
        await restoreBodyOoxml(record.afterOoxml);
      } else {
        // 首次应用（例如自动应用失败后手动点击“接受”）
        await performApply(record);
      }
      record.applied = true;
      record.hadApplied = true;
      if (!record.isLatest) editLog.push(record);
      if (editLog.length > MAX_EDITS) deleteBookmark(editLog.shift());
      updateButtonFor(record);
      toast(isReapply ? '已重新应用到文档' : '已应用到文档');
    } catch (err) {
      toast('应用失败：' + err.message, true);
    }
  }

  async function autoApplyReply(bubble) {
    const text = lastAssistantText;
    if (!text) return;
    const extracted = extractInsertText(text);
    // 只允许把回复中的代码块内容应用到文档。
    // 回复里的说明性文字（建议、解释等）不属于正文，绝不能写入文档，
    // 因此没有代码块时不显示“接受”按钮，防止误插入。
    if (extracted === null) return;
    // 上一条编辑降级为“之前的回复”：已应用的进入回退列表，未应用的移除按钮
    if (latestEdit) {
      if (latestEdit.applied) {
        latestEdit.isLatest = false;
        editLog.push(latestEdit);
        if (editLog.length > MAX_EDITS) deleteBookmark(editLog.shift());
      } else {
        const wrap = latestEdit.bubbleEl ? latestEdit.bubbleEl.querySelector('.reply-actions') : null;
        if (wrap) wrap.remove();
        deleteBookmark(latestEdit);
      }
    }
    const record = {
      insertedText: normalizeText(extracted),
      insertedHtml: '<p style="font-family:Consolas,monospace">' + escapeHtml(extracted).replace(/\n/g, '<br>') + '</p>',
      replacedText: '',
      applied: false,
      hadApplied: false,
      bubbleEl: bubble,
      isLatest: true,
      replyContent: text,
    };
    latestEdit = record;
    appendToggleButton(bubble, record);
    if (!inOffice()) return;
    try {
      await performApply(record);
      record.applied = true;
      record.hadApplied = true;
      updateButtonFor(record);
    } catch (err) {
      toast('自动应用失败：' + err.message, true);
    }
  }

  async function undoLatest() {
    const rec = latestEdit;
    if (!rec || !rec.applied) return;
    if (!inOffice()) {
      toast('当前不在 Word 中运行，无法撤销', true);
      return;
    }
    if (!rec.beforeOoxml && !rec.bookmarkName) {
      toast('没有存档点，无法撤销', true);
      return;
    }
    try {
      await restoreEdit(rec, true);
      rec.applied = false;
      updateButtonFor(rec);
      toast('已撤销，可点击“接受”重新应用');
    } catch (err) {
      toast('撤销失败：' + err.message, true);
    }
  }

  async function rollbackTo(record) {
    const idx = editLog.indexOf(record);
    if (idx < 0) return;
    try {
      // 先逆序撤销该回复之后的全部修改（这些记录将被丢弃，顺带删除其书签），
      // 再撤销该回复本身（保留书签，便于之后点“接受”重新应用）
      for (let i = editLog.length - 1; i > idx; i--) {
        await restoreEdit(editLog[i], false);
      }
      if (latestEdit && latestEdit.applied && latestEdit !== record) {
        await restoreEdit(latestEdit, false);
      }
      await restoreEdit(record, true);
      // 删除该回复之后的对话记录与气泡
      const keepIdx = conversation.findIndex(
        (m) => m.role === 'assistant' && record.replyContent && m.content === record.replyContent
      );
      if (keepIdx >= 0) {
        conversation = conversation.slice(0, keepIdx + 1);
        let node = record.bubbleEl;
        while (node && node.nextSibling) {
          node.parentNode.removeChild(node.nextSibling);
        }
      }
      record.applied = false;
      if (latestEdit && latestEdit !== record) latestEdit.applied = false;
      editLog = editLog.slice(0, idx);
      latestEdit = record;
      record.isLatest = true;
      updateButtonFor(record);
      saveCurrentConversation();
      toast('已回退到该回复之前，并删除其后的对话');
    } catch (err) {
      toast('回退失败：' + err.message, true);
    }
  }

  function showConfirmRollback(record) {
    pendingRollback = record;
    const idx = editLog.indexOf(record);
    const later = (idx >= 0 ? editLog.length - idx - 1 : 0) + (latestEdit && latestEdit.applied ? 1 : 0);
    els.confirmMsg.textContent = later > 0
      ? '将撤销该回复写入的内容及其后的 ' + later + ' 条修改，并删除该回复之后的对话记录。确定回退吗？'
      : '将撤销该回复写入的内容，并删除该回复之后的对话记录。确定回退吗？';
    els.confirmModal.classList.remove('hidden');
  }

  function closeConfirm() {
    els.confirmModal.classList.add('hidden');
    pendingRollback = null;
  }

  /* ---------- 对话逻辑 ---------- */

  async function streamChat(bubble) {
    lastAssistantText = '';
    lastReasoningText = '';
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
          reasoning_effort: reasoningEffort,
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
      lastReasoningText = (j.choices && j.choices[0] && j.choices[0].message.reasoning_content) || '';
      bubble.innerHTML = renderAssistantHtml(lastReasoningText, lastAssistantText);
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
            const d = j.choices && j.choices[0] ? j.choices[0].delta || {} : {};
            const delta = d.content || '';
            const rdelta = d.reasoning_content || '';
            if (rdelta || delta) {
              if (rdelta) lastReasoningText += rdelta;
              if (delta) lastAssistantText += delta;
              renderStreamingBubble(bubble, lastReasoningText, lastAssistantText);
              scrollToBottom();
            }
          } catch {
            /* 忽略无法解析的片段 */
          }
        }
      }
    }

    if (!lastAssistantText && !lastReasoningText) {
      throw new Error('未收到模型回复');
    }
    renderStreamingBubble(bubble, lastReasoningText, lastAssistantText);
    scrollToBottom();
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
      conversation.push({ role: 'assistant', content: lastAssistantText, reasoning: lastReasoningText });
      // 控制历史长度
      if (conversation.length > MAX_HISTORY * 2) {
        conversation = conversation.slice(conversation.length - MAX_HISTORY * 2);
      }
      saveCurrentConversation();
      await autoApplyReply(bubble);
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
    lastReasoningText = '';
    currentConvId = null;
    els.messages.innerHTML = '';
    if (latestEdit) deleteBookmark(latestEdit);
    for (const rec of editLog) deleteBookmark(rec);
    latestEdit = null;
    editLog = [];
    pendingRollback = null;
    renderWelcome();
    updateSelectionInfo();
  }

  /* ---------- 设置 ---------- */

  function openSettings() {
    els.apiKeyInput.value = apiKey;
    els.customModelInput.value = LEGACY_MODELS.includes(model) ? '' : model;
    els.effortSelect.value = reasoningEffort;
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
    reasoningEffort = els.effortSelect.value;
    localStorage.setItem('ds_api_key', apiKey);
    localStorage.setItem('ds_model', model);
    localStorage.setItem('ds_reasoning_effort', reasoningEffort);
    closeSettings();
    refreshStatus();
    toast('设置已保存：' + model + '（思考强度 ' + reasoningEffort + '）');
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
    els.effortSelect.addEventListener('change', () => {
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
    els.confirmOkBtn.addEventListener('click', () => {
      const rec = pendingRollback;
      closeConfirm();
      if (rec) rollbackTo(rec);
    });
    els.confirmCancelBtn.addEventListener('click', closeConfirm);
    els.confirmModal.addEventListener('click', (e) => {
      if (e.target === els.confirmModal) closeConfirm();
    });
    els.refreshCtxBtn.addEventListener('click', updateSelectionInfo);
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
