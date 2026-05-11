(function () {
  'use strict';

  const API = '/api';
  const token = localStorage.getItem('kb_token') || '';

  // DOM refs
  const backBtn = document.getElementById('backBtn');
  const clearBtn = document.getElementById('clearBtn');
  const askBtn = document.getElementById('askBtn');
  const textarea = document.getElementById('questionInput');
  const messagesArea = document.getElementById('messagesArea');
  const welcome = document.getElementById('welcome');
  const modelChip = document.getElementById('modelChip');
  const fileBindSelect = document.getElementById('fileBindSelect');

  // Conversation history [{role, content}]
  let history = [];
  let isStreaming = false;
  let selectedFileId = '';

  // ──────────────── Markdown / escape ────────────────

  function renderMarkdown(text) {
    if (typeof marked !== 'undefined') {
      try {
        marked.setOptions({ breaks: true, gfm: true });
        return marked.parse(text || '');
      } catch (e) {}
    }
    return escapeHtml(text || '').replace(/\n/g, '<br>');
  }

  function escapeHtml(str) {
    return (str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ──────────────── Model info ────────────────

  async function loadModelInfo() {
    if (!token) return;
    try {
      const res = await fetch(API + '/qa/model-info', {
        headers: { Authorization: 'Bearer ' + token }
      });
      if (!res.ok) return;
      const data = await res.json();
      if (data.has_llm && data.model) {
        modelChip.textContent = '⚡ ' + (data.name || data.model);
        modelChip.classList.remove('no-model');
      } else {
        modelChip.textContent = '未配置 AI 模型';
        modelChip.classList.add('no-model');
      }
    } catch (e) {
      modelChip.textContent = '无法获取模型';
      modelChip.classList.add('no-model');
    }
  }

  async function loadFileOptions() {
    if (!token || !fileBindSelect) return;
    try {
      const res = await fetch(API + '/files', {
        headers: { Authorization: 'Bearer ' + token }
      });
      if (!res.ok) return;
      const files = await res.json();
      const current = fileBindSelect.value;
      fileBindSelect.innerHTML = '<option value="">全部文件（自动检索）</option>';
      (files || []).forEach(f => {
        const opt = document.createElement('option');
        opt.value = f.id || '';
        opt.textContent = f.filename || f.id || '未命名文件';
        fileBindSelect.appendChild(opt);
      });
      if (current) fileBindSelect.value = current;
      selectedFileId = fileBindSelect.value || '';
    } catch (e) {}
  }

  // ──────────────── UI helpers ────────────────

  function scrollToBottom(smooth) {
    messagesArea.scrollTo({
      top: messagesArea.scrollHeight,
      behavior: smooth ? 'smooth' : 'auto'
    });
  }

  function showWelcome(show) {
    welcome.style.display = show ? '' : 'none';
    clearBtn.classList.toggle('hidden', show);
  }

  function getFileIcon(type) {
    const map = { '.pdf': '📄', '.docx': '📝', '.doc': '📝', '.xlsx': '📊', '.xls': '📊',
      '.csv': '📊', '.txt': '📃', '.md': '📋', '.png': '🖼', '.jpg': '🖼', '.jpeg': '🖼' };
    return map[(type || '').toLowerCase()] || '📄';
  }

  // ──────────────── Message rendering ────────────────

  function appendUserMessage(text) {
    showWelcome(false);
    const div = document.createElement('div');
    div.className = 'msg msg-user';
    div.innerHTML = `<div class="msg-bubble">${escapeHtml(text)}</div>`;
    messagesArea.appendChild(div);
    scrollToBottom(true);
  }

  function appendAiPlaceholder() {
    const div = document.createElement('div');
    div.className = 'msg msg-ai';
    div.innerHTML = `
      <div class="msg-avatar">🤖</div>
      <div class="msg-body">
        <div class="msg-bubble-ai">
          <div class="searching-text">
            <span class="searching-spinner"></span>正在检索知识库…
          </div>
        </div>
      </div>`;
    messagesArea.appendChild(div);
    scrollToBottom(true);
    return div;
  }

  function updateAiMessageSources(msgEl, sources) {
    const bubble = msgEl.querySelector('.msg-bubble-ai');
    // Replace searching indicator with streaming cursor
    bubble.innerHTML = `<div class="msg-content streaming-cursor"></div>`;
    if (sources && sources.length) {
      const sourcesDiv = document.createElement('div');
      sourcesDiv.className = 'msg-sources';
      sourcesDiv.innerHTML = `<div class="sources-label">📚 参考来源</div>` +
        sources.map(s => `
          <div class="source-item">
            <span class="source-icon">${getFileIcon(s.type)}</span>
            <div class="source-info">
              <div class="source-name">${escapeHtml(s.filename || '未命名文件')}</div>
              <div class="source-snip">${escapeHtml(s.snippet || '')}</div>
            </div>
          </div>`).join('');
      msgEl.querySelector('.msg-body').appendChild(sourcesDiv);
    }
    scrollToBottom(false);
  }

  function appendToken(msgEl, rawText) {
    const contentEl = msgEl.querySelector('.msg-content');
    if (contentEl) {
      // Show plain text while streaming
      contentEl.textContent = rawText;
      scrollToBottom(false);
    }
  }

  function finalizeAiMessage(msgEl, fullText, sources) {
    const bubble = msgEl.querySelector('.msg-bubble-ai');
    const contentEl = bubble.querySelector('.msg-content');
    if (contentEl) {
      contentEl.classList.remove('streaming-cursor');
      contentEl.innerHTML = renderMarkdown(fullText);
    }
    scrollToBottom(true);
  }

  function showAiError(msgEl, message) {
    const bubble = msgEl.querySelector('.msg-bubble-ai');
    bubble.innerHTML = `<div class="msg-error">⚠️ ${escapeHtml(message)}</div>`;
  }

  // ──────────────── Core ask logic (SSE) ────────────────

  async function ask() {
    const question = (textarea.value || '').trim();
    if (!question || isStreaming) return;

    if (!token) {
      alert('登录已失效，请返回主页重新登录');
      return;
    }

    isStreaming = true;
    askBtn.disabled = true;
    textarea.value = '';
    textarea.style.height = '';

    appendUserMessage(question);
    const aiMsgEl = appendAiPlaceholder();

    // Build payload with history
    const payload = {
      question,
      messages: history.map(m => ({ role: m.role, content: m.content })),
      selected_file_id: selectedFileId || null
    };

    let accumulated = '';
    let sources = [];
    let gotSources = false;
    let finalized = false;

    try {
      const res = await fetch(API + '/qa/ask/stream', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: 'Bearer ' + token
        },
        body: JSON.stringify(payload)
      });

      if (res.status === 401) {
        showAiError(aiMsgEl, '登录已失效，请返回主页重新登录');
        return;
      }

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        showAiError(aiMsgEl, err.detail || '请求失败');
        return;
      }

      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop(); // keep incomplete line

        for (const line of lines) {
          const normalized = line.replace(/\r$/, '');
          if (!normalized.startsWith('data: ')) continue;
          const dataStr = normalized.slice(6).trim();
          if (!dataStr || dataStr === '[DONE]') continue;

          let evt;
          try { evt = JSON.parse(dataStr); } catch (e) { continue; }

          if (evt.type === 'sources') {
            sources = evt.sources || [];
            gotSources = true;
            updateAiMessageSources(aiMsgEl, sources);
          } else if (evt.type === 'token') {
            if (!gotSources) {
              updateAiMessageSources(aiMsgEl, []);
              gotSources = true;
            }
            accumulated += evt.content;
            // 知识库直接答案（source=kb）一次性到达，直接渲染 Markdown；LLM 流式用纯文本追加
            if (evt.source === 'kb') {
              finalizeAiMessage(aiMsgEl, accumulated, sources);
              finalized = true;
              history.push({ role: 'user', content: question });
              history.push({ role: 'assistant', content: accumulated });
            } else {
              appendToken(aiMsgEl, accumulated);
            }
          } else if (evt.type === 'done') {
            if (!finalized) {
              finalized = true;
              finalizeAiMessage(aiMsgEl, accumulated, sources);
              history.push({ role: 'user', content: question });
              history.push({ role: 'assistant', content: accumulated });
            }
          } else if (evt.type === 'error') {
            showAiError(aiMsgEl, evt.message || '未知错误');
          }
        }
      }

      // If stream ended without done event
      if (!finalized) {
        if (!accumulated && !aiMsgEl.querySelector('.msg-error')) {
          showAiError(aiMsgEl, '未收到有效回复，请重试');
        } else if (accumulated) {
          finalizeAiMessage(aiMsgEl, accumulated, sources);
          history.push({ role: 'user', content: question });
          history.push({ role: 'assistant', content: accumulated });
        }
      }

    } catch (e) {
      if (e.name === 'AbortError') return;
      showAiError(aiMsgEl, '请求失败：' + (e.message || '网络错误'));
    } finally {
      isStreaming = false;
      askBtn.disabled = false;
      textarea.focus();
    }
  }

  // ──────────────── Clear conversation ────────────────

  function clearConversation() {
    history = [];
    // Remove all message elements
    const msgs = messagesArea.querySelectorAll('.msg');
    msgs.forEach(m => m.remove());
    showWelcome(true);
    textarea.focus();
  }

  // ──────────────── Auto-resize textarea ────────────────

  function autoResize() {
    textarea.style.height = 'auto';
    textarea.style.height = Math.min(textarea.scrollHeight, 160) + 'px';
  }

  // ──────────────── Tip clicks ────────────────

  document.querySelectorAll('.tip').forEach(tip => {
    tip.addEventListener('click', () => {
      const q = tip.getAttribute('data-q');
      if (q && !isStreaming) {
        textarea.value = q;
        autoResize();
        ask();
      }
    });
  });

  // ──────────────── Events ────────────────

  askBtn.addEventListener('click', ask);

  textarea.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      ask();
    }
  });

  textarea.addEventListener('input', autoResize);

  clearBtn.addEventListener('click', clearConversation);
  if (fileBindSelect) {
    fileBindSelect.addEventListener('change', () => {
      selectedFileId = fileBindSelect.value || '';
    });
  }

  backBtn.addEventListener('click', () => {
    const origin = window.location.protocol === 'file:'
      ? 'http://localhost:8000'
      : window.location.origin;
    window.location.href = origin + '/';
  });

  // ──────────────── Init ────────────────

  loadModelInfo();
  loadFileOptions();

  if (!token) {
    modelChip.textContent = '未登录';
    modelChip.classList.add('no-model');
  }

})();
