// ==UserScript==
// @name         XP420B Auto Print (1-day dedupe + no duplicates + cleanup badge)
// @namespace    xp420b-autoprint
// @match        https://turbo-pvz.ozon.ru/*
// @run-at       document-end
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_addValueChangeListener
// @connect      127.0.0.1
// @version      1.0.3
// @downloadURL  https://github.com/voronovmaksim57-dotcom/Print/raw/refs/heads/main/xp420b-autoprint.user.js
// @updateURL    https://github.com/voronovmaksim57-dotcom/Print/raw/refs/heads/main/xp420b-autoprint.user.js
// ==/UserScript==

(function () {
  'use strict';

  const SHELF_CLASS = '_shelfTag_1tkm1_2';
  const TIME_CLASS_HINT = '_time_';
  const CODE_RE = /^\d+-\d+$/;

  const STORAGE_KEY = 'xp420b_handledCodeTimes_v2';
  const LAST_PRINTED_KEY = 'xp420b_lastPrinted';

  // Храним не более суток
  const MAX_AGE_DAYS = 1;
  const MAX_AGE_MS = MAX_AGE_DAYS * 24 * 60 * 60 * 1000;

  // handled = [{ key: "код||время", ts: timestamp }]
  let handled = JSON.parse(GM_getValue(STORAGE_KEY, '[]') || '[]');
  let lastPrinted = GM_getValue(LAST_PRINTED_KEY, null);

  console.log(
      '[XP420B] Loaded handled entries =',
      handled.length,
      'lastPrinted =',
      lastPrinted
  );

  // ====== синхронизация между вкладками ======

  GM_addValueChangeListener(STORAGE_KEY, (name, oldValue, newValue, remote) => {
      if (!remote) return;
      handled = JSON.parse(newValue || '[]');
  });

  GM_addValueChangeListener(LAST_PRINTED_KEY, (name, oldValue, newValue, remote) => {
      if (!remote) return;
      lastPrinted = newValue;
  });

  // ===== UI: маленький индикатор очистки в углу =====

  let cleanupBadge = null;
  let cleanupBadgeTimer = null;

  function getCleanupBadge() {
    if (cleanupBadge) return cleanupBadge;

    const div = document.createElement('div');
    div.id = 'xp420b-cleanup-badge';
    div.style.position = 'fixed';
    div.style.bottom = '10px';
    div.style.right = '10px';
    div.style.zIndex = '999999';
    div.style.background = 'rgba(0,0,0,0.8)';
    div.style.color = '#fff';
    div.style.padding = '6px 10px';
    div.style.borderRadius = '6px';
    div.style.fontSize = '12px';
    div.style.fontFamily = 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
    div.style.boxShadow = '0 2px 6px rgba(0,0,0,0.4)';
    div.style.pointerEvents = 'none';
    div.style.opacity = '0';
    div.style.transition = 'opacity 0.2s ease-out';

    document.body.appendChild(div);
    cleanupBadge = div;
    return div;
  }

  function showCleanupBadge(removed, total) {
    if (removed <= 0) return;
    const badge = getCleanupBadge();
    badge.textContent = `XP420B: удалено ${removed} записей (осталось ${total})`;
    badge.style.opacity = '1';

    if (cleanupBadgeTimer) {
      clearTimeout(cleanupBadgeTimer);
    }

    cleanupBadgeTimer = setTimeout(() => {
      badge.style.opacity = '0';
    }, 2500);
  }

  // ====== служебные функции хранения ======

  function saveHandled() {
    GM_setValue(STORAGE_KEY, JSON.stringify(handled));
  }

  function saveLastPrinted(code) {
    lastPrinted = code;
    GM_setValue(LAST_PRINTED_KEY, code);
  }

  // Очистка старых записей
  function cleanupOld() {
    const now = Date.now();
    const before = handled.length;

    handled = handled.filter(entry => (now - entry.ts) < MAX_AGE_MS);

    const after = handled.length;
    const removed = before - after;

    if (removed > 0) {
      console.log(`[XP420B] Cleanup: removed ${removed} old entries, left ${after}`);
      saveHandled();
      showCleanupBadge(removed, after);
    }
  }

  cleanupOld(); // выполняем при старте скрипта

  function wasHandled(key) {
    return handled.some(e => e.key === key);
  }

  function markHandled(key) {
    // Чтобы сильно не раздувать массив, не дублируем одинаковые key
    if (!wasHandled(key)) {
      handled.push({ key, ts: Date.now() });
      saveHandled();
    }
  }

  // ====== печать ======

  function extractCode(el) {
    if (!el || !el.textContent) return null;
    const text = el.textContent.trim();
    return CODE_RE.test(text) ? text : null;
  }

  function looksLikeTime(text) {
  if (!text) return false;
  const t = text.trim();

  // 12:34, 9:05, 12:34:56
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(t)) return true;

  // 12 мин, 3 ч, 1 час назад и т.п. — если у тебя на странице бывают такие форматы
  if (/\d+\s*(мин|ч|час|часа|часов)/i.test(t)) return true;

  return false;
}

function findTimeTextForShelf(el) {
  let node = el;

  for (let i = 0; i < 12 && node; i += 1) {
    if (node.querySelectorAll) {
      const candidates = Array.from(node.querySelectorAll('*'));

      for (const c of candidates) {
        const text = c.textContent?.trim();
        const className = typeof c.className === 'string' ? c.className : '';

        // 1) частичное совпадение по классу
        if (className.includes(TIME_CLASS_HINT) && text) {
          return text;
        }

        // 2) просто похоже на время
        if (looksLikeTime(text)) {
          return text;
        }
      }
    }

    node = node.parentElement;
  }

  return null;
}
  function sendToPrinter(code) {
    console.log('[XP420B] Print request for', code);
    GM_xmlhttpRequest({
      method: 'POST',
      url: 'http://127.0.0.1:9123/print',
      headers: { 'Content-Type': 'application/json' },
      data: JSON.stringify({ label: code }),
      onload: res => {
        console.log('[XP420B] Print response', res.status, res.responseText);
      },
      onerror: err => console.error('[XP420B] Print error', err)
    });
  }

  function handleShelfElement(el, retry = 0) {
  const code = extractCode(el);
  if (!code) return;

  const timeText = findTimeTextForShelf(el);

  if (!timeText) {
    if (retry < 5) {
      setTimeout(() => handleShelfElement(el, retry + 1), 300);
    } else {
      console.log('[XP420B] No time found for', code);
    }
    return;
  }

  const key = code + '||' + timeText;

  if (wasHandled(key)) {
    console.log('[XP420B] Already handled', key);
    return;
  }

  if (lastPrinted === code) {
    console.log('[XP420B] Skip consecutive duplicate', code);
    markHandled(key);
    return;
  }

  sendToPrinter(code);
  markHandled(key);
  saveLastPrinted(code);

  console.log('[XP420B] Printed and set lastPrinted =', code);
}

  function processShelfTags(nodes) {
    nodes.forEach(handleShelfElement);
  }

  // ====== наблюдение за DOM ======

  function initialScan() {
    const tags = Array.from(document.querySelectorAll('.' + SHELF_CLASS));
    if (tags.length === 0) {
      console.log('[XP420B] No tags on initial scan');
      return;
    }
    console.log('[XP420B] Initial scan found', tags.length, 'tags');
    processShelfTags(tags);
  }

  function observe() {
    const observer = new MutationObserver(muts => {
      const found = [];
      for (const m of muts) {
        m.addedNodes.forEach(n => {
          if (!(n instanceof HTMLElement)) return;

          if (n.classList?.contains(SHELF_CLASS)) found.push(n);
          n.querySelectorAll?.('.' + SHELF_CLASS)
            .forEach(el => found.push(el));
        });

        if (m.type === 'characterData') {
          const el = m.target?.parentElement;
          if (el?.classList?.contains(SHELF_CLASS)) found.push(el);
        }
      }

      if (found.length) processShelfTags(found);
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: true
    });

    console.log('[XP420B] MutationObserver started');
  }

  initialScan();
  observe();
})();
