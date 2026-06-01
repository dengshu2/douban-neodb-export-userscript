// ==UserScript==
// @name         豆瓣书影音游导出 NeoDB Excel
// @namespace    https://github.com/dengshu2/douban-neodb-export-userscript
// @version      0.1.2
// @description  将当前登录豆瓣账号的书、影、音、游和舞台剧记录导出为 NeoDB 兼容 Excel。
// @author       dengshu2
// @license      MIT
// @homepageURL  https://github.com/dengshu2/douban-neodb-export-userscript
// @supportURL   https://github.com/dengshu2/douban-neodb-export-userscript/issues
// @downloadURL  https://raw.githubusercontent.com/dengshu2/douban-neodb-export-userscript/main/douban-neodb-export.user.js
// @updateURL    https://raw.githubusercontent.com/dengshu2/douban-neodb-export-userscript/main/douban-neodb-export.user.js
// @match        *://douban.com/*
// @match        *://*.douban.com/*
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @grant        GM_cookie
// @grant        GM_registerMenuCommand
// @grant        GM_xmlhttpRequest
// @connect      douban.com
// @run-at       document-idle
// @noframes
// ==/UserScript==

(function () {
  'use strict';

  const PAGE_SIZE = 50;
  const REQUEST_INTERVAL = 1000;
  const WIDGET_ID = 'douban-interest-export-widget';
  const API_BASE = 'https://m.douban.com/rexxar/api/v2';
  const INTEREST_TYPES = [
    {
      key: 'movie',
      name: '电影',
      statuses: [
        ['done', '看过'],
        ['doing', '在看'],
        ['mark', '想看'],
      ],
    },
    {
      key: 'music',
      name: '音乐',
      statuses: [
        ['done', '听过'],
        ['doing', '在听'],
        ['mark', '想听'],
      ],
    },
    {
      key: 'book',
      name: '图书',
      statuses: [
        ['done', '读过'],
        ['doing', '在读'],
        ['mark', '想读'],
      ],
    },
    {
      key: 'game',
      name: '游戏',
      statuses: [
        ['done', '玩过'],
        ['doing', '在玩'],
        ['mark', '想玩'],
      ],
    },
    {
      key: 'drama',
      name: '舞台剧',
      statuses: [
        ['done', '看过的舞台剧'],
        ['mark', '想看的舞台剧'],
      ],
    },
  ];

  const state = {
    cancelled: false,
    currentRequest: null,
    lastRequestAt: 0,
    running: false,
    widget: null,
  };

  function sleep(milliseconds) {
    return new Promise(resolve => setTimeout(resolve, milliseconds));
  }

  function assertNotCancelled() {
    if (state.cancelled) {
      throw new Error('导出已取消。');
    }
  }

  function escapeSpreadsheetText(value) {
    if (value === null || value === undefined) {
      return '';
    }
    const text = Array.isArray(value) ? value.join(', ') : String(value);
    return /^[=+\-@]/.test(text) ? `'${text}` : text;
  }

  function parseCookieString(cookieString) {
    return cookieString.split(';').reduce((cookies, item) => {
      const separatorIndex = item.indexOf('=');
      if (separatorIndex === -1) {
        return cookies;
      }
      const name = item.slice(0, separatorIndex).trim();
      const value = item.slice(separatorIndex + 1).trim();
      cookies[name] = decodeURIComponent(value);
      return cookies;
    }, {});
  }

  async function getCookie(name) {
    const pageCookie = parseCookieString(document.cookie)[name];
    if (pageCookie) {
      return pageCookie;
    }

    if (typeof GM_cookie === 'undefined' || typeof GM_cookie.list !== 'function') {
      return '';
    }

    return await new Promise(resolve => {
      GM_cookie.list({ url: 'https://m.douban.com/', name }, (cookies, error) => {
        if (error || !cookies || cookies.length === 0) {
          resolve('');
          return;
        }
        resolve(cookies[0].value || '');
      });
    });
  }

  function request(details) {
    return new Promise((resolve, reject) => {
      assertNotCancelled();
      state.currentRequest = GM_xmlhttpRequest({
        method: 'GET',
        timeout: 30000,
        ...details,
        onabort() {
          state.currentRequest = null;
          reject(new Error('导出已取消。'));
        },
        onerror() {
          state.currentRequest = null;
          reject(new Error('网络请求失败，请稍后重试。'));
        },
        ontimeout() {
          state.currentRequest = null;
          reject(new Error('网络请求超时，请稍后重试。'));
        },
        onload(response) {
          state.currentRequest = null;
          if (response.status < 200 || response.status >= 300) {
            reject(new Error(`豆瓣接口返回 HTTP ${response.status}。`));
            return;
          }
          resolve(response);
        },
      });
    });
  }

  async function politeRequest(details) {
    assertNotCancelled();
    const remaining = state.lastRequestAt + REQUEST_INTERVAL - Date.now();
    if (remaining > 0) {
      await sleep(remaining);
    }
    assertNotCancelled();
    state.lastRequestAt = Date.now();
    return await request(details);
  }

  function parseJsonResponse(response) {
    if (response.response && typeof response.response === 'object') {
      return response.response;
    }
    return JSON.parse(response.responseText);
  }

  async function getCurrentUser() {
    const response = await politeRequest({
      url: 'https://m.douban.com/mine/',
      headers: { Referer: 'https://m.douban.com/' },
    });
    const page = new DOMParser().parseFromString(response.responseText, 'text/html');
    const userInput = page.querySelector('#user');
    const homepageLink = page.querySelector('.profile .detail .basic-info > a');

    if (!userInput || !homepageLink) {
      throw new Error('没有识别到豆瓣登录账号。请先登录豆瓣，再重新导出。');
    }

    const homepageMatch = homepageLink.getAttribute('href').match(/\/people\/([^/]+)/);
    return {
      id: userInput.getAttribute('value'),
      name: userInput.getAttribute('data-name') || '豆瓣用户',
      symbol: homepageMatch ? homepageMatch[1] : '',
    };
  }

  function createInterestUrl(userId, ck, type, status, start) {
    const params = new URLSearchParams({
      type,
      status,
      start: String(start),
      count: String(PAGE_SIZE),
      ck,
      for_mobile: '1',
    });
    return `${API_BASE}/user/${encodeURIComponent(userId)}/interests?${params}`;
  }

  async function fetchInterestGroup(userId, ck, type, status, onProgress) {
    const records = [];
    let start = 0;
    let total = null;

    do {
      const response = await politeRequest({
        url: createInterestUrl(userId, ck, type, status, start),
        headers: {
          Accept: 'application/json',
          Referer: `https://m.douban.com/mine/${type}`,
        },
        responseType: 'json',
      });
      const payload = parseJsonResponse(response);
      const interests = Array.isArray(payload.interests) ? payload.interests : [];
      total = Number(payload.total) || 0;
      records.push(...interests);
      start += interests.length;
      onProgress(records.length, total);

      if (interests.length === 0) {
        break;
      }
    } while (start < total);

    return records;
  }

  function interestToRow(interest) {
    const subject = interest.subject || {};
    const subjectRating = subject.rating && subject.rating.value !== undefined
      ? subject.rating.value
      : subject.null_rating_reason || '';
    return [
      escapeSpreadsheetText(subject.title),
      escapeSpreadsheetText(subject.card_subtitle),
      subjectRating,
      escapeSpreadsheetText(subject.url),
      escapeSpreadsheetText(interest.create_time),
      interest.rating && interest.rating.value !== undefined ? interest.rating.value : '',
      escapeSpreadsheetText(interest.tags),
      escapeSpreadsheetText(interest.comment),
      interest.is_private ? 'private' : 'public',
    ];
  }

  function appendSheet(workbook, name, rows) {
    const worksheet = XLSX.utils.aoa_to_sheet(rows, { dense: true });
    worksheet['!cols'] = [
      { wch: 36 },
      { wch: 52 },
      { wch: 12 },
      { wch: 48 },
      { wch: 20 },
      { wch: 12 },
      { wch: 24 },
      { wch: 60 },
      { wch: 12 },
    ];
    XLSX.utils.book_append_sheet(workbook, worksheet, name);
  }

  function buildWorkbook(user, groups) {
    const workbook = XLSX.utils.book_new();
    const createdAt = new Date().toLocaleString();
    const total = groups.reduce((sum, group) => sum + group.records.length, 0);
    const summaryRows = [
      ['豆瓣书影音游导出'],
      ['账号', escapeSpreadsheetText(user.name)],
      ['豆瓣 ID', escapeSpreadsheetText(user.symbol)],
      ['导出时间', createdAt],
      ['总记录数', total],
      [],
      ['分类', '记录数'],
      ...groups.map(group => [group.sheetName, group.records.length]),
    ];
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(summaryRows, { dense: true }),
      '汇总',
    );

    for (const group of groups) {
      const rows = [
        ['标题', '简介', '豆瓣评分', '链接', '创建时间', '我的评分', '标签', '评论', '可见性'],
        ...group.records.map(interestToRow),
      ];
      appendSheet(workbook, group.sheetName, rows);
    }

    return workbook;
  }

  function getWidget() {
    if (state.widget) {
      return state.widget;
    }

    const widget = document.createElement('section');
    widget.id = WIDGET_ID;
    widget.innerHTML = `
      <button class="douban-export-start" type="button">导出豆瓣记录</button>
      <div class="douban-export-progress" hidden>
        <div class="douban-export-message">准备导出...</div>
        <progress value="0" max="100"></progress>
        <button class="douban-export-cancel" type="button">取消</button>
      </div>
    `;

    const style = document.createElement('style');
    style.textContent = `
      #${WIDGET_ID} {
        position: fixed;
        right: 20px;
        bottom: 20px;
        z-index: 2147483647;
        width: 260px;
        padding: 12px;
        border: 1px solid #d8e5dc;
        border-radius: 8px;
        background: #fff;
        box-shadow: 0 6px 24px rgba(0, 0, 0, .15);
        color: #555;
        font: 14px/1.5 Arial, sans-serif;
      }
      #${WIDGET_ID} button {
        cursor: pointer;
        border: 1px solid #007722;
        border-radius: 4px;
        padding: 7px 12px;
        background: #fff;
        color: #007722;
      }
      #${WIDGET_ID} .douban-export-start {
        width: 100%;
        background: #007722;
        color: #fff;
      }
      #${WIDGET_ID} .douban-export-progress {
        margin-top: 10px;
      }
      #${WIDGET_ID} progress {
        display: block;
        width: 100%;
        margin: 8px 0;
      }
    `;

    (document.head || document.documentElement).appendChild(style);
    document.body.appendChild(widget);
    widget.querySelector('.douban-export-start').addEventListener('click', runExport);
    widget.querySelector('.douban-export-cancel').addEventListener('click', cancelExport);
    state.widget = widget;
    return widget;
  }

  function updateWidget(message, percentage) {
    const widget = getWidget();
    widget.querySelector('.douban-export-message').textContent = message;
    widget.querySelector('progress').value = Math.max(0, Math.min(100, percentage));
  }

  function mountWidget() {
    if (!document.body || document.getElementById(WIDGET_ID)) {
      return;
    }
    getWidget();
  }

  function setRunning(running) {
    const widget = getWidget();
    widget.querySelector('.douban-export-start').disabled = running;
    widget.querySelector('.douban-export-progress').hidden = !running;
  }

  function cancelExport() {
    state.cancelled = true;
    if (state.currentRequest && typeof state.currentRequest.abort === 'function') {
      state.currentRequest.abort();
    }
  }

  async function runExport() {
    if (state.running) {
      return;
    }

    state.running = true;
    state.cancelled = false;
    state.lastRequestAt = 0;
    setRunning(true);

    try {
      updateWidget('正在识别登录账号...', 0);
      const ck = await getCookie('ck');
      if (!ck) {
        throw new Error('没有读取到登录凭据 ck。请先登录豆瓣，再重新导出。');
      }
      const user = await getCurrentUser();
      const groups = [];
      const combinations = INTEREST_TYPES.flatMap(type => (
        type.statuses.map(([status, sheetName]) => ({
          sheetName,
          status,
          type: type.key,
        }))
      ));

      for (let index = 0; index < combinations.length; index += 1) {
        assertNotCancelled();
        const combination = combinations[index];
        const basePercentage = Math.floor(index / combinations.length * 90);
        updateWidget(`正在读取 ${combination.sheetName}...`, basePercentage);
        const records = await fetchInterestGroup(
          user.id,
          ck,
          combination.type,
          combination.status,
          (loaded, total) => {
            updateWidget(
              `正在读取 ${combination.sheetName}：${loaded}/${total}`,
              basePercentage,
            );
          },
        );
        groups.push({ ...combination, records });
      }

      assertNotCancelled();
      updateWidget('正在生成 Excel...', 95);
      const workbook = buildWorkbook(user, groups);
      XLSX.writeFile(workbook, `豆瓣书影音游(${user.symbol || user.id}).xlsx`, {
        compression: true,
      });
      updateWidget('导出完成。', 100);
      alert('豆瓣记录已导出为 Excel。');
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      updateWidget(message, 0);
      alert(message);
    } finally {
      state.running = false;
      state.currentRequest = null;
      setRunning(false);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', mountWidget, { once: true });
  } else {
    mountWidget();
  }
  GM_registerMenuCommand('导出豆瓣书影音游记录', runExport);
})();
