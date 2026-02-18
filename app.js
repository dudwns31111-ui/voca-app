(() => {
  'use strict';

  const DB_NAME = 'vocabDB';
  const DB_VERSION = 2;
  const WORDS_STORE = 'words';
  const META_STORE = 'appMeta';
  const BACKUP_FILE_NAME = 'vocab_backup.json';
  const AUTO_BACKUP_DEBOUNCE_MS = 10_000;
  const PAGE_SIZE = 100;

  const state = {
    db: null,
    words: [],
    filteredWords: [],
    searchTerm: '',
    currentPage: 1,
    reviewPool: [],
    reviewIndex: 0,
    backupTimer: null,
    backupInFlight: false,
    pendingBackupReason: null,
    backupHandle: null
  };

  const els = {
    mainScreen: document.getElementById('mainScreen'),
    listScreen: document.getElementById('listScreen'),
    reviewScreen: document.getElementById('reviewScreen'),
    searchInput: document.getElementById('searchInput'),
    wordInput: document.getElementById('wordInput'),
    meaningInput: document.getElementById('meaningInput'),
    exampleInput: document.getElementById('exampleInput'),
    saveBtn: document.getElementById('saveBtn'),
    reviewModeBtn: document.getElementById('reviewModeBtn'),
    wordListBtn: document.getElementById('wordListBtn'),
    backToMainBtn: document.getElementById('backToMainBtn'),
    exportBtn: document.getElementById('exportBtn'),
    importBtn: document.getElementById('importBtn'),
    importFile: document.getElementById('importFile'),
    backupLinkBtn: document.getElementById('backupLinkBtn'),
    status: document.getElementById('status'),
    totalCount: document.getElementById('totalCount'),
    wordList: document.getElementById('wordList'),
    paginationControls: document.getElementById('paginationControls'),
    prevPageBtn: document.getElementById('prevPageBtn'),
    nextPageBtn: document.getElementById('nextPageBtn'),
    pageIndicator: document.getElementById('pageIndicator'),
    reviewSection: document.getElementById('reviewSection'),
    reviewEmpty: document.getElementById('reviewEmpty'),
    reviewContent: document.getElementById('reviewContent'),
    reviewWord: document.getElementById('reviewWord'),
    reviewMeaning: document.getElementById('reviewMeaning'),
    reviewExample: document.getElementById('reviewExample'),
    showMeaningBtn: document.getElementById('showMeaningBtn'),
    knownBtn: document.getElementById('knownBtn'),
    unknownBtn: document.getElementById('unknownBtn'),
    exitReviewBtn: document.getElementById('exitReviewBtn')
  };

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  function requestToPromise(req) {
    return new Promise((resolve, reject) => {
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  function txDone(tx) {
    return new Promise((resolve, reject) => {
      tx.oncomplete = resolve;
      tx.onabort = () => reject(tx.error || new Error('Transaction aborted'));
      tx.onerror = () => reject(tx.error);
    });
  }

  async function openDB() {
    if (state.db) return state.db;

    state.db = await new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(WORDS_STORE)) {
          const words = db.createObjectStore(WORDS_STORE, { keyPath: 'id', autoIncrement: true });
          words.createIndex('createdAt', 'createdAt', { unique: false });
          words.createIndex('word', 'word', { unique: false });
        }
        if (!db.objectStoreNames.contains(META_STORE)) {
          db.createObjectStore(META_STORE, { keyPath: 'key' });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });

    return state.db;
  }

  async function getAllWords() {
    const db = await openDB();
    const tx = db.transaction(WORDS_STORE, 'readonly');
    const all = await requestToPromise(tx.objectStore(WORDS_STORE).getAll());
    await txDone(tx);
    return all;
  }

  async function getWords(offset, limit) {
    const db = await openDB();
    const tx = db.transaction(WORDS_STORE, 'readonly');
    const index = tx.objectStore(WORDS_STORE).index('createdAt');
    const direction = 'prev';
    const rows = [];
    let skipped = 0;

    await new Promise((resolve, reject) => {
      const req = index.openCursor(null, direction);
      req.onerror = () => reject(req.error);
      req.onsuccess = () => {
        const cursor = req.result;
        if (!cursor || rows.length >= limit) {
          resolve();
          return;
        }

        if (skipped < offset) {
          const jump = Math.min(offset - skipped, 500);
          skipped += jump;
          cursor.advance(jump);
          return;
        }

        rows.push(cursor.value);
        cursor.continue();
      };
    });

    await txDone(tx);
    return rows;
  }

  async function addWord(row) {
    const db = await openDB();
    const tx = db.transaction(WORDS_STORE, 'readwrite');
    tx.objectStore(WORDS_STORE).add(row);
    await txDone(tx);
  }

  async function putWord(row) {
    const db = await openDB();
    const tx = db.transaction(WORDS_STORE, 'readwrite');
    tx.objectStore(WORDS_STORE).put(row);
    await txDone(tx);
  }

  async function deleteWordById(id) {
    const db = await openDB();
    const tx = db.transaction(WORDS_STORE, 'readwrite');
    tx.objectStore(WORDS_STORE).delete(id);
    await txDone(tx);
  }

  async function setMeta(key, value) {
    const db = await openDB();
    const tx = db.transaction(META_STORE, 'readwrite');
    tx.objectStore(META_STORE).put({ key, value });
    await txDone(tx);
  }

  async function getMeta(key) {
    const db = await openDB();
    const tx = db.transaction(META_STORE, 'readonly');
    const row = await requestToPromise(tx.objectStore(META_STORE).get(key));
    await txDone(tx);
    return row?.value ?? null;
  }

  function norm(v) {
    return String(v || '').trim().toLowerCase();
  }

  function formatDate(ts) {
    return ts ? new Date(ts).toLocaleString() : '-';
  }

  function setStatus(message, holdMs = 1700) {
    els.status.textContent = message;
    if (!holdMs) return;
    clearTimeout(setStatus.timer);
    setStatus.timer = setTimeout(() => {
      if (els.status.textContent === message) els.status.textContent = '';
    }, holdMs);
  }

  function hideAllScreens() {
    els.mainScreen.classList.add('hidden');
    els.listScreen.classList.add('hidden');
    els.reviewScreen.classList.add('hidden');
  }

  function showMainScreen() {
    hideAllScreens();
    els.mainScreen.classList.remove('hidden');
  }

  async function showListScreen() {
    hideAllScreens();
    els.listScreen.classList.remove('hidden');
    await renderWordListPage();
  }

  function showReviewScreen() {
    hideAllScreens();
    els.reviewScreen.classList.remove('hidden');
  }

  function applySearch() {
    const t = norm(state.searchTerm);
    state.filteredWords = t
      ? state.words.filter((row) => norm(row.word).includes(t) || norm(row.meaning).includes(t))
      : [];
    state.currentPage = 1;
  }

  function getPageCount(totalItems) {
    return Math.max(1, Math.ceil(totalItems / PAGE_SIZE));
  }

  function renderTotalCount() {
    const count = state.words.length;
    els.totalCount.textContent = `${count.toLocaleString()} ${count === 1 ? 'word' : 'words'}`;
  }

  function makeWordRow(row) {
    const wrap = document.createElement('article');
    wrap.className = 'word-item';

    const main = document.createElement('div');
    main.className = 'word-main';

    const word = document.createElement('p');
    word.className = 'word-label';
    word.textContent = row.word;

    const meaning = document.createElement('p');
    meaning.className = 'meaning hidden';
    meaning.textContent = row.meaning;

    const ex = document.createElement('p');
    ex.className = 'example';
    ex.textContent = row.example || '';

    const meta = document.createElement('p');
    meta.className = 'meta';
    meta.textContent = `Added: ${formatDate(row.createdAt)} • Reviews: ${row.reviewCount || 0}`;

    word.addEventListener('click', () => meaning.classList.toggle('hidden'));

    main.append(word, meaning);
    if (row.example) main.appendChild(ex);
    main.appendChild(meta);

    const del = document.createElement('button');
    del.className = 'delete-btn';
    del.textContent = 'Delete Word';
    del.addEventListener('click', async () => {
      if (!confirm(`Delete "${row.word}"?`)) return;
      await deleteWordById(row.id);
      await reloadWords();

      const totalItems = state.searchTerm ? state.filteredWords.length : state.words.length;
      const pageCount = getPageCount(totalItems);
      if (state.currentPage > pageCount) {
        state.currentPage = pageCount;
      }

      await renderWordListPage();
      scheduleAutoBackup('delete');
      setStatus('Word deleted.');
    });

    wrap.append(main, del);
    return wrap;
  }

  async function renderWordListPage() {
    let list = [];
    let totalItems = state.words.length;

    if (state.searchTerm) {
      list = state.filteredWords;
      totalItems = list.length;
      const offset = (state.currentPage - 1) * PAGE_SIZE;
      list = list.slice(offset, offset + PAGE_SIZE);
    } else {
      const offset = (state.currentPage - 1) * PAGE_SIZE;
      list = await getWords(offset, PAGE_SIZE);
    }

    const pageCount = getPageCount(totalItems);
    if (state.currentPage > pageCount) {
      state.currentPage = pageCount;
      return renderWordListPage();
    }

    if (!totalItems) {
      const empty = document.createElement('p');
      empty.className = 'muted';
      empty.textContent = state.searchTerm ? 'No matching words.' : 'No words saved yet.';
      els.wordList.replaceChildren(empty);
      els.paginationControls.classList.add('hidden');
      return;
    }

    const frag = document.createDocumentFragment();
    for (const row of list) {
      frag.appendChild(makeWordRow(row));
    }

    els.wordList.replaceChildren(frag);

    els.paginationControls.classList.remove('hidden');
    els.pageIndicator.textContent = `Page ${state.currentPage.toLocaleString()} of ${pageCount.toLocaleString()}`;
    els.prevPageBtn.disabled = state.currentPage <= 1;
    els.nextPageBtn.disabled = state.currentPage >= pageCount;
  }

  async function reloadWords() {
    const rows = await getAllWords();
    rows.sort((a, b) => b.createdAt - a.createdAt);
    state.words = rows;
    renderTotalCount();
    applySearch();
  }

  function asBackupArray(rows) {
    return rows.map((r) => ({
      id: r.id,
      word: r.word,
      meaning: r.meaning,
      example: r.example || '',
      createdAt: r.createdAt,
      reviewCount: Number(r.reviewCount) || 0,
      lastReviewedAt: r.lastReviewedAt || null
    }));
  }

  async function backupToLinkedFile(reason = 'auto') {
    if (!state.backupHandle) return false;

    if (state.backupInFlight) {
      state.pendingBackupReason = reason;
      return true;
    }

    state.backupInFlight = true;
    try {
      const data = JSON.stringify(asBackupArray(state.words));
      const writable = await state.backupHandle.createWritable();
      await writable.write(data);
      await writable.close();
      setStatus(`Backup synced (${reason})`, 1200);
      return true;
    } catch (err) {
      console.error('Backup write failed', err);
      setStatus('Auto backup failed. Relink backup file.', 2200);
      return false;
    } finally {
      state.backupInFlight = false;
      if (state.pendingBackupReason) {
        const next = state.pendingBackupReason;
        state.pendingBackupReason = null;
        queueMicrotask(() => backupToLinkedFile(next));
      }
    }
  }

  async function ensureBackupPermission() {
    if (!state.backupHandle) return false;
    if (!state.backupHandle.queryPermission) return true;
    const current = await state.backupHandle.queryPermission({ mode: 'readwrite' });
    if (current === 'granted') return true;
    const asked = await state.backupHandle.requestPermission({ mode: 'readwrite' });
    return asked === 'granted';
  }

  function scheduleAutoBackup(reason = 'change') {
    clearTimeout(state.backupTimer);
    state.backupTimer = setTimeout(async () => {
      const ok = await ensureBackupPermission();
      if (ok) await backupToLinkedFile(reason);
    }, AUTO_BACKUP_DEBOUNCE_MS);
  }

  async function doImmediateExitBackup(reason) {
    clearTimeout(state.backupTimer);
    const ok = await ensureBackupPermission();
    if (ok) await backupToLinkedFile(reason);
  }

  async function linkBackupFile() {
    if (!window.showSaveFilePicker) {
      setStatus('Browser does not support linked overwrite backup. Use Export JSON manually.', 2600);
      return;
    }

    const handle = await window.showSaveFilePicker({
      suggestedName: BACKUP_FILE_NAME,
      types: [{
        description: 'JSON backup',
        accept: { 'application/json': ['.json'] }
      }]
    });

    state.backupHandle = handle;
    await setMeta('backupFileHandle', handle);
    await doImmediateExitBackup('linked');
  }

  function manualExport() {
    const data = JSON.stringify(asBackupArray(state.words), null, 2);
    const blob = new Blob([data], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = BACKUP_FILE_NAME;
    a.click();
    URL.revokeObjectURL(url);
    setStatus('Exported vocab_backup.json');
  }

  async function importBackup(file) {
    const raw = await file.text();
    let rows;
    try {
      rows = JSON.parse(raw);
    } catch {
      setStatus('Invalid JSON file.', 2200);
      return;
    }

    if (!Array.isArray(rows)) {
      setStatus('Backup format must be a JSON array.', 2200);
      return;
    }

    const existing = new Set(state.words.map((r) => `${norm(r.word)}|${norm(r.meaning)}`));
    let inserted = 0;

    for (const row of rows) {
      const word = String(row?.word || '').trim();
      const meaning = String(row?.meaning || '').trim();
      if (!word || !meaning) continue;

      const key = `${norm(word)}|${norm(meaning)}`;
      if (existing.has(key)) continue;

      await addWord({
        word,
        meaning,
        example: String(row?.example || '').trim(),
        createdAt: Number(row?.createdAt) || Date.now(),
        reviewCount: Number(row?.reviewCount) || 0,
        lastReviewedAt: Number(row?.lastReviewedAt) || null
      });
      existing.add(key);
      inserted += 1;

      if (inserted % 1500 === 0) await sleep(0);
    }

    await reloadWords();
    await renderWordListPage();
    scheduleAutoBackup('import');
    setStatus(`Imported ${inserted.toLocaleString()} new words.`, 2200);
  }

  async function saveWord() {
    const word = els.wordInput.value.trim();
    const meaning = els.meaningInput.value.trim();
    const example = els.exampleInput.value.trim();

    if (!word || !meaning) {
      setStatus('Word and meaning are required.', 2200);
      (!word ? els.wordInput : els.meaningInput).focus();
      return;
    }

    const started = performance.now();
    await addWord({
      word,
      meaning,
      example,
      createdAt: Date.now(),
      reviewCount: 0,
      lastReviewedAt: null
    });

    els.wordInput.value = '';
    els.meaningInput.value = '';
    els.exampleInput.value = '';
    els.wordInput.focus();

    await reloadWords();
    if (!els.listScreen.classList.contains('hidden')) {
      await renderWordListPage();
    }
    scheduleAutoBackup('save');

    setStatus(`Saved in ${(performance.now() - started).toFixed(1)} ms.`);
  }

  function startReviewMode() {
    state.reviewPool = [...state.words];
    state.reviewIndex = 0;
    state.reviewPool.sort(() => Math.random() - 0.5);
    showReviewScreen();

    if (!state.reviewPool.length) {
      els.reviewEmpty.classList.remove('hidden');
      els.reviewContent.classList.add('hidden');
      return;
    }

    els.reviewEmpty.classList.add('hidden');
    els.reviewContent.classList.remove('hidden');
    renderReviewWord();
  }

  function closeReviewMode() {
    showMainScreen();
  }

  function renderReviewWord() {
    if (!state.reviewPool.length) {
      els.reviewEmpty.classList.remove('hidden');
      els.reviewContent.classList.add('hidden');
      return;
    }

    if (state.reviewIndex >= state.reviewPool.length) {
      state.reviewPool.sort(() => Math.random() - 0.5);
      state.reviewIndex = 0;
    }

    const row = state.reviewPool[state.reviewIndex];
    els.reviewWord.textContent = row.word;
    els.reviewMeaning.textContent = row.meaning;
    els.reviewMeaning.classList.add('hidden');
    els.reviewExample.textContent = row.example ? `Example: ${row.example}` : '';
  }

  async function reviewStep(isKnown) {
    if (!state.reviewPool.length) return;

    const current = state.reviewPool[state.reviewIndex];
    const updated = {
      ...current,
      reviewCount: (current.reviewCount || 0) + 1,
      lastReviewedAt: Date.now()
    };

    await putWord(updated);
    state.reviewPool[state.reviewIndex] = updated;
    state.reviewIndex += 1;
    renderReviewWord();

    await reloadWords();
    if (!els.listScreen.classList.contains('hidden')) {
      await renderWordListPage();
    }
    scheduleAutoBackup('review');
    setStatus(isKnown ? 'Marked known.' : 'Marked unknown.');
  }

  function bindEvents() {
    els.saveBtn.addEventListener('click', () => saveWord().catch(console.error));
    els.meaningInput.addEventListener('keydown', (event) => {
      if (event.key !== 'Enter') return;
      event.preventDefault();
      saveWord().catch(console.error);
    });

    let searchTimer = null;
    els.searchInput.addEventListener('input', () => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => {
        state.searchTerm = els.searchInput.value;
        applySearch();
        if (!els.listScreen.classList.contains('hidden')) {
          renderWordListPage().catch(console.error);
        }
      }, 60);
    });

    els.reviewModeBtn.addEventListener('click', startReviewMode);
    els.exitReviewBtn.addEventListener('click', closeReviewMode);
    els.wordListBtn.addEventListener('click', () => showListScreen().catch(console.error));
    els.backToMainBtn.addEventListener('click', showMainScreen);

    els.showMeaningBtn.addEventListener('click', () => els.reviewMeaning.classList.remove('hidden'));
    els.knownBtn.addEventListener('click', () => reviewStep(true).catch(console.error));
    els.unknownBtn.addEventListener('click', () => reviewStep(false).catch(console.error));

    els.prevPageBtn.addEventListener('click', () => {
      if (state.currentPage <= 1) return;
      state.currentPage -= 1;
      renderWordListPage().catch(console.error);
    });

    els.nextPageBtn.addEventListener('click', () => {
      state.currentPage += 1;
      renderWordListPage().catch(console.error);
    });

    els.exportBtn.addEventListener('click', manualExport);
    els.importBtn.addEventListener('click', () => els.importFile.click());
    els.importFile.addEventListener('change', () => {
      const [file] = els.importFile.files;
      if (!file) return;
      importBackup(file).catch((error) => {
        console.error(error);
        setStatus('Import failed.', 2200);
      }).finally(() => {
        els.importFile.value = '';
      });
    });

    els.backupLinkBtn.addEventListener('click', () => {
      linkBackupFile().catch((error) => {
        console.error(error);
        setStatus('Unable to link backup file.', 2200);
      });
    });

    window.addEventListener('beforeunload', () => {
      void doImmediateExitBackup('beforeunload');
    });

    window.addEventListener('pagehide', () => {
      void doImmediateExitBackup('pagehide');
    });

    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'hidden') {
        void doImmediateExitBackup('visibilitychange');
      }
    });
  }

  async function registerSW() {
    if (!('serviceWorker' in navigator)) return;
    try {
      await navigator.serviceWorker.register('./service-worker.js');
    } catch (error) {
      console.error('Service worker registration failed', error);
    }
  }

  async function initBackupHandle() {
    state.backupHandle = await getMeta('backupFileHandle');
    if (!state.backupHandle) {
      setStatus('Tip: Link vocab_backup.json once for true overwrite auto-backup to OneDrive.', 2800);
      return;
    }

    const ok = await ensureBackupPermission();
    if (ok) await backupToLinkedFile('startup');
  }

  async function init() {
    await openDB();
    bindEvents();
    await reloadWords();
    await initBackupHandle();
    await registerSW();
    showMainScreen();
    els.wordInput.focus();
  }

  init().catch((error) => {
    console.error(error);
    setStatus('App failed to initialize.', 2600);
  });
})();
