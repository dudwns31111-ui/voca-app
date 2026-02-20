(() => {
  'use strict';

  const DB_NAME = 'vocabDB';
  const DB_VERSION = 2;
  const WORDS_STORE = 'words';
  const META_STORE = 'appMeta';
  const BACKUP_FILE_NAME = 'vocab_backup.json';
  const AUTO_BACKUP_DEBOUNCE_MS = 10_000;
  const DEFAULT_PAGE_SIZE = 100;
  const MAX_PAGE_SIZE = 200;
  let displayOrder = 'word-first';

  const state = {
    db: null,
    words: [],
    filteredWords: [],
    sortedWordsCache: [],
    sortedWordsCacheKey: '',
    wordsVersion: 0,
    searchTerm: '',
    currentPage: 1,
    pageSize: DEFAULT_PAGE_SIZE,
    sortBy: 'newest',
    reviewPool: [],
    currentReviewDirection: 'word-first',
    backupTimer: null,
    backupInFlight: false,
    pendingBackupReason: null,
    backupHandle: null
  };

  const els = {
    mainScreen: document.getElementById('mainScreen'),
    reviewOverlay: document.getElementById('reviewOverlay'),
    searchInput: document.getElementById('searchInput'),
    wordInput: document.getElementById('wordInput'),
    meaningInput: document.getElementById('meaningInput'),
    exampleInput: document.getElementById('exampleInput'),
    saveBtn: document.getElementById('saveBtn'),
    reviewModeBtn: document.getElementById('reviewModeBtn'),
    mobilePrimaryActions: document.getElementById('mobilePrimaryActions'),
    addWordToggleBtn: document.getElementById('addWordToggleBtn'),
    mobileReviewBtn: document.getElementById('mobileReviewBtn'),
    addWordForm: document.getElementById('addWordForm'),
    displayOrderSelect: document.getElementById('displayOrderSelect'),
    sortBySelect: document.getElementById('sortBySelect'),
    pageSizeSelect: document.getElementById('pageSizeSelect'),
    exportBtn: document.getElementById('exportBtn'),
    importBtn: document.getElementById('importBtn'),
    backupToggleBtn: document.getElementById('backupToggleBtn'),
    backupMenu: document.getElementById('backupMenu'),
    importFile: document.getElementById('importFile'),
    linkBtn: document.getElementById('linkBtn'),
    status: document.getElementById('status'),
    totalCount: document.getElementById('totalCount'),
    wordList: document.getElementById('wordList'),
    paginationControls: document.getElementById('paginationControls'),
    prevPageBtn: document.getElementById('prevPageBtn'),
    nextPageBtn: document.getElementById('nextPageBtn'),
    pageIndicator: document.getElementById('pageIndicator'),
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

  function isMobile() {
    return window.matchMedia('(max-width: 600px)').matches;
  }

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

  async function putWords(rows) {
    if (!rows.length) return;
    const db = await openDB();
    const tx = db.transaction(WORDS_STORE, 'readwrite');
    const store = tx.objectStore(WORDS_STORE);
    for (const row of rows) {
      store.put(row);
    }
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

  function getNextReviewAt(intervalDays) {
    const now = new Date();
    const next = new Date(now);
    next.setHours(0, 0, 0, 0);
    next.setDate(next.getDate() + intervalDays);
    return next.getTime();
  }

  function setStatus(message, holdMs = 1700) {
    els.status.textContent = message;
    if (!holdMs) return;
    clearTimeout(setStatus.timer);
    setStatus.timer = setTimeout(() => {
      if (els.status.textContent === message) els.status.textContent = '';
    }, holdMs);
  }

  function isReviewModeOpen() {
    return !els.reviewOverlay.classList.contains('hidden');
  }

  function showMainScreen() {
    els.mainScreen.classList.remove('hidden');
    els.reviewOverlay.classList.add('hidden');
  }

  function showReviewOverlay() {
    els.reviewOverlay.classList.remove('hidden');
  }

  function migrationDefaults(row, now = Date.now()) {
    const migrated = { ...row };
    let changed = false;

    if (!Number.isFinite(migrated.createdAt) || migrated.createdAt <= 0) {
      migrated.createdAt = now;
      changed = true;
    }

    if (!Number.isFinite(migrated.reviewCount) || migrated.reviewCount < 0) {
      migrated.reviewCount = 0;
      changed = true;
    }

    if (!Number.isFinite(migrated.interval) || migrated.interval < 1) {
      migrated.interval = 1;
      changed = true;
    }

    if (!Number.isFinite(migrated.nextReviewAt) || migrated.nextReviewAt <= 0) {
      migrated.nextReviewAt = now;
      changed = true;
    }

    if (!Number.isFinite(migrated.lastReviewedAt) || migrated.lastReviewedAt < 0) {
      migrated.lastReviewedAt = 0;
      changed = true;
    }

    return { row: migrated, changed };
  }

  async function migrateWordsIfNeeded(rows) {
    const now = Date.now();
    const migratedRows = [];
    const changedRows = [];

    for (const row of rows) {
      const migrated = migrationDefaults(row, now);
      migratedRows.push(migrated.row);
      if (migrated.changed) changedRows.push(migrated.row);
    }

    if (changedRows.length) {
      await putWords(changedRows);
    }

    return migratedRows;
  }

  function applySearch() {
    const t = norm(state.searchTerm);
    state.filteredWords = t
      ? state.words.filter((row) => norm(row.word).includes(t) || norm(row.meaning).includes(t))
      : [];
    state.currentPage = 1;
    state.sortedWordsCacheKey = '';
  }

  function getPageCount(totalItems) {
    return Math.max(1, Math.ceil(totalItems / state.pageSize));
  }

  function renderTotalCount() {
    const count = state.words.length;
    els.totalCount.textContent = `${count.toLocaleString()} ${count === 1 ? 'word' : 'words'}`;
  }

  function compareWords(a, b) {
    switch (state.sortBy) {
      case 'oldest':
        return a.createdAt - b.createdAt;
      case 'mostReviewed':
        return (b.reviewCount || 0) - (a.reviewCount || 0) || b.createdAt - a.createdAt;
      case 'leastReviewed':
        return (a.reviewCount || 0) - (b.reviewCount || 0) || b.createdAt - a.createdAt;
      case 'alphabetical':
        return a.word.localeCompare(b.word, undefined, { sensitivity: 'base' }) || b.createdAt - a.createdAt;
      case 'newest':
      default:
        return b.createdAt - a.createdAt;
    }
  }

  function getActiveSource() {
    return state.searchTerm ? state.filteredWords : state.words;
  }

  function getSortedWordsForRender() {
    const source = getActiveSource();
    const cacheKey = `${state.wordsVersion}|${state.sortBy}|${state.pageSize}|${norm(state.searchTerm)}`;
    if (state.sortedWordsCacheKey === cacheKey) {
      return state.sortedWordsCache;
    }

    const sorted = source.slice();
    sorted.sort(compareWords);
    state.sortedWordsCache = sorted;
    state.sortedWordsCacheKey = cacheKey;
    return sorted;
  }

  function getDueWords() {
    const now = Date.now();
    return state.words.filter((row) => row.nextReviewAt <= now);
  }

  function refreshReviewButtonLabel() {
    const dueCount = getDueWords().length;
    const label = `Review (${dueCount.toLocaleString()} due)`;
    els.reviewModeBtn.textContent = `Review Mode (${dueCount.toLocaleString()} due)`;
    if (els.mobileReviewBtn) els.mobileReviewBtn.textContent = label;
  }

  function setReviewButtonsEnabled(enabled) {
    els.showMeaningBtn.disabled = !enabled;
    els.knownBtn.disabled = !enabled;
    els.unknownBtn.disabled = !enabled;
  }

  function makeWordRow(row) {
    const wrap = document.createElement('article');
    wrap.className = 'word-item word-card';

    const main = document.createElement('div');
    main.className = 'word-main';

    const primary = document.createElement('p');
    primary.className = 'word-label word-primary';
    primary.textContent = displayOrder === 'meaning-first' ? row.meaning : row.word;

    const secondary = document.createElement('p');
    secondary.className = 'meaning hidden word-secondary';
    secondary.textContent = displayOrder === 'meaning-first' ? row.word : row.meaning;

    const ex = document.createElement('p');
    ex.className = 'example';
    ex.textContent = row.example || '';

    const meta = document.createElement('p');
    meta.className = 'meta word-meta';
    meta.textContent = `Reviews: ${row.reviewCount || 0} • Added: ${formatDate(row.createdAt)}`;

    primary.addEventListener('click', () => secondary.classList.toggle('hidden'));

    main.append(primary, secondary);
    if (row.example) main.appendChild(ex);
    main.appendChild(meta);

    const del = document.createElement('button');
    del.className = 'delete-btn';
    del.textContent = 'Delete Word';
    del.addEventListener('click', async () => {
      if (!confirm(`Delete "${row.word}"?`)) return;
      await deleteWordById(row.id);
      await reloadWords();
      scheduleAutoBackup('delete');
      setStatus('Word deleted.');
    });

    wrap.append(main, del);
    return wrap;
  }

  async function renderWordListPage() {
    const sorted = getSortedWordsForRender();
    const totalItems = sorted.length;
    const pageCount = getPageCount(totalItems);

    if (state.currentPage > pageCount) {
      state.currentPage = pageCount;
    }

    if (!totalItems) {
      const empty = document.createElement('p');
      empty.className = 'muted';
      empty.textContent = state.searchTerm ? 'No matching words.' : 'No words saved yet.';
      els.wordList.replaceChildren(empty);
      els.paginationControls.classList.add('hidden');
      return;
    }

    const offset = (state.currentPage - 1) * state.pageSize;
    const list = sorted.slice(offset, offset + state.pageSize);

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
    state.words = await migrateWordsIfNeeded(rows);
    state.wordsVersion += 1;
    state.sortedWordsCacheKey = '';
    renderTotalCount();
    refreshReviewButtonLabel();
    applySearch();
    await renderWordListPage();
  }

  function asBackupArray(rows) {
    return rows.map((r) => ({
      id: r.id,
      word: r.word,
      meaning: r.meaning,
      example: r.example || '',
      createdAt: r.createdAt,
      reviewCount: Number(r.reviewCount) || 0,
      interval: Number(r.interval) || 1,
      nextReviewAt: Number(r.nextReviewAt) || Date.now(),
      lastReviewedAt: Number(r.lastReviewedAt) || 0
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

      const now = Date.now();
      const migrated = migrationDefaults({
        word,
        meaning,
        example: String(row?.example || '').trim(),
        createdAt: Number(row?.createdAt) || now,
        reviewCount: Number(row?.reviewCount) || 0,
        interval: Number(row?.interval),
        nextReviewAt: Number(row?.nextReviewAt),
        lastReviewedAt: Number(row?.lastReviewedAt)
      }, now).row;

      await addWord(migrated);
      existing.add(key);
      inserted += 1;

      if (inserted % 1500 === 0) await sleep(0);
    }

    await reloadWords();
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
      interval: 1,
      nextReviewAt: Date.now(),
      lastReviewedAt: 0
    });

    els.wordInput.value = '';
    els.meaningInput.value = '';
    els.exampleInput.value = '';
    els.addWordForm.classList.remove('active');
    els.wordInput.focus();

    await reloadWords();
    scheduleAutoBackup('save');

    setStatus(`Saved in ${(performance.now() - started).toFixed(1)} ms.`);
  }

  function startReviewMode() {
    state.reviewPool = getDueWords();
    state.reviewPool.sort(() => Math.random() - 0.5);
    showReviewOverlay();

    if (!state.reviewPool.length) {
      els.reviewEmpty.textContent = 'No words due for review 🎉';
      els.reviewEmpty.classList.remove('hidden');
      els.reviewContent.classList.add('hidden');
      setReviewButtonsEnabled(false);
      return;
    }

    els.reviewEmpty.classList.add('hidden');
    els.reviewContent.classList.remove('hidden');
    setReviewButtonsEnabled(true);
    renderReviewWord();
  }

  function closeReviewMode() {
    showMainScreen();
  }

  function renderReviewWord() {
    if (!state.reviewPool.length) {
      els.reviewEmpty.textContent = 'No words due for review 🎉';
      els.reviewEmpty.classList.remove('hidden');
      els.reviewContent.classList.add('hidden');
      setReviewButtonsEnabled(false);
      els.showMeaningBtn.textContent = 'Show Meaning';
      refreshReviewButtonLabel();
      return;
    }

    const row = state.reviewPool[0];
    state.currentReviewDirection = Math.random() < 0.7 ? 'meaning-first' : 'word-first';

    if (state.currentReviewDirection === 'meaning-first') {
      els.reviewWord.textContent = row.meaning;
      els.reviewMeaning.textContent = row.word;
      els.showMeaningBtn.textContent = 'Show Word';
    } else {
      els.reviewWord.textContent = row.word;
      els.reviewMeaning.textContent = row.meaning;
      els.showMeaningBtn.textContent = 'Show Meaning';
    }

    els.reviewWord.classList.remove('hidden');
    els.reviewMeaning.classList.add('hidden');
    els.reviewExample.textContent = row.example ? `Example: ${row.example}` : '';
  }

  function revealHiddenReviewSide() {
    if (els.showMeaningBtn.disabled) return;
    els.reviewMeaning.classList.remove('hidden');
  }

  async function reviewStep(isKnown) {
    if (!state.reviewPool.length) return;

    const current = state.reviewPool[0];
    const now = Date.now();
    const nextInterval = isKnown
      ? Math.max(1, Math.round((current.interval || 1) * 2))
      : 1;

    const updated = {
      ...current,
      interval: nextInterval,
      lastReviewedAt: now,
      nextReviewAt: isKnown ? getNextReviewAt(nextInterval) : getNextReviewAt(1),
      reviewCount: isKnown ? (current.reviewCount || 0) + 1 : (current.reviewCount || 0)
    };

    await putWord(updated);

    const i = state.words.findIndex((row) => row.id === updated.id);
    if (i >= 0) {
      state.words[i] = updated;
      state.wordsVersion += 1;
      state.sortedWordsCacheKey = '';
    }

    state.reviewPool.shift();
    renderReviewWord();
    refreshReviewButtonLabel();

    if (!state.searchTerm) {
      await renderWordListPage();
    } else {
      applySearch();
      await renderWordListPage();
    }

    scheduleAutoBackup('review');
    setStatus(isKnown ? 'Marked known.' : 'Marked unknown.');
  }

  function handleReviewShortcuts(event) {
    if (!isReviewModeOpen()) return;

    if (event.key === 'Escape') {
      event.preventDefault();
      closeReviewMode();
      return;
    }

    if (event.code === 'Space' || event.key === ' ') {
      event.preventDefault();
      revealHiddenReviewSide();
      return;
    }

    if (event.key === 'ArrowRight') {
      event.preventDefault();
      if (!els.knownBtn.disabled) reviewStep(true).catch(console.error);
      return;
    }

    if (event.key === 'ArrowLeft') {
      event.preventDefault();
      if (!els.unknownBtn.disabled) reviewStep(false).catch(console.error);
    }
  }

  function bindEvents() {
    els.saveBtn.addEventListener('click', () => saveWord().catch(console.error));
    els.meaningInput.addEventListener('keydown', (event) => {
      if (event.key !== 'Enter') return;
      event.preventDefault();
      saveWord().catch(console.error);
    });

    if (els.addWordToggleBtn) {
      els.addWordToggleBtn.addEventListener('click', () => {
        if (!window.matchMedia('(max-width: 600px)').matches) return;
        const isActive = els.addWordForm.classList.toggle('active');
        if (isActive) els.wordInput.focus();
      });
    }
    
    if (els.mobileReviewBtn) {
      els.mobileReviewBtn.addEventListener('click', startReviewMode);
    }

    let searchTimer = null;
    els.searchInput.addEventListener('input', () => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => {
        state.searchTerm = els.searchInput.value;
        applySearch();
        renderWordListPage().catch(console.error);
      }, 60);
    });

    els.sortBySelect.addEventListener('change', () => {
      state.sortBy = els.sortBySelect.value;
      state.currentPage = 1;
      state.sortedWordsCacheKey = '';
      renderWordListPage().catch(console.error);
    });

    els.displayOrderSelect.addEventListener('change', () => {
      displayOrder = els.displayOrderSelect.value;
      renderWordListPage().catch(console.error);
    });

    els.pageSizeSelect.addEventListener('change', () => {
      const next = Math.min(MAX_PAGE_SIZE, Math.max(1, Number(els.pageSizeSelect.value) || DEFAULT_PAGE_SIZE));
      state.pageSize = next;
      els.pageSizeSelect.value = String(next);
      state.currentPage = 1;
      renderWordListPage().catch(console.error);
    });

    els.reviewModeBtn.addEventListener('click', startReviewMode);
    els.exitReviewBtn.addEventListener('click', closeReviewMode);

    els.showMeaningBtn.addEventListener('click', () => {
      revealHiddenReviewSide();
    });
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

    els.linkBtn.addEventListener('click', () => {
      linkBackupFile().catch((error) => {
        console.error(error);
        setStatus('Unable to link backup file.', 2200);
      });
    });

    if (els.backupToggleBtn && els.backupMenu) {
      els.backupToggleBtn.addEventListener('click', () => {
        if (!window.matchMedia('(max-width: 600px)').matches) return;
        els.backupMenu.classList.toggle('active');
      });
    }

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

    document.addEventListener('keydown', handleReviewShortcuts);
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

  function syncMobileUI() {
    if (!els.mobilePrimaryActions || !els.addWordForm) return;

    const mobile = isMobile();
    els.mobilePrimaryActions.classList.toggle('hidden', !mobile);

    if (!mobile) {
      els.addWordForm.classList.remove('active');
      return;
    }

    els.addWordForm.classList.remove('active');
  }

  async function init() {
    await openDB();
    bindEvents();
    els.displayOrderSelect.value = displayOrder;
    els.sortBySelect.value = state.sortBy;
    els.pageSizeSelect.value = String(state.pageSize);
    syncMobileUI();
    window.matchMedia('(max-width: 600px)').addEventListener('change', syncMobileUI);
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

