let currentDate = new Date();
let currentMeetingId = null;
let currentMeetingData = null;
let currentTab = 'calendar';
let actions = [];
let actionSearchQuery = '';

// Keyboard navigation state
let selectedCardIndex = -1;
let eventCards = [];

// Auto-save state
let autoSaveTimeout = null;
let isSaving = false;

// Action card expanded state (in-memory)
let expandedActions = new Set();

// Folder expanded state
let expandedFolders = new Set();

// Currently expanded folder for smart creation
let currentExpandedFolder = null;

// Countdown timer
let countdownInterval = null;

// Week view state
let currentWeekStart = null; // Will be set to Monday of current week
let highlightEventAtHour = null; // Hour to highlight when navigating from heatmap
let highlightAppliedAt = null; // Timestamp when highlight was first applied
let highlightTimeout = null; // Pending timeout for applying highlight

// Theme state
let currentTheme = 'ocean-blue';

// Tab order state
let tabOrder = ['calendar', 'weekview', 'actions'];

// ===== SESSION CACHE (in-memory) =====
// Keeps events in memory for the entire session - avoids redundant Outlook queries
const sessionCache = {
  events: new Map(), // Key: 'YYYY-MM-DD', Value: { events: [...], timestamp: Date }
  lastOutlookSync: null, // Last time we queried Outlook
  CACHE_TTL: 15 * 60 * 1000, // 15 minutes - only re-query Outlook if older than this
  weekViewLoaded: false, // Lazy load flag for week view

  set(dateKey, events) {
    this.events.set(dateKey, {
      events: events,
      timestamp: Date.now()
    });
  },

  get(dateKey) {
    return this.events.get(dateKey);
  },

  has(dateKey) {
    return this.events.has(dateKey);
  },

  isStale(dateKey) {
    const cached = this.get(dateKey);
    if (!cached) return true;
    return (Date.now() - cached.timestamp) > this.CACHE_TTL;
  },

  shouldSyncOutlook() {
    if (!this.lastOutlookSync) return true;
    return (Date.now() - this.lastOutlookSync) > this.CACHE_TTL;
  },

  markOutlookSynced() {
    this.lastOutlookSync = Date.now();
  },

  clear() {
    this.events.clear();
    this.lastOutlookSync = null;
  }
};

// Initialize the app
async function init() {
  // Load saved theme first
  loadTheme();
  // Restore tab order before rendering (reorders DOM nodes)
  loadTabOrder();

  // Setup listeners synchronously (fast)
  setupEventListeners();
  setupKeyboardNavigation();
  setupScrollDetection();
  setupAutoSave();
  setupFocusManagement();
  // Enable drag-to-reorder for tab buttons
  initTabDragAndDrop();

  // Don't await - let it run in background
  fastStartup();
}

// ===== FOCUS MANAGEMENT =====
// Prevent focus capture issues that can interfere with other applications
function setupFocusManagement() {
  // Release focus from any contentEditable elements when clicking outside the widget
  document.addEventListener('blur', () => {
    // Find any active contentEditable elements
    const editingElements = document.querySelectorAll('[contenteditable="true"]');
    editingElements.forEach(el => {
      // Trigger blur to save/cancel editing
      el.blur();
    });
  });

  // Also handle when the entire window loses focus
  window.addEventListener('blur', () => {
    // Cancel any ongoing inline edits
    const editingElements = document.querySelectorAll('.editing');
    editingElements.forEach(el => {
      if (el.contentEditable === 'true') {
        // Restore original text if available
        if (el.dataset.originalText) {
          el.textContent = el.dataset.originalText;
        }
        el.contentEditable = false;
        el.classList.remove('editing');
        delete el.dataset.originalText;
      }
    });
  });
}

// ── Fast startup helpers ────────────────────────────────────────────────────

function formatDateKey(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function setRefreshSpinning(on) {
  const btn = document.getElementById('refresh-btn');
  if (btn) btn.classList.toggle('spinning', on);
}

// Step 1: show UI instantly from disk cache; Step 2: refresh from Outlook in background
async function fastStartup() {
  // Show main screen IMMEDIATELY - don't wait for anything
  showMainScreen();
  updateDateDisplay();

  const dateKey = formatDateKey(currentDate);
  const eventsList = document.getElementById('events-list');

  // Priority 1: Check session cache (fastest - already in memory)
  if (sessionCache.has(dateKey)) {
    const cached = sessionCache.get(dateKey);
    displayEvents(cached.events);
    console.log(`✓ Instant display from session cache`);

    // Only refresh from Outlook if cache is stale
    if (!sessionCache.isStale(dateKey)) {
      console.log(`✓ Cache is fresh (< 15 min), skipping Outlook query`);
      // Still pre-fetch surrounding days in background if we haven't synced recently
      if (sessionCache.shouldSyncOutlook()) {
        setTimeout(() => backgroundPrefetchRange(currentDate), 2000);
      }
      return;
    }
  }

  // Priority 2: Try disk cache (fast - no Outlook query needed)
  let hasCacheForToday = false;
  try {
    const cacheResult = await window.electronAPI.getCachedEvents(dateKey);
    if (cacheResult.success && Array.isArray(cacheResult.events) && cacheResult.fromCache) {
      displayEvents(cacheResult.events);
      sessionCache.set(dateKey, cacheResult.events);
      hasCacheForToday = true;
      console.log(`✓ Fast display from disk cache`);
    } else {
      eventsList.innerHTML = '<div class="no-events">Loading events…</div>';
    }
  } catch (error) {
    console.error('Cache read error:', error);
    eventsList.innerHTML = '<div class="no-events">Loading events…</div>';
  }

  // Load actions synchronously from disk (instant, no IPC delay)
  loadActionsSync();

  // Priority 3: Background refresh from Outlook (only if cache is stale or missing)
  if (!hasCacheForToday || sessionCache.shouldSyncOutlook()) {
    backgroundRefresh(dateKey, hasCacheForToday);
  } else {
    console.log(`✓ Skipping Outlook sync - cache is fresh`);
  }
}

// Synchronous actions load (non-blocking)
function loadActionsSync() {
  window.electronAPI.getActions().then(actionsResult => {
    if (actionsResult.success) {
      actions = actionsResult.actions;
      console.log(`✓ Actions loaded (${actions.length} items)`);
    }
  }).catch(err => console.error('Actions load error:', err));
}

async function backgroundRefresh(dateKey, hasCachedData) {
  setRefreshSpinning(true);
  console.log(`⟳ Syncing with Outlook...`);

  const startDate = new Date(currentDate);
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(currentDate);
  endDate.setHours(23, 59, 59, 999);

  try {
    const result = await window.electronAPI.checkAndGetEvents(startDate, endDate);

    if (result.success && result.available) {
      // Preserve highlight state before rebuilding DOM
      const shouldHighlight = highlightEventAtHour !== null;
      const highlightHour = highlightEventAtHour;

      // Update session cache
      sessionCache.set(dateKey, result.events);
      sessionCache.markOutlookSynced();

      // Cache to disk
      window.electronAPI.saveEventsCache(result.events, dateKey);

      // Only update display if we're still on the calendar tab and same date
      if (currentTab === 'calendar' && formatDateKey(currentDate) === dateKey) {
        // Restore highlight state if it was set
        if (shouldHighlight) {
          highlightEventAtHour = highlightHour;
        }
        displayEvents(result.events);
      }

      console.log(`✓ Outlook sync complete`);

      // Pre-fetch surrounding days in background (non-blocking, delayed)
      setTimeout(() => backgroundPrefetchRange(currentDate), 1000);
    } else if (!result.available) {
      if (!hasCachedData) {
        document.getElementById('events-list').innerHTML =
          '<div class="no-events">Outlook not available<br>' +
          '<span class="status-hint">Open Outlook and click ↻ to retry</span></div>';
        // One auto-retry after 15 s in case Outlook just hasn't started yet
        setTimeout(() => backgroundRefresh(dateKey, false), 15000);
      }
      // If we already showed cache, keep showing it silently
    }
  } catch (error) {
    console.error('Background refresh error:', error);
    if (!hasCachedData) {
      document.getElementById('events-list').innerHTML =
        '<div class="no-events">Could not connect to Outlook</div>';
    }
  } finally {
    setRefreshSpinning(false);
  }
}

// Pre-fetch events for ±5 days + next 2 weeks around the given center date (runs in background)
async function backgroundPrefetchRange(centerDate) {
  try {
    console.log('⟳ Pre-fetching 3 weeks of events...');

    const rangeStart = new Date(centerDate);
    rangeStart.setDate(rangeStart.getDate() - 5);
    rangeStart.setHours(0, 0, 0, 0);

    const rangeEnd = new Date(centerDate);
    rangeEnd.setDate(rangeEnd.getDate() + 19); // Current + 2 weeks ahead
    rangeEnd.setHours(23, 59, 59, 999);

    // Single bulk fetch for ~24 days
    const result = await window.electronAPI.getEvents(rangeStart, rangeEnd);

    if (result.success && result.events) {
      // Group events by date
      const eventsMap = {};

      result.events.forEach(event => {
        const eventDate = new Date(event.start);
        const key = formatDateKey(eventDate);

        if (!eventsMap[key]) {
          eventsMap[key] = [];
        }
        eventsMap[key].push(event);
      });

      // Update session cache for all fetched dates
      Object.keys(eventsMap).forEach(dateKey => {
        sessionCache.set(dateKey, eventsMap[dateKey]);
      });

      // Save to disk cache (fire-and-forget)
      await window.electronAPI.saveEventsRangeCache(eventsMap);

      console.log(`✓ Pre-fetched ${Object.keys(eventsMap).length} days`);
      sessionCache.markOutlookSynced();
    }
  } catch (error) {
    console.error('Background pre-fetch error (non-critical):', error);
  }
}

function setupEventListeners() {
  // Window controls
  document.getElementById('minimize-btn').addEventListener('click', () => {
    window.electronAPI.minimizeWindow();
  });

  document.getElementById('close-btn').addEventListener('click', () => {
    window.electronAPI.closeWindow();
  });

  // Retry connection
  document.getElementById('retry-btn').addEventListener('click', checkOutlookConnection);

  // Tab switching
  document.getElementById('calendar-tab').addEventListener('click', () => switchTab('calendar'));
  document.getElementById('actions-tab').addEventListener('click', () => switchTab('actions'));
  document.getElementById('weekview-tab').addEventListener('click', () => switchTab('weekview'));

  // Date navigation
  document.getElementById('prev-day').addEventListener('click', () => {
    // Clear any active highlight when manually navigating
    if (highlightTimeout) clearTimeout(highlightTimeout);
    highlightEventAtHour = null;
    highlightAppliedAt = null;
    highlightTimeout = null;
    currentDate.setDate(currentDate.getDate() - 1);
    loadEvents();
  });

  document.getElementById('next-day').addEventListener('click', () => {
    // Clear any active highlight when manually navigating
    if (highlightTimeout) clearTimeout(highlightTimeout);
    highlightEventAtHour = null;
    highlightAppliedAt = null;
    highlightTimeout = null;
    currentDate.setDate(currentDate.getDate() + 1);
    loadEvents();
  });

  document.getElementById('refresh-btn').addEventListener('click', async () => {
    setRefreshSpinning(true);

    if (currentTab === 'weekview') {
      // Force refresh: mark cache as stale
      sessionCache.lastOutlookSync = null;
      await populateHeatmap();
    } else {
      // Force refresh: mark current date as stale and sync
      const dateKey = formatDateKey(currentDate);
      sessionCache.lastOutlookSync = null; // Force Outlook sync
      await loadEvents(true); // Pass forceRefresh flag
    }

    setRefreshSpinning(false);
  });

  // Week view navigation
  document.getElementById('prev-week').addEventListener('click', () => {
    navigateWeek(-1);
  });

  document.getElementById('next-week').addEventListener('click', () => {
    navigateWeek(1);
  });

  // Notes
  document.getElementById('close-notes-btn').addEventListener('click', closeNotes);
  document.getElementById('save-notes-btn').addEventListener('click', saveNotes);

  // Actions
  document.getElementById('add-action-btn').addEventListener('click', () => {
    clearExpandedFolderContext(); // Create at root level
    openAddActionModal();
  });
  document.getElementById('cancel-action-btn').addEventListener('click', closeAddActionModal);
  document.getElementById('action-form').addEventListener('submit', handleAddAction);

  // Folders
  document.getElementById('add-folder-btn').addEventListener('click', () => {
    clearExpandedFolderContext(); // Create at root level
    openAddFolderModal();
  });
  document.getElementById('cancel-folder-btn').addEventListener('click', closeAddFolderModal);
  document.getElementById('folder-form').addEventListener('submit', handleAddFolder);

  // Actions search
  document.getElementById('action-search-input').addEventListener('input', (e) => {
    actionSearchQuery = e.target.value;
    displayActions();
  });

  // Close modal on overlay click
  document.getElementById('action-modal').addEventListener('click', (e) => {
    if (e.target.id === 'action-modal') {
      closeAddActionModal();
    }
  });

  document.getElementById('folder-modal').addEventListener('click', (e) => {
    if (e.target.id === 'folder-modal') {
      closeAddFolderModal();
    }
  });

  // URL/Path mutual exclusivity in action modal
  document.getElementById('action-url-input').addEventListener('input', (e) => {
    const pathInput = document.getElementById('action-path-input');
    if (e.target.value.trim()) {
      pathInput.disabled = true;
      pathInput.placeholder = 'URL is set - clear URL to use path';
    } else {
      pathInput.disabled = false;
      pathInput.placeholder = 'C:\\Users\\...';
    }
  });

  document.getElementById('action-path-input').addEventListener('input', (e) => {
    // Strip quotes from pasted paths (Windows File Explorer adds quotes)
    let pathValue = e.target.value.trim();
    if ((pathValue.startsWith('"') && pathValue.endsWith('"')) ||
        (pathValue.startsWith("'") && pathValue.endsWith("'"))) {
      pathValue = pathValue.slice(1, -1);
      e.target.value = pathValue;
    }

    const urlInput = document.getElementById('action-url-input');
    if (pathValue) {
      urlInput.disabled = true;
      urlInput.placeholder = 'Path is set - clear path to use URL';
    } else {
      urlInput.disabled = false;
      urlInput.placeholder = 'https://...';
    }
  });

  // Theme picker
  document.getElementById('theme-btn').addEventListener('click', (e) => {
    e.stopPropagation();
    toggleThemePicker();
  });

  // Close theme picker when clicking outside
  document.addEventListener('click', (e) => {
    // Only handle clicks within our window
    if (!document.hasFocus()) {
      return;
    }

    const themePicker = document.getElementById('theme-picker');
    const themeBtn = document.getElementById('theme-btn');
    if (!themePicker.contains(e.target) && e.target !== themeBtn && !themeBtn.contains(e.target)) {
      themePicker.classList.add('hidden');
    }
  });

  // Theme options
  document.querySelectorAll('.theme-option').forEach(option => {
    option.addEventListener('click', () => {
      const theme = option.dataset.theme;
      setTheme(theme);
      document.getElementById('theme-picker').classList.add('hidden');
    });
  });
}

// ===== KEYBOARD NAVIGATION =====

function setupKeyboardNavigation() {
  document.addEventListener('keydown', (e) => {
    // CRITICAL: Only handle keyboard events if the widget window has focus
    // This prevents interfering with other applications
    if (!document.hasFocus()) {
      return;
    }

    // Don't capture events if user is typing in an input/textarea/contentEditable
    const activeElement = document.activeElement;
    if (
      activeElement.tagName === 'INPUT' ||
      activeElement.tagName === 'TEXTAREA' ||
      activeElement.contentEditable === 'true'
    ) {
      return;
    }

    // Handle Escape key
    if (e.key === 'Escape') {
      handleEscapeKey();
      return;
    }

    // Only handle navigation keys when in calendar tab with events
    if (currentTab === 'calendar' && !isNotesViewVisible()) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        navigateCards('down');
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        navigateCards('up');
      } else if (e.key === 'Enter') {
        e.preventDefault();
        openSelectedCard();
      }
    }
  });
}

function navigateCards(direction) {
  updateEventCardsArray();

  if (eventCards.length === 0) return;

  // Calculate new index
  if (direction === 'down') {
    selectedCardIndex = Math.min(selectedCardIndex + 1, eventCards.length - 1);
  } else if (direction === 'up') {
    selectedCardIndex = Math.max(selectedCardIndex - 1, 0);
  }

  // Update visual selection
  updateCardSelection();
}

function updateEventCardsArray() {
  eventCards = Array.from(document.querySelectorAll('.event-card'));
}

function updateCardSelection() {
  // Remove previous selection
  eventCards.forEach(card => card.classList.remove('selected'));

  // Add selection to current card
  if (selectedCardIndex >= 0 && selectedCardIndex < eventCards.length) {
    const selectedCard = eventCards[selectedCardIndex];
    selectedCard.classList.add('selected');

    // Smooth scroll into view
    selectedCard.scrollIntoView({
      behavior: 'smooth',
      block: 'nearest',
      inline: 'nearest'
    });
  }
}

function openSelectedCard() {
  if (selectedCardIndex >= 0 && selectedCardIndex < eventCards.length) {
    eventCards[selectedCardIndex].click();
  }
}

function resetCardSelection() {
  selectedCardIndex = -1;
  eventCards.forEach(card => card.classList.remove('selected'));
}

function handleEscapeKey() {
  // Check if action modal is open
  const actionModal = document.getElementById('action-modal');
  if (actionModal && !actionModal.classList.contains('hidden')) {
    closeAddActionModal();
    return;
  }

  // Check if folder modal is open
  const folderModal = document.getElementById('folder-modal');
  if (folderModal && !folderModal.classList.contains('hidden')) {
    closeAddFolderModal();
    return;
  }

  // Check if notes view is open
  if (isNotesViewVisible()) {
    closeNotes();
    return;
  }

  // Otherwise, clear card selection
  resetCardSelection();
}

function isNotesViewVisible() {
  const notesView = document.getElementById('notes-view');
  return notesView && !notesView.classList.contains('hidden');
}

// ===== SCROLL DETECTION =====

function setupScrollDetection() {
  // Monitor events list scrolling
  const eventsList = document.getElementById('events-list');
  if (eventsList) {
    eventsList.addEventListener('scroll', () => updateScrollIndicators(eventsList));
  }

  // Monitor notes editor scrolling
  const notesEditor = document.getElementById('notes-editor');
  if (notesEditor) {
    notesEditor.addEventListener('scroll', () => updateScrollIndicators(notesEditor));
  }

  // Also check on window resize
  window.addEventListener('resize', () => {
    updateScrollIndicators(eventsList);
    updateScrollIndicators(notesEditor);
  });
}

function updateScrollIndicators(element) {
  if (!element) return;

  const isScrollable = element.scrollHeight > element.clientHeight;
  const hasScrollTop = element.scrollTop > 0;
  const hasScrollBottom = element.scrollTop + element.clientHeight < element.scrollHeight - 1;

  element.classList.toggle('is-scrollable', isScrollable);
  element.classList.toggle('has-scroll-top', hasScrollTop);
  element.classList.toggle('has-scroll-bottom', hasScrollBottom);
}

function checkScrollIndicators() {
  const eventsList = document.getElementById('events-list');
  const notesEditor = document.getElementById('notes-editor');

  updateScrollIndicators(eventsList);
  updateScrollIndicators(notesEditor);
}

// ===== AUTO-SAVE =====

function setupAutoSave() {
  const notesEditor = document.getElementById('notes-editor');
  if (notesEditor) {
    notesEditor.addEventListener('input', handleNotesInput);
  }
}

function handleNotesInput() {
  // Clear existing timeout
  if (autoSaveTimeout) {
    clearTimeout(autoSaveTimeout);
  }

  // Set status to indicate typing
  const statusEl = document.getElementById('notes-status');
  if (!isSaving) {
    statusEl.textContent = '';
  }

  // Debounce: save 2 seconds after user stops typing
  autoSaveTimeout = setTimeout(() => {
    autoSaveNotes();
  }, 2000);
}

async function autoSaveNotes() {
  if (!currentMeetingId || !currentMeetingData || isSaving) return;

  const noteContent = document.getElementById('notes-editor').value;
  const statusEl = document.getElementById('notes-status');

  isSaving = true;
  statusEl.textContent = 'Saving...';

  try {
    const result = await window.electronAPI.saveNote(
      currentMeetingId,
      noteContent,
      currentMeetingData
    );

    if (result.success) {
      statusEl.textContent = 'Saved';

      // Fade out after 2 seconds
      setTimeout(() => {
        if (statusEl.textContent === 'Saved') {
          statusEl.textContent = '';
        }
      }, 2000);
    } else {
      statusEl.textContent = 'Error: ' + result.error;
    }
  } catch (error) {
    statusEl.textContent = 'Error: ' + error.message;
  } finally {
    isSaving = false;
  }
}

// ===== MEETING STATE DETECTION =====

function getMeetingState(event) {
  const now = new Date();
  const startTime = new Date(event.start);
  const endTime = new Date(event.end);

  // Past: meeting has ended
  if (now > endTime) {
    return 'past';
  }

  // In progress: meeting has started but not ended
  if (now >= startTime && now <= endTime) {
    return 'in-progress';
  }

  // Upcoming soon: starts within 15 minutes
  const fifteenMinutes = 15 * 60 * 1000;
  if (startTime - now <= fifteenMinutes && startTime > now) {
    return 'upcoming';
  }

  // Future: starts more than 15 minutes away
  return 'future';
}

async function checkOutlookConnection() {
  const statusMessage = document.getElementById('status-message');
  const retryBtn = document.getElementById('retry-btn');

  statusMessage.innerHTML = 'Connecting to Outlook...';
  retryBtn.classList.add('hidden');

  try {
    const result = await window.electronAPI.checkOutlook();

    if (result.success && result.available) {
      showMainScreen();
      await loadEvents();
      await loadActions();
    } else {
      statusMessage.innerHTML = 'Outlook not detected<br><span class="status-hint">Open Outlook and click Retry</span>';
      retryBtn.classList.remove('hidden');

      // Auto-retry every 10 seconds
      setTimeout(async () => {
        if (!retryBtn.classList.contains('hidden')) {
          await checkOutlookConnection();
        }
      }, 10000);
    }
  } catch (error) {
    statusMessage.innerHTML = `Connection error<br><span class="status-hint">${error.message}</span>`;
    retryBtn.classList.remove('hidden');

    // Auto-retry on error
    setTimeout(async () => {
      if (!retryBtn.classList.contains('hidden')) {
        await checkOutlookConnection();
      }
    }, 10000);
  }
}

function showLoginScreen() {
  document.getElementById('login-screen').classList.remove('hidden');
  document.getElementById('main-screen').classList.add('hidden');
}

function showMainScreen() {
  document.getElementById('login-screen').classList.add('hidden');
  document.getElementById('main-screen').classList.remove('hidden');
}

function switchTab(tab) {
  currentTab = tab;

  // Update tab buttons
  document.getElementById('calendar-tab').classList.toggle('active', tab === 'calendar');
  document.getElementById('actions-tab').classList.toggle('active', tab === 'actions');
  document.getElementById('weekview-tab').classList.toggle('active', tab === 'weekview');

  // Update content
  document.getElementById('calendar-content').classList.toggle('hidden', tab !== 'calendar');
  document.getElementById('actions-content').classList.toggle('hidden', tab !== 'actions');
  document.getElementById('weekview-content').classList.toggle('hidden', tab !== 'weekview');

  // Reset card selection when switching tabs
  if (tab !== 'calendar') {
    resetCardSelection();
  }

  // Load data if needed
  if (tab === 'actions') {
    displayActions();
  } else if (tab === 'weekview') {
    // Lazy load: only initialize week view the first time user opens it
    if (!sessionCache.weekViewLoaded) {
      console.log('⟳ First-time week view load...');
      sessionCache.weekViewLoaded = true;
      initWeekView();
    } else {
      // Just refresh the display with cached data
      initWeekView();
    }
  }
}

// ===== CALENDAR FUNCTIONS =====

async function loadEvents(forceRefresh = false) {
  const eventsList = document.getElementById('events-list');
  const dateKey = formatDateKey(currentDate);

  // Reset card selection
  resetCardSelection();
  updateDateDisplay();

  // Step 1: Check session cache first (instant) - unless forcing refresh
  if (!forceRefresh && sessionCache.has(dateKey)) {
    const cached = sessionCache.get(dateKey);
    displayEvents(cached.events);
    console.log(`✓ Instant load from session (${dateKey})`);

    // Only refresh if stale
    if (!sessionCache.isStale(dateKey)) {
      return; // Fresh enough, no need to query Outlook
    }
  }

  // Step 2: Check disk cache (fast) - unless forcing refresh
  if (!forceRefresh) {
    const cacheResult = await window.electronAPI.getCachedEvents(dateKey);

    if (cacheResult.success && Array.isArray(cacheResult.events) && cacheResult.fromCache) {
      // Show cached data instantly
      displayEvents(cacheResult.events);
      sessionCache.set(dateKey, cacheResult.events);
      console.log(`✓ Fast load from disk (${dateKey})`);

      // Only refresh if session cache says we should sync
      if (!sessionCache.shouldSyncOutlook()) {
        return;
      }
    } else {
      // No cache, show loading
      eventsList.innerHTML = '<div class="no-events">Loading events...</div>';
    }
  } else {
    // Force refresh: show loading immediately
    eventsList.innerHTML = '<div class="no-events">Refreshing...</div>';
  }

  // Step 3: Refresh from Outlook (only if cache is stale or missing)
  const startDate = new Date(currentDate);
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(currentDate);
  endDate.setHours(23, 59, 59, 999);

  setRefreshSpinning(true);
  try {
    const result = await window.electronAPI.getEvents(startDate, endDate);

    if (result.success) {
      displayEvents(result.events);

      // Update session cache
      sessionCache.set(dateKey, result.events);
      sessionCache.markOutlookSynced();

      // Update disk cache
      window.electronAPI.saveEventsCache(result.events, dateKey);

      console.log(`✓ Outlook sync (${dateKey})`);
    } else {
      // Only show error if we didn't have cached data
      if (!cacheResult.fromCache) {
        eventsList.innerHTML = '<div class="no-events">Error loading events: ' + result.error + '</div>';
      }
    }
  } catch (error) {
    console.error('Error refreshing events:', error);
    if (!cacheResult.fromCache) {
      eventsList.innerHTML = '<div class="no-events">Error: ' + error.message + '</div>';
    }
  } finally {
    setRefreshSpinning(false);
  }
}

function displayEvents(events) {
  const eventsList = document.getElementById('events-list');

  if (events.length === 0) {
    eventsList.innerHTML = '<div class="no-events">No meetings scheduled for this day</div>';
    resetCardSelection();
    if (highlightTimeout) clearTimeout(highlightTimeout);
    highlightEventAtHour = null; // Clear highlight target
    highlightAppliedAt = null;
    highlightTimeout = null;
    return;
  }

  eventsList.innerHTML = '';

  events.forEach(event => {
    const eventCard = createEventCard(event);
    eventsList.appendChild(eventCard);
  });

  // Update scroll indicators after adding cards
  setTimeout(() => checkScrollIndicators(), 100);

  // Update keyboard navigation array
  updateEventCardsArray();

  // Start / restart the per-minute countdown refresh
  startCountdownTimer();

  // Apply highlight if navigating from heatmap
  if (highlightEventAtHour !== null) {
    // Clear any pending highlight timeout to avoid applying to stale DOM
    if (highlightTimeout) {
      console.log(`🔄 Clearing pending highlight timeout (DOM rebuild)`);
      clearTimeout(highlightTimeout);
    }

    const isFirstApplication = highlightAppliedAt === null;
    const delay = isFirstApplication ? 200 : 0; // No delay for re-applications (from backgroundRefresh)

    console.log(`✨ Scheduling highlight for hour ${highlightEventAtHour} (delay: ${delay}ms, first: ${isFirstApplication})`);

    highlightTimeout = setTimeout(() => {
      highlightTimeout = null; // Clear the timeout reference

      if (isFirstApplication) {
        highlightAppliedAt = Date.now();
      }

      console.log(`🎯 Applying highlight to hour ${highlightEventAtHour}`);
      applyEventHighlight(highlightEventAtHour);

      // Clear highlight state after 5s (measured from first application)
      if (isFirstApplication) {
        setTimeout(() => {
          console.log(`⏰ Clearing highlight state after 5s`);
          highlightEventAtHour = null;
          highlightAppliedAt = null;
        }, 5000);
      }
    }, delay);
  }
}

// ===== COUNTDOWN HELPERS =====

function getCountdownText(start, end) {
  const now = new Date();
  const startDate = new Date(start);
  const endDate = new Date(end);

  if (now >= endDate) return { text: 'Done', cls: 'done' };
  if (now >= startDate) return { text: 'Now', cls: 'now' };

  const diffMins = Math.round((startDate - now) / 60000);
  const hours = Math.floor(diffMins / 60);
  const mins = diffMins % 60;

  let text;
  if (hours >= 24) {
    text = `Join in ${Math.floor(hours / 24)}d`;
  } else if (hours > 0) {
    text = mins > 0 ? `Join in ${hours}h ${mins}m` : `Join in ${hours}h`;
  } else {
    text = `Join in ${mins}m`;
  }

  return { text, cls: diffMins <= 15 ? 'soon' : 'upcoming' };
}

function startCountdownTimer() {
  if (countdownInterval) clearInterval(countdownInterval);
  countdownInterval = setInterval(updateAllCountdowns, 60000);
}

function updateAllCountdowns() {
  document.querySelectorAll('.event-card[data-event-start]').forEach(card => {
    const el = card.querySelector('.event-countdown');
    if (!el) return;
    const { text, cls } = getCountdownText(card.dataset.eventStart, card.dataset.eventEnd);
    el.textContent = text;
    el.className = `event-countdown event-countdown--${cls}`;
  });
}

function createEventCard(event) {
  const card = document.createElement('div');
  card.className = 'event-card';

  const state = getMeetingState(event);
  card.setAttribute('data-state', state);

  const startTime = new Date(event.start).toLocaleTimeString('en-US', {
    hour: 'numeric',
    minute: '2-digit',
    hour12: true
  });
  const endTime = new Date(event.end).toLocaleTimeString('en-US', {
    hour: 'numeric',
    minute: '2-digit',
    hour12: true
  });

  card.dataset.eventStart = event.start;
  card.dataset.eventEnd = event.end;

  const countdown = getCountdownText(event.start, event.end);

  card.innerHTML = `
    <div class="event-card-body">
      <div class="event-time-row">
        <div class="event-time">${startTime} – ${endTime}</div>
        <span class="event-countdown event-countdown--${countdown.cls}">${countdown.text}</span>
      </div>
      <div class="event-title">${escapeHtml(event.subject)}</div>
      ${(event.location || event.organizer) ? `
        <div class="event-meta">
          ${event.location ? `
            <span class="event-meta-item">
              <svg width="11" height="11" aria-hidden="true"><use href="#icon-location"/></svg>
              ${escapeHtml(event.location)}
            </span>
          ` : ''}
          ${event.organizer ? `
            <span class="event-meta-item">
              <svg width="11" height="11" aria-hidden="true"><use href="#icon-user"/></svg>
              ${escapeHtml(event.organizer)}
            </span>
          ` : ''}
        </div>
      ` : ''}
      ${event.teamsLink || event.accountName ? `
        <div class="event-teams-row">
          ${event.teamsLink ? `
            <a href="#" class="teams-join-link" data-teams-url="${escapeHtml(event.teamsLink)}">
              <svg width="12" height="12" aria-hidden="true"><use href="#icon-video"/></svg>
              Join Teams
            </a>
          ` : ''}
          ${event.accountName ? `
            <span class="event-account-name">${escapeHtml(event.accountName)}</span>
          ` : ''}
        </div>
      ` : ''}
    </div>
  `;

  card.addEventListener('click', () => openNotes(event));

  // Handle Teams link click - convert to msteams:// protocol and stop propagation
  if (event.teamsLink) {
    const teamsLink = card.querySelector('.teams-join-link');
    if (teamsLink) {
      teamsLink.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        openTeamsLink(event.teamsLink);
      });
    }
  }

  // Update selected index on mouse click
  card.addEventListener('mousedown', () => {
    updateEventCardsArray();
    selectedCardIndex = eventCards.indexOf(card);
    updateCardSelection();
  });

  return card;
}

function updateDateDisplay() {
  const options = { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' };
  const dateString = currentDate.toLocaleDateString('en-US', options);
  document.getElementById('current-date').textContent = dateString;
}

// ===== NOTES FUNCTIONS =====

async function openNotes(event) {
  currentMeetingId = event.id;
  currentMeetingData = event;

  document.getElementById('notes-meeting-title').textContent = event.subject;

  const metadata = document.getElementById('notes-metadata');
  const startTime = new Date(event.start).toLocaleString('en-US', {
    weekday: 'short',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    hour12: true
  });

  metadata.innerHTML = `
    <div><strong>Time:</strong> ${startTime}</div>
    ${event.location ? `<div><strong>Location:</strong> ${escapeHtml(event.location)}</div>` : ''}
    ${event.teamsLink ? `
      <div style="margin-top: 8px;">
        <a href="#" class="teams-join-link" data-teams-url="${escapeHtml(event.teamsLink)}">
          <svg width="12" height="12" aria-hidden="true"><use href="#icon-video"/></svg>
          Join Teams
        </a>
      </div>
    ` : ''}
    <div><strong>Organizer:</strong> ${escapeHtml(event.organizer)}</div>
    ${event.accountName ? `<div style="margin-top: 4px; opacity: 0.5;"><strong>Account:</strong> ${escapeHtml(event.accountName)}</div>` : ''}
  `;

  // Handle Teams link in metadata
  if (event.teamsLink) {
    const teamsMetaLink = metadata.querySelector('.teams-join-link');
    if (teamsMetaLink) {
      teamsMetaLink.addEventListener('click', (e) => {
        e.preventDefault();
        openTeamsLink(event.teamsLink);
      });
    }
  }

  const result = await window.electronAPI.getNote(event.id);
  document.getElementById('notes-editor').value = result.content || '';

  // Clear any existing auto-save timeout
  if (autoSaveTimeout) {
    clearTimeout(autoSaveTimeout);
    autoSaveTimeout = null;
  }

  // Clear status
  document.getElementById('notes-status').textContent = '';

  document.getElementById('calendar-content').classList.add('hidden');
  document.getElementById('notes-view').classList.remove('hidden');

  // Check scroll indicators for notes editor
  setTimeout(() => {
    const notesEditor = document.getElementById('notes-editor');
    updateScrollIndicators(notesEditor);
  }, 100);
}

function closeNotes() {
  // Clear any pending auto-save
  if (autoSaveTimeout) {
    clearTimeout(autoSaveTimeout);
    autoSaveTimeout = null;
  }

  document.getElementById('calendar-content').classList.remove('hidden');
  document.getElementById('notes-view').classList.add('hidden');
  currentMeetingId = null;
  currentMeetingData = null;
}

async function saveNotes() {
  if (!currentMeetingId || !currentMeetingData) return;

  // Cancel any pending auto-save
  if (autoSaveTimeout) {
    clearTimeout(autoSaveTimeout);
    autoSaveTimeout = null;
  }

  const noteContent = document.getElementById('notes-editor').value;
  const statusEl = document.getElementById('notes-status');

  isSaving = true;
  statusEl.textContent = 'Saving...';

  try {
    const result = await window.electronAPI.saveNote(
      currentMeetingId,
      noteContent,
      currentMeetingData
    );

    if (result.success) {
      statusEl.textContent = 'Saved';
      setTimeout(() => {
        if (statusEl.textContent === 'Saved') {
          statusEl.textContent = '';
        }
      }, 2000);
    } else {
      statusEl.textContent = 'Error: ' + result.error;
    }
  } catch (error) {
    statusEl.textContent = 'Error: ' + error.message;
  } finally {
    isSaving = false;
  }
}

// ===== ACTIONS FUNCTIONS =====

async function loadActions() {
  try {
    const result = await window.electronAPI.getActions();
    if (result.success) {
      actions = result.actions;
      if (currentTab === 'actions') {
        displayActions();
      }
    }
  } catch (error) {
    console.error('Error loading actions:', error);
  }
}

// Helper: Get breadcrumb path for an item (returns array of folder objects)
function getBreadcrumbPath(itemId) {
  const path = [];
  const visited = new Set(); // Prevent circular references
  let currentId = itemId;
  const item = actions.find(a => a.id === currentId);
  if (!item) return [];

  // Walk up the parent chain
  let parentId = item.parentId;
  while (parentId) {
    // Prevent infinite loops from circular references
    if (visited.has(parentId)) {
      console.error('Circular reference detected in folder hierarchy');
      break;
    }
    visited.add(parentId);

    const parent = actions.find(a => a.id === parentId);
    if (!parent) break;
    path.unshift({ title: parent.title, level: path.length });
    parentId = parent.parentId;

    // Safety: max 10 levels deep
    if (path.length >= 10) break;
  }

  return path;
}

// Helper: Render breadcrumb pills with increasing opacity
function renderBreadcrumbPills(breadcrumbPath) {
  if (!breadcrumbPath || breadcrumbPath.length === 0) return '';

  const pills = breadcrumbPath.map((folder, index) => {
    // Opacity increases with depth: 30%, 50%, 70%, 90%
    const opacity = Math.min(0.3 + (index * 0.2), 0.9);
    return `<span class="breadcrumb-pill" style="background: rgba(var(--accent-rgb), ${opacity});">${escapeHtml(folder.title)}</span>`;
  }).join('<svg class="breadcrumb-arrow" width="12" height="12" viewBox="0 0 24 24"><path d="M9 18l6-6-6-6" stroke="currentColor" fill="none" stroke-width="2"/></svg>');

  return `<div class="breadcrumb-container">${pills}</div>`;
}

// Helper: Sort items (pinned first, then by creation date)
function sortItems(items) {
  return [...items].sort((a, b) => {
    // Pinned items first
    if (a.pinned && !b.pinned) return -1;
    if (!a.pinned && b.pinned) return 1;

    // Then by creation date (newest first)
    return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
  });
}

// Render items recursively with hierarchy
function renderItemsRecursive(parentId, level = 0, visited = new Set()) {
  const fragment = document.createDocumentFragment();

  // Safety: prevent infinite recursion
  if (level > 10) {
    console.error('Maximum nesting depth reached');
    return fragment;
  }

  // Prevent circular references
  if (parentId && visited.has(parentId)) {
    console.error('Circular reference detected in folder hierarchy');
    return fragment;
  }

  if (parentId) visited.add(parentId);

  // Get children of this parent
  const children = actions.filter(a => a.parentId === parentId);

  // Sort children
  const sorted = sortItems(children);

  sorted.forEach(item => {
    if (item.type === 'folder') {
      const folderCard = createFolderCard(item, level);
      fragment.appendChild(folderCard);

      // If folder is expanded, render its children
      if (expandedFolders.has(item.id)) {
        const childrenContainer = document.createElement('div');
        childrenContainer.className = 'folder-children';
        const newVisited = new Set(visited); // Clone visited set for this branch
        childrenContainer.appendChild(renderItemsRecursive(item.id, level + 1, newVisited));
        fragment.appendChild(childrenContainer);
      }
    } else {
      const actionCard = createActionCard(item, level);
      fragment.appendChild(actionCard);
    }
  });

  return fragment;
}

function displayActions() {
  const actionsList = document.getElementById('actions-list');

  if (actions.length === 0) {
    actionsList.innerHTML = '<div class="no-actions">No actions or folders yet.<br>Click "+" for actions or "📁" for folders!</div>';
    return;
  }

  const q = actionSearchQuery.trim().toLowerCase();

  // If searching, show flat list with breadcrumbs
  if (q) {
    const filtered = actions.filter(a =>
      a.title.toLowerCase().includes(q) ||
      (a.note && a.note.toLowerCase().includes(q))
    );

    if (filtered.length === 0) {
      actionsList.innerHTML = '<div class="no-actions">No actions match your search.</div>';
      return;
    }

    actionsList.innerHTML = '';

    // Sort and display with breadcrumbs
    const sorted = sortItems(filtered);
    sorted.forEach(item => {
      const breadcrumbPath = getBreadcrumbPath(item.id);
      const breadcrumbHtml = renderBreadcrumbPills(breadcrumbPath);

      if (item.type === 'folder') {
        const folderCard = createFolderCard(item, 0, breadcrumbHtml);
        actionsList.appendChild(folderCard);
      } else {
        const actionCard = createActionCard(item, 0, breadcrumbHtml);
        actionsList.appendChild(actionCard);
      }
    });
  } else {
    // Normal hierarchical view
    actionsList.innerHTML = '';
    actionsList.appendChild(renderItemsRecursive(null, 0));
  }
}

function createFolderCard(folder, level = 0, breadcrumbHtml = '') {
  const card = document.createElement('div');
  card.className = 'folder-card';
  card.dataset.folderId = folder.id;
  card.dataset.level = level;

  const isExpanded = expandedFolders.has(folder.id);

  // Count children
  const childCount = actions.filter(a => a.parentId === folder.id).length;

  card.innerHTML = `
    <div class="folder-header">
      ${breadcrumbHtml ? breadcrumbHtml : ''}
      <div class="folder-header-content">
        <svg class="icon folder-expand-icon ${isExpanded ? 'expanded' : ''}" width="14" height="14" aria-hidden="true">
          <use href="#icon-chevron-right"/>
        </svg>
        <svg class="icon folder-icon" width="16" height="16" aria-hidden="true">
          <use href="#icon-folder"/>
        </svg>
        <div class="folder-title">${escapeHtml(folder.title)} (${childCount})</div>
        <div class="folder-buttons">
          <button class="folder-add-action-btn" aria-label="Add action to this folder" title="Add action here">
            <svg class="icon" width="14" height="14" aria-hidden="true"><use href="#icon-plus"/></svg>
          </button>
          <button class="folder-add-subfolder-btn" aria-label="Add subfolder" title="Add subfolder">
            <svg class="icon" width="14" height="14" aria-hidden="true"><use href="#icon-folder"/></svg>
          </button>
          <button class="folder-pin-btn${folder.pinned ? ' pinned' : ''}" aria-label="${folder.pinned ? 'Unpin' : 'Pin to top'}" title="${folder.pinned ? 'Unpin' : 'Pin to top'}">
            <svg class="icon" width="14" height="14" aria-hidden="true"><use href="#icon-pin"/></svg>
          </button>
          <button class="folder-delete-btn" aria-label="Delete folder" title="Delete folder">
            <svg class="icon" width="13" height="13" aria-hidden="true"><use href="#icon-x"/></svg>
          </button>
        </div>
      </div>
    </div>
  `;

  // Add event listeners
  const folderHeader = card.querySelector('.folder-header-content');
  const folderTitle = card.querySelector('.folder-title');
  const addActionBtn = card.querySelector('.folder-add-action-btn');
  const addSubfolderBtn = card.querySelector('.folder-add-subfolder-btn');
  const pinBtn = card.querySelector('.folder-pin-btn');
  const deleteBtn = card.querySelector('.folder-delete-btn');

  // Click to expand/collapse
  folderHeader.addEventListener('click', (e) => {
    // Don't expand if clicking on title or buttons
    if (!e.target.closest('button') && !e.target.closest('.folder-title')) {
      toggleFolderExpand(folder.id);
    }
  });

  // Double-click title to edit
  folderTitle.addEventListener('dblclick', (e) => {
    e.stopPropagation();
    makeEditableTitle(folderTitle, folder.id, 'folder');
  });

  // Add action to this folder
  addActionBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    currentExpandedFolder = folder.id; // Set context for creation
    openAddActionModal();
  });

  // Add subfolder
  addSubfolderBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    currentExpandedFolder = folder.id; // Set context for creation
    openAddFolderModal();
  });

  // Pin/unpin handler
  pinBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    pinAction(folder.id); // Reuse same function
  });

  // Delete folder handler
  deleteBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    deleteAction(folder.id); // Reuse same function (it has folder validation)
  });

  return card;
}

function createActionCard(action, level = 0, breadcrumbHtml = '') {
  const card = document.createElement('div');
  card.className = 'action-card';
  card.dataset.actionId = action.id;
  card.dataset.level = level;

  // Check if this action is expanded
  if (expandedActions.has(action.id)) {
    card.classList.add('expanded');
  }

  const hasUrl = action.url && action.url.trim();
  const hasPath = action.path && action.path.trim();

  card.innerHTML = `
    ${breadcrumbHtml ? breadcrumbHtml : ''}
    <div class="action-header">
      <div class="action-content">
        <div class="action-title">${escapeHtml(action.title)}</div>
      </div>
      <div class="action-buttons">
        <button class="action-save-btn" aria-label="Save changes" title="Save">
          <svg class="icon" width="14" height="14" aria-hidden="true"><use href="#icon-save"/></svg>
        </button>
        <button class="action-pin-btn${action.pinned ? ' pinned' : ''}" aria-label="${action.pinned ? 'Unpin' : 'Pin to top'}" title="${action.pinned ? 'Unpin' : 'Pin to top'}">
          <svg class="icon" width="14" height="14" aria-hidden="true"><use href="#icon-pin"/></svg>
        </button>
        ${hasUrl || hasPath ? '<button class="action-open-btn">Open</button>' : ''}
        <button class="action-delete-btn" aria-label="Delete action" title="Delete action">
          <svg class="icon" width="13" height="13" aria-hidden="true"><use href="#icon-x"/></svg>
        </button>
      </div>
    </div>
    <div class="action-details">
      <div class="form-group">
        <label>
          <svg class="icon field-icon" width="12" height="12" aria-hidden="true"><use href="#icon-link"/></svg>
          URL
        </label>
        <input type="url" class="action-url-input" placeholder="https://..." value="${escapeHtml(action.url || '')}">
      </div>
      <div class="form-group">
        <label>
          <svg class="icon field-icon" width="12" height="12" aria-hidden="true"><use href="#icon-file"/></svg>
          File Path
        </label>
        <input type="text" class="action-path-input" placeholder="C:\\Users\\..." value="${escapeHtml(action.path || '')}">
        <span class="path-validation-msg"></span>
      </div>
      <div class="form-group">
        <label>Note</label>
        <textarea class="action-note" placeholder="Add notes about this action..." rows="3">${escapeHtml(action.note || '')}</textarea>
      </div>
    </div>
  `;

  // Add event listeners (not using onclick to properly handle events)
  const saveBtn = card.querySelector('.action-save-btn');
  const pinBtn = card.querySelector('.action-pin-btn');
  const openBtn = card.querySelector('.action-open-btn');
  const deleteBtn = card.querySelector('.action-delete-btn');
  const noteArea = card.querySelector('.action-note');
  const urlInput = card.querySelector('.action-url-input');
  const pathInput = card.querySelector('.action-path-input');
  const pathValidationMsg = card.querySelector('.path-validation-msg');
  const actionHeader = card.querySelector('.action-header');
  const actionTitle = card.querySelector('.action-title');

  // Click anywhere on card header to expand/collapse
  actionHeader.addEventListener('click', (e) => {
    // Only expand if not clicking on a button or breadcrumb
    if (!e.target.closest('button') && !e.target.closest('.item-breadcrumb') && !e.target.closest('.action-title')) {
      toggleActionExpand(action.id, card);
    }
  });

  // Double-click title to edit
  actionTitle.addEventListener('dblclick', (e) => {
    e.stopPropagation();
    makeEditableTitle(actionTitle, action.id, 'action');
  });

  // Save button - manually save all fields immediately
  saveBtn.addEventListener('click', async (e) => {
    e.stopPropagation();

    // Save all fields immediately (bypass debounce)
    action.url = urlInput.value.trim();
    action.path = pathInput.value.trim();
    action.note = noteArea.value.trim();

    const result = await window.electronAPI.saveActions(actions);

    // Visual feedback
    if (result.success) {
      saveBtn.style.color = '#10b981'; // Green
      saveBtn.style.transform = 'scale(1.1)';
      setTimeout(() => {
        saveBtn.style.color = '';
        saveBtn.style.transform = '';
      }, 500);
    } else {
      saveBtn.style.color = '#ef4444'; // Red
      setTimeout(() => {
        saveBtn.style.color = '';
      }, 1000);
    }
  });

  // Pin/unpin handler
  pinBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    pinAction(action.id);
  });

  // Open action handler - handles both URL and path
  if (openBtn) {
    openBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const currentUrl = urlInput.value.trim();
      const currentPath = pathInput.value.trim();

      if (currentUrl) {
        openAction(currentUrl);
      } else if (currentPath) {
        openAction(currentPath);
      }
    });
  }

  // Delete action handler with proper error handling
  deleteBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    deleteAction(action.id);
  });

  // URL/Path mutual exclusivity
  urlInput.addEventListener('input', (e) => {
    e.stopPropagation();
    if (e.target.value.trim()) {
      pathInput.disabled = true;
    } else {
      pathInput.disabled = false;
    }
  });

  pathInput.addEventListener('input', (e) => {
    e.stopPropagation();
    if (e.target.value.trim()) {
      urlInput.disabled = true;
    } else {
      urlInput.disabled = false;
    }
  });

  // Set initial disabled state
  if (urlInput.value.trim()) {
    pathInput.disabled = true;
  } else if (pathInput.value.trim()) {
    urlInput.disabled = true;
  }

  // Save URL on change (debounced)
  let urlTimeout;
  urlInput.addEventListener('input', (e) => {
    e.stopPropagation();
    clearTimeout(urlTimeout);
    urlTimeout = setTimeout(() => {
      saveActionField(action.id, 'url', e.target.value);
    }, 1000);
  });

  // Save path on change (debounced) with validation
  let pathTimeout;
  pathInput.addEventListener('input', (e) => {
    e.stopPropagation();

    // Strip quotes from pasted paths (Windows File Explorer adds quotes)
    let pathValue = e.target.value.trim();
    if ((pathValue.startsWith('"') && pathValue.endsWith('"')) ||
        (pathValue.startsWith("'") && pathValue.endsWith("'"))) {
      pathValue = pathValue.slice(1, -1);
      e.target.value = pathValue;
    }

    clearTimeout(pathTimeout);
    pathTimeout = setTimeout(async () => {
      saveActionField(action.id, 'path', pathValue);

      // Validate path if not empty
      if (pathValue) {
        const validation = await window.electronAPI.validatePath(pathValue);
        if (validation.success && !validation.exists) {
          pathValidationMsg.textContent = '⚠ Path does not exist';
          pathValidationMsg.style.color = '#f59e0b';
        } else {
          pathValidationMsg.textContent = '';
        }
      } else {
        pathValidationMsg.textContent = '';
      }
    }, 1000);
  });

  // Save note on change (debounced)
  let noteTimeout;
  noteArea.addEventListener('input', (e) => {
    e.stopPropagation();
    clearTimeout(noteTimeout);
    noteTimeout = setTimeout(() => {
      saveActionField(action.id, 'note', e.target.value);
    }, 1000);
  });

  return card;
}

function openAddActionModal() {
  document.getElementById('action-modal').classList.remove('hidden');
  document.getElementById('action-title-input').value = '';
  document.getElementById('action-url-input').value = '';
  document.getElementById('action-path-input').value = '';
  const noteInput = document.getElementById('action-note-input');
  if (noteInput) noteInput.value = '';

  // Re-enable both fields
  document.getElementById('action-url-input').disabled = false;
  document.getElementById('action-path-input').disabled = false;
  document.getElementById('action-url-input').placeholder = 'https://...';
  document.getElementById('action-path-input').placeholder = 'C:\\Users\\...';
  document.getElementById('path-validation-msg').textContent = '';

  document.getElementById('action-title-input').focus();
}

function closeAddActionModal() {
  document.getElementById('action-modal').classList.add('hidden');
}

async function handleAddAction(e) {
  e.preventDefault();

  const title = document.getElementById('action-title-input').value.trim();
  const url = document.getElementById('action-url-input').value.trim();
  const path = document.getElementById('action-path-input').value.trim();
  const note = document.getElementById('action-note-input')?.value.trim() || '';

  if (!title) {
    alert('Please enter a title');
    return;
  }

  // Smart creation: use currentExpandedFolder if available
  const parentId = currentExpandedFolder;

  try {
    const result = await window.electronAPI.addAction({ title, url, path, note, parentId });

    if (result.success) {
      actions.push(result.action);
      displayActions();
      closeAddActionModal();
    } else {
      alert('Error adding action: ' + result.error);
    }
  } catch (error) {
    alert('Error: ' + error.message);
  }
}

function openAddFolderModal() {
  const modal = document.getElementById('folder-modal');
  const input = document.getElementById('folder-title-input');

  modal.classList.remove('hidden');
  input.value = '';

  // Delay focus slightly to ensure modal is fully rendered
  setTimeout(() => {
    input.focus();
  }, 100);
}

function closeAddFolderModal() {
  document.getElementById('folder-modal').classList.add('hidden');
}

async function handleAddFolder(e) {
  e.preventDefault();

  const title = document.getElementById('folder-title-input').value.trim();

  if (!title) {
    alert('Please enter a folder name');
    return;
  }

  // Smart creation: use currentExpandedFolder if available
  const parentId = currentExpandedFolder;

  try {
    const result = await window.electronAPI.addFolder({ title, parentId });

    if (result.success) {
      actions.push(result.folder);
      displayActions();
      closeAddFolderModal();
    } else {
      alert('Error creating folder: ' + result.error);
    }
  } catch (error) {
    alert('Error: ' + error.message);
  }
}

async function deleteAction(actionId) {
  const item = actions.find(a => a.id === actionId);
  const itemType = item?.type === 'folder' ? 'folder' : 'action';

  if (!confirm(`Delete this ${itemType}?`)) return;

  try {
    const result = await window.electronAPI.deleteAction(actionId);

    if (result.success) {
      // Clean up expanded state
      expandedActions.delete(actionId);
      expandedFolders.delete(actionId);

      // Clear currentExpandedFolder if we're deleting it
      if (currentExpandedFolder === actionId) {
        currentExpandedFolder = null;
      }

      // Remove from actions array
      actions = actions.filter(a => a.id !== actionId);

      // Re-render the list
      displayActions();
    } else {
      alert(`Error deleting ${itemType}: ` + result.error);
    }
  } catch (error) {
    console.error(`Error deleting ${itemType}:`, error);
    alert('Error: ' + error.message);
  }
}

function openAction(url) {
  if (!url || !url.trim()) {
    alert('No URL or path specified for this action');
    return;
  }
  window.electronAPI.openExternal(url.trim());
}

// Toggle action card expand/collapse
function toggleActionExpand(actionId, cardElement) {
  if (expandedActions.has(actionId)) {
    expandedActions.delete(actionId);
    cardElement.classList.remove('expanded');
  } else {
    expandedActions.add(actionId);
    cardElement.classList.add('expanded');
  }
}

// Toggle folder expand/collapse
function toggleFolderExpand(folderId) {
  if (expandedFolders.has(folderId)) {
    expandedFolders.delete(folderId);
    // If closing the current expanded folder, clear it
    if (currentExpandedFolder === folderId) {
      currentExpandedFolder = null;
    }
  } else {
    expandedFolders.add(folderId);
    // Set as current expanded folder for smart creation
    currentExpandedFolder = folderId;
  }

  // Re-render to show/hide children
  displayActions();
}

// Clear currentExpandedFolder when creating at root level
function clearExpandedFolderContext() {
  currentExpandedFolder = null;
}

// Pin / unpin action and re-sort
async function pinAction(actionId) {
  const action = actions.find(a => a.id === actionId);
  if (!action) return;

  const previousState = action.pinned;
  action.pinned = !action.pinned;

  try {
    const result = await window.electronAPI.saveActions(actions);
    if (!result.success) {
      // Revert on failure
      action.pinned = previousState;
      console.error('Failed to save pin state:', result.error);
    }
  } catch (error) {
    // Revert on error
    action.pinned = previousState;
    console.error('Error pinning action:', error);
  }

  displayActions();
}

// Save action field (unified function for url, path, note)
async function saveActionField(actionId, field, value) {
  try {
    // Find the action in the array
    const action = actions.find(a => a.id === actionId);
    if (!action) return;

    // Update the field
    action[field] = value;

    // Save to backend
    const result = await window.electronAPI.saveActions(actions);

    if (!result.success) {
      console.error(`Error saving action ${field}:`, result.error);
    }

    // Update Open button visibility without full re-render
    if (field === 'url' || field === 'path') {
      updateOpenButtonVisibility(actionId);
    }
  } catch (error) {
    console.error(`Error saving action ${field}:`, error);
  }
}

// Make title editable inline (for both actions and folders)
function makeEditableTitle(titleElement, itemId, itemType) {
  const originalText = titleElement.textContent.trim();

  // Store original for cancel
  titleElement.dataset.originalText = originalText;

  // Make editable
  titleElement.contentEditable = true;
  titleElement.classList.add('editing');
  titleElement.focus();

  // Select all text
  const range = document.createRange();
  range.selectNodeContents(titleElement);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);

  // Save handler
  const saveEdit = async () => {
    const newText = titleElement.textContent.trim();

    // Enforce 50 character limit
    if (newText.length > 50) {
      titleElement.textContent = newText.substring(0, 50);
      return; // Don't save, just truncate
    }

    if (newText && newText !== originalText) {
      // Update the item
      const item = actions.find(a => a.id === itemId);
      if (item) {
        item.title = newText;

        // Save to backend
        try {
          const result = await window.electronAPI.saveActions(actions);
          if (!result.success) {
            console.error('Error saving title:', result.error);
            titleElement.textContent = originalText; // Revert
          }
        } catch (error) {
          console.error('Error saving title:', error);
          titleElement.textContent = originalText; // Revert
        }
      }
    } else if (!newText) {
      // Empty title, revert
      titleElement.textContent = originalText;
    }

    // Clean up
    titleElement.contentEditable = false;
    titleElement.classList.remove('editing');
    delete titleElement.dataset.originalText;
  };

  // Cancel handler
  const cancelEdit = () => {
    titleElement.textContent = originalText;
    titleElement.contentEditable = false;
    titleElement.classList.remove('editing');
    delete titleElement.dataset.originalText;
  };

  // Listen for blur (save)
  const blurHandler = () => {
    saveEdit();
  };

  // Listen for Enter (save) and Escape (cancel)
  const keyHandler = (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      saveEdit();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      cancelEdit();
    } else if (titleElement.textContent.length >= 50 && e.key !== 'Backspace' && e.key !== 'Delete' && !e.ctrlKey && !e.metaKey) {
      // Prevent typing beyond 50 characters (allow Ctrl/Cmd shortcuts)
      e.preventDefault();
    }
  };

  // Handle input event to enforce limit on paste
  const inputHandler = () => {
    if (titleElement.textContent.length > 50) {
      const cursorPos = window.getSelection().getRangeAt(0).startOffset;
      titleElement.textContent = titleElement.textContent.substring(0, 50);

      // Restore cursor position
      const range = document.createRange();
      const sel = window.getSelection();
      try {
        range.setStart(titleElement.childNodes[0], Math.min(cursorPos, 50));
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
      } catch (e) {
        // Cursor positioning failed, just focus at end
        titleElement.focus();
      }
    }
  };

  titleElement.addEventListener('blur', blurHandler, { once: true });
  titleElement.addEventListener('keydown', keyHandler);
  titleElement.addEventListener('input', inputHandler);

  // Clean up listeners after blur
  titleElement.addEventListener('blur', () => {
    titleElement.removeEventListener('keydown', keyHandler);
    titleElement.removeEventListener('input', inputHandler);
  }, { once: true });
}

// Update Open button visibility for a specific action without re-rendering
function updateOpenButtonVisibility(actionId) {
  const action = actions.find(a => a.id === actionId);
  if (!action) return;

  const card = document.querySelector(`.action-card[data-action-id="${actionId}"]`);
  if (!card) return;

  const hasUrl = action.url && action.url.trim();
  const hasPath = action.path && action.path.trim();
  const openBtn = card.querySelector('.action-open-btn');

  if (hasUrl || hasPath) {
    if (!openBtn) {
      // Create Open button if it doesn't exist
      const newBtn = document.createElement('button');
      newBtn.className = 'action-open-btn';
      newBtn.textContent = 'Open';
      const actionButtons = card.querySelector('.action-buttons');
      const pinBtn = actionButtons.querySelector('.action-pin-btn');
      actionButtons.insertBefore(newBtn, pinBtn.nextSibling);

      // Add click handler
      newBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        const currentUrl = card.querySelector('.action-url-input').value.trim();
        const currentPath = card.querySelector('.action-path-input').value.trim();
        if (currentUrl) {
          openAction(currentUrl);
        } else if (currentPath) {
          openAction(currentPath);
        }
      });
    }
  } else {
    // Remove Open button if URL and path are both empty
    if (openBtn) {
      openBtn.remove();
    }
  }
}

// Open Teams meeting link - converts to msteams:// protocol
function openTeamsLink(teamsUrl) {
  try {
    let finalUrl = teamsUrl;

    // Convert https://teams.microsoft.com/ URLs to msteams:// protocol
    if (teamsUrl.startsWith('https://teams.microsoft.com/')) {
      finalUrl = teamsUrl.replace('https://', 'msteams://');
    }

    // Open via OS — launches Teams app or falls back to browser
    window.electronAPI.openExternal(finalUrl);
  } catch (error) {
    console.error('Error opening Teams link:', error);
    window.electronAPI.openExternal(teamsUrl);
  }
}

// ===== WEEK VIEW / HEATMAP FUNCTIONS =====

// Initialize week view - set to current week and build grid
async function initWeekView() {
  if (!currentWeekStart) {
    // Set to Monday of current week
    currentWeekStart = getMondayOfWeek(new Date());
  }
  updateWeekDisplay();
  buildHeatmapGrid();
  await populateHeatmap();
}

// Get Monday of the week for a given date
function getMondayOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1); // Adjust when day is Sunday
  const monday = new Date(d.setDate(diff));
  monday.setHours(0, 0, 0, 0);
  return monday;
}

// Navigate forward or backward by weeks
async function navigateWeek(direction) {
  const newWeekStart = new Date(currentWeekStart);
  newWeekStart.setDate(newWeekStart.getDate() + (direction * 7));
  currentWeekStart = newWeekStart;
  updateWeekDisplay();
  buildHeatmapGrid(); // Rebuild grid to update current hour highlighting

  // Show loading state briefly
  setRefreshSpinning(true);
  await populateHeatmap();
  setRefreshSpinning(false);
}

// Update the week display header
function updateWeekDisplay() {
  const weekEnd = new Date(currentWeekStart);
  weekEnd.setDate(weekEnd.getDate() + 6);

  const startStr = currentWeekStart.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  const endStr = weekEnd.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

  document.getElementById('current-week').textContent = `Week of ${startStr} - ${endStr}`;
}

// Build the heatmap grid structure
function buildHeatmapGrid() {
  const grid = document.getElementById('heatmap-grid');
  grid.innerHTML = '';

  // Days of week headers (Mon - Sun)
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Empty corner cell
  const cornerCell = document.createElement('div');
  cornerCell.style.gridColumn = '1';
  cornerCell.style.gridRow = '1';
  grid.appendChild(cornerCell);

  // Day headers
  days.forEach((day, index) => {
    const dayHeader = document.createElement('div');
    dayHeader.className = 'heatmap-day-header';
    dayHeader.style.gridColumn = index + 2;
    dayHeader.textContent = day;

    // Check if this day is today
    const dayDate = new Date(currentWeekStart);
    dayDate.setDate(dayDate.getDate() + index);
    if (dayDate.getTime() === today.getTime()) {
      dayHeader.classList.add('today');
    }

    grid.appendChild(dayHeader);
  });

  // Time labels and blocks (8am - 5pm = 10 hours)
  const hours = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17]; // 8am to 5pm

  hours.forEach((hour, rowIndex) => {
    // Time label
    const timeLabel = document.createElement('div');
    timeLabel.className = 'heatmap-time-label';
    timeLabel.style.gridRow = rowIndex + 2;
    const isPM = hour >= 12;
    const displayHour = hour > 12 ? hour - 12 : hour;
    timeLabel.textContent = `${displayHour}${isPM ? 'p' : 'a'}`;
    grid.appendChild(timeLabel);

    // Blocks for each day
    days.forEach((day, colIndex) => {
      const block = document.createElement('div');
      block.className = 'heatmap-block empty';
      block.style.gridRow = rowIndex + 2;
      block.style.gridColumn = colIndex + 2;
      block.tabIndex = 0;
      block.setAttribute('role', 'button');

      // Calculate the date for this block
      const blockDate = new Date(currentWeekStart);
      blockDate.setDate(blockDate.getDate() + colIndex);
      blockDate.setHours(hour, 0, 0, 0);

      // Store date in dataset
      block.dataset.date = blockDate.toISOString();
      block.dataset.hour = hour;

      // Check if this is the current hour today
      const now = new Date();
      if (
        blockDate.getDate() === now.getDate() &&
        blockDate.getMonth() === now.getMonth() &&
        blockDate.getFullYear() === now.getFullYear() &&
        hour === now.getHours()
      ) {
        block.classList.add('current-hour');
      }

      // Check if this is a past time
      if (blockDate < now) {
        block.classList.add('past');
      }

      // Add click handler
      block.addEventListener('click', () => handleHeatmapBlockClick(blockDate));

      grid.appendChild(block);
    });
  });
}

// Populate heatmap with meeting data from session/disk cache (lazy Outlook query)
async function populateHeatmap() {
  const days = 7;
  const hours = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17];

  // Clear existing meeting classes from all blocks
  document.querySelectorAll('.heatmap-block').forEach(block => {
    block.classList.remove('has-meeting', 'short-meeting');
    if (!block.classList.contains('current-hour')) {
      block.classList.add('empty');
    }
  });

  const weekStart = new Date(currentWeekStart);
  weekStart.setHours(0, 0, 0, 0);
  const weekEnd = new Date(currentWeekStart);
  weekEnd.setDate(weekEnd.getDate() + 6);
  weekEnd.setHours(23, 59, 59, 999);

  // Collect all events for the week from session/disk cache
  let allEvents = [];
  let needsOutlookSync = false;

  // Check each day in session cache first
  for (let dayOffset = 0; dayOffset < days; dayOffset++) {
    const date = new Date(currentWeekStart);
    date.setDate(date.getDate() + dayOffset);
    const dateKey = formatDateKey(date);

    if (sessionCache.has(dateKey)) {
      const cached = sessionCache.get(dateKey);
      allEvents.push(...cached.events);
      if (sessionCache.isStale(dateKey)) {
        needsOutlookSync = true;
      }
    } else {
      // Try disk cache
      const cacheResult = await window.electronAPI.getCachedEvents(dateKey);
      if (cacheResult.success && Array.isArray(cacheResult.events) && cacheResult.fromCache) {
        allEvents.push(...cacheResult.events);
        sessionCache.set(dateKey, cacheResult.events);
      }
      needsOutlookSync = true; // Missing from session, need to sync
    }
  }

  // Display cached events immediately
  if (allEvents.length > 0) {
    console.log(`✓ Week view: ${allEvents.length} events from cache`);
    renderHeatmapEvents(allEvents, weekStart, hours);
  }

  // Only query Outlook if cache is stale/missing AND we should sync
  if (needsOutlookSync && sessionCache.shouldSyncOutlook()) {
    console.log(`⟳ Week view: syncing with Outlook...`);

    const result = await window.electronAPI.getEvents(weekStart, weekEnd);

    if (result.success && Array.isArray(result.events)) {
      // Cache the events by day
      const eventsByDay = {};
      result.events.forEach(event => {
        const eventDate = new Date(event.start);
        const dateKey = formatDateKey(eventDate);
        if (!eventsByDay[dateKey]) {
          eventsByDay[dateKey] = [];
        }
        eventsByDay[dateKey].push(event);
      });

      // Update session cache
      Object.keys(eventsByDay).forEach(dateKey => {
        sessionCache.set(dateKey, eventsByDay[dateKey]);
      });
      sessionCache.markOutlookSynced();

      // Save to disk
      await window.electronAPI.saveEventsRangeCache(eventsByDay);

      // Re-render with fresh data
      renderHeatmapEvents(result.events, weekStart, hours);
      console.log(`✓ Week view: synced ${result.events.length} events`);
    }
  }
}

// Helper: render events on heatmap grid
function renderHeatmapEvents(events, weekStart, hours) {
  events.forEach(event => {
    const startTime = new Date(event.start);
    const endTime = new Date(event.end);
    const durationMinutes = (endTime - startTime) / 60000;

    // Calculate which day of the week this is
    const daysDiff = Math.floor((startTime - weekStart) / (1000 * 60 * 60 * 24));
    if (daysDiff < 0 || daysDiff > 6) return;

    // Mark the hour blocks this meeting occupies (8am-5pm)
    hours.forEach(hour => {
      const blockStart = new Date(startTime);
      blockStart.setHours(hour, 0, 0, 0);
      const blockEnd = new Date(blockStart);
      blockEnd.setHours(hour + 1, 0, 0, 0);

      const meetingOverlapsBlock = startTime < blockEnd && endTime > blockStart;

      if (meetingOverlapsBlock) {
        const dateKey = formatDateKey(startTime);
        const block = document.querySelector(
          `.heatmap-block[data-hour="${hour}"][data-date^="${dateKey}"]`
        );

        if (block) {
          const isShortMeeting = durationMinutes < 60;

          if (isShortMeeting) {
            if (!block.classList.contains('has-meeting')) {
              block.classList.remove('empty');
              block.classList.add('short-meeting');
            }
          } else {
            block.classList.remove('empty', 'short-meeting');
            block.classList.add('has-meeting');
          }
        }
      }
    });
  });
}

// Handle clicking on a heatmap block - jump to that day in calendar
async function handleHeatmapBlockClick(blockDate) {
  // Store the clicked hour for highlighting
  highlightEventAtHour = blockDate.getHours();

  // Set the current date to the clicked block's date
  currentDate = new Date(blockDate);
  currentDate.setHours(0, 0, 0, 0);

  // Switch to calendar tab
  switchTab('calendar');

  // Load events for that day (displayEvents will handle highlighting)
  await loadEvents();
}

// Apply highlight to event at specific hour
function applyEventHighlight(clickedHour) {
  const eventCards = document.querySelectorAll('.event-card');
  let foundEvent = false;

  console.log(`🔍 Searching ${eventCards.length} cards for event at hour ${clickedHour}`);

  eventCards.forEach(card => {
    const eventStart = card.dataset.eventStart;
    if (eventStart) {
      const startDate = new Date(eventStart);
      const eventEnd = card.dataset.eventEnd;
      const endDate = new Date(eventEnd);

      // Use same overlap logic as heatmap: event overlaps the clicked hour block
      // (handles sub-hour meetings where startHour === endHour)
      const startMins = startDate.getHours() * 60 + startDate.getMinutes();
      const endMins = endDate.getHours() * 60 + endDate.getMinutes();
      const blockStartMins = clickedHour * 60;
      const blockEndMins = (clickedHour + 1) * 60;

      if (startMins < blockEndMins && endMins > blockStartMins) {
        // Add glow effect
        card.classList.add('event-glow');
        foundEvent = true;

        console.log(`✅ Found event: "${card.querySelector('.event-title')?.textContent}" (${Math.floor(startMins/60)}:${String(startMins%60).padStart(2,'0')}-${Math.floor(endMins/60)}:${String(endMins%60).padStart(2,'0')}`);

        // Scroll to the event
        card.scrollIntoView({ behavior: 'smooth', block: 'center' });

        // Remove glow after 5 seconds
        setTimeout(() => {
          console.log(`🌟 Removing glow from event`);
          card.classList.remove('event-glow');
        }, 5000);
      }
    }
  });

  if (!foundEvent) {
    console.log(`❌ No event found for hour ${clickedHour}`);
  }
}

// ===== THEME FUNCTIONS =====

function loadTheme() {
  const savedTheme = localStorage.getItem('calendar-theme') || 'ocean-blue';
  setTheme(savedTheme, false); // Don't save on initial load
}

function setTheme(themeName, save = true) {
  currentTheme = themeName;
  document.documentElement.setAttribute('data-theme', themeName);

  // Update active state in picker
  document.querySelectorAll('.theme-option').forEach(opt => {
    opt.classList.toggle('active', opt.dataset.theme === themeName);
  });

  // Save to localStorage
  if (save) {
    localStorage.setItem('calendar-theme', themeName);
  }
}

function toggleThemePicker() {
  const picker = document.getElementById('theme-picker');
  picker.classList.toggle('hidden');
}

// ===== TAB ORDER FUNCTIONS =====

function loadTabOrder() {
  const saved = localStorage.getItem('calendar-tab-order');
  if (saved) {
    try {
      const parsed = JSON.parse(saved);
      const validTabs = ['calendar', 'weekview', 'actions'];
      if (Array.isArray(parsed) && parsed.length === 3 &&
          validTabs.every(t => parsed.includes(t))) {
        tabOrder = parsed;
      }
    } catch (e) {
      // Invalid data — keep default order
    }
  }
  applyTabOrder();
}

function applyTabOrder() {
  const tabNav = document.querySelector('.tab-nav');
  const refreshBtn = document.getElementById('refresh-btn');
  tabOrder.forEach(tabName => {
    const btn = document.getElementById(`${tabName}-tab`);
    if (btn) tabNav.insertBefore(btn, refreshBtn);
  });
}

function saveTabOrder() {
  const tabNav = document.querySelector('.tab-nav');
  const buttons = tabNav.querySelectorAll('.tab-btn:not(.tab-btn-icon-only)');
  tabOrder = Array.from(buttons).map(btn => btn.id.replace('-tab', ''));
  localStorage.setItem('calendar-tab-order', JSON.stringify(tabOrder));
}

function initTabDragAndDrop() {
  const tabNav = document.querySelector('.tab-nav');
  const getDraggableTabs = () => tabNav.querySelectorAll('.tab-btn:not(.tab-btn-icon-only)');
  let draggedTab = null;

  function clearDragIndicators() {
    getDraggableTabs().forEach(btn => btn.classList.remove('drag-over-left', 'drag-over-right'));
  }

  getDraggableTabs().forEach(tab => {
    tab.setAttribute('draggable', 'true');

    tab.addEventListener('dragstart', (e) => {
      draggedTab = tab;
      tab.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', tab.id);
    });

    tab.addEventListener('dragend', () => {
      tab.classList.remove('dragging');
      clearDragIndicators();
      draggedTab = null;
    });

    tab.addEventListener('dragover', (e) => {
      e.preventDefault();
      if (!draggedTab || draggedTab === tab) return;
      e.dataTransfer.dropEffect = 'move';
      clearDragIndicators();
      const rect = tab.getBoundingClientRect();
      if (e.clientX < rect.left + rect.width / 2) {
        tab.classList.add('drag-over-left');
      } else {
        tab.classList.add('drag-over-right');
      }
    });

    tab.addEventListener('dragleave', (e) => {
      // Only clear if leaving to outside the tab nav
      if (!tabNav.contains(e.relatedTarget)) {
        clearDragIndicators();
      }
      tab.classList.remove('drag-over-left', 'drag-over-right');
    });

    tab.addEventListener('drop', (e) => {
      e.preventDefault();
      if (!draggedTab || draggedTab === tab) return;
      const rect = tab.getBoundingClientRect();
      const insertBefore = e.clientX < rect.left + rect.width / 2;
      if (insertBefore) {
        tabNav.insertBefore(draggedTab, tab);
      } else {
        tabNav.insertBefore(draggedTab, tab.nextElementSibling);
      }
      clearDragIndicators();
      saveTabOrder();
    });
  });
}

// ===== UTILITY FUNCTIONS =====

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// Make functions globally available for onclick handlers
window.openAction = openAction;
window.deleteAction = deleteAction;

// Initialize on load
init();
