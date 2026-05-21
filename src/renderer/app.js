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

// Countdown timer
let countdownInterval = null;

// Week view state
let currentWeekStart = null; // Will be set to Monday of current week
let highlightEventAtHour = null; // Hour to highlight when navigating from heatmap

// Initialize the app
async function init() {
  // Setup listeners synchronously (fast)
  setupEventListeners();
  setupKeyboardNavigation();
  setupScrollDetection();
  setupAutoSave();

  // Don't await - let it run in background
  fastStartup();
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

  // Show loading state immediately
  eventsList.innerHTML = '<div class="no-events">Loading events…</div>';

  // Try cache first (synchronous fast path)
  let hasCacheForToday = false;
  try {
    const cacheResult = await window.electronAPI.getCachedEvents(dateKey);
    if (cacheResult.success && Array.isArray(cacheResult.events) && cacheResult.fromCache) {
      displayEvents(cacheResult.events);
      hasCacheForToday = true;
    }
  } catch (error) {
    console.error('Cache read error:', error);
  }

  // Load actions in parallel (don't block event display)
  window.electronAPI.getActions().then(actionsResult => {
    if (actionsResult.success) {
      actions = actionsResult.actions;
    }
  }).catch(err => console.error('Actions load error:', err));

  // Background: refresh today, then pre-fetch surrounding days
  backgroundRefresh(dateKey, hasCacheForToday);
}

async function backgroundRefresh(dateKey, hasCachedData) {
  setRefreshSpinning(true);

  const startDate = new Date(currentDate);
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(currentDate);
  endDate.setHours(23, 59, 59, 999);

  try {
    const result = await window.electronAPI.checkAndGetEvents(startDate, endDate);

    if (result.success && result.available) {
      displayEvents(result.events);
      // Cache today's fresh data
      window.electronAPI.saveEventsCache(result.events, dateKey);

      // Pre-fetch surrounding ±5 days in background (non-blocking)
      backgroundPrefetchRange(currentDate);
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
    console.log('Pre-fetching current week + next 2 weeks for instant navigation...');

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

      // Save to cache (fire-and-forget)
      await window.electronAPI.saveEventsRangeCache(eventsMap);
      console.log(`Cached ${Object.keys(eventsMap).length} days for instant navigation`);
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
    currentDate.setDate(currentDate.getDate() - 1);
    loadEvents();
  });

  document.getElementById('next-day').addEventListener('click', () => {
    currentDate.setDate(currentDate.getDate() + 1);
    loadEvents();
  });

  document.getElementById('refresh-btn').addEventListener('click', async () => {
    setRefreshSpinning(true);

    if (currentTab === 'weekview') {
      // Refresh heatmap
      await populateHeatmap();
    } else {
      // Refresh calendar
      await loadEvents();
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
  document.getElementById('add-action-btn').addEventListener('click', openAddActionModal);
  document.getElementById('cancel-action-btn').addEventListener('click', closeAddActionModal);
  document.getElementById('action-form').addEventListener('submit', handleAddAction);

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
}

// ===== KEYBOARD NAVIGATION =====

function setupKeyboardNavigation() {
  document.addEventListener('keydown', (e) => {
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
  // Check if modal is open
  const modal = document.getElementById('action-modal');
  if (modal && !modal.classList.contains('hidden')) {
    closeAddActionModal();
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
    initWeekView();
  }
}

// ===== CALENDAR FUNCTIONS =====

async function loadEvents() {
  const eventsList = document.getElementById('events-list');
  const dateKey = formatDateKey(currentDate);

  // Reset card selection
  resetCardSelection();
  updateDateDisplay();

  // Step 1: Check cache first for instant display
  const cacheResult = await window.electronAPI.getCachedEvents(dateKey);

  if (cacheResult.success && Array.isArray(cacheResult.events) && cacheResult.fromCache) {
    // Show cached data instantly
    displayEvents(cacheResult.events);
  } else {
    // No cache, show loading
    eventsList.innerHTML = '<div class="no-events">Loading events...</div>';
  }

  // Step 2: Always refresh from Outlook in background for latest data
  const startDate = new Date(currentDate);
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(currentDate);
  endDate.setHours(23, 59, 59, 999);

  setRefreshSpinning(true);
  try {
    const result = await window.electronAPI.getEvents(startDate, endDate);

    if (result.success) {
      displayEvents(result.events);
      // Update cache with fresh data
      window.electronAPI.saveEventsCache(result.events, dateKey);
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
    highlightEventAtHour = null; // Clear highlight target
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
    setTimeout(() => {
      applyEventHighlight(highlightEventAtHour);
      highlightEventAtHour = null; // Clear after applying
    }, 200);
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

  if (event.accountColor) {
    card.style.setProperty('--account-color', event.accountColor);
  }
  if (event.accountName) {
    card.setAttribute('data-account', event.accountName);
  }

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
      ${event.teamsLink ? `
        <a href="#" class="teams-join-link" data-teams-url="${escapeHtml(event.teamsLink)}">
          <svg width="12" height="12" aria-hidden="true"><use href="#icon-video"/></svg>
          Join Teams
        </a>
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
    ${event.teamsLink ? `<div><strong>Teams:</strong> <a href="#" class="teams-meeting-link" data-teams-url="${escapeHtml(event.teamsLink)}">Join Meeting</a></div>` : ''}
    <div><strong>Organizer:</strong> ${escapeHtml(event.organizer)}</div>
  `;

  // Handle Teams link in metadata
  if (event.teamsLink) {
    const teamsMetaLink = metadata.querySelector('.teams-meeting-link');
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

function displayActions() {
  const actionsList = document.getElementById('actions-list');

  if (actions.length === 0) {
    actionsList.innerHTML = '<div class="no-actions">No actions yet. Click "+" to create one!</div>';
    return;
  }

  const q = actionSearchQuery.trim().toLowerCase();
  const filtered = q
    ? actions.filter(a =>
        a.title.toLowerCase().includes(q) ||
        (a.note && a.note.toLowerCase().includes(q))
      )
    : actions;

  if (filtered.length === 0) {
    actionsList.innerHTML = '<div class="no-actions">No actions match your search.</div>';
    return;
  }

  actionsList.innerHTML = '';

  // Pinned actions float to the top
  const sorted = [...filtered].sort((a, b) => (b.pinned ? 1 : 0) - (a.pinned ? 1 : 0));

  sorted.forEach(action => {
    const actionCard = createActionCard(action);
    actionsList.appendChild(actionCard);
  });
}

function createActionCard(action) {
  const card = document.createElement('div');
  card.className = 'action-card';
  card.dataset.actionId = action.id;

  // Check if this action is expanded
  if (expandedActions.has(action.id)) {
    card.classList.add('expanded');
  }

  card.innerHTML = `
    <div class="action-header">
      <div class="action-content">
        <div class="action-title">${escapeHtml(action.title)}</div>
      </div>
      <div class="action-buttons">
        <button class="action-expand-btn" aria-label="Expand/Collapse" title="Expand/Collapse">
          <svg class="icon" width="16" height="16" aria-hidden="true"><use href="#icon-chevron-down"/></svg>
        </button>
        <button class="action-pin-btn${action.pinned ? ' pinned' : ''}" aria-label="${action.pinned ? 'Unpin' : 'Pin to top'}" title="${action.pinned ? 'Unpin' : 'Pin to top'}">
          <svg class="icon" width="14" height="14" aria-hidden="true"><use href="#icon-pin"/></svg>
        </button>
        ${action.url ? '<button class="action-open-btn">Open</button>' : ''}
        <button class="action-delete-btn" aria-label="Delete action" title="Delete action">
          <svg class="icon" width="13" height="13" aria-hidden="true"><use href="#icon-x"/></svg>
        </button>
      </div>
    </div>
    <div class="action-details">
      ${action.url ? `<div class="action-url">${escapeHtml(action.url)}</div>` : ''}
      <div class="form-group">
        <label>Note</label>
        <textarea class="action-note" placeholder="Add notes about this action..." rows="3">${escapeHtml(action.note || '')}</textarea>
      </div>
    </div>
  `;

  // Add event listeners (not using onclick to properly handle events)
  const expandBtn = card.querySelector('.action-expand-btn');
  const pinBtn   = card.querySelector('.action-pin-btn');
  const openBtn = card.querySelector('.action-open-btn');
  const deleteBtn = card.querySelector('.action-delete-btn');
  const noteArea = card.querySelector('.action-note');

  // Expand/collapse handler
  expandBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    toggleActionExpand(action.id, card);
  });

  // Pin/unpin handler
  pinBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    pinAction(action.id);
  });

  // Open action handler (only if URL exists)
  if (openBtn && action.url) {
    openBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      openAction(action.url);
    });
  }

  // Delete action handler with proper error handling
  deleteBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    deleteAction(action.id);
  });

  // Save note on change (debounced)
  let noteTimeout;
  noteArea.addEventListener('input', (e) => {
    e.stopPropagation();
    clearTimeout(noteTimeout);
    noteTimeout = setTimeout(() => {
      saveActionNote(action.id, e.target.value);
    }, 1000);
  });

  return card;
}

function openAddActionModal() {
  document.getElementById('action-modal').classList.remove('hidden');
  document.getElementById('action-title-input').value = '';
  document.getElementById('action-url-input').value = '';
  const noteInput = document.getElementById('action-note-input');
  if (noteInput) noteInput.value = '';
  document.getElementById('action-title-input').focus();
}

function closeAddActionModal() {
  document.getElementById('action-modal').classList.add('hidden');
}

async function handleAddAction(e) {
  e.preventDefault();

  const title = document.getElementById('action-title-input').value.trim();
  const url = document.getElementById('action-url-input').value.trim();
  const note = document.getElementById('action-note-input')?.value.trim() || '';

  if (!title) {
    alert('Please enter a title');
    return;
  }

  try {
    const result = await window.electronAPI.addAction({ title, url, note });

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

async function deleteAction(actionId) {
  if (!confirm('Delete this action?')) return;

  try {
    const result = await window.electronAPI.deleteAction(actionId);

    if (result.success) {
      // Remove from expanded set
      expandedActions.delete(actionId);
      // Remove from actions array
      actions = actions.filter(a => a.id !== actionId);
      // Re-render the list
      displayActions();
    } else {
      alert('Error deleting action: ' + result.error);
    }
  } catch (error) {
    console.error('Error deleting action:', error);
    alert('Error: ' + error.message);
  }
}

function openAction(url) {
  window.electronAPI.openExternal(url);
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

// Save action note
async function saveActionNote(actionId, note) {
  try {
    // Find the action in the array
    const action = actions.find(a => a.id === actionId);
    if (!action) return;

    // Update the note
    action.note = note;

    // Save to backend
    const result = await window.electronAPI.saveActions(actions);

    if (!result.success) {
      console.error('Error saving action note:', result.error);
    }
  } catch (error) {
    console.error('Error saving action note:', error);
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

// Populate heatmap with meeting data from cache and Outlook
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

  // Fetch events for the entire week from Outlook
  const weekStart = new Date(currentWeekStart);
  weekStart.setHours(0, 0, 0, 0);
  const weekEnd = new Date(currentWeekStart);
  weekEnd.setDate(weekEnd.getDate() + 6);
  weekEnd.setHours(23, 59, 59, 999);

  // Try to get events from Outlook for this week
  const result = await window.electronAPI.getEvents(weekStart, weekEnd);

  console.log(`Fetched ${result.events?.length || 0} events for week ${formatDateKey(weekStart)}`);

  // Log each event for debugging
  if (result.events) {
    result.events.forEach(evt => {
      const date = new Date(evt.start);
      console.log(`  - ${evt.subject} on ${date.toLocaleDateString()} at ${date.toLocaleTimeString()} | Organizer: ${evt.organizer} | Status: ${evt.meetingStatus}/${evt.responseStatus}`);
    });
  }

  if (result.success && Array.isArray(result.events)) {
    // Cache the events by day for faster future access
    const eventsByDay = {};
    result.events.forEach(event => {
      const eventDate = new Date(event.start);
      const dateKey = formatDateKey(eventDate);
      if (!eventsByDay[dateKey]) {
        eventsByDay[dateKey] = [];
      }
      eventsByDay[dateKey].push(event);
    });

    // Save to cache
    await window.electronAPI.saveEventsRangeCache(eventsByDay);

    // Process all events
    result.events.forEach(event => {
      const startTime = new Date(event.start);
      const endTime = new Date(event.end);
      const startHour = startTime.getHours();
      const endHour = endTime.getHours();
      const durationMinutes = (endTime - startTime) / 60000;

      // Calculate which day of the week this is
      const daysDiff = Math.floor((startTime - weekStart) / (1000 * 60 * 60 * 24));
      if (daysDiff < 0 || daysDiff > 6) return;

      // Mark the hour blocks this meeting occupies (8am-5pm)
      hours.forEach(hour => {
        // Check if this hour block overlaps with the meeting
        // Meeting can start before 8am or end after 5pm, but we only show 8am-5pm blocks
        const blockStart = new Date(startTime);
        blockStart.setHours(hour, 0, 0, 0);
        const blockEnd = new Date(blockStart);
        blockEnd.setHours(hour + 1, 0, 0, 0);

        // Check if meeting overlaps with this hour block
        const meetingOverlapsBlock = startTime < blockEnd && endTime > blockStart;

        if (meetingOverlapsBlock) {
          const dateKey = formatDateKey(startTime);
          const block = document.querySelector(
            `.heatmap-block[data-hour="${hour}"][data-date^="${dateKey}"]`
          );

          if (block) {
            // Determine if this is a short meeting
            const isShortMeeting = durationMinutes < 60;

            if (isShortMeeting) {
              // Only mark as short if not already marked as full
              if (!block.classList.contains('has-meeting')) {
                block.classList.remove('empty');
                block.classList.add('short-meeting');
              }
            } else {
              // Full meeting or multiple meetings - always mark as has-meeting
              block.classList.remove('empty', 'short-meeting');
              block.classList.add('has-meeting');
            }
          }
        }
      });
    });
  } else {
    // Fallback to cached data if Outlook not available
    for (let dayOffset = 0; dayOffset < days; dayOffset++) {
      const date = new Date(currentWeekStart);
      date.setDate(date.getDate() + dayOffset);
      const dateKey = formatDateKey(date);

      const cacheResult = await window.electronAPI.getCachedEvents(dateKey);

      if (cacheResult.success && Array.isArray(cacheResult.events)) {
        cacheResult.events.forEach(event => {
          const startTime = new Date(event.start);
          const endTime = new Date(event.end);
          const durationMinutes = (endTime - startTime) / 60000;

          hours.forEach(hour => {
            // Check if this hour block overlaps with the meeting
            const blockStart = new Date(startTime);
            blockStart.setHours(hour, 0, 0, 0);
            const blockEnd = new Date(blockStart);
            blockEnd.setHours(hour + 1, 0, 0, 0);

            const meetingOverlapsBlock = startTime < blockEnd && endTime > blockStart;

            if (meetingOverlapsBlock) {
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
    }
  }
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

  eventCards.forEach(card => {
    const eventStart = card.dataset.eventStart;
    if (eventStart) {
      const startDate = new Date(eventStart);
      const startHour = startDate.getHours();

      // Check if this event includes the clicked hour
      const eventEnd = card.dataset.eventEnd;
      const endDate = new Date(eventEnd);
      const endHour = endDate.getHours();

      if (clickedHour >= startHour && clickedHour < endHour) {
        // Add glow effect
        card.classList.add('event-glow');
        foundEvent = true;

        // Scroll to the event
        card.scrollIntoView({ behavior: 'smooth', block: 'center' });

        // Remove glow after 5 seconds
        setTimeout(() => {
          card.classList.remove('event-glow');
        }, 5000);
      }
    }
  });

  if (!foundEvent) {
    console.log(`No event found for hour ${clickedHour}`);
  }
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
