let currentDate = new Date();
let currentMeetingId = null;
let currentMeetingData = null;
let currentTab = 'calendar';
let actions = [];

// Initialize the app
async function init() {
  setupEventListeners();
  await checkOutlookConnection();
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

  // Date navigation
  document.getElementById('prev-day').addEventListener('click', () => {
    currentDate.setDate(currentDate.getDate() - 1);
    loadEvents();
  });

  document.getElementById('next-day').addEventListener('click', () => {
    currentDate.setDate(currentDate.getDate() + 1);
    loadEvents();
  });

  document.getElementById('refresh-btn').addEventListener('click', loadEvents);

  // Notes
  document.getElementById('close-notes-btn').addEventListener('click', closeNotes);
  document.getElementById('save-notes-btn').addEventListener('click', saveNotes);

  // Actions
  document.getElementById('add-action-btn').addEventListener('click', openAddActionModal);
  document.getElementById('cancel-action-btn').addEventListener('click', closeAddActionModal);
  document.getElementById('action-form').addEventListener('submit', handleAddAction);

  // Close modal on overlay click
  document.getElementById('action-modal').addEventListener('click', (e) => {
    if (e.target.id === 'action-modal') {
      closeAddActionModal();
    }
  });
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

  // Update content
  document.getElementById('calendar-content').classList.toggle('hidden', tab !== 'calendar');
  document.getElementById('actions-content').classList.toggle('hidden', tab !== 'actions');

  // Load data if needed
  if (tab === 'actions') {
    displayActions();
  }
}

// ===== CALENDAR FUNCTIONS =====

async function loadEvents() {
  const eventsList = document.getElementById('events-list');
  eventsList.innerHTML = '<div class="no-events">Loading events...</div>';

  updateDateDisplay();

  const startDate = new Date(currentDate);
  startDate.setHours(0, 0, 0, 0);

  const endDate = new Date(currentDate);
  endDate.setHours(23, 59, 59, 999);

  try {
    const result = await window.electronAPI.getEvents(startDate, endDate);

    if (result.success) {
      displayEvents(result.events);
    } else {
      eventsList.innerHTML = '<div class="no-events">Error loading events: ' + result.error + '</div>';
    }
  } catch (error) {
    eventsList.innerHTML = '<div class="no-events">Error: ' + error.message + '</div>';
  }
}

function displayEvents(events) {
  const eventsList = document.getElementById('events-list');

  if (events.length === 0) {
    eventsList.innerHTML = '<div class="no-events">No meetings scheduled for this day</div>';
    return;
  }

  eventsList.innerHTML = '';

  events.forEach(event => {
    const eventCard = createEventCard(event);
    eventsList.appendChild(eventCard);
  });
}

function createEventCard(event) {
  const card = document.createElement('div');
  card.className = 'event-card';

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

  card.innerHTML = `
    <div class="event-time">${startTime} - ${endTime}</div>
    <div class="event-title">${escapeHtml(event.subject)}</div>
    <div class="event-details">
      ${event.location ? `<div>📍 ${escapeHtml(event.location)}</div>` : ''}
      ${event.organizer ? `<div>👤 ${escapeHtml(event.organizer)}</div>` : ''}
      ${event.teamsLink ? `<a href="${event.teamsLink}" class="teams-link" onclick="event.stopPropagation()">🎥 Join Teams Meeting</a>` : ''}
    </div>
  `;

  card.addEventListener('click', () => openNotes(event));

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
    ${event.teamsLink ? `<div><strong>Teams:</strong> <a href="${event.teamsLink}" target="_blank">Join Meeting</a></div>` : ''}
    <div><strong>Organizer:</strong> ${escapeHtml(event.organizer)}</div>
  `;

  const result = await window.electronAPI.getNote(event.id);
  document.getElementById('notes-editor').value = result.content || '';

  document.getElementById('calendar-content').classList.add('hidden');
  document.getElementById('notes-view').classList.remove('hidden');
}

function closeNotes() {
  document.getElementById('calendar-content').classList.remove('hidden');
  document.getElementById('notes-view').classList.add('hidden');
  currentMeetingId = null;
  currentMeetingData = null;
}

async function saveNotes() {
  if (!currentMeetingId || !currentMeetingData) return;

  const noteContent = document.getElementById('notes-editor').value;
  const statusEl = document.getElementById('notes-status');

  statusEl.textContent = 'Saving...';

  try {
    const result = await window.electronAPI.saveNote(
      currentMeetingId,
      noteContent,
      currentMeetingData
    );

    if (result.success) {
      statusEl.textContent = '✓ Saved';
      setTimeout(() => {
        statusEl.textContent = '';
      }, 2000);
    } else {
      statusEl.textContent = '✗ Error: ' + result.error;
    }
  } catch (error) {
    statusEl.textContent = '✗ Error: ' + error.message;
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
    actionsList.innerHTML = '<div class="no-actions">No actions yet. Click "+ Add Action" to create one!</div>';
    return;
  }

  actionsList.innerHTML = '';

  actions.forEach(action => {
    const actionCard = createActionCard(action);
    actionsList.appendChild(actionCard);
  });
}

function createActionCard(action) {
  const card = document.createElement('div');
  card.className = 'action-card';

  card.innerHTML = `
    <div class="action-content">
      <div class="action-title">${escapeHtml(action.title)}</div>
      <div class="action-url">${escapeHtml(action.url)}</div>
    </div>
    <div class="action-buttons">
      <button class="action-open-btn" onclick="openAction('${escapeHtml(action.url)}')">Open</button>
      <button class="action-delete-btn" onclick="deleteAction('${action.id}')">Delete</button>
    </div>
  `;

  return card;
}

function openAddActionModal() {
  document.getElementById('action-modal').classList.remove('hidden');
  document.getElementById('action-title-input').value = '';
  document.getElementById('action-url-input').value = '';
  document.getElementById('action-title-input').focus();
}

function closeAddActionModal() {
  document.getElementById('action-modal').classList.add('hidden');
}

async function handleAddAction(e) {
  e.preventDefault();

  const title = document.getElementById('action-title-input').value.trim();
  const url = document.getElementById('action-url-input').value.trim();

  if (!title || !url) return;

  try {
    const result = await window.electronAPI.addAction({ title, url });

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
      actions = actions.filter(a => a.id !== actionId);
      displayActions();
    } else {
      alert('Error deleting action: ' + result.error);
    }
  } catch (error) {
    alert('Error: ' + error.message);
  }
}

function openAction(url) {
  window.open(url, '_blank');
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
