let currentDate = new Date();
let currentMeetingId = null;
let currentMeetingData = null;

// Initialize the app
async function init() {
  setupEventListeners();

  const isAuthenticated = await window.electronAPI.getAuthStatus();

  if (isAuthenticated) {
    showMainScreen();
    await loadEvents();
  } else {
    showLoginScreen();
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

  // Login
  document.getElementById('login-btn').addEventListener('click', handleLogin);

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
}

function showLoginScreen() {
  document.getElementById('login-screen').classList.remove('hidden');
  document.getElementById('main-screen').classList.add('hidden');
}

function showMainScreen() {
  document.getElementById('login-screen').classList.add('hidden');
  document.getElementById('main-screen').classList.remove('hidden');
}

async function handleLogin() {
  const loginBtn = document.getElementById('login-btn');
  loginBtn.textContent = 'Signing in...';
  loginBtn.disabled = true;

  try {
    const result = await window.electronAPI.login();

    if (result.success) {
      showMainScreen();
      await loadEvents();
    } else {
      alert('Login failed: ' + result.error);
      loginBtn.textContent = 'Sign in with Microsoft';
      loginBtn.disabled = false;
    }
  } catch (error) {
    alert('Login error: ' + error.message);
    loginBtn.textContent = 'Sign in with Microsoft';
    loginBtn.disabled = false;
  }
}

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

  document.getElementById('calendar-view').classList.add('hidden');
  document.getElementById('notes-view').classList.remove('hidden');
}

function closeNotes() {
  document.getElementById('calendar-view').classList.remove('hidden');
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

function updateDateDisplay() {
  const options = { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' };
  const dateString = currentDate.toLocaleDateString('en-US', options);
  document.getElementById('current-date').textContent = dateString;
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// Initialize on load
init();
