// ===== Data Store =====
const STORAGE_KEYS = {
    tasks: 'pa_tasks',
    archivedTasks: 'pa_tasks_archived',
    events1p: 'pa_events_1p',
    events3p: 'pa_events_3p',
    products: 'pa_products',
    personalTasks: 'pa_personal_tasks',
    archivedPersonalTasks: 'pa_personal_tasks_archived',
    familyEvents: 'pa_family_events'
};

const GITHUB_REPO = 'marcobellinimsft/Personal-Assistant';
const GITHUB_DATA_FILE = 'data.json';
const GITHUB_BACKUP_FILE = 'data-backup.json';
const GITHUB_BRANCH = 'main';
let githubSha = null; // track file SHA for updates
let githubBackupSha = null;
let syncPending = false;
let syncTimer = null;
let lastSyncTime = null;

function loadData(key, defaults) {
    try {
        const saved = localStorage.getItem(key);
        return saved ? JSON.parse(saved) : defaults;
    } catch { return defaults; }
}

function saveData(key, data) {
    localStorage.setItem(key, JSON.stringify(data));
    // Also keep a localStorage backup snapshot
    saveLocalBackup();
    scheduleSyncToGitHub();
}

function saveLocalBackup() {
    try {
        localStorage.setItem('pa_local_backup', JSON.stringify(getAllData()));
        localStorage.setItem('pa_local_backup_time', new Date().toISOString());
    } catch { /* quota exceeded — ok */ }
}

function scheduleSyncToGitHub() {
    syncPending = true;
    updateSyncIndicator('pending');
    clearTimeout(syncTimer);
    syncTimer = setTimeout(() => syncToGitHub(), 2000);
}

function getAllData() {
    const data = {};
    for (const [name, key] of Object.entries(STORAGE_KEYS)) {
        try { data[name] = JSON.parse(localStorage.getItem(key)) || []; }
        catch { data[name] = []; }
    }
    data._savedAt = new Date().toISOString();
    return data;
}

function applyAllData(data) {
    for (const [name, key] of Object.entries(STORAGE_KEYS)) {
        if (data[name] !== undefined) {
            localStorage.setItem(key, JSON.stringify(data[name]));
        }
    }
}

function getGitHubToken() {
    return localStorage.getItem('pa_github_token') || '';
}

function setGitHubToken(token) {
    localStorage.setItem('pa_github_token', token);
}

function updateSyncIndicator(state) {
    const el = document.getElementById('sync-status');
    if (!el) return;
    const timeStr = lastSyncTime ? new Date(lastSyncTime).toLocaleTimeString() : '';
    if (state === 'syncing') {
        el.innerHTML = '<span class="material-icons-outlined spin">sync</span>';
        el.title = 'Syncing to GitHub...';
    } else if (state === 'ok') {
        el.innerHTML = '<span class="material-icons-outlined" style="color:#86efac">cloud_done</span>';
        el.title = 'Synced to GitHub' + (timeStr ? ' at ' + timeStr : '');
    } else if (state === 'pending') {
        el.innerHTML = '<span class="material-icons-outlined" style="color:#fbbf24">cloud_upload</span>';
        el.title = 'Changes pending sync...';
    } else if (state === 'error') {
        el.innerHTML = '<span class="material-icons-outlined" style="color:#f87171">cloud_off</span>';
        el.title = 'Sync failed — check token in settings';
    } else if (state === 'notoken') {
        el.innerHTML = '<span class="material-icons-outlined" style="color:#888">cloud_off</span>';
        el.title = 'No GitHub token — click settings to configure';
    }
    // Update last sync display
    const tsEl = document.getElementById('last-sync-time');
    if (tsEl) tsEl.textContent = timeStr ? 'Last sync: ' + timeStr : '';
}

async function syncToGitHub() {
    const token = getGitHubToken();
    if (!token) { updateSyncIndicator('notoken'); return; }
    updateSyncIndicator('syncing');
    try {
        const content = btoa(unescape(encodeURIComponent(JSON.stringify(getAllData(), null, 2))));
        const body = {
            message: 'Auto-sync data ' + new Date().toISOString(),
            content: content,
            branch: GITHUB_BRANCH
        };
        if (githubSha) body.sha = githubSha;
        const resp = await fetch(`https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_DATA_FILE}`, {
            method: 'PUT',
            headers: {
                'Authorization': `token ${token}`,
                'Content-Type': 'application/json',
                'Accept': 'application/vnd.github.v3+json'
            },
            body: JSON.stringify(body)
        });
        if (resp.ok) {
            const result = await resp.json();
            githubSha = result.content.sha;
            syncPending = false;
            lastSyncTime = new Date().toISOString();
            localStorage.setItem('pa_last_sync', lastSyncTime);
            updateSyncIndicator('ok');
            // Also save a rolling backup to GitHub (async, don't block)
            saveGitHubBackup();
        } else if (resp.status === 409) {
            // Conflict — re-fetch SHA and retry
            await loadFromGitHub();
            await syncToGitHub();
        } else {
            console.error('GitHub sync failed:', resp.status, await resp.text());
            updateSyncIndicator('error');
        }
    } catch (err) {
        console.error('GitHub sync error:', err);
        updateSyncIndicator('error');
    }
}

async function loadFromGitHub() {
    const token = getGitHubToken();
    if (!token) { updateSyncIndicator('notoken'); return false; }
    try {
        const resp = await fetch(`https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_DATA_FILE}?ref=${GITHUB_BRANCH}`, {
            headers: {
                'Authorization': `token ${token}`,
                'Accept': 'application/vnd.github.v3+json'
            }
        });
        if (resp.ok) {
            const file = await resp.json();
            githubSha = file.sha;
            const decoded = decodeURIComponent(escape(atob(file.content.replace(/\n/g, ''))));
            const data = JSON.parse(decoded);
            applyAllData(data);
            updateSyncIndicator('ok');
            return true;
        } else if (resp.status === 404) {
            // File doesn't exist yet — will be created on first save
            githubSha = null;
            updateSyncIndicator('ok');
            return false;
        } else {
            updateSyncIndicator('error');
            return false;
        }
    } catch (err) {
        console.error('GitHub load error:', err);
        updateSyncIndicator('error');
        return false;
    }
}

function exportData() {
    const data = getAllData();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `personal-assistant-backup-${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
}

async function saveGitHubBackup() {
    const token = getGitHubToken();
    if (!token) return;
    try {
        // Get current backup SHA
        if (!githubBackupSha) {
            const check = await fetch(`https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_BACKUP_FILE}?ref=${GITHUB_BRANCH}`, {
                headers: { 'Authorization': `token ${token}`, 'Accept': 'application/vnd.github.v3+json' }
            });
            if (check.ok) {
                const f = await check.json();
                githubBackupSha = f.sha;
            }
        }
        const content = btoa(unescape(encodeURIComponent(JSON.stringify(getAllData(), null, 2))));
        const body = {
            message: 'Auto-backup ' + new Date().toISOString(),
            content: content,
            branch: GITHUB_BRANCH
        };
        if (githubBackupSha) body.sha = githubBackupSha;
        const resp = await fetch(`https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_BACKUP_FILE}`, {
            method: 'PUT',
            headers: {
                'Authorization': `token ${token}`,
                'Content-Type': 'application/json',
                'Accept': 'application/vnd.github.v3+json'
            },
            body: JSON.stringify(body)
        });
        if (resp.ok) {
            const result = await resp.json();
            githubBackupSha = result.content.sha;
        }
    } catch (err) {
        console.error('Backup save error:', err);
    }
}

function importData() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.onchange = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            try {
                const data = JSON.parse(ev.target.result);
                applyAllData(data);
                reloadAllState();
                scheduleSyncToGitHub();
                alert('Data imported successfully!');
            } catch { alert('Invalid JSON file.'); }
        };
        reader.readAsText(file);
    };
    input.click();
}

function reloadAllState() {
    tasks = loadData(STORAGE_KEYS.tasks, DEFAULT_TASKS);
    archivedTasks = loadData(STORAGE_KEYS.archivedTasks, []);
    events1p = loadData(STORAGE_KEYS.events1p, DEFAULT_EVENTS_1P);
    events3p = loadData(STORAGE_KEYS.events3p, DEFAULT_EVENTS_3P);
    products = loadData(STORAGE_KEYS.products, DEFAULT_PRODUCTS);
    personalTasks = loadData(STORAGE_KEYS.personalTasks, []);
    archivedPersonalTasks = loadData(STORAGE_KEYS.archivedPersonalTasks, []);
    familyEvents = loadData(STORAGE_KEYS.familyEvents, []);
    showPage('welcome');
}

function showSettings() {
    const modal = document.getElementById('settings-modal');
    const tokenInput = document.getElementById('github-token-input');
    tokenInput.value = getGitHubToken();
    modal.style.display = 'flex';
}

function closeSettings() {
    document.getElementById('settings-modal').style.display = 'none';
}

function saveSettings() {
    const token = document.getElementById('github-token-input').value.trim();
    setGitHubToken(token);
    closeSettings();
    if (token) {
        loadFromGitHub().then(loaded => {
            if (loaded) reloadAllState();
            else syncToGitHub(); // push current data if file doesn't exist
        });
    } else {
        updateSyncIndicator('notoken');
    }
}

// ===== Helpers =====
function getNextFriday() {
    const d = new Date();
    const diff = (5 - d.getDay() + 7) % 7 || 7;
    d.setDate(d.getDate() + diff);
    return d.toISOString().split('T')[0];
}

function esc(str) {
    if (str == null) return '';
    return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

let autoSaveTimers = {};
function debounceSave(key, data) {
    clearTimeout(autoSaveTimers[key]);
    autoSaveTimers[key] = setTimeout(() => saveData(key, data), 400);
}

// ===== Default Data =====
const DEFAULT_TASKS = [
    { id: 1, name: 'Official landing page', who: 'Mindy Bomonti', date: getNextFriday(), link: '', vendor: '' },
    { id: 2, name: 'Playlist 5', who: 'Larry Larsen', date: getNextFriday(), link: 'CAIP Marketing Planning Greenlight .pptx', vendor: '' },
    { id: 3, name: '', who: 'Mindy Bomonti', date: '', link: '', vendor: '' }
];

const DEFAULT_EVENTS_1P = [
    { id: 1, date: '2026-05-15', name: 'M&M Summit - Spring', hero: 'Azure Arc', pmm: '', contact: '', plan: '', activity: '', notes: '' },
    { id: 2, date: '2026-09-20', name: 'Windows Server Summit', hero: 'Windows Server', pmm: 'Jennifer Yuan', contact: '', plan: '', activity: '', notes: '' },
    { id: 3, date: '2026-11-17', name: 'Ignite 2026', hero: 'Multiple', pmm: '', contact: '', plan: '', activity: '', notes: 'Main keynote + breakout sessions' },
    { id: 4, date: '2026-06-10', name: 'M&M Summit - Security', hero: 'Defender for Cloud', pmm: 'Shirleyse Haley', contact: '', plan: '', activity: '', notes: '' }
];

const DEFAULT_EVENTS_3P = [];

const DEFAULT_PRODUCTS = [
    { name: 'App Modernization (AKS, App Service)', pmm: 'Mike Weber, Ashley Adelberg, Mayunk Jain', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Windows Server', pmm: 'Jennifer Yuan', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'SQL Server', pmm: 'Debbi Lyons, Govanna Flores', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Azure SQL', pmm: 'Govanna Flores', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Defender for Cloud', pmm: 'Shirleyse Haley', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Azure VMware Solution', pmm: 'Kirsten Megahan, Britney Cretella', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Azure Arc', pmm: 'Jyoti Sharma, Antonio Ortoll', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Azure CoPilot', pmm: 'Jyoti Sharma, Antonio Ortoll, Arti Gulwadi', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Linux', pmm: 'Enrico Fuiano, Naga Surendran', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Security', pmm: 'Shirleyse Haley, Molina Sharma, Sean Whalen', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Azure Database for PostgreSQL', pmm: 'Pooja Yarabothu, Teneil Lawrence', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Azure MySQL', pmm: 'Teneil Lawrence, Pooja Yarabothu', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'Oracle', pmm: 'Alex Stairs, Sparsh Agrawat', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] },
    { name: 'SAP on Azure', pmm: 'Ankita Bhalla, Sanjay Satheesh', pmmManager: '', sme: '', globalSkilling: '', updates: [], events: [] }
];

// ===== State =====
let tasks = loadData(STORAGE_KEYS.tasks, DEFAULT_TASKS);
let archivedTasks = loadData(STORAGE_KEYS.archivedTasks, []);
let events1p = loadData(STORAGE_KEYS.events1p, DEFAULT_EVENTS_1P);
let events3p = loadData(STORAGE_KEYS.events3p, DEFAULT_EVENTS_3P);
let products = loadData(STORAGE_KEYS.products, DEFAULT_PRODUCTS);
let personalTasks = loadData(STORAGE_KEYS.personalTasks, DEFAULT_PERSONAL_TASKS);
let archivedPersonalTasks = loadData(STORAGE_KEYS.archivedPersonalTasks, []);
let familyEvents = loadData(STORAGE_KEYS.familyEvents, DEFAULT_FAMILY_EVENTS);
let currentSort = { table: null, column: null, dir: 'asc' };

// ===== Page Navigation =====
function showPage(pageId) {
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    const page = document.getElementById('page-' + pageId);
    if (page) {
        page.classList.add('active');
        if (pageId === 'tasks') renderTasks();
        if (pageId === 'hero-products') renderProducts();
        if (pageId === '1p-events') renderEvents('1p');
        if (pageId === '3p-events') renderEvents('3p');
        if (pageId === 'personal-tasks') renderPersonalTasks();
        if (pageId === 'family-events') renderFamilyEvents();
        if (pageId === 'welcome') updateStats();
    }
}

// ===== Sidebar Toggle =====
function toggleGroup(el) {
    el.classList.toggle('collapsed');
    el.nextElementSibling.classList.toggle('open');
}

// ===== Task Categories & Colors =====
const TASK_CATEGORIES = [
    { value: 'Content', color: '#2e7d32', bg: '#e8f5e9' },
    { value: 'GTM', color: '#e67e22', bg: '#fef5ec' },
    { value: 'Planning', color: '#9b59b6', bg: '#f5eef8' },
    { value: 'Sync', color: '#27ae60', bg: '#eafaf1' },
    { value: 'Event', color: '#e74c3c', bg: '#fdedec' },
    { value: 'Presentations', color: '#c2185b', bg: '#fce4ec' },
    { value: 'Other', color: '#7f8c8d', bg: '#f2f4f4' }
];

const VENDOR_NAMES = ['Amy', 'Mindy', 'Erica'];

const PERSONAL_CATEGORIES = [
    { value: 'Event', color: '#e74c3c', bg: '#fdedec' },
    { value: 'Personal', color: '#3498db', bg: '#ebf5fb' },
    { value: 'Finance', color: '#27ae60', bg: '#eafaf1' },
    { value: 'Church', color: '#8e44ad', bg: '#f5eef8' },
    { value: 'Entertainment', color: '#e67e22', bg: '#fef5ec' },
    { value: 'Family', color: '#2980b9', bg: '#d6eaf8' },
    { value: 'UW', color: '#4b2e83', bg: '#ece3f5' },
    { value: 'Moglie', color: '#c2185b', bg: '#fce4ec' }
];

const PERSONAL_WHO = ['Marco', 'Daniela', 'David', 'Andrew', 'Nicholas', 'Simon', 'Sara', 'Jackson', 'Maria', 'Egidio', 'Liana', 'Papi', 'Altri'];

const DEFAULT_PERSONAL_TASKS = [];
const DEFAULT_FAMILY_EVENTS = [];

const URGENCY_LEVELS = [
    { value: '1', label: 'Urgent 1', color: '#e74c3c', bg: '#fdedec' },
    { value: '2', label: 'Urgent 2', color: '#f39c12', bg: '#fef9e7' },
    { value: '3', label: 'Urgent 3', color: '#27ae60', bg: '#eafaf1' },
    { value: '4', label: 'Urgent 4', color: '#3498db', bg: '#ebf5fb' }
];

function getUrgencyInfo(val) {
    return URGENCY_LEVELS.find(u => u.value === val) || URGENCY_LEVELS[3];
}

function buildUrgencySelect(selected, idx) {
    const u = getUrgencyInfo(selected);
    return `<select class="urgency-select" style="background:${u.color};color:#fff" onchange="updateTask(${idx},'urgency',this.value); this.style.background=getUrgencyInfo(this.value).color;">
        ${URGENCY_LEVELS.map(l => `<option value="${l.value}" ${l.value === selected ? 'selected' : ''} style="background:#fff;color:#333">${l.label}</option>`).join('')}
    </select>`;
}

function getCategoryInfo(cat) {
    return TASK_CATEGORIES.find(c => c.value === cat) || TASK_CATEGORIES[5];
}

function getDateAlert(dateStr) {
    if (!dateStr) return '';
    const today = new Date(); today.setHours(0,0,0,0);
    const due = new Date(dateStr + 'T00:00:00');
    const diff = (due - today) / (1000 * 60 * 60 * 24);
    if (diff < 0) return '<span class="date-alert date-overdue" title="Overdue!"><span class="material-icons-outlined">error</span></span>';
    if (diff <= 7) return '<span class="date-alert date-soon" title="Due within a week"><span class="material-icons-outlined">warning</span></span>';
    return '';
}

function buildCategorySelect(selected, idx) {
    const cat = getCategoryInfo(selected);
    return `<select class="cat-select" style="background:${cat.color};color:#fff" onchange="updateTask(${idx},'category',this.value); this.style.background=getCategoryInfo(this.value).color;">
        ${TASK_CATEGORIES.map(c => `<option value="${c.value}" ${c.value === selected ? 'selected' : ''} style="background:#fff;color:#333">${c.value}</option>`).join('')}
    </select>`;
}

function buildVendorSelect(selected, idx) {
    return `<select class="vendor-select" onchange="updateTask(${idx},'vendor',this.value)">
        <option value="" ${!selected ? 'selected' : ''}>—</option>
        ${VENDOR_NAMES.map(v => `<option value="${v}" ${v === selected ? 'selected' : ''}>${v}</option>`).join('')}
    </select>`;
}

// ===== Tasks =====
function renderTasks() {
    const tbody = document.getElementById('tasks-body');
    tbody.innerHTML = '';
    // Sort tasks by urgency (1 on top, then 2, 3, 4)
    const sorted = tasks.map((t, i) => ({...t, _idx: i})).sort((a, b) => {
        const ua = parseInt(a.urgency || '4');
        const ub = parseInt(b.urgency || '4');
        return ua - ub;
    });
    sorted.forEach((task) => {
        const idx = task._idx;
        const cat = getCategoryInfo(task.category);
        const alert = getDateAlert(task.date);
        const tr = document.createElement('tr');
        tr.style.background = task.category ? cat.bg : '';
        tr.innerHTML = `
            <td class="col-check"><input type="checkbox" class="task-checkbox" onchange="completeTask(${idx})" title="Mark complete"></td>
            <td>${buildUrgencySelect(task.urgency || '4', idx)}</td>
            <td>${buildCategorySelect(task.category || 'Other', idx)}</td>
            <td><input type="text" value="${esc(task.name)}" placeholder="Task name..." onchange="updateTask(${idx},'name',this.value)"></td>
            <td><input type="text" value="${esc(task.who)}" placeholder="Person..." onchange="updateTask(${idx},'who',this.value)"></td>
            <td class="date-cell"><input type="date" value="${esc(task.date)}" onchange="updateTask(${idx},'date',this.value)">${alert}</td>
            <td><input type="text" value="${esc(task.link)}" placeholder="Link..." onchange="updateTask(${idx},'link',this.value)"></td>
            <td>${buildVendorSelect(task.vendor, idx)}</td>
            <td><textarea rows="1" placeholder="Notes..." onchange="updateTask(${idx},'notes',this.value)">${esc(task.notes || '')}</textarea></td>
            <td><button class="btn-delete" onclick="deleteTask(${idx})" title="Delete"><span class="material-icons-outlined">delete</span></button></td>
        `;
        tbody.appendChild(tr);
    });
    updateStats();
    renderArchivedToggle();
}

function completeTask(idx) {
    const task = tasks.splice(idx, 1)[0];
    task.completedDate = new Date().toISOString().split('T')[0];
    archivedTasks.unshift(task);
    saveData(STORAGE_KEYS.tasks, tasks);
    saveData(STORAGE_KEYS.archivedTasks, archivedTasks);
    renderTasks();
}

function resumeTask(idx) {
    const task = archivedTasks.splice(idx, 1)[0];
    delete task.completedDate;
    tasks.push(task);
    saveData(STORAGE_KEYS.tasks, tasks);
    saveData(STORAGE_KEYS.archivedTasks, archivedTasks);
    renderTasks();
}

function deleteArchivedTask(idx) {
    archivedTasks.splice(idx, 1);
    saveData(STORAGE_KEYS.archivedTasks, archivedTasks);
    renderTasks();
}

function renderArchivedToggle() {
    let section = document.getElementById('archived-section');
    if (!section) {
        section = document.createElement('div');
        section.id = 'archived-section';
        document.getElementById('page-tasks').appendChild(section);
    }
    if (archivedTasks.length === 0) {
        section.innerHTML = '';
        return;
    }
    section.innerHTML = `
        <button class="btn-archived-toggle" onclick="toggleArchived()">
            <span class="material-icons-outlined">archive</span>
            Archived (${archivedTasks.length})
            <span class="material-icons-outlined chevron-arch" id="arch-chevron">expand_more</span>
        </button>
        <div class="archived-table-wrap" id="archived-table-wrap">
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Urgency</th>
                            <th>Category</th>
                            <th>Task Name</th>
                            <th>Who</th>
                            <th>Date Due</th>
                            <th>Completed</th>
                            <th>Link</th>
                            <th>Vendor</th>
                            <th>Notes</th>
                            <th class="col-actions"></th>
                        </tr>
                    </thead>
                    <tbody>
                        ${archivedTasks.map((t, i) => {
                            const aCat = getCategoryInfo(t.category);
                            return `
                            <tr class="archived-row" style="background:${t.category ? aCat.bg : ''}">
                                <td><span class="urgency-badge" style="background:${getUrgencyInfo(t.urgency || '4').color}">${getUrgencyInfo(t.urgency || '4').label}</span></td>
                                <td><span class="cat-badge" style="background:${aCat.color}">${esc(t.category || 'Other')}</span></td>
                                <td>${esc(t.name)}</td>
                                <td>${esc(t.who)}</td>
                                <td>${esc(t.date)}</td>
                                <td><span class="completed-badge">${esc(t.completedDate)}</span></td>
                                <td>${esc(t.link)}</td>
                                <td>${esc(t.vendor)}</td>
                                <td>${esc(t.notes || '')}</td>
                                <td>
                                    <button class="btn-resume" onclick="resumeTask(${i})" title="Resume task"><span class="material-icons-outlined">replay</span></button>
                                    <button class="btn-delete" onclick="deleteArchivedTask(${i})" title="Delete"><span class="material-icons-outlined">delete</span></button>
                                </td>
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;
}

function toggleArchived() {
    const wrap = document.getElementById('archived-table-wrap');
    const chevron = document.getElementById('arch-chevron');
    wrap.classList.toggle('open');
    chevron.style.transform = wrap.classList.contains('open') ? 'rotate(180deg)' : '';
}

function addTask() {
    tasks.push({ id: Date.now(), name: '', who: '', date: '', link: '', vendor: '', category: 'Other', notes: '', urgency: '4' });
    saveData(STORAGE_KEYS.tasks, tasks);
    renderTasks();
    const inputs = document.querySelectorAll('#tasks-body tr:last-child input');
    if (inputs.length) inputs[0].focus();
}

function updateTask(idx, field, value) {
    tasks[idx][field] = value;
    debounceSave(STORAGE_KEYS.tasks, tasks);
}

function deleteTask(idx) {
    tasks.splice(idx, 1);
    saveData(STORAGE_KEYS.tasks, tasks);
    renderTasks();
}

// ===== Personal Tasks =====
function getPersonalCategoryInfo(cat) {
    return PERSONAL_CATEGORIES.find(c => c.value === cat) || PERSONAL_CATEGORIES[1];
}

function buildPersonalCategorySelect(selected, idx) {
    const cat = getPersonalCategoryInfo(selected);
    return `<select class="cat-select" style="background:${cat.color};color:#fff" onchange="updatePersonalTask(${idx},'category',this.value); this.style.background=getPersonalCategoryInfo(this.value).color;">
        ${PERSONAL_CATEGORIES.map(c => `<option value="${c.value}" ${c.value === selected ? 'selected' : ''} style="background:#fff;color:#333">${c.value}</option>`).join('')}
    </select>`;
}

function buildPersonalUrgencySelect(selected, idx) {
    const u = getUrgencyInfo(selected);
    return `<select class="urgency-select" style="background:${u.color};color:#fff" onchange="updatePersonalTask(${idx},'urgency',this.value); this.style.background=getUrgencyInfo(this.value).color;">
        ${URGENCY_LEVELS.map(l => `<option value="${l.value}" ${l.value === selected ? 'selected' : ''} style="background:#fff;color:#333">${l.label}</option>`).join('')}
    </select>`;
}

function buildPersonalWhoSelect(selected, idx) {
    return `<select class="vendor-select" onchange="updatePersonalTask(${idx},'who',this.value)">
        <option value="" ${!selected ? 'selected' : ''}>—</option>
        ${PERSONAL_WHO.map(w => `<option value="${w}" ${w === selected ? 'selected' : ''}>${w}</option>`).join('')}
    </select>`;
}

function renderPersonalTasks() {
    const tbody = document.getElementById('personal-tasks-body');
    tbody.innerHTML = '';
    const sorted = personalTasks.map((t, i) => ({...t, _idx: i})).sort((a, b) => {
        const ua = parseInt(a.urgency || '4');
        const ub = parseInt(b.urgency || '4');
        return ua - ub;
    });
    sorted.forEach((task) => {
        const idx = task._idx;
        const cat = getPersonalCategoryInfo(task.category);
        const alert = getDateAlert(task.date);
        const tr = document.createElement('tr');
        tr.style.background = task.category ? cat.bg : '';
        tr.innerHTML = `
            <td class="col-check"><input type="checkbox" class="task-checkbox" onchange="completePersonalTask(${idx})" title="Mark complete"></td>
            <td>${buildPersonalUrgencySelect(task.urgency || '4', idx)}</td>
            <td>${buildPersonalCategorySelect(task.category || 'Personal', idx)}</td>
            <td><input type="text" value="${esc(task.name)}" placeholder="Task name..." onchange="updatePersonalTask(${idx},'name',this.value)"></td>
            <td class="date-cell"><input type="date" value="${esc(task.date)}" onchange="updatePersonalTask(${idx},'date',this.value)">${alert}</td>
            <td><input type="text" value="${esc(task.link)}" placeholder="Link..." onchange="updatePersonalTask(${idx},'link',this.value)"></td>
            <td>${buildPersonalWhoSelect(task.who, idx)}</td>
            <td><textarea rows="1" placeholder="Notes..." onchange="updatePersonalTask(${idx},'notes',this.value)">${esc(task.notes || '')}</textarea></td>
            <td><button class="btn-delete" onclick="deletePersonalTask(${idx})" title="Delete"><span class="material-icons-outlined">delete</span></button></td>
        `;
        tbody.appendChild(tr);
    });
    renderPersonalArchivedToggle();
}

function completePersonalTask(idx) {
    const task = personalTasks.splice(idx, 1)[0];
    task.completedDate = new Date().toISOString().split('T')[0];
    archivedPersonalTasks.unshift(task);
    saveData(STORAGE_KEYS.personalTasks, personalTasks);
    saveData(STORAGE_KEYS.archivedPersonalTasks, archivedPersonalTasks);
    renderPersonalTasks();
}

function resumePersonalTask(idx) {
    const task = archivedPersonalTasks.splice(idx, 1)[0];
    delete task.completedDate;
    personalTasks.push(task);
    saveData(STORAGE_KEYS.personalTasks, personalTasks);
    saveData(STORAGE_KEYS.archivedPersonalTasks, archivedPersonalTasks);
    renderPersonalTasks();
}

function deleteArchivedPersonalTask(idx) {
    archivedPersonalTasks.splice(idx, 1);
    saveData(STORAGE_KEYS.archivedPersonalTasks, archivedPersonalTasks);
    renderPersonalTasks();
}

function renderPersonalArchivedToggle() {
    let section = document.getElementById('personal-archived-section');
    if (!section) {
        section = document.createElement('div');
        section.id = 'personal-archived-section';
        document.getElementById('page-personal-tasks').appendChild(section);
    }
    if (archivedPersonalTasks.length === 0) { section.innerHTML = ''; return; }
    section.innerHTML = `
        <div class="archived-header" onclick="togglePersonalArchived()">
            <span class="material-icons-outlined">inventory_2</span> Completed Tasks (${archivedPersonalTasks.length})
            <span class="material-icons-outlined" id="personal-arch-chevron" style="margin-left:auto; transition:transform 0.3s">expand_more</span>
        </div>
        <div id="personal-archived-table-wrap" class="archived-table-wrap">
            <div class="table-container">
                <table>
                    <thead><tr>
                        <th>Urgency</th><th>Category</th><th>Task</th><th>Date</th><th>Completed</th><th>Link</th><th>Who</th><th>Notes</th><th></th>
                    </tr></thead>
                    <tbody>
                        ${archivedPersonalTasks.map((t, i) => {
                            const aCat = getPersonalCategoryInfo(t.category);
                            return `
                            <tr class="archived-row" style="background:${t.category ? aCat.bg : ''}">
                                <td><span class="urgency-badge" style="background:${getUrgencyInfo(t.urgency || '4').color}">${getUrgencyInfo(t.urgency || '4').label}</span></td>
                                <td><span class="cat-badge" style="background:${aCat.color}">${esc(t.category || 'Personal')}</span></td>
                                <td>${esc(t.name)}</td>
                                <td>${esc(t.date)}</td>
                                <td><span class="completed-badge">${esc(t.completedDate)}</span></td>
                                <td>${esc(t.link)}</td>
                                <td>${esc(t.who)}</td>
                                <td>${esc(t.notes || '')}</td>
                                <td>
                                    <button class="btn-resume" onclick="resumePersonalTask(${i})" title="Resume"><span class="material-icons-outlined">replay</span></button>
                                    <button class="btn-delete" onclick="deleteArchivedPersonalTask(${i})" title="Delete"><span class="material-icons-outlined">delete</span></button>
                                </td>
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;
}

function togglePersonalArchived() {
    const wrap = document.getElementById('personal-archived-table-wrap');
    const chevron = document.getElementById('personal-arch-chevron');
    wrap.classList.toggle('open');
    chevron.style.transform = wrap.classList.contains('open') ? 'rotate(180deg)' : '';
}

function addPersonalTask() {
    personalTasks.push({ id: Date.now(), name: '', who: '', date: '', link: '', category: 'Personal', notes: '', urgency: '4' });
    saveData(STORAGE_KEYS.personalTasks, personalTasks);
    renderPersonalTasks();
    const inputs = document.querySelectorAll('#personal-tasks-body tr:last-child input');
    if (inputs.length) inputs[0].focus();
}

function updatePersonalTask(idx, field, value) {
    personalTasks[idx][field] = value;
    debounceSave(STORAGE_KEYS.personalTasks, personalTasks);
}

function deletePersonalTask(idx) {
    personalTasks.splice(idx, 1);
    saveData(STORAGE_KEYS.personalTasks, personalTasks);
    renderPersonalTasks();
}

// ===== Family Events =====
function renderFamilyEvents() {
    const tbody = document.getElementById('family-events-body');
    tbody.innerHTML = '';
    familyEvents.forEach((ev, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="date" value="${esc(ev.dateFrom)}" onchange="updateFamilyEvent(${idx},'dateFrom',this.value)"></td>
            <td><input type="date" value="${esc(ev.dateTo)}" onchange="updateFamilyEvent(${idx},'dateTo',this.value)"></td>
            <td><input type="text" value="${esc(ev.name)}" placeholder="Event name..." onchange="updateFamilyEvent(${idx},'name',this.value)"></td>
            <td><input type="text" value="${esc(ev.where)}" placeholder="Location..." onchange="updateFamilyEvent(${idx},'where',this.value)"></td>
            <td><input type="text" value="${esc(ev.hotel)}" placeholder="Hotel..." onchange="updateFamilyEvent(${idx},'hotel',this.value)"></td>
            <td><input type="text" value="${esc(ev.car)}" placeholder="Car..." onchange="updateFamilyEvent(${idx},'car',this.value)"></td>
            <td><input type="text" value="${esc(ev.transportation)}" placeholder="Details..." onchange="updateFamilyEvent(${idx},'transportation',this.value)"></td>
            <td><input type="text" value="${esc(ev.who)}" placeholder="Who..." onchange="updateFamilyEvent(${idx},'who',this.value)"></td>
            <td><textarea rows="1" placeholder="Notes..." onchange="updateFamilyEvent(${idx},'notes',this.value)">${esc(ev.notes)}</textarea></td>
            <td><button class="btn-delete" onclick="deleteFamilyEvent(${idx})" title="Delete"><span class="material-icons-outlined">delete</span></button></td>
        `;
        tbody.appendChild(tr);
    });
}

function addFamilyEvent() {
    familyEvents.push({ id: Date.now(), dateFrom: '', dateTo: '', name: '', where: '', hotel: '', car: '', transportation: '', who: '', notes: '' });
    saveData(STORAGE_KEYS.familyEvents, familyEvents);
    renderFamilyEvents();
    const inputs = document.querySelectorAll('#family-events-body tr:last-child input');
    if (inputs.length) inputs[0].focus();
}

function updateFamilyEvent(idx, field, value) {
    familyEvents[idx][field] = value;
    debounceSave(STORAGE_KEYS.familyEvents, familyEvents);
}

function deleteFamilyEvent(idx) {
    familyEvents.splice(idx, 1);
    saveData(STORAGE_KEYS.familyEvents, familyEvents);
    renderFamilyEvents();
}

// ===== Events =====
function renderEvents(type) {
    const data = type === '1p' ? events1p : events3p;
    const tbody = document.getElementById(`events-${type}-body`);
    tbody.innerHTML = '';
    const today = new Date(); today.setHours(0,0,0,0);
    const threeMonths = new Date(today); threeMonths.setMonth(threeMonths.getMonth() + 3);
    data.forEach((ev, idx) => {
        // Only show events in the next 3 months (or with no date set yet)
        if (ev.date) {
            const d = new Date(ev.date + 'T00:00:00');
            if (d < today || d > threeMonths) return;
        }
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="date" value="${esc(ev.date)}" onchange="updateEvent('${type}',${idx},'date',this.value)"></td>
            <td><input type="text" value="${esc(ev.name)}" placeholder="Event name..." onchange="updateEvent('${type}',${idx},'name',this.value)"></td>
            <td><input type="text" value="${esc(ev.hero)}" placeholder="Product..." onchange="updateEvent('${type}',${idx},'hero',this.value)"></td>
            <td><input type="text" value="${esc(ev.pmm)}" placeholder="PMM..." onchange="updateEvent('${type}',${idx},'pmm',this.value)"></td>
            <td><input type="text" value="${esc(ev.contact)}" placeholder="Contact..." onchange="updateEvent('${type}',${idx},'contact',this.value)"></td>
            <td><input type="text" value="${esc(ev.plan)}" placeholder="Plan..." onchange="updateEvent('${type}',${idx},'plan',this.value)"></td>
            <td><input type="text" value="${esc(ev.activity)}" placeholder="Activity..." onchange="updateEvent('${type}',${idx},'activity',this.value)"></td>
            <td><textarea rows="1" placeholder="Notes..." onchange="updateEvent('${type}',${idx},'notes',this.value)">${esc(ev.notes)}</textarea></td>
            <td><button class="btn-delete" onclick="deleteEvent('${type}',${idx})" title="Delete"><span class="material-icons-outlined">delete</span></button></td>
        `;
        tbody.appendChild(tr);
    });
    updateStats();
}

function addEvent(type) {
    const ev = { id: Date.now(), date: '', name: '', hero: '', pmm: '', contact: '', plan: '', activity: '', notes: '' };
    if (type === '1p') { events1p.push(ev); saveData(STORAGE_KEYS.events1p, events1p); }
    else { events3p.push(ev); saveData(STORAGE_KEYS.events3p, events3p); }
    renderEvents(type);
    const inputs = document.querySelectorAll(`#events-${type}-body tr:last-child input`);
    if (inputs.length) inputs[0].focus();
}

function updateEvent(type, idx, field, value) {
    if (type === '1p') { events1p[idx][field] = value; debounceSave(STORAGE_KEYS.events1p, events1p); }
    else { events3p[idx][field] = value; debounceSave(STORAGE_KEYS.events3p, events3p); }
}

function deleteEvent(type, idx) {
    if (type === '1p') { events1p.splice(idx, 1); saveData(STORAGE_KEYS.events1p, events1p); renderEvents('1p'); }
    else { events3p.splice(idx, 1); saveData(STORAGE_KEYS.events3p, events3p); renderEvents('3p'); }
}

// ===== Hero Products =====
function getPmmList(pmmStr) {
    if (!pmmStr) return [];
    return pmmStr.split(',').map(s => s.trim()).filter(Boolean);
}

function renderBubbles(names, prodIdx, field) {
    return names.map((name, i) =>
        `<span class="bubble">${esc(name)}<button class="bubble-x" onclick="event.stopPropagation(); removePmm(${prodIdx},'${field}',${i})" title="Remove">&times;</button></span>`
    ).join('');
}

function removePmm(prodIdx, field, nameIdx) {
    const list = getPmmList(products[prodIdx][field]);
    list.splice(nameIdx, 1);
    products[prodIdx][field] = list.join(', ');
    saveData(STORAGE_KEYS.products, products);
    renderProducts();
}

function addPmmToProduct() {
    const sel = document.getElementById('pmm-product-select');
    const input = document.getElementById('pmm-name-input');
    const idx = parseInt(sel.value);
    const name = input.value.trim();
    if (isNaN(idx) || !name) return;
    const list = getPmmList(products[idx].pmm);
    list.push(name);
    products[idx].pmm = list.join(', ');
    saveData(STORAGE_KEYS.products, products);
    input.value = '';
    renderProducts();
}

function addPmmFromDetail(idx) {
    const input = document.getElementById('detail-pmm-input');
    const name = input.value.trim();
    if (!name) return;
    const list = getPmmList(products[idx].pmm);
    list.push(name);
    products[idx].pmm = list.join(', ');
    saveData(STORAGE_KEYS.products, products);
    input.value = '';
    showProductDetail(idx);
}

function renderProducts() {
    const grid = document.getElementById('products-grid');
    grid.innerHTML = '';
    products.forEach((prod, idx) => {
        const pmmNames = getPmmList(prod.pmm);
        const managerName = prod.pmmManager || '';
        const card = document.createElement('div');
        card.className = 'product-card';
        card.onclick = () => showProductDetail(idx);
        card.innerHTML = `
            <h3>${esc(prod.name)}</h3>
            ${managerName ? `<div class="card-manager"><span class="material-icons-outlined">manage_accounts</span> ${esc(managerName)}</div>` : ''}
            <div class="card-label">PMM</div>
            <div class="bubble-wrap">${renderBubbles(pmmNames, idx, 'pmm')}</div>
        `;
        grid.appendChild(card);
    });

    // Populate product selector
    const sel = document.getElementById('pmm-product-select');
    if (sel) {
        sel.innerHTML = products.map((p, i) => `<option value="${i}">${esc(p.name)}</option>`).join('');
    }
}

function showProductDetail(idx) {
    const prod = products[idx];
    document.getElementById('product-detail-title').textContent = prod.name;

    const content = document.getElementById('product-detail-content');
    const pmmBubbles = renderBubbles(getPmmList(prod.pmm), idx, 'pmm');
    content.innerHTML = `
        <div class="detail-section">
            <h3><span class="material-icons-outlined">people</span> Team</h3>
            <div class="detail-fields">
                <div class="detail-field"><label>PMM Manager</label><input value="${esc(prod.pmmManager)}" onchange="updateProduct(${idx},'pmmManager',this.value)"></div>
                <div class="detail-field"><label>SME</label><input value="${esc(prod.sme)}" onchange="updateProduct(${idx},'sme',this.value)"></div>
                <div class="detail-field"><label>Global Skilling</label><input value="${esc(prod.globalSkilling)}" onchange="updateProduct(${idx},'globalSkilling',this.value)"></div>
            </div>
            <div class="detail-field" style="margin-top:16px">
                <label>PMMs</label>
                <div class="bubble-wrap">${pmmBubbles}</div>
                <div class="pmm-add-inline">
                    <input type="text" id="detail-pmm-input" placeholder="Add PMM name...">
                    <button class="btn-primary btn-sm" onclick="addPmmFromDetail(${idx})"><span class="material-icons-outlined">add</span></button>
                </div>
            </div>
        </div>
        <div class="detail-section">
            <h3><span class="material-icons-outlined">update</span> Product Updates</h3>
            <div class="table-container"><table><thead><tr><th>Name</th><th>Date</th><th>Details</th><th class="col-actions"></th></tr></thead><tbody id="product-updates-body"></tbody></table></div>
            <button class="btn-primary" onclick="addProductUpdate(${idx})"><span class="material-icons-outlined">add</span> Add Update</button>
        </div>
        <div class="detail-section">
            <h3><span class="material-icons-outlined">event</span> Events</h3>
            <div class="table-container"><table><thead><tr><th>Event Name</th><th>Location</th><th>Date</th><th>Plan of Record</th><th class="col-actions"></th></tr></thead><tbody id="product-events-body"></tbody></table></div>
            <button class="btn-primary" onclick="addProductEvent(${idx})"><span class="material-icons-outlined">add</span> Add Event</button>
        </div>
    `;

    renderProductUpdates(idx);
    renderProductEvents(idx);
    showPage('product-detail');
}

function renderProductUpdates(idx) {
    const tbody = document.getElementById('product-updates-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    (products[idx].updates || []).forEach((upd, ui) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input value="${esc(upd.name)}" placeholder="Update name..." onchange="updateProductSub(${idx},'updates',${ui},'name',this.value)"></td>
            <td><input type="date" value="${esc(upd.date)}" onchange="updateProductSub(${idx},'updates',${ui},'date',this.value)"></td>
            <td><textarea rows="2" placeholder="Details..." onchange="updateProductSub(${idx},'updates',${ui},'details',this.value)">${esc(upd.details)}</textarea></td>
            <td><button class="btn-delete" onclick="deleteProductSub(${idx},'updates',${ui})"><span class="material-icons-outlined">delete</span></button></td>
        `;
        tbody.appendChild(tr);
    });
}

function renderProductEvents(idx) {
    const tbody = document.getElementById('product-events-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    (products[idx].events || []).forEach((ev, ei) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input value="${esc(ev.name)}" placeholder="Event..." onchange="updateProductSub(${idx},'events',${ei},'name',this.value)"></td>
            <td><input value="${esc(ev.location)}" placeholder="Location..." onchange="updateProductSub(${idx},'events',${ei},'location',this.value)"></td>
            <td><input type="date" value="${esc(ev.date)}" onchange="updateProductSub(${idx},'events',${ei},'date',this.value)"></td>
            <td><input value="${esc(ev.plan)}" placeholder="Plan..." onchange="updateProductSub(${idx},'events',${ei},'plan',this.value)"></td>
            <td><button class="btn-delete" onclick="deleteProductSub(${idx},'events',${ei})"><span class="material-icons-outlined">delete</span></button></td>
        `;
        tbody.appendChild(tr);
    });
}

function updateProduct(idx, field, value) {
    products[idx][field] = value;
    debounceSave(STORAGE_KEYS.products, products);
}

function updateProductSub(prodIdx, arr, subIdx, field, value) {
    products[prodIdx][arr][subIdx][field] = value;
    debounceSave(STORAGE_KEYS.products, products);
}

function addProductUpdate(idx) {
    if (!products[idx].updates) products[idx].updates = [];
    products[idx].updates.push({ name: '', date: '', details: '' });
    saveData(STORAGE_KEYS.products, products);
    renderProductUpdates(idx);
}

function addProductEvent(idx) {
    if (!products[idx].events) products[idx].events = [];
    products[idx].events.push({ name: '', location: '', date: '', plan: '' });
    saveData(STORAGE_KEYS.products, products);
    renderProductEvents(idx);
}

function deleteProductSub(prodIdx, arr, subIdx) {
    products[prodIdx][arr].splice(subIdx, 1);
    saveData(STORAGE_KEYS.products, products);
    if (arr === 'updates') renderProductUpdates(prodIdx);
    else renderProductEvents(prodIdx);
}

// ===== Sorting =====
document.addEventListener('click', function(e) {
    const th = e.target.closest('th[data-sort]');
    if (!th) return;
    const table = th.closest('table');
    const col = th.dataset.sort;
    const tableId = table.id;

    if (currentSort.table === tableId && currentSort.column === col) {
        currentSort.dir = currentSort.dir === 'asc' ? 'desc' : 'asc';
    } else {
        currentSort = { table: tableId, column: col, dir: 'asc' };
    }

    let dataArr, key;
    if (tableId === 'tasks-table') { dataArr = tasks; key = STORAGE_KEYS.tasks; }
    else if (tableId === 'events-1p-table') { dataArr = events1p; key = STORAGE_KEYS.events1p; }
    else if (tableId === 'events-3p-table') { dataArr = events3p; key = STORAGE_KEYS.events3p; }
    else return;

    dataArr.sort((a, b) => {
        const va = (a[col] || '').toString().toLowerCase();
        const vb = (b[col] || '').toString().toLowerCase();
        return currentSort.dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    });

    saveData(key, dataArr);
    if (tableId === 'tasks-table') renderTasks();
    else if (tableId === 'events-1p-table') renderEvents('1p');
    else if (tableId === 'events-3p-table') renderEvents('3p');
});

// ===== Stats & Home =====
function updateStats() {
    renderHomeDashboard();
}

function renderHomeDashboard() {
    const today = new Date(); today.setHours(0,0,0,0);

    // Urgent tasks: urgency 1-2 or due within 7 days or overdue
    const urgentTasks = tasks.filter(t => {
        if (!t.name) return false;
        const isHighUrgency = t.urgency === '1' || t.urgency === '2';
        let isDueSoon = false;
        if (t.date) {
            const due = new Date(t.date + 'T00:00:00');
            isDueSoon = (due - today) / (1000*60*60*24) <= 7;
        }
        return isHighUrgency || isDueSoon;
    }).sort((a, b) => {
        const ua = parseInt(a.urgency || '4');
        const ub = parseInt(b.urgency || '4');
        if (ua !== ub) return ua - ub;
        return (a.date || '').localeCompare(b.date || '');
    });

    const urgentList = document.getElementById('home-urgent-list');
    if (urgentList) {
        if (urgentTasks.length === 0) {
            urgentList.innerHTML = '<div class="home-empty">No urgent tasks — nice!</div>';
        } else {
            urgentList.innerHTML = urgentTasks.map(t => {
                const u = getUrgencyInfo(t.urgency || '4');
                const alert = getDateAlert(t.date);
                const dateStr = t.date ? new Date(t.date + 'T00:00:00').toLocaleDateString('en-US', {month:'short', day:'numeric'}) : '';
                return `<div class="home-task-row">
                    <span class="urgency-dot" style="background:${u.color}" title="${u.label}"></span>
                    <span class="home-task-name">${esc(t.name)}</span>
                    <span class="home-task-who">${esc(t.who || '')}</span>
                    <span class="home-task-date">${dateStr} ${alert}</span>
                </div>`;
            }).join('');
        }
    }

    // Upcoming events (next 3 months, combined 1P + 3P)
    const threeMonths = new Date(today);
    threeMonths.setMonth(threeMonths.getMonth() + 3);
    const allEvents = [...events1p.map(e => ({...e, type:'1P'})), ...events3p.map(e => ({...e, type:'3P'}))];
    const upcoming = allEvents.filter(e => {
        if (!e.date || !e.name) return false;
        const d = new Date(e.date + 'T00:00:00');
        return d >= today && d <= threeMonths;
    }).sort((a, b) => (a.date || '').localeCompare(b.date || ''));

    const eventsList = document.getElementById('home-events-list');
    if (eventsList) {
        if (upcoming.length === 0) {
            eventsList.innerHTML = '<div class="home-empty">No events in the next 3 months</div>';
        } else {
            eventsList.innerHTML = upcoming.map(e => {
                const dateStr = new Date(e.date + 'T00:00:00').toLocaleDateString('en-US', {month:'short', day:'numeric'});
                return `<div class="home-event-row">
                    <span class="home-event-badge">${e.type}</span>
                    <span class="home-event-name">${esc(e.name)}</span>
                    <span class="home-event-date">${dateStr}</span>
                </div>`;
            }).join('');
        }
    }

    // Personal tasks recap
    const personalRecap = document.getElementById('home-personal-list');
    if (personalRecap) {
        const urgentPersonal = personalTasks.filter(t => {
            if (!t.name) return false;
            const isHigh = t.urgency === '1' || t.urgency === '2';
            let isDueSoon = false;
            if (t.date) {
                const due = new Date(t.date + 'T00:00:00');
                isDueSoon = (due - today) / (1000*60*60*24) <= 7;
            }
            return isHigh || isDueSoon;
        }).sort((a, b) => {
            const ua = parseInt(a.urgency || '4');
            const ub = parseInt(b.urgency || '4');
            if (ua !== ub) return ua - ub;
            return (a.date || '').localeCompare(b.date || '');
        });
        if (urgentPersonal.length === 0) {
            personalRecap.innerHTML = '<div class="home-empty">No urgent personal tasks</div>';
        } else {
            personalRecap.innerHTML = urgentPersonal.map(t => {
                const u = getUrgencyInfo(t.urgency || '4');
                const cat = getPersonalCategoryInfo(t.category);
                const dateStr = t.date ? new Date(t.date + 'T00:00:00').toLocaleDateString('en-US', {month:'short', day:'numeric'}) : '';
                return `<div class="home-task-row">
                    <span class="urgency-dot" style="background:${u.color}" title="${u.label}"></span>
                    <span class="home-event-badge" style="background:${cat.color}">${esc(t.category || 'Personal')}</span>
                    <span class="home-task-name">${esc(t.name)}</span>
                    <span class="home-task-date">${dateStr}</span>
                </div>`;
            }).join('');
        }
    }

    // Family events on home
    const familyList = document.getElementById('home-family-events-list');
    if (familyList) {
        const upcomingFamily = familyEvents.filter(e => {
            if (!e.name) return false;
            if (!e.dateFrom) return true; // show events without dates
            const d = new Date(e.dateFrom + 'T00:00:00');
            return d >= today;
        }).sort((a, b) => (a.dateFrom || '').localeCompare(b.dateFrom || ''));
        if (upcomingFamily.length === 0) {
            familyList.innerHTML = '<div class="home-empty">No upcoming family events</div>';
        } else {
            familyList.innerHTML = upcomingFamily.map(e => {
                let dateStr = '';
                if (e.dateFrom) {
                    dateStr = new Date(e.dateFrom + 'T00:00:00').toLocaleDateString('en-US', {month:'short', day:'numeric'});
                    if (e.dateTo) dateStr += ' – ' + new Date(e.dateTo + 'T00:00:00').toLocaleDateString('en-US', {month:'short', day:'numeric'});
                }
                return `<div class="home-event-row">
                    <span class="home-event-badge" style="background:#2980b9">Family</span>
                    <span class="home-event-name">${esc(e.name)}</span>
                    <span class="home-event-date">${dateStr}</span>
                </div>`;
            }).join('');
        }
    }
}

// ===== Open in Browser =====
document.addEventListener('click', function(e) {
    const link = e.target.closest('#open-in-browser');
    if (link) {
        e.preventDefault();
        const url = window.location.href;
        navigator.clipboard.writeText(url).then(() => {
            link.innerHTML = '<span class="material-icons-outlined" style="color:var(--neon-green)">check</span>';
            setTimeout(() => { link.innerHTML = '<span class="material-icons-outlined">open_in_new</span>'; }, 2000);
        }).catch(() => { prompt('Copy this URL:', url); });
    }
});

// ===== Greeting =====
function setGreeting() {
    // Convert to PST (UTC-7 / UTC-8 depending on DST)
    const now = new Date();
    const pstOptions = { timeZone: 'America/Los_Angeles', hour: 'numeric', hour12: false };
    const pstHour = parseInt(new Intl.DateTimeFormat('en-US', pstOptions).format(now));
    let greeting;
    if (pstHour >= 21 || pstHour < 5) greeting = 'Buonanotte';
    else if (pstHour >= 17) greeting = 'Buonasera';
    else greeting = 'Buongiorno';
    const el = document.getElementById('greeting-time');
    if (el) el.textContent = greeting;

    const dateEl = document.getElementById('current-date');
    if (dateEl) {
        dateEl.textContent = now.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    }
}

// ===== Top News via RSS =====
async function renderNews() {
    const container = document.getElementById('home-news-list');
    if (!container) return;
    container.innerHTML = '<div class="home-empty">Loading news...</div>';
    try {
        const feeds = [
            { name: 'Google', rss: 'https://news.google.com/rss/topics/CAAqJggKIiBDQkFTRWdvSUwyMHZNRGRqTVhZU0FtVnVHZ0pWVXlnQVAB?hl=en-US&gl=US&ceid=US:en' },
            { name: 'BBC', rss: 'https://feeds.bbci.co.uk/news/technology/rss.xml' }
        ];
        const articles = [];
        for (const feed of feeds) {
            try {
                const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(feed.rss)}`;
                const resp = await fetch(proxyUrl);
                if (!resp.ok) continue;
                const text = await resp.text();
                const parser = new DOMParser();
                const xml = parser.parseFromString(text, 'text/xml');
                const items = xml.querySelectorAll('item');
                items.forEach((item, i) => {
                    if (i >= 4) return;
                    const title = item.querySelector('title')?.textContent || '';
                    const link = item.querySelector('link')?.textContent || '';
                    const pubDate = item.querySelector('pubDate')?.textContent || '';
                    if (title) articles.push({ title, link, source: feed.name, date: pubDate });
                });
            } catch { /* skip failed feed */ }
        }
        if (articles.length === 0) {
            container.innerHTML = '<div class="home-empty">Unable to load news — <a href="https://news.google.com/topics/CAAqJggKIiBDQkFTRWdvSUwyMHZNRGRqTVhZU0FtVnVHZ0pWVXlnQVAB" target="_blank" style="color:var(--accent)">open Google News</a></div>';
            return;
        }
        container.innerHTML = articles.slice(0, 8).map(a =>
            `<a href="${esc(a.link)}" target="_blank" class="home-news-row">
                <span class="home-news-source">${esc(a.source)}</span>
                <span class="home-news-title">${esc(a.title)}</span>
            </a>`
        ).join('');
    } catch {
        container.innerHTML = '<div class="home-empty">Unable to load news</div>';
    }
}

// ===== Tech Stocks =====
async function renderStocks() {
    const container = document.getElementById('home-stocks-list');
    if (!container) return;
    const stocks = [
        { symbol: 'MSFT', name: 'Microsoft' },
        { symbol: 'AAPL', name: 'Apple' },
        { symbol: 'GOOGL', name: 'Alphabet' },
        { symbol: 'AMZN', name: 'Amazon' },
        { symbol: 'NVDA', name: 'NVIDIA' },
        { symbol: 'META', name: 'Meta' },
        { symbol: 'TSLA', name: 'Tesla' }
    ];

    // Try fetching live quotes from Yahoo via allorigins proxy
    try {
        const symbols = stocks.map(s => s.symbol).join(',');
        const yUrl = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${symbols}&fields=regularMarketPrice,regularMarketChange,regularMarketChangePercent`;
        const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(yUrl)}`;
        const resp = await fetch(proxyUrl);
        if (resp.ok) {
            const data = await resp.json();
            if (data.quoteResponse && data.quoteResponse.result) {
                const quotes = data.quoteResponse.result;
                container.innerHTML = stocks.map(s => {
                    const q = quotes.find(r => r.symbol === s.symbol);
                    if (!q) return renderStockFallback(s);
                    const price = q.regularMarketPrice?.toFixed(2) || '—';
                    const change = q.regularMarketChange?.toFixed(2) || '0';
                    const pct = q.regularMarketChangePercent?.toFixed(2) || '0';
                    const isUp = parseFloat(change) >= 0;
                    const color = isUp ? '#22c55e' : '#ef4444';
                    const arrow = isUp ? '▲' : '▼';
                    return `<a href="https://finance.yahoo.com/quote/${s.symbol}" target="_blank" class="home-stock-row">
                        <span class="home-stock-symbol">${s.symbol}</span>
                        <span class="home-stock-name">${s.name}</span>
                        <span class="home-stock-price">$${price}</span>
                        <span class="home-stock-change" style="color:${color}">${arrow} ${change} (${pct}%)</span>
                    </a>`;
                }).join('');
                return;
            }
        }
    } catch { /* fallback below */ }

    // Fallback: static links to Yahoo Finance
    container.innerHTML = stocks.map(s => renderStockFallback(s)).join('');
}

function renderStockFallback(s) {
    return `<a href="https://finance.yahoo.com/quote/${s.symbol}" target="_blank" class="home-stock-row">
        <span class="home-stock-symbol">${s.symbol}</span>
        <span class="home-stock-name">${s.name}</span>
        <span class="home-stock-price" style="color:#888">View →</span>
    </a>`;
}

// ===== Init =====
document.addEventListener('DOMContentLoaded', async () => {
    document.querySelectorAll('.link-group-header').forEach((h, i) => {
        if (i > 0) h.classList.add('collapsed');
    });

    // Restore last sync time
    lastSyncTime = localStorage.getItem('pa_last_sync');

    // Try loading from GitHub first
    const token = getGitHubToken();
    if (token) {
        const loaded = await loadFromGitHub();
        if (loaded) reloadAllState();
    } else {
        updateSyncIndicator('notoken');
    }

    setGreeting();
    showPage('welcome');
    renderNews();
    renderStocks();

    // Periodic auto-sync every 60 seconds
    setInterval(() => {
        if (syncPending) syncToGitHub();
    }, 60000);

    // Sync before page close
    window.addEventListener('beforeunload', () => {
        if (syncPending) {
            // Synchronous localStorage backup is always safe
            saveLocalBackup();
            // Best effort sync via sendBeacon
            const token = getGitHubToken();
            if (token) {
                const data = JSON.stringify({
                    message: 'Auto-sync on close ' + new Date().toISOString(),
                    content: btoa(unescape(encodeURIComponent(JSON.stringify(getAllData(), null, 2)))),
                    branch: GITHUB_BRANCH,
                    ...(githubSha ? { sha: githubSha } : {})
                });
                navigator.sendBeacon && navigator.sendBeacon(
                    `https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_DATA_FILE}`,
                    new Blob([data], { type: 'application/json' })
                );
            }
        }
    });

    // Sync when tab becomes visible again
    document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible' && getGitHubToken()) {
            loadFromGitHub().then(loaded => {
                if (loaded) reloadAllState();
            });
        }
    });
});
