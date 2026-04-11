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

let syncPending = false;
let syncTimer = null;
let lastSyncTime = null;
let isAuthenticated = false;

function loadData(key, defaults) {
    try {
        const saved = localStorage.getItem(key);
        return saved ? JSON.parse(saved) : defaults;
    } catch { return defaults; }
}

function saveData(key, data) {
    localStorage.setItem(key, JSON.stringify(data));
    saveLocalBackup();
    scheduleSyncToServer();
}

function saveLocalBackup() {
    try {
        localStorage.setItem('pa_local_backup', JSON.stringify(getAllData()));
        localStorage.setItem('pa_local_backup_time', new Date().toISOString());
    } catch { /* quota exceeded — ok */ }
}

function scheduleSyncToServer() {
    if (!isAuthenticated) return;
    syncPending = true;
    updateSyncIndicator('pending');
    clearTimeout(syncTimer);
    syncTimer = setTimeout(() => {
        syncToServer().catch(err => console.error('Sync error:', err));
    }, 1000);
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

function updateSyncIndicator(state) {
    const el = document.getElementById('sync-status');
    if (!el) return;
    const timeStr = lastSyncTime ? new Date(lastSyncTime).toLocaleTimeString() : '';
    if (state === 'syncing') {
        el.innerHTML = '<span class="material-icons-outlined spin">sync</span>';
        el.title = 'Syncing to server...';
    } else if (state === 'ok') {
        el.innerHTML = '<span class="material-icons-outlined" style="color:#86efac">cloud_done</span>';
        el.title = 'Synced' + (timeStr ? ' at ' + timeStr : '');
    } else if (state === 'pending') {
        el.innerHTML = '<span class="material-icons-outlined" style="color:#fbbf24">cloud_upload</span>';
        el.title = 'Changes pending sync...';
    } else if (state === 'error') {
        el.innerHTML = '<span class="material-icons-outlined" style="color:#f87171">cloud_off</span>';
        el.title = 'Sync failed';
    }
    const tsEl = document.getElementById('last-sync-time');
    if (tsEl) tsEl.textContent = timeStr ? 'Last sync: ' + timeStr : '';
}

async function syncToServer() {
    if (!isAuthenticated) return;
    updateSyncIndicator('syncing');
    try {
        const localData = getAllData();
        // Data loss protection: check if we'd overwrite substantial server data with empty local data
        const localItemCount = Object.entries(STORAGE_KEYS).reduce((sum, [name]) => {
            const arr = localData[name];
            return sum + (Array.isArray(arr) ? arr.length : 0);
        }, 0);
        if (localItemCount === 0) {
            // Don't sync completely empty data — likely a fresh/broken session
            console.warn('Sync blocked: local data is completely empty, refusing to overwrite server');
            updateSyncIndicator('ok');
            return;
        }
        const resp = await fetch('/api/data', {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(localData)
        });
        if (resp.ok) {
            const result = await resp.json();
            syncPending = false;
            lastSyncTime = result.savedAt || new Date().toISOString();
            localStorage.setItem('pa_last_sync', lastSyncTime);
            updateSyncIndicator('ok');
        } else if (resp.status === 401) {
            showLoginScreen();
        } else {
            console.error('Sync failed:', resp.status);
            updateSyncIndicator('error');
        }
    } catch (err) {
        console.error('Sync error:', err);
        updateSyncIndicator('error');
    }
}

async function loadFromServer() {
    if (!isAuthenticated) return false;
    try {
        const resp = await fetch('/api/data');
        if (resp.ok) {
            const data = await resp.json();
            if (data && Object.keys(data).length > 0 && data.tasks) {
                // Only apply server data if it's newer than local backup
                const localBackupTime = localStorage.getItem('pa_local_backup_time');
                const serverTime = data._savedAt;
                if (localBackupTime && serverTime && new Date(localBackupTime) > new Date(serverTime)) {
                    // Local is newer — push local data to server instead
                    updateSyncIndicator('pending');
                    syncToServer();
                    return false;
                }
                applyAllData(data);
                updateSyncIndicator('ok');
                return true;
            }
            return false;
        } else if (resp.status === 401) {
            showLoginScreen();
            return false;
        } else {
            updateSyncIndicator('error');
            return false;
        }
    } catch (err) {
        console.error('Load error:', err);
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
                scheduleSyncToServer();
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
    modal.style.display = 'flex';
}

function closeSettings() {
    document.getElementById('settings-modal').style.display = 'none';
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

const DEFAULT_PERSONAL_TASKS = [];
const DEFAULT_FAMILY_EVENTS = [];

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

const URGENCY_LEVELS = [
    { value: '1', label: 'Urgent 1', color: '#e74c3c', bg: '#fdedec' },
    { value: '2', label: 'Urgent 2', color: '#f39c12', bg: '#fef9e7' },
    { value: '3', label: 'Urgent 3', color: '#27ae60', bg: '#eafaf1' },
    { value: '4', label: 'Urgent 4', color: '#3498db', bg: '#ebf5fb' }
];

const EVENT_PRIORITIES = [
    { value: 'P1', label: 'P1', color: '#e74c3c', bg: '#fdedec' },
    { value: 'P2', label: 'P2', color: '#f39c12', bg: '#fef9e7' },
    { value: 'P3', label: 'P3', color: '#27ae60', bg: '#eafaf1' }
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
function getEventPriorityInfo(val) {
    return EVENT_PRIORITIES.find(p => p.value === val) || EVENT_PRIORITIES[2];
}

function buildEventPrioritySelect(type, idx, selected) {
    const p = getEventPriorityInfo(selected);
    return `<select class="urgency-select" style="background:${p.color};color:#fff" onchange="updateEvent('${type}',${idx},'priority',this.value); this.style.background=getEventPriorityInfo(this.value).color;">
        ${EVENT_PRIORITIES.map(pr => `<option value="${pr.value}" ${pr.value === selected ? 'selected' : ''} style="background:#fff;color:#333">${pr.label}</option>`).join('')}
    </select>`;
}

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
            <td>${buildEventPrioritySelect(type, idx, ev.priority || 'P3')}</td>
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
    const ev = { id: Date.now(), priority: 'P3', date: '', name: '', hero: '', pmm: '', contact: '', plan: '', activity: '', notes: '' };
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
                const pri = getEventPriorityInfo(e.priority || 'P3');
                return `<div class="home-event-row">
                    <span class="urgency-dot" style="background:${pri.color}" title="${pri.label}"></span>
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

// ===== Top News =====
function renderNews() {
    // News is now rendered via embedded RSS widget in HTML — no JS needed
}

// ===== Tech Stocks via TradingView =====
function renderStocks() {
    const container = document.getElementById('home-stocks-list');
    if (!container) return;
    const stocks = ['MSFT', 'AAPL', 'GOOGL', 'AMZN', 'NVDA', 'META', 'TSLA'];
    // Create a fresh widget container
    container.innerHTML = '';
    const widgetDiv = document.createElement('div');
    widgetDiv.className = 'tradingview-widget-container';
    const innerDiv = document.createElement('div');
    widgetDiv.appendChild(innerDiv);
    container.appendChild(widgetDiv);

    const config = {
        symbols: stocks.map(s => ({ proName: 'NASDAQ:' + s, title: s })),
        showSymbolLogo: true,
        colorTheme: 'light',
        isTransparent: true,
        displayMode: 'regular',
        locale: 'en'
    };

    const script = document.createElement('script');
    script.type = 'text/javascript';
    script.src = 'https://s3.tradingview.com/external-embedding/embed-widget-ticker-tape.js';
    script.async = true;
    script.textContent = JSON.stringify(config);
    widgetDiv.appendChild(script);
}

// ===== Auth =====
async function checkAuth() {
    try {
        const resp = await fetch('/api/verify');
        if (resp.ok) {
            const data = await resp.json();
            return data.authenticated === true;
        }
        return false;
    } catch {
        return false;
    }
}

function showLoginScreen() {
    isAuthenticated = false;
    document.getElementById('login-screen').style.display = 'flex';
    document.getElementById('app-container').style.display = 'none';
}

function showApp() {
    isAuthenticated = true;
    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('app-container').style.display = '';
}

async function handleLogin(e) {
    e.preventDefault();
    const btn = document.getElementById('login-btn');
    const errEl = document.getElementById('login-error');
    const password = document.getElementById('login-password').value;

    btn.disabled = true;
    errEl.style.display = 'none';

    try {
        const resp = await fetch('/api/login', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ password })
        });
        if (resp.ok) {
            showApp();
            await initApp();
        } else {
            errEl.textContent = 'Invalid password. Please try again.';
            errEl.style.display = 'block';
        }
    } catch (err) {
        errEl.textContent = 'Connection error. Please try again.';
        errEl.style.display = 'block';
    }

    btn.disabled = false;
}

async function handleLogout() {
    try {
        await fetch('/api/logout', { method: 'POST' });
    } catch { /* ignore */ }
    showLoginScreen();
}

// ===== Init =====
async function initApp() {
    document.querySelectorAll('.link-group-header').forEach((h, i) => {
        if (i > 0) h.classList.add('collapsed');
    });

    lastSyncTime = localStorage.getItem('pa_last_sync');

    // Load data from server
    const loaded = await loadFromServer();
    if (loaded) reloadAllState();

    setGreeting();
    showPage('welcome');
    renderNews();
    renderStocks();

    // Periodic auto-sync every 60 seconds
    setInterval(() => {
        if (syncPending) syncToServer();
    }, 60000);

    // Save to localStorage backup before page close
    window.addEventListener('beforeunload', () => {
        saveLocalBackup();
        if (syncPending) {
            // Data loss protection: don't sendBeacon if local data is empty
            const data = getAllData();
            const itemCount = Object.entries(STORAGE_KEYS).reduce((sum, [name]) => {
                const arr = data[name];
                return sum + (Array.isArray(arr) ? arr.length : 0);
            }, 0);
            if (itemCount > 0) {
                const blob = new Blob([JSON.stringify(data)], { type: 'application/json' });
                navigator.sendBeacon('/api/data', blob);
            }
        }
    });
}

document.addEventListener('DOMContentLoaded', async () => {
    const authed = await checkAuth();
    if (authed) {
        showApp();
        await initApp();
    } else {
        showLoginScreen();
    }
});
