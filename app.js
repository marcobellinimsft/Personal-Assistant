// ===== Data Store =====
const STORAGE_KEYS = {
    tasks: 'pa_tasks',
    events1p: 'pa_events_1p',
    events3p: 'pa_events_3p',
    products: 'pa_products'
};

function loadData(key, defaults) {
    try {
        const saved = localStorage.getItem(key);
        return saved ? JSON.parse(saved) : defaults;
    } catch { return defaults; }
}

function saveData(key, data) {
    localStorage.setItem(key, JSON.stringify(data));
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
let events1p = loadData(STORAGE_KEYS.events1p, DEFAULT_EVENTS_1P);
let events3p = loadData(STORAGE_KEYS.events3p, DEFAULT_EVENTS_3P);
let products = loadData(STORAGE_KEYS.products, DEFAULT_PRODUCTS);
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
        if (pageId === 'welcome') updateStats();
    }
}

// ===== Sidebar Toggle =====
function toggleGroup(el) {
    el.classList.toggle('collapsed');
    el.nextElementSibling.classList.toggle('open');
}

// ===== Tasks =====
function renderTasks() {
    const tbody = document.getElementById('tasks-body');
    tbody.innerHTML = '';
    tasks.forEach((task, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" value="${esc(task.name)}" placeholder="Task name..." onchange="updateTask(${idx},'name',this.value)"></td>
            <td><input type="text" value="${esc(task.who)}" placeholder="Person..." onchange="updateTask(${idx},'who',this.value)"></td>
            <td><input type="date" value="${esc(task.date)}" onchange="updateTask(${idx},'date',this.value)"></td>
            <td><input type="text" value="${esc(task.link)}" placeholder="Link..." onchange="updateTask(${idx},'link',this.value)"></td>
            <td><input type="text" value="${esc(task.vendor)}" placeholder="Vendor..." onchange="updateTask(${idx},'vendor',this.value)"></td>
            <td><button class="btn-delete" onclick="deleteTask(${idx})" title="Delete"><span class="material-icons-outlined">delete</span></button></td>
        `;
        tbody.appendChild(tr);
    });
    updateStats();
}

function addTask() {
    tasks.push({ id: Date.now(), name: '', who: '', date: '', link: '', vendor: '' });
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

// ===== Events =====
function renderEvents(type) {
    const data = type === '1p' ? events1p : events3p;
    const tbody = document.getElementById(`events-${type}-body`);
    tbody.innerHTML = '';
    data.forEach((ev, idx) => {
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
function renderProducts() {
    const grid = document.getElementById('products-grid');
    grid.innerHTML = '';
    products.forEach((prod, idx) => {
        const card = document.createElement('div');
        card.className = 'product-card';
        card.onclick = () => showProductDetail(idx);
        card.innerHTML = `<h3>${esc(prod.name)}</h3><div class="pmm-list"><strong>PMM:</strong> ${esc(prod.pmm)}</div>`;
        grid.appendChild(card);
    });
}

function showProductDetail(idx) {
    const prod = products[idx];
    document.getElementById('product-detail-title').textContent = prod.name;

    const content = document.getElementById('product-detail-content');
    content.innerHTML = `
        <div class="detail-section">
            <h3><span class="material-icons-outlined">people</span> Team</h3>
            <div class="detail-fields">
                <div class="detail-field"><label>PMM</label><input value="${esc(prod.pmm)}" onchange="updateProduct(${idx},'pmm',this.value)"></div>
                <div class="detail-field"><label>PMM Manager</label><input value="${esc(prod.pmmManager)}" onchange="updateProduct(${idx},'pmmManager',this.value)"></div>
                <div class="detail-field"><label>SME</label><input value="${esc(prod.sme)}" onchange="updateProduct(${idx},'sme',this.value)"></div>
                <div class="detail-field"><label>Global Skilling</label><input value="${esc(prod.globalSkilling)}" onchange="updateProduct(${idx},'globalSkilling',this.value)"></div>
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

// ===== Stats =====
function updateStats() {
    const el = id => document.getElementById(id);
    if (el('stat-tasks')) el('stat-tasks').textContent = tasks.filter(t => t.name).length;
    if (el('stat-events-1p')) el('stat-events-1p').textContent = events1p.length;
    if (el('stat-events-3p')) el('stat-events-3p').textContent = events3p.length;
    if (el('stat-products')) el('stat-products').textContent = products.length;
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
    const h = new Date().getHours();
    const greeting = h < 12 ? 'morning' : h < 17 ? 'afternoon' : 'evening';
    const el = document.getElementById('greeting-time');
    if (el) el.textContent = greeting;

    const dateEl = document.getElementById('current-date');
    if (dateEl) {
        dateEl.textContent = new Date().toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    }
}

// ===== Init =====
document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('.link-group-header').forEach((h, i) => {
        if (i > 0) h.classList.add('collapsed');
    });

    setGreeting();
    showPage('welcome');
});
