// ===== Data Store =====
const STORAGE_KEYS = {
    tasks: 'pa_tasks',
    archivedTasks: 'pa_tasks_archived',
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
let archivedTasks = loadData(STORAGE_KEYS.archivedTasks, []);
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

// ===== Task Categories & Colors =====
const TASK_CATEGORIES = [
    { value: 'Content', color: '#4a90d9', bg: '#eaf1fb' },
    { value: 'GTM', color: '#e67e22', bg: '#fef5ec' },
    { value: 'Planning', color: '#9b59b6', bg: '#f5eef8' },
    { value: 'Sync', color: '#27ae60', bg: '#eafaf1' },
    { value: 'Event', color: '#e74c3c', bg: '#fdedec' },
    { value: 'Other', color: '#7f8c8d', bg: '#f2f4f4' }
];

const VENDOR_NAMES = ['Amy', 'Mindy', 'Erica'];

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
    tasks.forEach((task, idx) => {
        const cat = getCategoryInfo(task.category);
        const alert = getDateAlert(task.date);
        const tr = document.createElement('tr');
        tr.style.background = task.category ? cat.bg : '';
        tr.innerHTML = `
            <td class="col-check"><input type="checkbox" class="task-checkbox" onchange="completeTask(${idx})" title="Mark complete"></td>
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
    tasks.push({ id: Date.now(), name: '', who: '', date: '', link: '', vendor: '', category: 'Other', notes: '' });
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
