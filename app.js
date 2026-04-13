// ===== Data Store =====
const STORAGE_KEYS = {
    tasks: 'pa_tasks',
    archivedTasks: 'pa_tasks_archived',
    events1p: 'pa_events_1p',
    events3p: 'pa_events_3p',
    products: 'pa_products',
    personalTasks: 'pa_personal_tasks',
    archivedPersonalTasks: 'pa_personal_tasks_archived',
    familyEvents: 'pa_family_events',
    financeRecords: 'pa_finance_records',
    sidebarLinks: 'pa_sidebar_links'
};

let syncPending = false;
let syncTimer = null;
let lastSyncTime = null;
let isAuthenticated = false;

function loadData(key, defaults) {
    try {
        const saved = localStorage.getItem(key);
        if (saved) return JSON.parse(saved);
        // Write defaults to localStorage so getAllData/sync never sees empty
        if (defaults && defaults.length > 0) {
            localStorage.setItem(key, JSON.stringify(defaults));
        }
        return defaults;
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
    // Use in-memory state variables — they always have data (including defaults)
    return {
        tasks, archivedTasks, events1p, events3p, products,
        personalTasks, archivedPersonalTasks, familyEvents, financeRecords, sidebarLinks,
        _savedAt: new Date().toISOString()
    };
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
            console.warn('Sync blocked: local data is completely empty, refusing to overwrite server');
            updateSyncIndicator('ok');
            return;
        }
        // Per-key protection: fetch server data and don't overwrite non-empty server keys with empty local
        try {
            const srvResp = await fetch('/api/data');
            if (srvResp.ok) {
                const srvData = await srvResp.json();
                for (const [name] of Object.entries(STORAGE_KEYS)) {
                    const localArr = localData[name];
                    const srvArr = srvData[name];
                    if ((!localArr || (Array.isArray(localArr) && localArr.length === 0))
                        && Array.isArray(srvArr) && srvArr.length > 0) {
                        console.warn(`Sync: preserving server ${name} (${srvArr.length} items) — local is empty`);
                        localData[name] = srvArr;
                    }
                }
            }
        } catch (e) { console.warn('Per-key merge check failed:', e); }
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
                // Always prefer server data on initial load
                applyAllData(data);
                lastSyncTime = data._savedAt || new Date().toISOString();
                localStorage.setItem('pa_last_sync', lastSyncTime);
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
    financeRecords = loadData(STORAGE_KEYS.financeRecords, []);
    sidebarLinks = loadData(STORAGE_KEYS.sidebarLinks, DEFAULT_SIDEBAR_LINKS);
    migrateSidebarLinks();
    migrateWorkEventDates();
    renderSidebar();
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

const DEFAULT_SIDEBAR_LINKS = [
    { id: 'g-skilling', icon: 'school', title: 'Skilling VTeam', items: [
        { id: 'l-1', icon: 'slideshow', label: 'Leadership Walking Deck', url: 'https://microsoft.sharepoint.com/:p:/t/AzureSolutionPlaysMarketingTeam/cQqb7i3ZkO8dSYjNOimJuNuREgUCtxLtWKpASy4q0N6PuYVj9g' },
        { id: 'l-2', icon: 'loop', label: 'Loop', url: 'https://loop.cloud.microsoft/' },
        { id: 'l-3', icon: 'code', label: 'GitHub Skilling Site', url: 'https://github.com/akiyaani2/azure-solutions-skilling' },
        { id: 'l-4', icon: 'folder_shared', label: 'Skilling SharePoint', url: 'https://microsoft.sharepoint.com/:f:/t/AzureSolutionPlaysMarketingTeam/IgDR9vsrGrzUQ5IKayqWVvOSAXFNoppTk8nl_cr8LGJLzr0?e=L5aufN' },
        { id: 'l-5', icon: 'filter_alt', label: 'Skilling Funnel', url: 'https://onedrive.cloud.microsoft/:p:/a@l77xg24h/S/cQpouaHRprJqRJfs340cI-DDEgUCBknL2J032D--OOApoDst4w' }
    ]},
    { id: 'g-rob', icon: 'assessment', title: 'ROB', items: [
        { id: 'l-6', icon: 'bar_chart', label: 'MMR', url: '#' },
        { id: 'l-7', icon: 'note', label: 'Takeshi Updates', url: 'onenote:https://microsoft.sharepoint.com/teams/AzureSolutionPlaysMarketingTeam/SiteAssets/Azure%20Solution%20Plays%20Marketing%20Team%20Notebook/Takeshi%20Updates.one' }
    ]},
    { id: 'g-org', icon: 'account_tree', title: 'Org Chart', items: [
        { id: 'l-8', icon: 'corporate_fare', label: 'CC&AI Marketing Org Chart', url: 'https://microsoft.sharepoint.com/:p:/t/CommercialCloudAIMktgROB/EYTIthwe7vpIjCnz_2kSsTYBCzKYpFngBc8jV0elDv--OQ' }
    ]},
    { id: 'g-planning', icon: 'edit_calendar', title: 'Planning', items: [
        { id: 'l-9', icon: 'table_chart', label: 'Current Plan XLS', url: 'https://microsoft-my.sharepoint.com/:x:/p/marcobellini/cQo9bmRfvmFmRovaZIOqipN6EgUC4aps4MjbWv1j5bSjLE_EYg' },
        { id: 'l-10', icon: 'slideshow', label: 'FY27 Greenlight One Pager', url: 'https://microsoft.sharepoint.com/:p:/r/sites/FY27PlanningCommercialCloudandAI/FY27%20Planning%20CCAI%20Documents/FY27%20Greenlight%20-%20CAIP%20Marketing/CAIP%20Marketing%20Planning%20Greenlight%20.pptx' },
        { id: 'l-11', icon: 'grid_view', label: 'FY27 Placemats', url: 'https://microsoft.sharepoint.com/:p:/t/AzureSolutionPlaysMarketingTeam/cQqtJWnaGOcGRLGzJ43zsZqFEgUCNLuS9OecmkX9wTT3M1FPZQ' },
        { id: 'l-12', icon: 'track_changes', label: 'FY27 Placemats Tracker', url: 'https://microsoft.sharepoint.com/:p:/t/AzureSolutionPlaysMarketingTeam/cQqtJWnaGOcGRLGzJ43zsZqFEgUCNLuS9OecmkX9wTT3M1FPZQ' }
    ]},
    { id: 'g-mm', icon: 'cloud', title: 'M&M Skilling', items: [
        { id: 'l-13', icon: 'web', label: 'Landing Page', url: 'https://aka.ms/MigrateModernize2026' },
        { id: 'l-14', icon: 'menu_book', label: 'Microsoft Learn', url: 'https://learn.microsoft.com/en-us/users/marcobellini-8438/' },
        { id: 'l-15', icon: 'videocam', label: 'VTD - Virtual Training Days', url: 'https://aka.ms/MIgrateModernizeVTD' },
        { id: 'l-16', icon: 'emoji_events', label: 'Challenges', url: 'https://learn.microsoft.com/en-us/users/marcobellini-8438/challenges?source=learn' },
        { id: 'l-17', icon: 'podcasts', label: 'Reactor Channel', url: '#' },
        { id: 'l-18', icon: 'edit_note', label: 'Blog Opportunities', url: 'https://onedrive.cloud.microsoft/:x:/a@8pv38am2/S/cQruuyZWS4ypQJnLxBHXlZglEgUCkDelWyndM_Hf9rLzW8kmvA' },
        { id: 'l-19', icon: 'event', label: '1P Events', page: '1p-events' },
        { id: 'l-20', icon: 'groups', label: '3P Events', page: '3p-events' }
    ]},
    { id: 'g-other', icon: 'lightbulb', title: 'Other Projects', items: [
        { id: 'l-21', icon: 'handshake', label: 'Mentoring Guide', url: 'https://onedrive.cloud.microsoft/:p:/a@l77xg24h/S/cQqoD0HIdw8fSohLs5nTfj-OEgUCzV6gis-Dh25_mALZ0OshuQ' }
    ]},
    { id: 'g-corp', icon: 'business', title: 'Corp Useful', items: [
        { id: 'l-22', icon: 'badge', label: 'In Office Profile', url: 'https://msit.powerbi.com/groups/me/reports/63620602-8a32-417d-bc33-65dc73e93db4' },
        { id: 'l-23', icon: 'auto_stories', label: 'Glossary', url: 'https://microsoft.sharepoint.com/SitePages/Glossary.aspx' },
        { id: 'l-24', icon: 'redeem', label: 'Free Things!', url: 'https://microsoft-my.sharepoint.com/personal/dacoulte_microsoft_com/_layouts/15/Doc.aspx?sourcedoc=%7b3df2290f-997a-4622-abb0-f6aa5e4d86b0%7d' }
    ]},
    { id: 'g-personal', icon: 'person', title: 'Personal', items: [
        { id: 'l-25', icon: 'checklist', label: 'Personal Tasks', page: 'personal-tasks' },
        { id: 'l-26', icon: 'family_restroom', label: 'Family Events', page: 'family-events' },
        { id: 'l-27', icon: 'account_balance_wallet', label: 'Finance Tracker', page: 'finance' },
        { id: 'l-28', icon: 'calendar_month', label: 'Calendar', page: 'calendar' },
        { id: 'l-29', icon: 'sports_esports', label: 'UW-TA', url: 'https://uw-ta.onrender.com/' },
        { id: 'l-30', icon: 'auto_fix_high', label: 'Kronoscript', url: 'https://www.kronoscript.net' },
        { id: 'l-31', icon: 'palette', label: 'Canvas', url: '#' }
    ]},
    { id: 'g-settings', icon: 'settings', title: 'Settings', items: [
        { id: 'l-32', icon: 'link', label: 'Link Management', page: 'link-management' }
    ]}
];

function guessIconForLink(label, url) {
    const s = ((label || '') + ' ' + (url || '')).toLowerCase();
    const rules = [
        [/github/, 'code'],
        [/youtube|youtu\.be|vimeo|video|\.mp4/, 'play_circle'],
        [/sharepoint|:f:|folder/, 'folder_shared'],
        [/onenote|\.one/, 'note'],
        [/loop\.cloud/, 'loop'],
        [/onedrive|drive\.google/, 'cloud'],
        [/mail|outlook/, 'mail'],
        [/teams\.microsoft/, 'groups'],
        [/calendar/, 'calendar_month'],
        [/\.pptx|slideshow|deck|presentation/, 'slideshow'],
        [/\.xlsx|\.xls|spreadsheet|sheets/, 'table_chart'],
        [/\.docx|\.doc|\.pdf|document/, 'description'],
        [/learn\.microsoft|microsoft learn/, 'menu_book'],
        [/linkedin/, 'business_center'],
        [/twitter|\bx\.com/, 'tag'],
        [/facebook/, 'thumb_up'],
        [/bank|finance|money|pay/, 'account_balance'],
        [/news|bbc|wsj|times|corriere|gazzetta/, 'newspaper'],
        [/shop|amazon|cart|store/, 'shopping_cart'],
        [/music|spotify|apple\.com\/music/, 'music_note'],
        [/map|maps\.google/, 'map'],
        [/game|play|xbox|playstation|nintendo/, 'sports_esports'],
        [/school|university|\.edu|washington/, 'school'],
        [/health|medical|doctor|hospital/, 'medical_services'],
        [/chart|report|dashboard|powerbi|msit/, 'bar_chart'],
        [/plan|roadmap|track/, 'edit_calendar'],
        [/search|google\.com|bing/, 'search'],
        [/wiki/, 'menu_book'],
        [/blog|substack|medium/, 'edit_note'],
        [/settings|admin|config/, 'settings']
    ];
    for (const [re, icon] of rules) if (re.test(s)) return icon;
    return 'link';
}

// ===== Task Categories & Colors =====
const TASK_CATEGORIES = [
    { value: 'Content', color: '#2e7d32', bg: '#e8f5e9' },
    { value: 'GTM', color: '#e67e22', bg: '#fef5ec' },
    { value: 'Planning', color: '#9b59b6', bg: '#f5eef8' },
    { value: 'Sync', color: '#27ae60', bg: '#eafaf1' },
    { value: 'Event', color: '#e74c3c', bg: '#fdedec' },
    { value: 'Presentations', color: '#c2185b', bg: '#fce4ec' },
    { value: 'ROB', color: '#1565c0', bg: '#e3f2fd' },
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
    { value: 'Moglie', color: '#c2185b', bg: '#fce4ec' },
    { value: 'Home', color: '#6d4c41', bg: '#efebe9' },
    { value: 'Car', color: '#455a64', bg: '#eceff1' },
    { value: 'Medical', color: '#d32f2f', bg: '#fdecea' }
];

const PERSONAL_WHO = ['Marco', 'Daniela', 'David', 'Andrew', 'Nicholas', 'Simon', 'Sara', 'Jackson', 'Maria', 'Egidio', 'Liana', 'Papi', 'Altri'];

const URGENCY_LEVELS = [
    { value: '1', label: 'P1', color: '#e74c3c', bg: '#fdedec' },
    { value: '2', label: 'P2', color: '#f39c12', bg: '#fef9e7' },
    { value: '3', label: 'P3', color: '#27ae60', bg: '#eafaf1' },
    { value: '4', label: 'P4', color: '#3498db', bg: '#ebf5fb' }
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
let financeRecords = loadData(STORAGE_KEYS.financeRecords, []);
let sidebarLinks = loadData(STORAGE_KEYS.sidebarLinks, DEFAULT_SIDEBAR_LINKS);
let currentSort = { table: null, column: null, dir: 'asc' };
let showHiddenTasks = localStorage.getItem('pa_show_hidden_tasks') === '1';

// ===== Page Navigation =====
function showPage(pageId) {
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    const page = document.getElementById('page-' + pageId);
    if (page) {
        page.classList.add('active');
        if (pageId === 'tasks') renderTasks();
        if (pageId === 'hero-products') renderProducts();
        if (pageId === '1p-events' || pageId === '3p-events' || pageId === 'work-events') {
            const target = document.getElementById('page-work-events');
            if (target) {
                document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
                target.classList.add('active');
                renderWorkEvents();
                return;
            }
            renderEvents(pageId === '3p-events' ? '3p' : '1p');
        }
        if (pageId === 'personal-tasks') renderPersonalTasks();
        if (pageId === 'family-events') renderFamilyEvents();
        if (pageId === 'finance') renderFinance();
        if (pageId === 'calendar') renderCalendar();
        if (pageId === 'link-management') renderLinkManagement();
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
    const toggleBtn = document.getElementById('tasks-show-hidden-btn');
    if (toggleBtn) {
        toggleBtn.textContent = showHiddenTasks ? 'Hide hidden' : 'Show hidden';
        toggleBtn.style.background = showHiddenTasks ? '#fff3cd' : '';
    }
    const sorted = tasks.map((t, i) => ({...t, _idx: i}))
        .filter(t => showHiddenTasks || !t.hidden)
        .sort((a, b) => {
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
        if (task.hidden) tr.style.opacity = '0.45';
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
            <td class="col-check"><input type="checkbox" ${task.hidden ? 'checked' : ''} onchange="toggleTaskHidden(${idx})" title="Hide row"></td>
        `;
        tbody.appendChild(tr);
    });
    updateStats();
    renderArchivedToggle();
}

function toggleTaskHidden(idx) {
    tasks[idx].hidden = !tasks[idx].hidden;
    saveData(STORAGE_KEYS.tasks, tasks);
    renderTasks();
}

function toggleShowHiddenTasks() {
    showHiddenTasks = !showHiddenTasks;
    localStorage.setItem('pa_show_hidden_tasks', showHiddenTasks ? '1' : '0');
    renderTasks();
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

function buildFamilyWhoSelect(selected, idx) {
    return `<select class="vendor-select" onchange="updateFamilyEvent(${idx},'who',this.value)">
        <option value="" ${!selected ? 'selected' : ''}>—</option>
        ${PERSONAL_WHO.map(w => `<option value="${w}" ${w === selected ? 'selected' : ''}>${w}</option>`).join('')}
    </select>`;
}

// ===== Family Events =====
const FAMILY_EVENT_COLORS = ['#2980b9','#e74c3c','#27ae60','#8e44ad','#e67e22','#c2185b','#00897b','#5c6bc0','#f57c00','#6d4c41'];

function hexToPastel(hex) {
    const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
    return `rgba(${r},${g},${b},0.12)`;
}

// Biweekly work-day pattern: cycle starts Thu Apr 9 2026, repeats every 14 days
// Days within cycle: offset 0(Thu),1(Fri),2(Sat),5(Tue),6(Wed)
const BIWEEKLY_ANCHOR = new Date(2026, 3, 9); // April 9, 2026
const BIWEEKLY_OFFSETS = [0, 1, 2, 5, 6];
function isBiweeklyDay(year, month, day) {
    const d = new Date(year, month, day);
    const diff = Math.round((d - BIWEEKLY_ANCHOR) / 86400000);
    if (diff < 0) return false;
    const cycleDay = diff % 14;
    return BIWEEKLY_OFFSETS.includes(cycleDay);
}

function getFamilyEventColors() {
    const eventColors = {};
    let colorIdx = 0;
    familyEvents.forEach(ev => {
        if (ev.name && !eventColors[ev.name]) {
            eventColors[ev.name] = FAMILY_EVENT_COLORS[colorIdx % FAMILY_EVENT_COLORS.length];
            colorIdx++;
        }
    });
    return eventColors;
}

function renderFamilyCalendar() {
    const container = document.getElementById('family-calendar-strip');
    if (!container) return;
    const now = new Date();
    const eventColors = getFamilyEventColors();

    let html = '<div class="fam-cal-strip">';
    for (let m = 0; m < 12; m++) {
        const viewDate = new Date(now.getFullYear(), now.getMonth() + m, 1);
        const year = viewDate.getFullYear();
        const month = viewDate.getMonth();
        const monthName = viewDate.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        const firstDay = new Date(year, month, 1).getDay();
        const today = new Date(); today.setHours(0,0,0,0);

        // Build day-to-events map
        const dayEvents = {};
        familyEvents.forEach(ev => {
            if (!ev.name || !ev.dateFrom) return;
            const from = new Date(ev.dateFrom + 'T00:00:00');
            const to = ev.dateTo ? new Date(ev.dateTo + 'T00:00:00') : from;
            for (let d = new Date(from); d <= to; d.setDate(d.getDate() + 1)) {
                if (d.getFullYear() === year && d.getMonth() === month) {
                    if (!dayEvents[d.getDate()]) dayEvents[d.getDate()] = [];
                    dayEvents[d.getDate()].push(ev);
                }
            }
        });

        html += `<div class="fam-cal-month">
            <div class="fam-cal-month-title">${monthName}</div>
            <div class="fam-cal-grid">
                <div class="fam-cal-dh">S</div><div class="fam-cal-dh">M</div><div class="fam-cal-dh">T</div><div class="fam-cal-dh">W</div><div class="fam-cal-dh">T</div><div class="fam-cal-dh">F</div><div class="fam-cal-dh">S</div>`;
        for (let i = 0; i < firstDay; i++) html += '<div class="fam-cal-day fam-cal-empty"></div>';
        for (let day = 1; day <= daysInMonth; day++) {
            const isToday = year === today.getFullYear() && month === today.getMonth() && day === today.getDate();
            const evts = dayEvents[day] || [];
            const bgColor = evts.length ? eventColors[evts[0].name] : '';
            const bw = isBiweeklyDay(year, month, day);
            let style = '';
            if (bgColor) style = `background:${bgColor};color:#fff`;
            html += `<div class="fam-cal-day${isToday ? ' fam-cal-today' : ''}${evts.length ? ' fam-cal-has-event' : ''}${bw ? ' fam-cal-bw' : ''}" ${style ? `style="${style}"` : ''}>
                ${day}
                ${evts.length ? `<div class="fam-cal-tip">${evts.map(e => `<div><strong>${esc(e.name)}</strong>${e.where ? '<br>' + esc(e.where) : ''}${e.who ? '<br><em>' + esc(e.who) + '</em>' : ''}</div>`).join('')}</div>` : ''}
            </div>`;
        }
        html += '</div></div>';
    }
    // Legend
    html += '<div class="fam-cal-legend">';
    for (const [name, color] of Object.entries(eventColors)) {
        html += `<span class="fam-cal-legend-item"><span class="fam-cal-legend-dot" style="background:${color}"></span>${esc(name)}</span>`;
    }
    html += '<span class="fam-cal-legend-item"><span class="fam-cal-legend-dot" style="background:transparent;border:2px solid #e74c3c"></span>Work days</span>';
    html += '</div></div>';
    container.innerHTML = html;
}

function renderFamilyEvents() {
    renderFamilyCalendar();
    const eventColors = getFamilyEventColors();
    const tbody = document.getElementById('family-events-body');
    tbody.innerHTML = '';
    familyEvents.forEach((ev, idx) => {
        const evColor = ev.name ? eventColors[ev.name] : '';
        const rowBg = evColor ? hexToPastel(evColor) : '';
        const tr = document.createElement('tr');
        if (rowBg) tr.style.background = rowBg;
        tr.innerHTML = `
            ${evColor ? `<td style="width:4px;padding:0;background:${evColor}"></td>` : '<td style="width:4px;padding:0"></td>'}
            <td><input type="date" value="${esc(ev.dateFrom)}" onchange="updateFamilyEvent(${idx},'dateFrom',this.value)"></td>
            <td><input type="date" value="${esc(ev.dateTo)}" onchange="updateFamilyEvent(${idx},'dateTo',this.value)"></td>
            <td><input type="text" value="${esc(ev.name)}" placeholder="Event name..." onchange="updateFamilyEvent(${idx},'name',this.value)"></td>
            <td><input type="text" value="${esc(ev.where)}" placeholder="Location..." onchange="updateFamilyEvent(${idx},'where',this.value)"></td>
            <td><input type="text" value="${esc(ev.hotel)}" placeholder="Hotel..." onchange="updateFamilyEvent(${idx},'hotel',this.value)"></td>
            <td><input type="text" value="${esc(ev.car)}" placeholder="Car..." onchange="updateFamilyEvent(${idx},'car',this.value)"></td>
            <td><input type="text" value="${esc(ev.transportation)}" placeholder="Details..." onchange="updateFamilyEvent(${idx},'transportation',this.value)"></td>
            <td>${buildFamilyWhoSelect(ev.who, idx)}</td>
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
    if (field === 'dateFrom' || field === 'dateTo' || field === 'name' || field === 'who' || field === 'where') {
        renderFamilyEvents();
    }
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

// ===== Work Events (unified 1P + 3P) =====
let workEventsFilter = { type: 'all', range: '3m', search: '' };

function getAllWorkEvents() {
    const hydrate = (e, i, t) => ({
        ...e,
        dateFrom: e.dateFrom || e.date || '',
        dateTo: e.dateTo || e.date || '',
        _type: t, _idx: i
    });
    return [
        ...events1p.map((e, i) => hydrate(e, i, '1p')),
        ...events3p.map((e, i) => hydrate(e, i, '3p'))
    ];
}

function migrateWorkEventDates() {
    let changed1p = false, changed3p = false;
    events1p.forEach(e => { if (!e.dateFrom && e.date) { e.dateFrom = e.date; e.dateTo = e.dateTo || e.date; changed1p = true; } });
    events3p.forEach(e => { if (!e.dateFrom && e.date) { e.dateFrom = e.date; e.dateTo = e.dateTo || e.date; changed3p = true; } });
    if (changed1p) saveData(STORAGE_KEYS.events1p, events1p);
    if (changed3p) saveData(STORAGE_KEYS.events3p, events3p);
}

const WORK_EVENT_PALETTE = ['#e74c3c', '#3498db', '#27ae60', '#9b59b6', '#e67e22', '#16a085', '#c2185b', '#2980b9', '#d35400', '#7f8c8d', '#8e44ad', '#00acc1'];
function getWorkEventColors() {
    const map = {};
    let i = 0;
    getAllWorkEvents().forEach(ev => {
        if (ev.name && !map[ev.name]) { map[ev.name] = WORK_EVENT_PALETTE[i % WORK_EVENT_PALETTE.length]; i++; }
    });
    return map;
}

function renderWorkEvents() {
    const tbody = document.getElementById('work-events-body');
    if (!tbody) return;
    const typeSel = document.getElementById('we-filter-type');
    const rangeSel = document.getElementById('we-filter-range');
    const searchInp = document.getElementById('we-filter-search');
    if (typeSel) workEventsFilter.type = typeSel.value;
    if (rangeSel) workEventsFilter.range = rangeSel.value;
    if (searchInp) workEventsFilter.search = searchInp.value;
    const today = new Date(); today.setHours(0, 0, 0, 0);
    let cutoffEnd = null, cutoffStart = today;
    if (workEventsFilter.range === '3m') { cutoffEnd = new Date(today); cutoffEnd.setMonth(cutoffEnd.getMonth() + 3); }
    else if (workEventsFilter.range === '6m') { cutoffEnd = new Date(today); cutoffEnd.setMonth(cutoffEnd.getMonth() + 6); }
    else if (workEventsFilter.range === '1y') { cutoffEnd = new Date(today); cutoffEnd.setFullYear(cutoffEnd.getFullYear() + 1); }
    else if (workEventsFilter.range === 'all') { cutoffStart = null; cutoffEnd = null; }
    const search = workEventsFilter.search.trim().toLowerCase();
    let list = getAllWorkEvents();
    if (workEventsFilter.type !== 'all') list = list.filter(e => e._type === workEventsFilter.type);
    list = list.filter(e => {
        const anchor = e.dateFrom || e.date;
        if (anchor) {
            const d = new Date(anchor + 'T00:00:00');
            if (cutoffStart && d < cutoffStart) return false;
            if (cutoffEnd && d > cutoffEnd) return false;
        }
        if (search) {
            const hay = `${e.name || ''} ${e.hero || ''} ${e.pmm || ''} ${e.contact || ''} ${e.notes || ''}`.toLowerCase();
            if (!hay.includes(search)) return false;
        }
        return true;
    });
    const sortCol = currentSort.table === 'work-events-table' ? currentSort.column : 'dateFrom';
    const sortDir = currentSort.table === 'work-events-table' ? currentSort.dir : 'asc';
    list.sort((a, b) => {
        const va = (a[sortCol] || '').toString().toLowerCase();
        const vb = (b[sortCol] || '').toString().toLowerCase();
        return sortDir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    });
    tbody.innerHTML = '';
    list.forEach(ev => {
        const t = ev._type, idx = ev._idx;
        const typeBubble = t === '1p'
            ? '<span style="background:#e67e22;color:#fff;font-size:11px;font-weight:700;padding:3px 9px;border-radius:10px">1P</span>'
            : '<span style="background:#fff3cd;color:#856404;font-size:11px;font-weight:700;padding:3px 9px;border-radius:10px;border:1px solid #ffeaa7">3P</span>';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${typeBubble}</td>
            <td>${buildEventPrioritySelect(t, idx, ev.priority || 'P3')}</td>
            <td><input type="date" value="${esc(ev.dateFrom)}" onchange="updateEvent('${t}',${idx},'dateFrom',this.value); renderWorkEvents()"></td>
            <td><input type="date" value="${esc(ev.dateTo)}" onchange="updateEvent('${t}',${idx},'dateTo',this.value); renderWorkEvents()"></td>
            <td><input type="text" value="${esc(ev.name)}" placeholder="Event name..." onchange="updateEvent('${t}',${idx},'name',this.value); renderWorkEvents()"></td>
            <td><input type="text" value="${esc(ev.hero)}" placeholder="Product..." onchange="updateEvent('${t}',${idx},'hero',this.value)"></td>
            <td><input type="text" value="${esc(ev.pmm)}" placeholder="PMM..." onchange="updateEvent('${t}',${idx},'pmm',this.value)"></td>
            <td><input type="text" value="${esc(ev.contact)}" placeholder="Contact..." onchange="updateEvent('${t}',${idx},'contact',this.value)"></td>
            <td><input type="text" value="${esc(ev.plan)}" placeholder="Plan..." onchange="updateEvent('${t}',${idx},'plan',this.value)"></td>
            <td><input type="text" value="${esc(ev.activity)}" placeholder="Activity..." onchange="updateEvent('${t}',${idx},'activity',this.value)"></td>
            <td><textarea rows="1" placeholder="Notes..." onchange="updateEvent('${t}',${idx},'notes',this.value)">${esc(ev.notes || '')}</textarea></td>
            <td><button class="btn-delete" onclick="deleteEvent('${t}',${idx}); renderWorkEvents()" title="Delete"><span class="material-icons-outlined">delete</span></button></td>
        `;
        tbody.appendChild(tr);
    });
    const countEl = document.getElementById('we-count');
    if (countEl) countEl.textContent = `${list.length} event${list.length === 1 ? '' : 's'}`;
    renderWorkEventsCalendar();
    updateStats();
}

function renderWorkEventsCalendar() {
    const container = document.getElementById('work-events-calendar');
    if (!container) return;
    const now = new Date();
    const colors = getWorkEventColors();
    const all = getAllWorkEvents();
    let html = '<div class="fam-cal-strip">';
    for (let m = 0; m < 12; m++) {
        const viewDate = new Date(now.getFullYear(), now.getMonth() + m, 1);
        const year = viewDate.getFullYear();
        const month = viewDate.getMonth();
        const monthName = viewDate.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        const firstDay = new Date(year, month, 1).getDay();
        const today = new Date(); today.setHours(0, 0, 0, 0);
        const dayEvents = {};
        all.forEach(ev => {
            if (!ev.name || !ev.dateFrom) return;
            const from = new Date(ev.dateFrom + 'T00:00:00');
            const to = ev.dateTo ? new Date(ev.dateTo + 'T00:00:00') : from;
            for (let d = new Date(from); d <= to; d.setDate(d.getDate() + 1)) {
                if (d.getFullYear() === year && d.getMonth() === month) {
                    if (!dayEvents[d.getDate()]) dayEvents[d.getDate()] = [];
                    dayEvents[d.getDate()].push(ev);
                }
            }
        });
        html += `<div class="fam-cal-month"><div class="fam-cal-month-title">${monthName}</div><div class="fam-cal-grid"><div class="fam-cal-dh">S</div><div class="fam-cal-dh">M</div><div class="fam-cal-dh">T</div><div class="fam-cal-dh">W</div><div class="fam-cal-dh">T</div><div class="fam-cal-dh">F</div><div class="fam-cal-dh">S</div>`;
        for (let i = 0; i < firstDay; i++) html += '<div class="fam-cal-day fam-cal-empty"></div>';
        for (let day = 1; day <= daysInMonth; day++) {
            const isToday = year === today.getFullYear() && month === today.getMonth() && day === today.getDate();
            const evts = dayEvents[day] || [];
            const bgColor = evts.length ? colors[evts[0].name] : '';
            const style = bgColor ? `background:${bgColor};color:#fff` : '';
            html += `<div class="fam-cal-day${isToday ? ' fam-cal-today' : ''}${evts.length ? ' fam-cal-has-event' : ''}" ${style ? `style="${style}"` : ''}>${day}${evts.length ? `<div class="fam-cal-tip">${evts.map(e => `<div><strong>${esc(e.name)}</strong> <span style="opacity:.8">[${e._type.toUpperCase()}]</span>${e.hero ? '<br>' + esc(e.hero) : ''}${e.pmm ? '<br><em>' + esc(e.pmm) + '</em>' : ''}</div>`).join('')}</div>` : ''}</div>`;
        }
        html += '</div></div>';
    }
    html += '<div class="fam-cal-legend">';
    for (const [name, color] of Object.entries(colors)) html += `<span class="fam-cal-legend-item"><span class="fam-cal-legend-dot" style="background:${color}"></span>${esc(name)}</span>`;
    html += '</div></div>';
    container.innerHTML = html;
}

function addWorkEvent() {
    const typeSel = document.getElementById('we-add-type');
    const t = typeSel ? typeSel.value : '1p';
    addEvent(t);
    renderWorkEvents();
}

// ===== Hero Products =====
function getPmmList(pmmStr) {
    if (!pmmStr) return [];
    return pmmStr.split(',').map(s => s.trim()).filter(Boolean);
}

function renderBubbles(names, prodIdx, field) {
    const color = getProductColor(prodIdx);
    return names.map((name, i) =>
        `<span class="bubble" style="background:${color.mid};color:#1e1e2d">${esc(name)}<button class="bubble-x" onclick="event.stopPropagation(); removePmm(${prodIdx},'${field}',${i})" title="Remove">&times;</button></span>`
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

const PRODUCT_PALETTE = [
    { bg: '#fde7e9', mid: '#f5c6cb', dark: '#c0392b' },
    { bg: '#fdebd0', mid: '#f6d5a6', dark: '#b9770e' },
    { bg: '#fef5d4', mid: '#f7e8a5', dark: '#9a7d0a' },
    { bg: '#e8f8e8', mid: '#c8ecc8', dark: '#1e7e34' },
    { bg: '#d4f1f4', mid: '#a8dfe6', dark: '#117a8b' },
    { bg: '#dbeafe', mid: '#b4cff8', dark: '#1d4ed8' },
    { bg: '#e7e0fa', mid: '#c9bdf0', dark: '#5b21b6' },
    { bg: '#f8d7da', mid: '#efb0b6', dark: '#a71d2a' },
    { bg: '#fce4ec', mid: '#f3bed1', dark: '#ad1457' },
    { bg: '#e0f2f1', mid: '#b2dfdc', dark: '#00695c' }
];
function getProductColor(idx) {
    const prod = products[idx];
    if (prod && typeof prod._paletteIdx === 'number') return PRODUCT_PALETTE[prod._paletteIdx % PRODUCT_PALETTE.length];
    return PRODUCT_PALETTE[idx % PRODUCT_PALETTE.length];
}

function getProductCategory(prod) { return prod.category || 'PMM'; }

function assignCardPaletteIdx(category) {
    const used = new Set(
        products.filter(p => getProductCategory(p) === category && typeof p._paletteIdx === 'number').map(p => p._paletteIdx)
    );
    for (let i = 0; i < PRODUCT_PALETTE.length; i++) if (!used.has(i)) return i;
    // all used — pick the least-used
    const counts = new Array(PRODUCT_PALETTE.length).fill(0);
    products.filter(p => getProductCategory(p) === category).forEach(p => {
        if (typeof p._paletteIdx === 'number') counts[p._paletteIdx]++;
    });
    let best = 0;
    for (let i = 1; i < counts.length; i++) if (counts[i] < counts[best]) best = i;
    return best;
}

function migrateProductsCategoryAndPalette() {
    let changed = false;
    const byCat = {};
    products.forEach((p) => {
        if (!p.category) { p.category = 'PMM'; changed = true; }
        if (typeof p._paletteIdx !== 'number') {
            const cat = getProductCategory(p);
            byCat[cat] = (byCat[cat] || 0);
            p._paletteIdx = byCat[cat] % PRODUCT_PALETTE.length;
            byCat[cat]++;
            changed = true;
        }
    });
    if (changed) saveData(STORAGE_KEYS.products, products);
}

function firstName(full) {
    if (!full) return '';
    return String(full).trim().split(/\s+/)[0];
}

function renderProducts() {
    migrateProductsCategoryAndPalette();
    const grid = document.getElementById('products-grid');
    grid.innerHTML = '';
    const categories = ['PMM', 'Other'];
    categories.forEach(cat => {
        const catProducts = products.map((p, i) => ({ p, i })).filter(({ p }) => getProductCategory(p) === cat);
        if (!catProducts.length && cat === 'Other') return;
        const section = document.createElement('div');
        section.className = 'product-category-section';
        section.style.cssText = 'grid-column:1/-1;margin-top:8px';
        section.innerHTML = `<h2 style="font-size:16px;font-weight:600;color:#555;margin:18px 0 10px 2px;text-transform:uppercase;letter-spacing:.5px">${esc(cat)}</h2>`;
        grid.appendChild(section);
        const inner = document.createElement('div');
        inner.className = 'products-grid-inner';
        inner.style.cssText = 'display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:16px;grid-column:1/-1';
        section.appendChild(inner);
        catProducts.forEach(({ p: prod, i: idx }) => {
            const pmmNames = getPmmList(prod.pmm);
            const managerNames = getPmmList(prod.pmmManager || '');
            const color = getProductColor(idx);
            const card = document.createElement('div');
            card.className = 'product-card';
            card.style.background = color.bg;
            card.style.borderColor = color.dark + '33';
            card.onclick = () => showProductDetail(idx);
            const managerBubbles = managerNames.map((name, mi) =>
                `<span class="bubble" style="background:${color.dark};color:#fff;font-weight:600">${esc(name)}<button class="bubble-x" onclick="event.stopPropagation(); removeManager(${idx},${mi})" title="Remove manager">&times;</button></span>`
            ).join('');
            card.innerHTML = `
                <h3>${esc(prod.name)}</h3>
                ${managerBubbles ? `<div class="card-label">Manager${managerNames.length > 1 ? 's' : ''}</div><div class="bubble-wrap" style="margin-bottom:10px">${managerBubbles}</div>` : ''}
                <div class="card-label">Who</div>
                <div class="bubble-wrap">${renderBubbles(pmmNames, idx, 'pmm')}</div>
            `;
            inner.appendChild(card);
        });
    });
    refreshProductSelects();
}

function refreshProductSelects() {
    const fill = (selId, catFilterId) => {
        const sel = document.getElementById(selId);
        if (!sel) return;
        const cat = (document.getElementById(catFilterId) || {}).value || 'PMM';
        const filtered = products.map((p, i) => ({ p, i })).filter(({ p }) => getProductCategory(p) === cat);
        sel.innerHTML = filtered.map(({ p, i }) => `<option value="${i}">${esc(p.name)}</option>`).join('') || '<option value="">(no cards in this category)</option>';
    };
    fill('pmm-product-select', 'pmm-category-filter');
    fill('mgr-product-select', 'mgr-category-filter');
}

function addManagerToProduct() {
    const sel = document.getElementById('mgr-product-select');
    const input = document.getElementById('mgr-name-input');
    const idx = parseInt(sel.value);
    const name = input.value.trim();
    if (isNaN(idx) || !name) return;
    const list = getPmmList(products[idx].pmmManager || '');
    list.push(name);
    products[idx].pmmManager = list.join(', ');
    saveData(STORAGE_KEYS.products, products);
    input.value = '';
    renderProducts();
}

function removeManager(idx, nameIdx) {
    const list = getPmmList(products[idx].pmmManager || '');
    if (typeof nameIdx === 'number') list.splice(nameIdx, 1);
    else list.length = 0;
    products[idx].pmmManager = list.join(', ');
    saveData(STORAGE_KEYS.products, products);
    renderProducts();
}

function addNewProductCard() {
    const title = (document.getElementById('new-card-title') || {}).value || '';
    const manager = (document.getElementById('new-card-manager') || {}).value || '';
    const pmms = (document.getElementById('new-card-pmms') || {}).value || '';
    const category = (document.getElementById('new-card-category') || {}).value || 'PMM';
    if (!title.trim()) { alert('New Title is required.'); return; }
    const paletteIdx = assignCardPaletteIdx(category);
    products.push({
        name: title.trim(),
        category,
        _paletteIdx: paletteIdx,
        pmm: pmms.split(',').map(s => s.trim()).filter(Boolean).join(', '),
        pmmManager: manager.trim(),
        sme: '', globalSkilling: '', updates: [], events: []
    });
    saveData(STORAGE_KEYS.products, products);
    document.getElementById('new-card-title').value = '';
    document.getElementById('new-card-manager').value = '';
    document.getElementById('new-card-pmms').value = '';
    renderProducts();
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

// ===== Finance Tracker =====
const fmtMoney = n => (n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function financeFilterRecords() {
    const whoVal = (document.getElementById('finance-filter-who') || {}).value || '';
    const rangeVal = (document.getElementById('finance-filter-range') || {}).value || 'all';
    const searchVal = ((document.getElementById('finance-filter-search') || {}).value || '').trim().toLowerCase();
    const now = new Date();
    let cutoff = null;
    if (rangeVal === 'ytd') cutoff = new Date(now.getFullYear(), 0, 1);
    else if (rangeVal === '1y') { cutoff = new Date(now); cutoff.setFullYear(cutoff.getFullYear() - 1); }
    else if (rangeVal === '30' || rangeVal === '60' || rangeVal === '90') {
        cutoff = new Date(now); cutoff.setDate(cutoff.getDate() - parseInt(rangeVal));
    }
    return financeRecords.map((rec, idx) => ({ rec, idx })).filter(({ rec }) => {
        if (whoVal && rec.who !== whoVal) return false;
        if (cutoff && rec.date) { if (new Date(rec.date) < cutoff) return false; }
        if (searchVal) {
            const hay = ((rec.notes || '') + ' ' + (rec.who || '')).toLowerCase();
            if (!hay.includes(searchVal)) return false;
        }
        return true;
    });
}

function renderFinance() {
    const filterWho = document.getElementById('finance-filter-who');
    if (filterWho) {
        const people = Array.from(new Set([...PERSONAL_WHO, ...financeRecords.map(r => r.who).filter(Boolean)]));
        const cur = filterWho.value;
        filterWho.innerHTML = '<option value="">All People</option>' + people.map(w => `<option value="${esc(w)}" ${w === cur ? 'selected' : ''}>${esc(w)}</option>`).join('');
    }
    const filtered = financeFilterRecords();
    const tbody = document.getElementById('finance-body');
    tbody.innerHTML = '';
    let totalDebit = 0, totalCredit = 0;
    filtered.forEach(({ rec, idx }) => {
        const amt = parseFloat(rec.amount) || 0;
        if (rec.type === 'Debit') totalDebit += amt; else totalCredit += amt;
        const mColor = rec.method === 'Cash' ? '#2980b9' : '#e67e22';
        const debitVal = rec.type === 'Debit' ? fmtMoney(amt) : '';
        const creditVal = rec.type === 'Credit' ? fmtMoney(amt) : '';
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${buildFinanceWhoSelect(rec.who, idx)}</td>
            <td><select class="cat-select" style="background:${mColor};color:#fff" onchange="updateFinance(${idx},'method',this.value)">
                <option value="Purchase" ${rec.method !== 'Cash' ? 'selected' : ''} style="background:#fff;color:#333">Purchase</option>
                <option value="Cash" ${rec.method === 'Cash' ? 'selected' : ''} style="background:#fff;color:#333">Cash</option>
            </select></td>
            <td style="text-align:right;white-space:nowrap;color:#e74c3c"><span style="color:#888">$</span><input type="text" inputmode="decimal" value="${debitVal}" placeholder="0.00" onchange="updateFinanceAmount(${idx},'Debit',this.value)" style="width:110px;text-align:right;color:#e74c3c;font-weight:600"></td>
            <td style="text-align:right;white-space:nowrap;color:#27ae60"><span style="color:#888">$</span><input type="text" inputmode="decimal" value="${creditVal}" placeholder="0.00" onchange="updateFinanceAmount(${idx},'Credit',this.value)" style="width:110px;text-align:right;color:#27ae60;font-weight:600"></td>
            <td><input type="date" value="${esc(rec.date)}" onchange="updateFinance(${idx},'date',this.value)"></td>
            <td><textarea rows="1" placeholder="Notes..." onchange="updateFinance(${idx},'notes',this.value)">${esc(rec.notes || '')}</textarea></td>
            <td><button class="btn-delete" onclick="deleteFinance(${idx})" title="Delete"><span class="material-icons-outlined">delete</span></button></td>
        `;
        tbody.appendChild(tr);
    });
    const totalEl = document.getElementById('finance-totals');
    if (totalEl) {
        const balance = totalCredit - totalDebit;
        const balColor = balance >= 0 ? '#27ae60' : '#e74c3c';
        totalEl.innerHTML = `<span style="color:#e74c3c;font-weight:700">Debit: $${fmtMoney(totalDebit)}</span> &nbsp;|&nbsp; <span style="color:#27ae60;font-weight:700">Credit: $${fmtMoney(totalCredit)}</span> &nbsp;|&nbsp; <span style="color:${balColor};font-weight:700">Balance: $${fmtMoney(balance)}</span> &nbsp;|&nbsp; <span style="color:#555">${filtered.length} record${filtered.length === 1 ? '' : 's'}</span>`;
    }
    renderFinanceCharts(filtered.map(f => f.rec));
}

function renderFinanceCharts(records) {
    const host = document.getElementById('finance-charts');
    if (!host) return;
    host.style.gridTemplateColumns = '1fr 1fr 1fr';
    const debits = records.filter(r => r.type === 'Debit');
    // Chart 1: spending by person
    const byPerson = {};
    debits.forEach(r => { const k = r.who || '—'; byPerson[k] = (byPerson[k] || 0) + (parseFloat(r.amount) || 0); });
    // Chart 2: spending by method
    const byMethod = {};
    debits.forEach(r => { const k = r.method || 'Purchase'; byMethod[k] = (byMethod[k] || 0) + (parseFloat(r.amount) || 0); });
    // Chart 3: debit vs credit by month
    const byMonth = {};
    records.forEach(r => {
        if (!r.date) return;
        const k = r.date.substring(0, 7);
        if (!byMonth[k]) byMonth[k] = { d: 0, c: 0 };
        const amt = parseFloat(r.amount) || 0;
        if (r.type === 'Debit') byMonth[k].d += amt; else byMonth[k].c += amt;
    });
    host.innerHTML =
        financeBarChart('Spending by Person', byPerson, '#e74c3c') +
        financeDonutChart('Spending by Method', byMethod, ['#e67e22', '#2980b9']) +
        financeMonthChart('Debit vs Credit by Month', byMonth);
}

function financeCard(title, inner) {
    return `<div style="background:#fff;border:1px solid #e8eaed;border-radius:8px;padding:14px"><div style="font-size:12px;font-weight:600;color:#555;text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">${title}</div>${inner}</div>`;
}

function financeBarChart(title, data, color) {
    const entries = Object.entries(data).sort((a, b) => b[1] - a[1]).slice(0, 6);
    if (!entries.length) return financeCard(title, '<div style="color:#999;font-size:13px">No data</div>');
    const max = Math.max(...entries.map(e => e[1])) || 1;
    const rows = entries.map(([k, v]) => {
        const pct = (v / max) * 100;
        return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:12px"><div style="width:90px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${esc(k)}">${esc(k)}</div><div style="flex:1;background:#f1f3f4;border-radius:3px;height:16px;position:relative"><div style="width:${pct}%;height:100%;background:${color};border-radius:3px"></div></div><div style="width:80px;text-align:right;font-weight:600;color:${color}">$${fmtMoney(v)}</div></div>`;
    }).join('');
    return financeCard(title, rows);
}

function financeDonutChart(title, data, palette) {
    const entries = Object.entries(data).filter(([, v]) => v > 0);
    if (!entries.length) return financeCard(title, '<div style="color:#999;font-size:13px">No data</div>');
    const total = entries.reduce((s, [, v]) => s + v, 0);
    const r = 52, c = 2 * Math.PI * r;
    let offset = 0;
    const segs = entries.map(([, v], i) => {
        const frac = v / total;
        const dash = frac * c;
        const el = `<circle cx="70" cy="70" r="${r}" fill="none" stroke="${palette[i % palette.length]}" stroke-width="20" stroke-dasharray="${dash} ${c - dash}" stroke-dashoffset="${-offset}" transform="rotate(-90 70 70)"/>`;
        offset += dash;
        return el;
    }).join('');
    const legend = entries.map(([k, v], i) => `<div style="display:flex;align-items:center;gap:6px;font-size:12px;margin-bottom:4px"><span style="width:10px;height:10px;background:${palette[i % palette.length]};border-radius:2px;display:inline-block"></span><span style="flex:1">${esc(k)}</span><span style="font-weight:600">$${fmtMoney(v)}</span></div>`).join('');
    return financeCard(title, `<div style="display:flex;align-items:center;gap:12px"><svg width="140" height="140" viewBox="0 0 140 140">${segs}<text x="70" y="68" text-anchor="middle" font-size="11" fill="#555">Total</text><text x="70" y="84" text-anchor="middle" font-size="13" font-weight="700" fill="#333">$${fmtMoney(total)}</text></svg><div style="flex:1;min-width:0">${legend}</div></div>`);
}

function financeMonthChart(title, byMonth) {
    const keys = Object.keys(byMonth).sort().slice(-6);
    if (!keys.length) return financeCard(title, '<div style="color:#999;font-size:13px">No data</div>');
    const max = Math.max(...keys.map(k => Math.max(byMonth[k].d, byMonth[k].c))) || 1;
    const w = 280, h = 140, pad = 22, bw = (w - pad * 2) / keys.length;
    const bars = keys.map((k, i) => {
        const x = pad + i * bw;
        const dh = (byMonth[k].d / max) * (h - pad - 20);
        const ch = (byMonth[k].c / max) * (h - pad - 20);
        const barW = (bw - 8) / 2;
        return `<rect x="${x + 2}" y="${h - pad - dh}" width="${barW}" height="${dh}" fill="#e74c3c"><title>Debit $${fmtMoney(byMonth[k].d)}</title></rect><rect x="${x + 2 + barW + 2}" y="${h - pad - ch}" width="${barW}" height="${ch}" fill="#27ae60"><title>Credit $${fmtMoney(byMonth[k].c)}</title></rect><text x="${x + bw / 2}" y="${h - 6}" text-anchor="middle" font-size="10" fill="#666">${k.substring(5)}/${k.substring(2, 4)}</text>`;
    }).join('');
    const legend = `<div style="display:flex;gap:14px;font-size:11px;margin-top:4px"><span><span style="display:inline-block;width:10px;height:10px;background:#e74c3c;border-radius:2px"></span> Debit</span><span><span style="display:inline-block;width:10px;height:10px;background:#27ae60;border-radius:2px"></span> Credit</span></div>`;
    return financeCard(title, `<svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}">${bars}</svg>${legend}`);
}

// ===== Sidebar (data-driven) =====
function renderSidebar() {
    const aside = document.querySelector('aside.sidebar');
    if (!aside) return;
    const html = ['<h2 class="sidebar-title">Quick Links</h2>'];
    sidebarLinks.forEach((group, gi) => {
        const openCls = gi === 0 ? 'open' : '';
        const collapsedCls = gi === 0 ? '' : 'collapsed';
        const items = group.items.map(it => {
            const href = it.page ? '#' : (it.url || '#');
            const click = it.page ? `onclick="showPage('${it.page}'); return false;"` : '';
            const target = it.page ? '' : 'target="_blank"';
            return `<a href="${esc(href)}" ${target} ${click} class="sidebar-link"><span class="material-icons-outlined">${esc(it.icon || 'link')}</span> ${esc(it.label)}</a>`;
        }).join('');
        html.push(`<div class="link-group"><div class="link-group-header ${collapsedCls}" onclick="toggleGroup(this)"><span class="material-icons-outlined">${esc(group.icon || 'folder')}</span> ${esc(group.title)}<span class="material-icons-outlined chevron">expand_more</span></div><div class="link-group-body ${openCls}">${items}</div></div>`);
    });
    aside.innerHTML = html.join('');
}

// ===== Link Management =====
function renderLinkManagement() {
    const host = document.getElementById('link-management-content');
    if (!host) return;
    const cards = sidebarLinks.map(group => {
        const rows = group.items.map(item => `
            <div style="display:flex;gap:8px;align-items:center;padding:8px;border-bottom:1px solid #f1f3f4">
                <span class="material-icons-outlined" style="color:#888;font-size:18px">${esc(item.icon || 'link')}</span>
                <input type="text" value="${esc(item.label)}" placeholder="Label" onchange="linkUpdate('${group.id}','${item.id}','label',this.value)" style="flex:1;min-width:120px;padding:5px 8px;border:1px solid #dadce0;border-radius:4px;font-size:12px">
                <input type="text" value="${esc(item.url || '')}" placeholder="${item.page ? 'internal page: ' + item.page : 'https://...'}" ${item.page ? 'disabled' : ''} onchange="linkUpdate('${group.id}','${item.id}','url',this.value)" style="flex:2;min-width:180px;padding:5px 8px;border:1px solid #dadce0;border-radius:4px;font-size:12px;${item.page ? 'background:#f5f5f5;color:#999' : ''}">
                <select onchange="linkMove('${group.id}','${item.id}',this.value)" style="padding:5px 8px;border:1px solid #dadce0;border-radius:4px;font-size:12px">
                    ${sidebarLinks.map(g => `<option value="${g.id}" ${g.id === group.id ? 'selected' : ''}>${esc(g.title)}</option>`).join('')}
                </select>
                <button class="btn-delete" onclick="linkDelete('${group.id}','${item.id}')" title="Delete link"><span class="material-icons-outlined">delete</span></button>
            </div>
        `).join('');
        return `
            <div style="background:#fff;border:1px solid #e8eaed;border-radius:8px;margin-bottom:16px">
                <div style="padding:12px 14px;background:#f8f9fa;border-bottom:1px solid #e8eaed;border-radius:8px 8px 0 0;display:flex;align-items:center;gap:8px">
                    <span class="material-icons-outlined">${esc(group.icon || 'folder')}</span>
                    <strong style="flex:1">${esc(group.title)}</strong>
                    <span style="color:#888;font-size:12px">${group.items.length} link${group.items.length === 1 ? '' : 's'}</span>
                </div>
                ${rows || '<div style="padding:12px;color:#999;font-size:13px">No links</div>'}
                <div style="padding:10px 12px;display:flex;gap:8px;align-items:center;background:#fafbfc;border-radius:0 0 8px 8px">
                    <input type="text" placeholder="New label" id="new-label-${group.id}" style="flex:1;padding:6px 8px;border:1px solid #dadce0;border-radius:4px;font-size:12px">
                    <input type="text" placeholder="https://..." id="new-url-${group.id}" style="flex:2;padding:6px 8px;border:1px solid #dadce0;border-radius:4px;font-size:12px">
                    <input type="text" placeholder="icon (optional)" id="new-icon-${group.id}" value="link" style="width:100px;padding:6px 8px;border:1px solid #dadce0;border-radius:4px;font-size:12px">
                    <button class="btn-primary" onclick="linkAdd('${group.id}')"><span class="material-icons-outlined">add</span> Add</button>
                </div>
            </div>
        `;
    }).join('');
    host.innerHTML = cards;
}

function linkUpdate(groupId, itemId, field, value) {
    const g = sidebarLinks.find(x => x.id === groupId);
    if (!g) return;
    const it = g.items.find(x => x.id === itemId);
    if (!it) return;
    it[field] = value;
    saveData(STORAGE_KEYS.sidebarLinks, sidebarLinks);
    renderSidebar();
}

function linkDelete(groupId, itemId) {
    if (!confirm('Delete this link?')) return;
    const g = sidebarLinks.find(x => x.id === groupId);
    if (!g) return;
    g.items = g.items.filter(x => x.id !== itemId);
    saveData(STORAGE_KEYS.sidebarLinks, sidebarLinks);
    renderSidebar();
    renderLinkManagement();
}

function linkMove(fromGroupId, itemId, toGroupId) {
    if (fromGroupId === toGroupId) return;
    const from = sidebarLinks.find(x => x.id === fromGroupId);
    const to = sidebarLinks.find(x => x.id === toGroupId);
    if (!from || !to) return;
    const idx = from.items.findIndex(x => x.id === itemId);
    if (idx < 0) return;
    const [item] = from.items.splice(idx, 1);
    to.items.push(item);
    saveData(STORAGE_KEYS.sidebarLinks, sidebarLinks);
    renderSidebar();
    renderLinkManagement();
}

function linkAdd(groupId) {
    const label = (document.getElementById('new-label-' + groupId) || {}).value || '';
    const url = (document.getElementById('new-url-' + groupId) || {}).value || '';
    const iconInput = ((document.getElementById('new-icon-' + groupId) || {}).value || '').trim();
    if (!label.trim() || !url.trim()) { alert('Label and URL are required.'); return; }
    const icon = iconInput && iconInput !== 'link' ? iconInput : guessIconForLink(label, url);
    const g = sidebarLinks.find(x => x.id === groupId);
    if (!g) return;
    g.items.push({ id: 'l-' + Date.now(), icon, label: label.trim(), url: url.trim() });
    saveData(STORAGE_KEYS.sidebarLinks, sidebarLinks);
    renderSidebar();
    renderLinkManagement();
}

function migrateSidebarLinks() {
    let changed = false;
    const personal = sidebarLinks.find(g => g.id === 'g-personal');
    if (personal) {
        const i = personal.items.findIndex(it => it.id === 'l-32' || it.page === 'link-management');
        if (i >= 0) { personal.items.splice(i, 1); changed = true; }
    }
    if (!sidebarLinks.find(g => g.id === 'g-settings')) {
        sidebarLinks.push({
            id: 'g-settings', icon: 'settings', title: 'Settings',
            items: [{ id: 'l-32', icon: 'link', label: 'Link Management', page: 'link-management' }]
        });
        changed = true;
    }
    if (changed) saveData(STORAGE_KEYS.sidebarLinks, sidebarLinks);
}

function updateFinanceAmount(idx, type, value) {
    const num = parseFloat(String(value).replace(/,/g, '')) || 0;
    financeRecords[idx].type = type;
    financeRecords[idx].amount = num;
    debounceSave(STORAGE_KEYS.financeRecords, financeRecords);
    renderFinance();
}

function buildFinanceWhoSelect(selected, idx) {
    return `<select class="vendor-select" onchange="updateFinance(${idx},'who',this.value)">
        <option value="" ${!selected ? 'selected' : ''}>—</option>
        ${PERSONAL_WHO.map(w => `<option value="${w}" ${w === selected ? 'selected' : ''}>${w}</option>`).join('')}
    </select>`;
}

function addFinanceRecord() {
    financeRecords.push({ id: Date.now(), who: '', type: 'Debit', method: 'Purchase', amount: '', date: new Date().toISOString().split('T')[0], notes: '' });
    saveData(STORAGE_KEYS.financeRecords, financeRecords);
    renderFinance();
}

function updateFinance(idx, field, value) {
    financeRecords[idx][field] = value;
    debounceSave(STORAGE_KEYS.financeRecords, financeRecords);
    if (field === 'type' || field === 'method' || field === 'amount' || field === 'who') renderFinance();
}

function deleteFinance(idx) {
    financeRecords.splice(idx, 1);
    saveData(STORAGE_KEYS.financeRecords, financeRecords);
    renderFinance();
}

function filterFinance() { renderFinance(); }

// ===== Calendar View =====
function renderCalendar() {
    const container = document.getElementById('calendar-container');
    if (!container) return;
    const now = new Date();
    const monthOffset = parseInt(container.dataset.monthOffset || '0');
    const viewDate = new Date(now.getFullYear(), now.getMonth() + monthOffset, 1);
    const year = viewDate.getFullYear();
    const month = viewDate.getMonth();
    const monthName = viewDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const firstDay = new Date(year, month, 1).getDay();

    // Gather all events for this month
    const monthEvents = {};
    const addToDay = (day, label, color) => {
        if (!monthEvents[day]) monthEvents[day] = [];
        monthEvents[day].push({ label, color });
    };
    // 1P/3P events
    [...events1p, ...events3p].forEach(e => {
        if (!e.date || !e.name) return;
        const d = new Date(e.date + 'T00:00:00');
        if (d.getFullYear() === year && d.getMonth() === month) addToDay(d.getDate(), e.name, '#0078d4');
    });
    // Family events (range)
    familyEvents.forEach(e => {
        if (!e.name) return;
        const from = e.dateFrom ? new Date(e.dateFrom + 'T00:00:00') : null;
        const to = e.dateTo ? new Date(e.dateTo + 'T00:00:00') : from;
        if (!from) return;
        for (let d = new Date(from); d <= (to || from); d.setDate(d.getDate() + 1)) {
            if (d.getFullYear() === year && d.getMonth() === month) addToDay(d.getDate(), e.name, '#2980b9');
        }
    });
    // Tasks with due dates
    tasks.forEach(t => {
        if (!t.date || !t.name) return;
        const d = new Date(t.date + 'T00:00:00');
        if (d.getFullYear() === year && d.getMonth() === month) addToDay(d.getDate(), t.name, '#e74c3c');
    });
    // Personal tasks
    personalTasks.forEach(t => {
        if (!t.date || !t.name) return;
        const d = new Date(t.date + 'T00:00:00');
        if (d.getFullYear() === year && d.getMonth() === month) addToDay(d.getDate(), t.name, '#8e44ad');
    });

    const today = new Date(); today.setHours(0,0,0,0);
    let html = `<div class="cal-header">
        <button class="cal-nav" onclick="calNav(-1)"><span class="material-icons-outlined">chevron_left</span></button>
        <h2 class="cal-month">${monthName}</h2>
        <button class="cal-nav" onclick="calNav(1)"><span class="material-icons-outlined">chevron_right</span></button>
    </div>
    <div class="cal-grid">
        <div class="cal-day-name">Sun</div><div class="cal-day-name">Mon</div><div class="cal-day-name">Tue</div><div class="cal-day-name">Wed</div><div class="cal-day-name">Thu</div><div class="cal-day-name">Fri</div><div class="cal-day-name">Sat</div>`;
    // Empty cells before first day
    for (let i = 0; i < firstDay; i++) html += '<div class="cal-cell cal-empty"></div>';
    for (let day = 1; day <= daysInMonth; day++) {
        const isToday = year === today.getFullYear() && month === today.getMonth() && day === today.getDate();
        const evts = monthEvents[day] || [];
        const tooltipLines = evts.map(e => e.label).join('\\n');
        const dotHtml = evts.slice(0, 3).map(e => `<span class="cal-dot" style="background:${e.color}"></span>`).join('');
        const extraDots = evts.length > 3 ? `<span class="cal-more">+${evts.length - 3}</span>` : '';
        html += `<div class="cal-cell${isToday ? ' cal-today' : ''}${evts.length ? ' cal-has-event' : ''}" ${evts.length ? `title="${tooltipLines}"` : ''}>
            <span class="cal-num">${day}</span>
            <div class="cal-dots">${dotHtml}${extraDots}</div>
            ${evts.length ? `<div class="cal-tooltip">${evts.map(e => `<div class="cal-tip-item" style="border-left:3px solid ${e.color}">${esc(e.label)}</div>`).join('')}</div>` : ''}
        </div>`;
    }
    html += '</div>';
    container.innerHTML = html;
}

function calNav(dir) {
    const container = document.getElementById('calendar-container');
    const cur = parseInt(container.dataset.monthOffset || '0');
    container.dataset.monthOffset = cur + dir;
    renderCalendar();
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

    let dataArr, key, rerender;
    if (tableId === 'tasks-table') { dataArr = tasks; key = STORAGE_KEYS.tasks; rerender = renderTasks; }
    else if (tableId === 'events-1p-table') { dataArr = events1p; key = STORAGE_KEYS.events1p; rerender = () => renderEvents('1p'); }
    else if (tableId === 'events-3p-table') { dataArr = events3p; key = STORAGE_KEYS.events3p; rerender = () => renderEvents('3p'); }
    else if (tableId === 'personal-tasks-table') { dataArr = personalTasks; key = STORAGE_KEYS.personalTasks; rerender = renderPersonalTasks; }
    else if (tableId === 'family-events-table') { dataArr = familyEvents; key = STORAGE_KEYS.familyEvents; rerender = renderFamilyEvents; }
    else if (tableId === 'finance-table') { dataArr = financeRecords; key = STORAGE_KEYS.financeRecords; rerender = renderFinance; }
    else if (tableId === 'work-events-table') { renderWorkEvents(); return; }
    else return;

    dataArr.sort((a, b) => {
        const va = (a[col] || '').toString().toLowerCase();
        const vb = (b[col] || '').toString().toLowerCase();
        return currentSort.dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    });

    saveData(key, dataArr);
    if (rerender) rerender();
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
    const allEvents = [
        ...events1p.map(e => ({...e, type:'1P', anchor: e.dateFrom || e.date})),
        ...events3p.map(e => ({...e, type:'3P', anchor: e.dateFrom || e.date}))
    ];
    const upcoming = allEvents.filter(e => {
        if (!e.anchor || !e.name) return false;
        const d = new Date(e.anchor + 'T00:00:00');
        return d >= today && d <= threeMonths;
    }).sort((a, b) => (a.anchor || '').localeCompare(b.anchor || ''));

    const eventsList = document.getElementById('home-events-list');
    if (eventsList) {
        if (upcoming.length === 0) {
            eventsList.innerHTML = '<div class="home-empty">No events in the next 3 months</div>';
        } else {
            eventsList.innerHTML = upcoming.map(e => {
                const dateStr = new Date(e.anchor + 'T00:00:00').toLocaleDateString('en-US', {month:'short', day:'numeric'});
                const pri = getEventPriorityInfo(e.priority || 'P3');
                const badgeStyle = e.type === '1P'
                    ? 'background:#e67e22;color:#fff;border:none'
                    : 'background:#fff3cd;color:#856404;border:1px solid #ffeaa7';
                return `<div class="home-event-row">
                    <span class="urgency-dot" style="background:${pri.color}" title="${pri.label}"></span>
                    <span class="home-event-badge" style="${badgeStyle};font-weight:700;padding:2px 8px;border-radius:10px;font-size:10px">${e.type}</span>
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
    migrateSidebarLinks();
    migrateWorkEventDates();
    renderSidebar();

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

// Global error handler - shows errors on screen for debugging
window.onerror = function(msg, url, line, col, error) {
    const d = document.createElement('div');
    d.style.cssText = 'position:fixed;bottom:0;left:0;right:0;background:red;color:white;padding:12px;font-size:13px;z-index:99999;font-family:monospace';
    d.textContent = 'JS ERROR: ' + msg + ' (line ' + line + ')';
    document.body.appendChild(d);
    console.error('Global error:', msg, url, line, col, error);
};

document.addEventListener('DOMContentLoaded', async () => {
    console.log('[PA] v4.12a DOMContentLoaded fired');
    try {
        const authed = await checkAuth();
        console.log('[PA] Auth check:', authed);
        if (authed) {
            showApp();
            await initApp();
            console.log('[PA] App initialized OK. Tasks:', tasks.length, 'Events1P:', events1p.length);
        } else {
            showLoginScreen();
            console.log('[PA] Login screen shown');
        }
    } catch (err) {
        console.error('[PA] Init error:', err);
        const d = document.createElement('div');
        d.style.cssText = 'position:fixed;bottom:0;left:0;right:0;background:red;color:white;padding:12px;font-size:13px;z-index:99999;font-family:monospace';
        d.textContent = 'INIT ERROR: ' + err.message;
        document.body.appendChild(d);
    }
});
