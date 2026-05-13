/**
 * EBM 課程落地頁 - 核心邏輯
 */
console.log('app.js 腳本已被載入');

// 全域狀態
const state = {
    user: null,
    activeTab: 'courses',
    courseFilter: 'all', // 預設顯示全部
    courseCategoryFilter: 'all', // 課程類別過濾
    promptFilter: 'all',
    courseSortOrder: 'asc',
    data: {
        courses: [],
        prompts: [],
        links: [],
        info: [],
        news: []
    },
    linkFilter: 'all',
    isInitialLoad: true, // 標記是否為初次載入
    isAdmin: false, // 是否為管理者旗標
    idToken: null, // 新增：儲存 Google ID Token
    showAllLogs: false // 是否顯示全量活動紀錄
};

/**
 * 讀取本地快取
 */
function loadFromCache() {
    const cachedData = localStorage.getItem('ebm_data_cache');
    if (cachedData) {
        try {
            const parsed = JSON.parse(cachedData);
            // 確保資料格式相容
            state.data = { ...state.data, ...parsed };
            console.log('已載入本地快取資料');
            return true;
        } catch (e) {
            console.warn('快取資料解析失敗');
        }
    }
    return false;
}

/**
 * 儲存資料至快取
 */
function saveToCache() {
    localStorage.setItem('ebm_data_cache', JSON.stringify(state.data));
    localStorage.setItem('ebm_cache_time', new Date().getTime());
}

// 初始化
function startApp() {
    console.log('DOM 已就緒，開始啟動應用程式...');
    initAuth();
    initTabs();
    initSubFilters(); // 初始化子濾鏡
    initEventListeners();
}

if (document.readyState === 'loading') {
    document.body ? startApp() : document.addEventListener('DOMContentLoaded', startApp);
} else {
    startApp();
}

/**
 * 身份驗證初始化
 */
function initAuth() {
    if (!CONFIG.GOOGLE_CLIENT_ID || CONFIG.GOOGLE_CLIENT_ID.includes('YOUR_GOOGLE')) {
        console.warn('請在 config.js 中設定有效的 GOOGLE_CLIENT_ID');
        return;
    }

    // 等待 Google Script 載入後進行初始化
    console.log('開始初始化登入驗證...');
    const checkGsi = setInterval(() => {
        if (typeof google !== 'undefined' && google.accounts.id) {
            console.log('Google Identity Services 已載入，目前 Origin:', window.location.origin);
            clearInterval(checkGsi);
            google.accounts.id.initialize({
                client_id: CONFIG.GOOGLE_CLIENT_ID,
                callback: handleCredentialResponse,
                auto_select: false,
                itp_support: true // 支援 iOS Safari 的智慧防追蹤 (ITP)
            });

            // 嘗試自動登入 (One Tap)
            google.accounts.id.prompt();

            // 檢查本地是否有快取的使用者資訊
            const cachedUser = localStorage.getItem('ebm_user');
            const cachedToken = localStorage.getItem('ebm_id_token');
            if (cachedUser && cachedToken) {
                try {
                    state.user = JSON.parse(cachedUser);
                    state.idToken = cachedToken;
                    updateAuthUI();

                    // 【優化】優先從快取載入並渲染，實現瞬時開屏
                    const hasCache = loadFromCache();
                    if (hasCache) {
                        renderNewsBanner();
                        renderData('courses');
                    }

                    // 背景同步最新資料
                    fetchInitialData();
                } catch (e) {
                    localStorage.removeItem('ebm_user');
                    localStorage.removeItem('ebm_id_token');
                }
            }

            // 渲染登入按鈕
            const signinBtn = document.getElementById('google-signin-btn');
            if (signinBtn) {
                console.log('正在渲染登入按鈕於 #google-signin-btn');
                google.accounts.id.renderButton(signinBtn, {
                    theme: 'outline',
                    size: 'large',
                    type: 'standard',
                    shape: 'rectangular',
                    text: 'signin_with',
                    logo_alignment: 'left',
                    ux_mode: 'popup' // 強制使用彈窗模式，避免行動裝置跳轉問題
                });
            } else {
                console.error('找不到登入按鈕容器 #google-signin-btn');
            }
        }
    }, 100);
}

// Google 登入回調
function handleCredentialResponse(response) {
    const idToken = response.credential;
    const payload = decodeJwtResponse(idToken);
    
    state.idToken = idToken;
    state.user = {
        name: payload.name,
        email: payload.email,
        picture: payload.picture
    };
    
    localStorage.setItem('ebm_user', JSON.stringify(state.user));
    localStorage.setItem('ebm_id_token', idToken);
    
    updateAuthUI();
    logActivity('login');
    fetchInitialData(); // 一次性抓取初始資料 (最新訊息 + 課程)
}

/**
 * 紀錄使用者活動
 */
async function logActivity(activity, details = '') {
    if (!state.user) return;
    try {
        await callApi('logActivity', {
            user: state.user.name || state.user.email,
            activity: activity,
            details: typeof details === 'object' ? JSON.stringify(details) : details
        });
    } catch (e) {
        console.warn('活動紀錄失敗', e);
    }
}

/**
 * 整合型初始資料抓取 (全量抓取 + 背景同步)
 */
async function fetchInitialData() {
    // 如果沒有快取資料，才顯示載入動畫
    const hasCourses = state.data.courses && state.data.courses.length > 0;
    if (!hasCourses) {
        const container = document.getElementById('courses-items');
        if (container) {
            container.innerHTML = `
                <div class="loader">
                    <div class="loading-text">正在同步最新資料...</div>
                </div>
            `;
        }
    }

    try {
        console.log('發起全量資料同步...');
        const result = await callApi('getInitialData', {
            sheetNames: ['最新訊息', '課程', 'Prompt', '資源連結', '競賽資訊'],
            userEmail: state.user ? state.user.email : null // 傳送 Email 供後端判定權限
        });

        if (result && result.success) {
            // 更新權限狀態
            state.isAdmin = !!result.isAdmin;
            updateAuthUI(); // 取得權限後更新 UI
            // 更新狀態
            state.data.news = result.news || [];
            state.data.courses = result.courses || [];
            state.data.prompts = result.prompts || [];
            state.data.links = result.links || [];
            state.data.info = result.info || [];

            // 存入快取
            saveToCache();

            // 重新渲染當前視圖
            renderNewsBanner();
            renderData(state.activeTab);
            console.log('資料同步完成');
        }
    } catch (e) {
        console.error('初始資料同步失敗:', e);
        // 如果完全沒資料（包含快取），才顯示錯誤
        if (!hasCourses) {
            showToast('同步失敗，請檢查網路連線', 'error');
        }
    } finally {
        state.isInitialLoad = false;
    }
}

function decodeJwtResponse(token) {
    let base64Url = token.split('.')[1];
    let base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    let jsonPayload = decodeURIComponent(atob(base64).split('').map(function (c) {
        return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
    }).join(''));
    return JSON.parse(jsonPayload);
}

/**
 * 從可能的 HTML 字串中提取純 URL (解決 GAS 傳回 <a> 標籤的問題)
 */
function getPureUrl(text) {
    if (!text) return '';
    // 如果包含 <a href="..."> 則提取 href 的內容
    const match = text.match(/href="([^"]+)"/) || text.match(/href='([^']+)'/);
    if (match && match[1]) return match[1];
    // 否則，如果包含 http 則提取第一個符合的 URL
    const urlMatch = text.match(/(https?:\/\/[^\s"'>]+)/);
    return urlMatch ? urlMatch[0] : text.trim();
}

/**
 * 將文字中的 URL 轉為可點擊連結 (避免重複包裹)
 */
function linkify(text) {
    if (!text) return '';
    // 如果已經包含 <a> 標籤，則不進行處理 (或進階處理)
    if (text.includes('<a ')) return text;

    const urlPattern = /(https?:\/\/[^\s\n<]+)/g;
    return text.replace(urlPattern, '<a href="$1" target="_blank">$1</a>');
}

/**
 * 格式化日期，去除非必要的 ISO 時間字串
 */
function formatDate(dateVal) {
    if (!dateVal) return '';
    let dateStr = String(dateVal);
    if (dateStr.includes('T')) {
        return dateStr.split('T')[0].replace(/-/g, '/');
    }
    if (dateVal instanceof Date) {
        const y = dateVal.getFullYear();
        const m = String(dateVal.getMonth() + 1).padStart(2, '0');
        const d = String(dateVal.getDate()).padStart(2, '0');
        return `${y}/${m}/${d}`;
    }
    return dateStr;
}

/**
 * 格式化日期時間 (YYYY/MM/DD HH:mm)
 */
function formatDateTime(dateVal) {
    if (!dateVal) return '';
    const date = new Date(dateVal);
    if (isNaN(date.getTime())) return String(dateVal);
    
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    const hh = String(date.getHours()).padStart(2, '0');
    const mm = String(date.getMinutes()).padStart(2, '0');
    return `${y}/${m}/${d} ${hh}:${mm}`;
}

/**
 * 人類語言化活動詳情
 */
function formatActivityDetail(action, detailsStr) {
    if (!detailsStr) return '';
    try {
        const details = JSON.parse(detailsStr);
        if (action === 'switch_tab') {
            const tabNames = { 
                courses: '課程資料', 
                prompts: 'Prompt 工具', 
                links: '資源連結', 
                info: '競賽資訊', 
                news: '管理設定',
                stats: '數據統計'
            };
            return `切換至「${tabNames[details.tab] || details.tab}」`;
        }
        if (action === 'check_in') return `簽到課程 ID: ${details.courseId}`;
        if (action === 'cancel_check_in') return `取消簽到課程 ID: ${details.courseId}`;
        return detailsStr;
    } catch (e) {
        return detailsStr;
    }
}

function updateAuthUI() {
    const loginSection = document.getElementById('login-section');
    const userInfo = document.getElementById('user-info');
    const adminBtn = document.getElementById('admin-add-btn');
    const mainContent = document.getElementById('main-content');
    const refreshBtn = document.getElementById('refresh-btn');

    if (state.user) {
        if (loginSection) loginSection.classList.add('hidden');
        userInfo.classList.remove('hidden');
        if (refreshBtn) refreshBtn.classList.remove('hidden');
        if (mainContent) mainContent.classList.remove('hidden');
        const stickyAuth = document.getElementById('sticky-auth-content');
        if (stickyAuth) stickyAuth.classList.remove('hidden');
        document.getElementById('user-name').textContent = state.user.name;

        // 修改：使用後端傳回的 isAdmin 旗標
        if (state.isAdmin) {
            if (adminBtn) adminBtn.classList.remove('hidden');
            const tabStats = document.getElementById('tab-stats');
            if (tabStats) tabStats.classList.remove('hidden');
        } else {
            if (adminBtn) adminBtn.classList.add('hidden');
            const tabStats = document.getElementById('tab-stats');
            if (tabStats) tabStats.classList.add('hidden');
        }
    } else {
        if (loginSection) loginSection.classList.remove('hidden');
        userInfo.classList.add('hidden');
        if (refreshBtn) refreshBtn.classList.add('hidden');
        adminBtn.classList.add('hidden');
        if (mainContent) mainContent.classList.add('hidden');
        const stickyAuth = document.getElementById('sticky-auth-content');
        if (stickyAuth) stickyAuth.classList.add('hidden');

        // 登出時清空內容
        clearTabContents();
    }
}

function clearTabContents() {
    const views = ['courses', 'prompts', 'links', 'info'];
    views.forEach(v => {
        const container = v === 'courses' ? document.getElementById('courses-items') : document.getElementById(`${v}-grid`);
        if (container) container.innerHTML = '<p class="empty-text">請先登入以查看內容</p>';
    });
}

function logout() {
    state.user = null;
    state.idToken = null;
    localStorage.removeItem('ebm_user');
    localStorage.removeItem('ebm_id_token');
    updateAuthUI();
    showToast('已登出');
    location.reload(); // 重新整理以清除狀態
}

/**
 * 取得最新訊息 (固定顯示)
 */
async function fetchNews() {
    const sheetMap = { 'news': '最新訊息' };
    const result = await callApi('getData', { type: sheetMap['news'] });
    if (result && result.success && result.data) {
        state.data.news = result.data;
        renderNewsBanner();
    } else {
        console.warn('公告 API 抓取失敗，使用模擬資料。');
        state.data.news = getDummyData('news');
        renderNewsBanner();
    }
}

function renderNewsBanner() {
    const container = document.getElementById('news-banner-container');
    if (!container) return;

    if (!state.data.news || state.data.news.length === 0) {
        container.innerHTML = '';
        return;
    }

    const now = new Date();
    // 1. 過濾：僅顯示日期已到或未設定日期的公告 (預約發佈功能)
    const availableNews = state.data.news.filter(item => {
        if (!item.建立日期) return true;
        const pubDate = new Date(item.建立日期);
        // 如果解析失敗則顯示，否則檢查日期是否已到 (忽略時分秒)
        if (isNaN(pubDate.getTime())) return true;

        const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        const compareDate = new Date(pubDate.getFullYear(), pubDate.getMonth(), pubDate.getDate());
        return compareDate <= today;
    });

    if (availableNews.length === 0) {
        container.innerHTML = '';
        return;
    }

    // 2. 排序：優先用日期，日期無效則用列號 (_row)
    const sortedNews = [...availableNews].sort((a, b) => {
        const dateA = new Date(a.建立日期).getTime();
        const dateB = new Date(b.建立日期).getTime();

        if (isNaN(dateA) || isNaN(dateB)) {
            return (b._row || 0) - (a._row || 0);
        }
        return dateB - dateA;
    });

    const latest = sortedNews[0];
    if (!latest) return;

    container.innerHTML = `
        <div class="news-banner">
            <span class="news-label">最新訊息</span>
            <div class="news-body">
                <div class="news-content">${linkify(latest.公告內容)}</div>
                <div class="news-date">📅 ${formatDate(latest.建立日期)}</div>
            </div>
        </div>
    `;
}

/**
 * API 請求封裝 (含 Token 驗證)
 */
async function callApi(action, payload = {}) {
    if (!CONFIG.GAS_URL || CONFIG.GAS_URL.includes('YOUR_GAS')) {
        showToast('API URL 未設定，請檢查 config.js', 'error');
        return null;
    }

    try {
        const idToken = state.idToken || localStorage.getItem('ebm_id_token');
        const response = await fetch(CONFIG.GAS_URL, {
            method: 'POST',
            body: JSON.stringify({
                idToken: idToken,
                action: action,
                ...payload
            })
        });

        // GAS Web App 會進行導向，fetch 預設會跟隨
        // 如果成功，解析 JSON
        const result = await response.json();
        return result;
    } catch (error) {
        console.error('API Error:', error);
        // 如果是 CORS 錯誤，通常是 GAS 端未正確處理或 Web App 權限問題
        return null;
    } finally {
        // 動態 Loader 會由 innerHTML 直接替換，不需另外隱藏全域 element
    }
}

/**
 * 標籤切換邏輯
 */
function initTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const tab = btn.dataset.tab;
            switchTab(tab);
        });
    });
}

function switchTab(tabId) {
    state.activeTab = tabId;

    // 更新按鈕樣式
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
    });

    // 更新內容視圖
    document.querySelectorAll('.content-view').forEach(view => {
        view.classList.toggle('active', view.id === `${tabId}-grid`);
    });

    // 【優化】因為已經全量抓取，直接渲染即可，不需要再等 API
    renderData(tabId);
    logActivity('switch_tab', { tab: tabId });
}

async function fetchData(type) {
    if (!state.user) {
        const container = type === 'courses' ? document.getElementById('courses-items') : document.getElementById(`${type}-grid`);
        if (container) container.innerHTML = '<p class="empty-text">請先登入以查看內容</p>';
        return;
    }

    // 先顯示載入中動畫
    const container = type === 'courses' ? document.getElementById('courses-items') : document.getElementById(`${type}-grid`);
    if (container) {
        container.innerHTML = `
            <div class="loader">
                <div class="loading-text">連線中...</div>
            </div>
        `;
    }

    // 呼叫 API (對應試算表名稱)
    const sheetMap = {
        'courses': '課程',
        'prompts': 'Prompt',
        'links': '資源連結',
        'info': '競賽資訊',
        'news': '最新訊息'
    };

    try {
        const result = await callApi('getData', { type: sheetMap[type] || type });

        if (result && result.success && result.data) {
            state.data[type] = result.data;
            renderData(type);
        } else {
            throw new Error('API return success false or no data');
        }
    } catch (error) {
        console.warn(`${type} API 抓取失敗，使用模擬資料。`, error);
        // 如果 API 失敗，使用模擬資料作為備援
        setTimeout(() => {
            const dummyData = getDummyData(type);
            state.data[type] = dummyData;
            renderData(type);
        }, 500);
    }
}

function renderData(type) {
    const grid = document.getElementById(`${type}-grid`);
    if (!grid) return;

    // 依照類型進行表格渲染
    switch (type) {
        case 'courses':
            renderCourseTable();
            break;
        case 'prompts':
            renderPromptTable();
            break;
        case 'links':
            renderLinkTable();
            break;
        case 'info':
            renderInfoTable();
            break;
        case 'statistics':
            fetchStatistics();
            break;
        default:
            grid.innerHTML = '<p class="empty-text">尚無資料</p>';
    }
}

function renderCourseTable() {
    const container = document.getElementById('courses-items');
    if (!container) return;

    let rawItems = state.data['courses'] || [];

    // 輔助函式：處理民國年轉西元
    const getYear = (item) => {
        if (!item.年) return new Date().getFullYear();
        const y = parseInt(item.年);
        return y < 1900 ? y + 1911 : y;
    };

    // 篩選邏輯
    let allItems = rawItems;

    // 1. 位置篩選 (實體/線上)
    if (state.courseFilter !== 'all') {
        allItems = allItems.filter(item => {
            if (state.courseFilter === 'physical') return item.課程地點 !== '線上';
            if (state.courseFilter === 'online') return item.課程地點 === '線上';
            return true;
        });
    }

    // 2. 類別篩選 (標籤過濾)
    if (state.courseCategoryFilter !== 'all') {
        allItems = allItems.filter(item => {
            if (!item.類別) return false;
            const cats = item.類別.split(/[、,，]/).map(c => c.trim());
            return cats.includes(state.courseCategoryFilter);
        });
    }

    const now = new Date();


    // 2. 課程分類
    const today = [];
    const upcoming = [];
    const past = [];

    allItems.forEach(item => {
        const year = getYear(item);
        const courseDate = new Date(year, parseInt(item.月) - 1, parseInt(item.日));
        const dateStr = courseDate.toDateString();
        const nowStr = now.toDateString();

        if (dateStr === nowStr) {
            today.push(item);
        } else if (courseDate > now) {
            upcoming.push(item);
        } else {
            past.push(item);
        }
    });

    // 排序
    const sortByDate = (a, b) => {
        const dA = new Date(getYear(a), parseInt(a.月) - 1, parseInt(a.日));
        const dB = new Date(getYear(b), parseInt(b.月) - 1, parseInt(b.日));
        return state.courseSortOrder === 'asc' ? dA - dB : dB - dA;
    };

    today.sort(sortByDate);
    upcoming.sort(sortByDate);
    // 已結束課程固定由新到舊排序
    past.sort((a, b) => {
        const dA = new Date(getYear(a), parseInt(a.月) - 1, parseInt(a.日));
        const dB = new Date(getYear(b), parseInt(b.月) - 1, parseInt(b.日));
        return dB - dA;
    });

    // 提取不重複的類別
    let rawCategories = [];
    rawItems.forEach(item => {
        if (item.類別) {
            const cats = item.類別.split(/[、,，]/);
            rawCategories.push(...cats.map(c => c.trim()));
        }
    });
    const categories = ['all', ...new Set(rawCategories)];

    let html = '';

    // 渲染類別標籤篩選列
    html += `
        <div class="category-filter-bar tag-cloud">
            ${categories.map(cat => `
                <button class="category-tag-btn ${state.courseCategoryFilter === cat ? 'active' : ''}" 
                        onclick="filterCourseByCategory('${cat}')">
                    ${cat === 'all' ? '全部類別' : cat}
                </button>
            `).join('')}
        </div>
    `;

    // 第一桶：今日課程
    if (today.length > 0) {
        html += `<div class="bucket bucket-today">
            <h3 class="bucket-title" onclick="toggleBucket(this.parentElement)">
                <span class="bucket-toggle"></span>
                <span class="bucket-icon">🔥</span> 即將開始課程
            </h3>
            <div class="bucket-content">
                ${renderGenericTable(today, true)}
            </div>
        </div>`;
    }

    // 第二桶：未來課程
    const filterText = state.courseFilter === 'physical' ? '實體課程' : (state.courseFilter === 'online' ? '線上課程' : '所有課程');
    if (upcoming.length > 0) {
        html += `<div class="bucket bucket-upcoming">
            <h3 class="bucket-title" onclick="toggleBucket(this.parentElement)">
                <span class="bucket-toggle"></span>
                <span class="bucket-icon">📅</span> ${filterText}
            </h3>
            <div class="bucket-content">
                ${renderGenericTable(upcoming)}
            </div>
        </div>`;
    }

    // 第三桶：已結束課程
    if (past.length > 0) {
        // 按年份分組
        const pastByYear = {};
        past.forEach(item => {
            const year = getYear(item);
            if (!pastByYear[year]) pastByYear[year] = [];
            pastByYear[year].push(item);
        });

        // 年份從新到舊排序
        const years = Object.keys(pastByYear).sort((a, b) => b - a);
        const currentYear = new Date().getFullYear();

        html += `
            <div class="bucket bucket-past">
                <h3 class="bucket-title" onclick="toggleBucket(this.parentElement)">
                    <span class="bucket-toggle"></span>
                    <span class="bucket-icon">📁</span> 已結束課程 (${past.length})
                </h3>
                <div class="bucket-content">
                    ${years.map(year => `
                        <div class="sub-bucket ${parseInt(year) === currentYear ? '' : 'collapsed'}">
                            <h4 class="sub-bucket-title" onclick="this.parentElement.classList.toggle('collapsed')">
                                <span class="bucket-toggle"></span> ${year} 年
                            </h4>
                            <div class="sub-bucket-content">
                                ${renderGenericTable(pastByYear[year], false, 'desc')}
                            </div>
                        </div>
                    `).join('')}
                </div>
            </div>
        `;
    }

    if (allItems.length === 0) {
        html = '<p class="empty-text">尚無課程資料</p>';
    }

    container.innerHTML = html;
}


function renderGenericTable(items, isUrgent = false, sortOrderOverride = null) {
    const now = new Date();
    const currentSortOrder = sortOrderOverride || state.courseSortOrder;

    // 輔助函式：處理民國年轉西元
    const getYear = (item) => {
        if (!item.年) return new Date().getFullYear();
        const y = parseInt(item.年);
        return y < 1900 ? y + 1911 : y;
    };

    return `
        <div class="course-table-container ${isUrgent ? 'urgent' : ''}">
            <table class="course-table">
                <thead>
                    <tr>
                        <th style="width: 150px; ${sortOrderOverride ? '' : 'cursor: pointer;'}" 
                            ${sortOrderOverride ? '' : 'onclick="toggleSortCourses()"'} >
                            課程日期 ${currentSortOrder === 'asc' ? '🔼' : '🔽'}
                        </th>
                        <th>課程內容</th>
                        <th style="width: 150px">出席人員</th>
                        <th style="width: 100px">操作</th>
                    </tr>
                </thead>
                <tbody>
                    ${items.map(item => {
        const courseDate = new Date(getYear(item), parseInt(item.月) - 1, parseInt(item.日));
        const days = ['日', '一', '二', '三', '四', '五', '六'];
        const dayName = isNaN(courseDate.getTime()) ? '' : days[courseDate.getDay()];
        const dateCode = (item.月 && item.日) ? `${item.月}/${item.日}${dayName ? ` (${dayName})` : ''}` : '日期待定';

        const isOnline = item.課程地點 === '線上';
        const attendees = item.出席人員 || [];
        const hasCheckedIn = state.user && (typeof attendees === 'string' ? attendees.split('、') : attendees).includes(state.user.name);

        const deadline = new Date(courseDate.getTime());
        deadline.setHours(10, 0, 0, 0);
        const isPastDeadline = now > deadline && now.toDateString() === courseDate.toDateString();
        const isPastDay = now > courseDate && now.toDateString() !== courseDate.toDateString(); // 是否已過當天

        // 判定前 7 天至當天
        const sevenDaysMillis = 7 * 24 * 60 * 60 * 1000;
        const startHighlightDate = new Date(courseDate.getTime() - sevenDaysMillis);
        const isUpcomingSoon = !isOnline && now >= startHighlightDate && !isPastDay;

        const hasPrePostTask = item['課前/課後事項'] && item['課前/課後事項'].trim() !== '';
        const hasVideo = item['課程影音'] && item['課程影音'].trim() !== '';

        return `
                            <tr class="course-row ${isOnline ? 'tr-online' : 'tr-physical'} ${isUpcomingSoon ? 'tr-upcoming-soon' : ''}">
                                <td data-label="課程日期">
                                    <div class="mobile-header-row">
                                        <span class="table-date">${dateCode}</span>
                                        <div class="category-stack flex-row">
                                            ${(item.類別 || '').split(/[、,，]/).map(cat => cat.trim()).filter(cat => cat).map(cat => `<span class="category-tag highlight clickable" onclick="filterCourseByCategory('${cat}', event)">${cat}</span>`).join('')}
                                        </div>
                                    </div>
                                </td>
                                <td data-label="課程內容">
                                    <div class="table-title">${item.課程名稱}${hasVideo ? `<a href="${getPureUrl(item['課程影音'])}" target="_blank" class="link-icon ml-1" title="課程影音">🔗</a>` : ''}</div>
                                    <div class="mobile-meta-grid">
                                        <div class="table-meta">
                                            <span>${item.講師 || '未定'}</span>
                                        </div>
                                        <div class="table-meta">
                                            <span class="location-tag ${isOnline ? 'online' : 'physical'}">📍 ${item.課程地點}</span>
                                        </div>
                                        ${item.課程時間 ? `
                                        <div class="table-meta">
                                            <span class="time-tag">⏰ ${item.課程時間}</span>
                                        </div>` : ''}
                                    </div>
                                    
                                    <!-- 課前/課後事項展開 -->
                                    <button class="btn-expand ${!hasPrePostTask ? 'btn-dimmed' : 'btn-has-content'}" 
                                            onclick="toggleDetails(this, '${item.編號}')" 
                                            ${!hasPrePostTask ? 'disabled' : ''}>
                                        查看課前/課後事項 <span>▼</span>
                                    </button>
                                    <div id="details-${item.編號}" class="course-details hidden">
                                        <div class="details-content">
                                            <div class="pre-wrap mt-1">${linkify(item['課前/課後事項'] || '目前尚無特殊需求。')}</div>
                                        </div>
                                    </div>
                                </td>
                                <td data-label="出席人員">
                                    <div class="table-attendees">
                                        ${isOnline ? '<span class="attendee-tag online-join">自行參加</span>' :
                (attendees.length > 0 ? (typeof attendees === 'string' ? attendees.split('、') : attendees).map(name => `<span class="attendee-tag">${name}</span>`).join('') : '<span class="empty-attendees">尚未確認</span>')}
                                    </div>
                                </td>
                                <td data-label="操作">
                                    ${isPastDay ?
                `<button class="btn-disabled" disabled>結束報名</button>` :
                (isOnline ?
                    (item.報名連結 ? `<button class="btn-secondary" onclick="window.open('${getPureUrl(item.報名連結)}', '_blank')"> 前往報名</button>` : '') :
                    `
                                                ${hasCheckedIn ? `<button class="btn-secondary btn-cancel-出席" onclick="cancelCheckIn('${item.編號}')" ${isPastDeadline ? 'disabled' : ''}>取消出席</button>` :
                        `<button class="btn-primary" onclick="checkIn('${item.編號}')" ${isPastDeadline ? 'disabled' : ''}>確認出席</button>`}
                                                ${isPastDeadline ? '<div class="deadline-text">今日已截止</div>' : ''}
                                            `
                )
            }
                                </td>
                            </tr>
                        `;
    }).join('')}
                </tbody>
            </table>
        </div>
    `;
}

function toggleBucket(bucketEl) {
    bucketEl.classList.toggle('collapsed');
}

function toggleDetails(btn, id) {
    const details = document.getElementById(`details-${id}`);
    const isHidden = details.classList.contains('hidden');

    // 關閉其他已展開的 (選用)
    // document.querySelectorAll('.course-details').forEach(d => d.classList.add('hidden'));

    if (isHidden) {
        details.classList.remove('hidden');
        btn.querySelector('span').textContent = '▲';
        btn.classList.add('active');
    } else {
        details.classList.add('hidden');
        btn.querySelector('span').textContent = '▼';
        btn.classList.remove('active');
    }
}

function renderPromptTable() {
    const container = document.getElementById('prompts-grid');
    if (!container) return;

    let allItems = state.data['prompts'] || [];
    if (allItems.length === 0) {
        container.innerHTML = '<p class="empty-text">尚無 Prompt 資料</p>';
        return;
    }

    // 1. 取得所有不重複類別 (處理「、」分隔的情形)
    let rawCategories = [];
    allItems.forEach(item => {
        if (item.類別) {
            const cats = item.類別.split(/[、,，]/);
            rawCategories.push(...cats.map(c => c.trim()));
        }
    });
    const categories = ['all', ...new Set(rawCategories)];

    // 2. 篩選資料
    const filteredItems = state.promptFilter === 'all'
        ? allItems
        : allItems.filter(item => {
            if (!item.類別) return false;
            const cats = item.類別.split(/[、,，]/).map(c => c.trim());
            return cats.includes(state.promptFilter);
        });

    const tableHtml = `
        <div class="prompt-tag-container tag-cloud">
            ${categories.map(cat => `
                <button class="sub-filter-btn ${state.promptFilter === cat ? 'active' : ''}" 
                        onclick="filterPrompts('${cat}')">
                    ${cat === 'all' ? '全部' : cat}
                </button>
            `).join('')}
        </div>
        <div class="course-table-container">
            <table class="course-table">
                <thead>
                    <tr>
                        <th style="width: 150px">類別與操作</th>
                        <th>Prompt 內容</th>
                    </tr>
                </thead>
                <tbody>
                    ${filteredItems.map(item => `
                        <tr>
                            <td data-label="類別/複製">
                                <div class="prompt-header-row">
                                    <div class="category-stack flex-row">
                                        ${(item.類別 || '一般').split(/[、,，]/).map(cat => cat.trim()).filter(cat => cat).map(cat => `<span class="category-tag highlight">${cat}</span>`).join('')}
                                    </div>
                                    <button class="icon-btn-copy" onclick="copyText(this, \`${item.Prompt.replace(/`/g, '\\`').replace(/\$/g, '\\$')}\`)" title="複製文字">
                                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>
                                    </button>
                                </div>
                            </td>
                            <td data-label="Prompt 內容" class="align-top">
                                <div class="prompt-content">${(item.Prompt || '無內容').trim()}</div>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
    container.innerHTML = tableHtml;
}

function filterPrompts(category) {
    state.promptFilter = category;
    renderPromptTable();
}

function renderLinkTable() {
    const container = document.getElementById('links-grid');
    if (!container) return;

    let allItems = state.data['links'] || [];
    if (allItems.length === 0) {
        container.innerHTML = '<p class="empty-text">尚無資源連結</p>';
        return;
    }

    // 1. 取得所有不重複類別
    let rawCategories = [];
    allItems.forEach(item => {
        if (item.類別) {
            const cats = item.類別.split(/[、,，]/);
            rawCategories.push(...cats.map(c => c.trim()));
        }
    });
    const categories = ['all', ...new Set(rawCategories)];

    // 2. 篩選資料
    const filteredItems = state.linkFilter === 'all'
        ? allItems
        : allItems.filter(item => {
            if (!item.類別) return false;
            const cats = item.類別.split(/[、,，]/).map(c => c.trim());
            return cats.includes(state.linkFilter);
        });

    // 依照建立日期排序 (新 > 舊)
    filteredItems.sort((a, b) => {
        const dateA = new Date(a.建立日期 || 0);
        const dateB = new Date(b.建立日期 || 0);
        return dateB - dateA;
    });

    const tableHtml = `
        <div class="prompt-tag-container">
            ${categories.map(cat => `
                <button class="sub-filter-btn ${state.linkFilter === cat ? 'active' : ''}" 
                        onclick="filterLinks('${cat}')">
                    ${cat === 'all' ? '全部' : cat}
                </button>
            `).join('')}
        </div>
        <div class="course-table-container">
            <table class="course-table">
                <thead>
                    <tr>
                        <th style="width: 120px">類別與連結</th>
                        <th>資源主題與內容</th>
                    </tr>
                </thead>
                <tbody>
                    ${filteredItems.map(item => {
        const url = getPureUrl(item.網址 || item.連結 || '');
        return `
                        <tr>
                            <td data-label="標頭">
                                <div class="prompt-header-row">
                                    <div class="category-stack flex-row">
                                        ${(item.類別 || '').split(/[、,，]/).map(cat => cat.trim()).filter(cat => cat).map(cat => `<span class="category-tag highlight">${cat}</span>`).join('')}
                                    </div>
                                </div>
                            </td>
                            <td data-label="資源內容">
                                <div class="resource-title-row">
                                    <span class="table-title">${item.名稱 || '無名稱'}${url ? `<a href="${url}" target="_blank" class="link-icon ml-1" title="開啟連結">🔗</a>` : ''}</span>
                                    ${item.建立日期 ? `<span class="table-date-sub opacity-50" style="font-size: 0.7rem; margin-left: auto;">${formatDate(item.建立日期)}</span>` : ''}
                                </div>
                                <div class="table-meta" style="margin-top: 4px;">
                                    <span>${item['單位/作者'] || item.單位 || ''}</span>
                                </div>
                                ${item.說明 ? `<div class="mt-1 opacity-75 pre-wrap" style="font-size: 0.85rem; line-height: 1.5">${linkify(item.說明)}</div>` : ''}
                            </td>
                        </tr>
                    `;
    }).join('')}
                </tbody>
            </table>
        </div>
    `;
    container.innerHTML = tableHtml;
}

function filterLinks(category) {
    state.linkFilter = category;
    renderLinkTable();
}

function renderInfoTable() {
    const container = document.getElementById('info-grid');
    if (!container) return;

    let items = state.data['info'] || [];
    if (items.length === 0) {
        container.innerHTML = '<p class="empty-text">尚無競賽資訊</p>';
        return;
    }

    // 依照建立日期排序 (新 > 舊)
    items.sort((a, b) => {
        const dateA = new Date(a.建立日期 || 0);
        const dateB = new Date(b.建立日期 || 0);
        return dateB - dateA;
    });

    const tableHtml = `
        <div class="course-table-container">
            <table class="course-table">
                <thead>
                    <tr>
                        <th style="width: 120px">發佈單位</th>
                        <th>資訊內容</th>
                    </tr>
                </thead>
                <tbody>
                    ${items.map(item => {
        const content = item.內容 || '無內容';
        const url = getPureUrl(item.連結 || item.網址 || content);
        const hasLink = (item.連結 || item.網址 || content.includes('<a'));

        return `
                        <tr>
                            <td data-label="單位"><span class="category-tag highlight">${item.單位 || ''}</span></td>
                            <td data-label="內容">
                                <div class="resource-title-row">
                                    <span class="table-title" style="color: var(--text-main); font-weight: 500;">${item.內容 || '無內容'}</span>
                                    <span class="table-date-sub">📅 ${formatDate(item.建立日期)}</span>
                                </div>
                                ${item.附註 ? `<div class="table-meta mt-1" style="font-size: 0.85rem; background: #f8fafc; padding: 4px 10px; border-radius: 6px;">💡 ${linkify(item.附註)}</div>` : ''}
                            </td>
                        </tr>
                    `}).join('')}
                </tbody>
            </table>
        </div>
    `;
    container.innerHTML = tableHtml;
}

function showModal(title, body, type = 'info') {
    document.getElementById('modal-title').textContent = title;
    document.getElementById('modal-body').innerHTML = body;
    document.getElementById('modal-overlay').classList.remove('hidden');

    // 如果是表單模式，調整按鈕顯示
    const confirmBtn = document.querySelector('.btn-confirm');
    const cancelBtn = document.querySelector('.btn-cancel');
    if (type === 'form') {
        confirmBtn.classList.remove('hidden');
        cancelBtn.textContent = '取消';
    } else {
        confirmBtn.classList.add('hidden');
        cancelBtn.textContent = '關閉';
    }
}

/**
 * 開啟管理者表單
 */
function openAdminForm(existingItem = null) {
    const type = state.activeTab;
    const title = existingItem ? `修改${getTabTitle(type)}` : `新增${getTabTitle(type)}`;

    let formHtml = `<form id="admin-form" class="admin-form">`;

    const fields = getFieldsForType(type);
    fields.forEach(field => {
        if (field === '編號' && type === 'courses' && !existingItem) {
            formHtml += `<div class="form-group">
                <label>編號</label>
                <input type="text" name="${field}" value="${generateCourseId()}" readonly class="input-readonly">
                <small>系統自動生成</small>
            </div>`;
        } else if (['建立日期', '最後更新', '出席人員', '編號'].includes(field)) {
            // 隱藏或唯讀欄位
            if (existingItem) {
                formHtml += `<input type="hidden" name="${field}" value="${existingItem[field] || ''}">`;
            }
        } else {
            const val = existingItem ? (existingItem[field] || '') : '';
            if (field === 'Prompt' || field === '內容' || field === '說明' || field === '課前/課後事項' || field === '附註' || field === '公告內容') {
                formHtml += `<div class="form-group">
                    <label>${field}</label>
                    <textarea name="${field}" rows="4">${val}</textarea>
                </div>`;
            } else {
                formHtml += `<div class="form-group">
                    <label>${field}</label>
                    <input type="text" name="${field}" value="${val}">
                </div>`;
            }
        }
    });

    formHtml += `</form>`;

    showModal(title, formHtml, 'form');

    // 儲存目前正在編輯的項目，供 save 時參考
    state.editingItem = existingItem;
}

function getTabTitle(type) {
    const titles = { courses: '課程', prompts: 'Prompt', links: '資源連結', info: '競賽資訊', news: '最新訊息' };
    return titles[type] || '資料';
}

function getFieldsForType(type) {
    const fieldMap = {
        courses: ['編號', '年', '月', '日', '星期', '課程時間', '課程地點', '學分類別', '院內學分', '課程類別', '課程名稱', '講師', '類別', '課前/課後事項', '報名連結', '課程影音', '附註', '最後更新'],
        prompts: ['建立日期', '類別', 'Prompt', '最後更新'],
        links: ['建立日期', '類別', '名稱', '單位/作者', '說明', '網址', '最後更新'],
        info: ['建立日期', '單位', '內容', '附註', '最後更新'],
        news: ['建立日期', '公告內容', '最後更新']
    };
    return fieldMap[type] || [];
}

function generateCourseId() {
    const year = new Date().getFullYear();
    const courses = state.data.courses || [];
    const yearPrefix = `A${year}`;
    const yearCourses = courses.filter(c => c.編號 && c.編號.startsWith(yearPrefix));

    let maxId = 0;
    yearCourses.forEach(c => {
        const num = parseInt(c.編號.replace(yearPrefix, ''));
        if (num > maxId) maxId = num;
    });

    return `${yearPrefix}${(maxId + 1).toString().padStart(2, '0')}`;
}

async function saveAdminData() {
    const form = document.getElementById('admin-form');
    if (!form) return;

    const formData = new FormData(form);
    const item = {};
    formData.forEach((value, key) => {
        item[key] = value;
    });

    const type = state.activeTab;
    const isNew = !state.editingItem;
    const nowStr = new Date().toLocaleDateString('zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit' }).replace(/\//g, '/');
    const timeStr = new Date().toLocaleTimeString('zh-TW', { hour12: false, hour: '2-digit', minute: '2-digit' });
    const fullTs = `${nowStr} ${timeStr}`;

    if (isNew) {
        if (getFieldsForType(type).includes('建立日期')) {
            item['建立日期'] = nowStr;
        }
    } else {
        // 修改時帶入最後更新
        if (getFieldsForType(type).includes('最後更新')) {
            item['最後更新'] = fullTs;
        }
        // 保留原有的建立日期
        if (state.editingItem['建立日期']) {
            item['建立日期'] = state.editingItem['建立日期'];
        }
        if (state.editingItem['出席人員']) {
            item['出席人員'] = state.editingItem['出席人員'];
        }
    }

    showToast('正在儲存...');
    const sheetMap = {
        'courses': '課程',
        'prompts': 'Prompt',
        'links': '資源連結',
        'info': '競賽資訊',
        'news': '最新訊息'
    };
    const result = await callApi('saveData', {
        type: sheetMap[type] || type,
        item: item,
        isNew: isNew,
        index: isNew ? -1 : state.data[type].indexOf(state.editingItem)
    });

    if (result && result.success) {
        showToast('儲存成功！');
        document.getElementById('modal-overlay').classList.add('hidden');
        // 【優化】儲存後同步全量資料
        fetchInitialData();
    } else {
        showToast('儲存失敗，請檢查權限或網路。', 'error');
        // 模擬更新
        if (isNew) {
            state.data[type].push(item);
        } else {
            const idx = state.data[type].indexOf(state.editingItem);
            if (idx > -1) state.data[type][idx] = item;
        }
        renderData(type);
        document.getElementById('modal-overlay').classList.add('hidden');
    }
}

function toggleSort() {
    state.courseSortOrder = state.courseSortOrder === 'asc' ? 'desc' : 'asc';
    renderData('courses');
}

function createCard(type, item) {
    // 課程由 renderCourseTable 處理，此處僅處理其他類型
    // Prompt 類型
    if (type === 'prompts') {
        return `
            <div class="card prompt-card">
                <div class="card-header">
                    <span class="category-tag">${item.類別 || '一般'}</span>
                </div>
                <h3>${item.編號}</h3>
                <div class="prompt-content">${item.Prompt || '無內容'}</div>
                <div class="card-footer">
                    <button class="btn-secondary" onclick="copyText(this, \`${item.Prompt}\`)">複製文字</button>
                </div>
            </div>
        `;
    }
    return `<div class="card"><h3>${item.title || item.name || item.編號 || '資料'}</h3></div>`;
}

function toggleSortCourses() {
    state.courseSortOrder = state.courseSortOrder === 'asc' ? 'desc' : 'asc';
    renderData('courses');
}

function filterCourseByCategory(category, event) {
    if (event) {
        event.stopPropagation(); // 避免觸發 tr 的點擊事件 (如果以後有的話)
    }
    state.courseCategoryFilter = category;
    renderData('courses');
}

function initSubFilters() {
    const subBtns = document.querySelectorAll('.sub-filter-btn');
    if (!subBtns.length) return;

    subBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            subBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.courseFilter = btn.dataset.location;
            renderData('courses');
        });
    });
}

function copyText(btn, text) {
    navigator.clipboard.writeText(text).then(() => {
        const originalHTML = btn.innerHTML;
        btn.innerHTML = '<span class="copied-text">已複製！</span>';
        btn.classList.add('copied');
        btn.disabled = true; // 避免連按
        setTimeout(() => {
            btn.innerHTML = originalHTML;
            btn.classList.remove('copied');
            btn.disabled = false;
        }, 2000);
    });
}

/**
 * 其他事件
 */
function initEventListeners() {
    const searchInput = document.getElementById('search-input');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase();
            // 實作搜尋過濾...
        });
    }

    const sortBtn = document.getElementById('toggle-sort');
    if (sortBtn) {
        sortBtn.remove();
    }

    // Modal 關閉
    document.querySelectorAll('.close-modal, .btn-cancel, .btn-confirm').forEach(btn => {
        btn.addEventListener('click', () => {
            document.getElementById('modal-overlay').classList.add('hidden');
        });
    });

    document.getElementById('modal-overlay').addEventListener('click', (e) => {
        if (e.target.id === 'modal-overlay') {
            e.target.classList.add('hidden');
        }
    });

    // 管理者按鈕
    const adminAddBtn = document.getElementById('admin-add-btn');
    if (adminAddBtn) {
        adminAddBtn.addEventListener('click', () => {
            openAdminForm();
        });
    }

    // Modal 確認按鈕
    const confirmBtn = document.querySelector('.btn-confirm');
    if (confirmBtn) {
        confirmBtn.onclick = () => {
            saveAdminData();
        };
    }

    // 重新整理按鈕
    const refreshBtn = document.getElementById('refresh-btn');
    if (refreshBtn) {
        refreshBtn.onclick = () => {
            showToast('正在同步最新資料...');
            fetchInitialData();
        };
    }
}

function showToast(message, type = 'info') {
    const toast = document.getElementById('toast');
    toast.textContent = message;
    toast.className = `toast ${type}`;
    toast.classList.remove('hidden');
    setTimeout(() => toast.classList.add('hidden'), 3000);
}

// 模擬資料生成器
function getDummyData(type) {
    const data = {
        courses: [
            {
                編號: 'A202601',
                課程名稱: 'SR文章篩選-Covidence工具介紹AI協助搜尋與評讀',
                講師: '大林慈濟醫院 何彥蓉醫師',
                類別: '搜尋、評讀',
                月: '3',
                日: '6',
                課程地點: '線上',
                課程時間: '15:00-17:00',
                報名連結: '',
                '課程影音': '',
                '課前/課後事項': '',
                出席人員: []
            },
            {
                編號: 'A202602',
                課程名稱: '搜尋-進階',
                講師: '成大圖書館 方靜如館員',
                類別: '搜尋',
                月: '4',
                日: '17',
                課程地點: '線上',
                課程時間: '14:00-16:00',
                報名連結: 'https://nlms.tzuchi.com.tw/tzuchi/edurd/register_course/course_opration.php?id=81538',
                '課程影音': 'https://youtu.be/r2HHSGpi1H0',
                '課前/課後事項': '',
                出席人員: ['李大華']
            },
            {
                編號: 'A202612',
                課程名稱: '(三)加強競賽著重項目',
                講師: '羅振旭',
                類別: '評讀',
                月: '6',
                日: '10',
                課程地點: 'EBM教室',
                課程時間: '18:00-20:00',
                '課前/課後事項': '1.評讀兩篇文章\n2.完成報告，題目在這裡',
                出席人員: ['可可']
            }
        ],
        prompts: [
            {
                類別: 'PICO',
                Prompt: '##你是一位EBM專家，擅長從臨床情境中快速萃取並拆解前景問題。\n-(臨床情境)\n-請用繁體中文以條列式呈現，顯示3-5個PICO；每項以完整句子開頭，並以(P..., I..., C..., O...)標示。'
            },
            {
                類別: '搜尋、評讀',
                Prompt: '請將以下問題使用PCIO模組拆解，需要有英文同義詞、Mesh及Emtree，並使用表格呈現。'
            }
        ],
        links: [
            { 建立日期: '2026/04/23', 類別: '搜尋', 名稱: 'PubMed搜尋小工具', '單位/作者': 'KUNFENG LEE', 說明: '', 網址: 'https://chatgpt.com/g/g-68987e6b6c8481918e739a943e040b07-pubmed-sou-xun-xiao-gong-ju', 最後更新: '2026/04/23' },
            { 建立日期: '2026/04/23', 類別: '評讀', 名稱: 'EBM - Appraisal Tool for SR/MA or RCT', '單位/作者': 'KUNFENG LEE', 說明: 'Based on CEBM Critical Appraisal Tool for SR/MA or RCT. Please choose language from the options below.', 網址: 'https://chatgpt.com/g/g-QVp1C1uar-ebm-appraisal-tool-for-sr-ma-or-rct', 最後更新: '2026/04/23' }
        ],
        info: [
            { 建立日期: '2026/01/30', 單位: '醫策會訊息', 內容: '2026年【NCMEA國家臨床醫學教育獎】活動辦法', 附註: '', 最後更新: '' },
            { 建立日期: '2026/04/22', 單位: '醫策會訊息', 內容: '臨床組參賽編號：2026C2019\n醫策會官方EBM競賽辦法請點我', 附註: '', 最後更新: '' },
            { 建立日期: '2026/04/23', 單位: '慈濟醫院', 內容: '慈濟跨院區競賽題目-臨床情境題', 附註: '*僅供內部練習', 最後更新: '' }
        ],
        news: [
            { 建立日期: '2026/04/24', 公告內容: '第一次課程在4/29，記得準時出席', 最後更新: '' }
        ]
    };
    return data[type] || [];
}

async function checkIn(courseId) {
    if (!state.user) {
        showToast('請先登入', 'warning');
        return;
    }

    const course = state.data.courses.find(c => c.編號 === courseId);
    if (!course) return;

    // 再次檢查截止時間
    const now = new Date();
    const y = parseInt(course.年);
    const year = y < 1900 ? y + 1911 : y;
    const courseDate = new Date(year, parseInt(course.月) - 1, parseInt(course.日));
    const deadline = new Date(courseDate.getTime());
    deadline.setHours(10, 0, 0, 0);

    if (now > deadline && now.toDateString() === courseDate.toDateString()) {
        showToast('今日 10:00 以後不可再辦理出席確認', 'error');
        return;
    }

    // 樂觀更新
    let attendees = course.出席人員 || '';
    let attendeeList = typeof attendees === 'string' ? (attendees ? attendees.split('、') : []) : attendees;

    if (!attendeeList.includes(state.user.name)) {
        attendeeList.push(state.user.name);
        course.出席人員 = attendeeList.join('、');
        renderData('courses');
    } else {
        showToast('您已在出席名單中');
        return;
    }

    await callApi('checkIn', { courseId: courseId, email: state.user.email, name: state.user.name });
    logActivity('check_in', { courseId: courseId });
}

async function cancelCheckIn(courseId) {
    if (!state.user) return;

    const course = state.data.courses.find(c => c.編號 === courseId);
    if (!course) return;

    // 檢查截止時間 (略)

    // 樂觀更新：移除名字
    let attendees = course.出席人員 || '';
    let attendeeList = typeof attendees === 'string' ? (attendees ? attendees.split('、') : []) : attendees;

    const index = attendeeList.indexOf(state.user.name);
    if (index > -1) {
        attendeeList.splice(index, 1);
        course.出席人員 = attendeeList.join('、');
        renderData('courses');
    }

    await callApi('cancelCheckIn', { courseId: courseId, email: state.user.email, name: state.user.name });
    logActivity('cancel_check_in', { courseId: courseId });
}

/**
 * 抓取統計數據
 */
async function fetchStatistics() {
    const container = document.getElementById('statistics-grid');
    if (!container) return;

    container.innerHTML = `
        <div class="loader">
            <div class="loading-text">正在分析統計數據...</div>
        </div>
    `;

    try {
        const result = await callApi('getStatistics');
        if (result && result.success && result.data) {
            renderStatisticsTable(result.data);
        } else {
            container.innerHTML = '<p class="empty-text">目前尚無統計數據，請稍後再試。</p>';
        }
    } catch (e) {
        console.error('統計數據抓取失敗', e);
        container.innerHTML = '<p class="empty-text">數據載入失敗</p>';
    }
}

/**
 * 渲染統計介面
 */
function renderStatisticsTable(rawData) {
    const container = document.getElementById('statistics-grid');
    if (!container || !rawData || rawData.length <= 1) {
        container.innerHTML = '<p class="empty-text">目前尚無紀錄</p>';
        return;
    }

    // 排除標題列並過濾無效列
    const logs = rawData.slice(1).filter(row => row[0]); 
    
    // 1. 基本數據計算
    const totalVisits = logs.filter(row => row[2] === 'login').length;
    const totalActions = logs.length;
    
    // 用戶活躍度
    const userMap = {};
    logs.forEach(row => {
        const user = row[1];
        userMap[user] = (userMap[user] || 0) + 1;
    });
    const topUsers = Object.entries(userMap)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

    // 計算一週前的時間點
    const oneWeekAgo = new Date();
    oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);

    // 過濾邏輯：如果沒開啟 showAllLogs，過濾掉一週前紀錄與 switch_tab
    let displayLogs = [...logs].reverse();
    if (!state.showAllLogs) {
        displayLogs = displayLogs.filter(row => {
            const logDate = new Date(row[0]);
            const isRecent = logDate >= oneWeekAgo;
            const isNotNoise = row[2] !== 'switch_tab';
            return isRecent && isNotNoise;
        });
    }
    // 僅取最近 50 筆避免太長 (如果全量模式則可視情況調整)
    const limit = state.showAllLogs ? 100 : 50;
    displayLogs = displayLogs.slice(0, limit);

    const actionNames = {
        login: '登入系統',
        switch_tab: '切換分頁',
        check_in: '課程簽到',
        cancel_check_in: '取消簽到',
        saveData: '儲存資料'
    };

    // 2. 構建 HTML
    let html = `
        <div class="stats-container">
            <div class="stats-section-title">📊 數據總覽</div>
            <div class="stats-summary">
                <div class="stats-card">
                    <div class="stats-value">${totalVisits}</div>
                    <div class="stats-label">總登入次數</div>
                </div>
                <div class="stats-card">
                    <div class="stats-value">${totalActions}</div>
                    <div class="stats-label">總操作行為</div>
                </div>
                <div class="stats-card">
                    <div class="stats-value">${Object.keys(userMap).length}</div>
                    <div class="stats-label">不重複使用者</div>
                </div>
            </div>

            <div class="stats-section-title">👤 最活躍使用者</div>
            <div class="stats-table-wrapper">
                <table class="stats-table">
                    <thead>
                        <tr>
                            <th>使用者</th>
                            <th>操作次數</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${topUsers.map(([user, count]) => `
                            <tr>
                                <td>${user}</td>
                                <td>${count} 次</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>

            <div class="stats-section-title">
                🕒 最近活動紀錄
                <button class="btn-toggle-logs ${state.showAllLogs ? 'active' : ''}" onclick="toggleShowAllLogs()">
                    ${state.showAllLogs ? '隱藏雜訊' : '顯示全量'}
                </button>
            </div>
            <div class="stats-table-wrapper">
                <table class="stats-table">
                    <thead>
                        <tr>
                            <th style="width: 160px">時間</th>
                            <th style="width: 100px">動作</th>
                            <th>使用者與詳情</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${displayLogs.map(row => {
                            const date = formatDateTime(row[0]);
                            const action = row[2];
                            const detail = formatActivityDetail(action, row[3]);
                            const tagClass = ['login', 'switch_tab', 'check_in'].includes(action) ? action : (action === 'cancel_check_in' ? 'cancel' : 'other');
                            return `
                                <tr>
                                    <td style="font-size: 0.85rem; color: #64748b;">${date}</td>
                                    <td><span class="activity-tag ${tagClass}">${actionNames[action] || action}</span></td>
                                    <td>
                                        <div style="font-weight: 600">${row[1]}</div>
                                        <div style="font-size: 0.8rem; opacity: 0.7">${detail}</div>
                                    </td>
                                </tr>
                            `;
                        }).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    container.innerHTML = html;
}

function toggleShowAllLogs() {
    state.showAllLogs = !state.showAllLogs;
    fetchStatistics(); // 重新抓取與渲染
}

