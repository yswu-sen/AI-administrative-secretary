/**
 * MODA AI 行政秘書 - 前端互動邏輯
 * 整合 Google Sheets 資料庫
 */

// ========================================
// API 設定（從 localStorage 讀取）
// ========================================
const API_CONFIG = {
    get geminiApiKey() {
        return localStorage.getItem('gemini_api_key') || '';
    },
    set geminiApiKey(value) {
        localStorage.setItem('gemini_api_key', value);
    },
    // 可用模型：gemini-2.5-flash, gemini-2.5-flash-lite, gemini-3-flash
    geminiModel: 'gemini-2.5-flash'
};

// ========================================
// Google Sheets 設定
// ========================================
const SHEETS_CONFIG = {
    spreadsheetId: '1GzW2xIWs9wUIS0mbB37LmXiTrfrJ6rtkGrOPmI-r2cU',
    appsScriptUrl: 'https://script.google.com/macros/s/AKfycbz4OnG65YDYIPQBJyPk3S82T9cJhnsxKaQynmIT0Cq2H816rmfVI_wQ2d3F_rzA7pM8qA/exec',
    sheets: {
        meetings: '01_會議工作清單',
        categories: '02_分類設定',
        organizations: '03_單位設定',
        staff: '04_人員設定',
        todos: '05_待辦追蹤'
    }
};

// ========================================
// Gemini AI 辨識功能
// ========================================
async function analyzeFileWithGemini(file) {
    if (!API_CONFIG.geminiApiKey) {
        showToast('請先至「系統設定」設定 Gemini API Key');
        return null;
    }
    
    showToast('AI 正在辨識檔案內容...');
    
    try {
        // 將檔案轉為 base64
        const base64Data = await fileToBase64(file);
        const mimeType = file.type || 'image/png';
        
        // 呼叫 Gemini API
        const response = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/${API_CONFIG.geminiModel}:generateContent?key=${API_CONFIG.geminiApiKey}`,
            {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{
                        parts: [
                            {
                                text: `請分析這份文件/截圖，提取以下資訊並以 JSON 格式回傳：
{
    "title": "會議或工作主題",
    "date": "日期（格式：YYYY-MM-DD）",
    "time": "時間（格式：HH:MM，若無則留空）",
    "category": "分類（計畫類/預算類/租稅優惠/國際人才/AI產業人才認定指引/其他）",
    "organization": "相關單位全銜（如：國家發展委員會、勞動部勞動力發展署）",
    "assignee": "負責人或承辦人姓名",
    "dueDate": "截止日期（格式：YYYY-MM-DD，若無則留空）",
    "summary": "簡短摘要（50字內）"
}
請只回傳 JSON，不要有其他文字。若無法辨識某欄位請填空字串。`
                            },
                            {
                                inline_data: {
                                    mime_type: mimeType,
                                    data: base64Data
                                }
                            }
                        ]
                    }]
                })
            }
        );
        
        const result = await response.json();
        
        if (result.candidates && result.candidates[0]?.content?.parts?.[0]?.text) {
            const text = result.candidates[0].content.parts[0].text;
            // 提取 JSON
            const jsonMatch = text.match(/\{[\s\S]*\}/);
            if (jsonMatch) {
                const parsed = JSON.parse(jsonMatch[0]);
                showToast('辨識完成！');
                return parsed;
            }
        }
        
        showToast('辨識失敗，請手動填寫');
        return null;
        
    } catch (error) {
        console.error('Gemini API 錯誤:', error);
        showToast('辨識失敗：' + error.message);
        return null;
    }
}

function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const base64 = reader.result.split(',')[1];
            resolve(base64);
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

// ========================================
// 寫入 Google Sheets
// ========================================
async function writeToSheet(action, data) {
    const params = new URLSearchParams({ action, ...data });
    
    try {
        const response = await fetch(`${SHEETS_CONFIG.appsScriptUrl}?${params.toString()}`, {
            method: 'GET',
            mode: 'cors'
        });
        
        const result = await response.json();
        return result;
    } catch (error) {
        console.error('寫入失敗:', error);
        return { success: false, error: error.toString() };
    }
}

async function addMeetingToSheet(meetingData) {
    const result = await writeToSheet('addMeeting', {
        title: meetingData.title,
        category: meetingData.category,
        organization: meetingData.organization,
        assignee: meetingData.assignee,
        assignDate: meetingData.assignDate,
        dueDate: meetingData.dueDate,
        status: meetingData.status || '待處理',
        note: meetingData.note || ''
    });
    
    if (result.success) {
        showToast(`會議已新增：${result.meetingId}`);
        await loadAllSheetsData(); // 重新載入資料
    } else {
        showToast('新增失敗：' + result.error);
    }
    
    return result;
}

async function addTodoToSheet(todoData) {
    const result = await writeToSheet('addTodo', {
        meetingId: todoData.meetingId,
        task: todoData.task,
        assignee: todoData.assignee,
        assigner: todoData.assigner,
        dueDate: todoData.dueDate,
        priority: todoData.priority || '中',
        status: todoData.status || '待處理'
    });
    
    if (result.success) {
        showToast(`待辦已新增：${result.todoId}`);
        await loadAllSheetsData();
    } else {
        showToast('新增失敗：' + result.error);
    }
    
    return result;
}

async function updateStatusInSheet(id, status, sheetName) {
    const result = await writeToSheet('updateStatus', {
        id: id,
        status: status,
        sheet: sheetName
    });
    
    if (result.success) {
        showToast('狀態已更新');
        await loadAllSheetsData();
    }
    
    return result;
}

// 資料快取
let sheetsData = {
    meetings: [],
    categories: [],
    organizations: [],
    staff: [],
    todos: []
};

// ========================================
// 初始化
// ========================================
document.addEventListener('DOMContentLoaded', async function() {
    initializeSidebar();
    initializeUploadZones();
    initializeGeneration();
    initializeModal();
    initializeSettingsModal();
    initializeTabs();
    initializeMobileNav();
    setDefaultDate();
    
    // 載入 Google Sheets 資料
    await loadAllSheetsData();
});

// ========================================
// 設定頁面
// ========================================
function initializeSettingsModal() {
    const settingsBtn = document.getElementById('settingsBtn');
    const settingsModalOverlay = document.getElementById('settingsModalOverlay');
    const settingsModalClose = document.getElementById('settingsModalClose');
    const settingsCancel = document.getElementById('settingsCancel');
    const settingsSave = document.getElementById('settingsSave');
    const geminiApiKeyInput = document.getElementById('geminiApiKey');
    const toggleApiKey = document.getElementById('toggleApiKey');
    const apiStatus = document.getElementById('apiStatus');
    
    if (!settingsModalOverlay) return;
    
    const openSettings = () => {
        settingsModalOverlay.classList.add('active');
        document.body.style.overflow = 'hidden';
        
        // 載入已儲存的 API Key
        if (API_CONFIG.geminiApiKey) {
            geminiApiKeyInput.value = API_CONFIG.geminiApiKey;
            updateApiStatus(true);
        } else {
            geminiApiKeyInput.value = '';
            updateApiStatus(false);
        }
    };
    
    const closeSettings = () => {
        settingsModalOverlay.classList.remove('active');
        document.body.style.overflow = '';
    };
    
    function updateApiStatus(configured) {
        if (configured) {
            apiStatus.className = 'api-status configured';
            apiStatus.querySelector('.status-text').textContent = '已設定 API Key';
        } else {
            apiStatus.className = 'api-status';
            apiStatus.querySelector('.status-text').textContent = '尚未設定';
        }
    }
    
    // 設定按鈕
    if (settingsBtn) {
        settingsBtn.addEventListener('click', (e) => {
            e.preventDefault();
            openSettings();
        });
    }
    
    // 關閉按鈕
    settingsModalClose?.addEventListener('click', closeSettings);
    settingsCancel?.addEventListener('click', closeSettings);
    
    // 點擊背景關閉
    settingsModalOverlay.addEventListener('click', (e) => {
        if (e.target === settingsModalOverlay) closeSettings();
    });
    
    // 顯示/隱藏密碼
    toggleApiKey?.addEventListener('click', () => {
        if (geminiApiKeyInput.type === 'password') {
            geminiApiKeyInput.type = 'text';
        } else {
            geminiApiKeyInput.type = 'password';
        }
    });
    
    // 儲存設定
    settingsSave?.addEventListener('click', () => {
        const apiKey = geminiApiKeyInput.value.trim();
        
        if (apiKey) {
            API_CONFIG.geminiApiKey = apiKey;
            showToast('設定已儲存');
            updateApiStatus(true);
            closeSettings();
        } else {
            localStorage.removeItem('gemini_api_key');
            updateApiStatus(false);
            showToast('API Key 已清除');
        }
    });
    
    // 頁面載入時檢查 API 狀態
    if (API_CONFIG.geminiApiKey) {
        updateApiStatus(true);
    }
}

// ========================================
// Google Sheets 資料讀取
// ========================================
async function fetchSheetData(sheetName) {
    const url = `https://docs.google.com/spreadsheets/d/${SHEETS_CONFIG.spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}`;
    
    try {
        const response = await fetch(url);
        const text = await response.text();
        
        // Google 回傳的是 JSONP 格式，需要解析
        const jsonStart = text.indexOf('{');
        const jsonEnd = text.lastIndexOf('}') + 1;
        const jsonString = text.substring(jsonStart, jsonEnd);
        const data = JSON.parse(jsonString);
        
        // 轉換為陣列格式
        const rows = [];
        const cols = data.table.cols.map(col => col.label || '');
        
        if (data.table.rows) {
            data.table.rows.forEach(row => {
                const rowData = {};
                row.c.forEach((cell, idx) => {
                    rowData[cols[idx]] = cell ? (cell.v || '') : '';
                });
                rows.push(rowData);
            });
        }
        
        return rows;
    } catch (error) {
        console.error(`讀取 ${sheetName} 失敗:`, error);
        return [];
    }
}

async function loadAllSheetsData() {
    showToast('正在載入資料庫...');
    
    try {
        // 並行讀取所有工作表
        const [meetings, categories, organizations, staff, todos] = await Promise.all([
            fetchSheetData(SHEETS_CONFIG.sheets.meetings),
            fetchSheetData(SHEETS_CONFIG.sheets.categories),
            fetchSheetData(SHEETS_CONFIG.sheets.organizations),
            fetchSheetData(SHEETS_CONFIG.sheets.staff),
            fetchSheetData(SHEETS_CONFIG.sheets.todos)
        ]);
        
        sheetsData = { meetings, categories, organizations, staff, todos };
        
        console.log('已載入資料:', sheetsData);
        
        // 更新 UI
        updateDropdowns();
        updateStatsCards();
        updateKanbanBoard();
        
        showToast('資料庫載入完成');
    } catch (error) {
        console.error('載入資料失敗:', error);
        showToast('載入失敗，請檢查網路連線');
    }
}

// ========================================
// 更新下拉選單
// ========================================
function updateDropdowns() {
    // 更新工作分類下拉選單
    const categorySelect = document.getElementById('modalMeetingType');
    if (categorySelect && sheetsData.categories.length > 0) {
        categorySelect.innerHTML = '<option value="">請選擇分類</option>';
        sheetsData.categories.forEach(cat => {
            if (cat['分類名稱'] && cat['啟用'] === '是') {
                const option = document.createElement('option');
                option.value = cat['分類代碼'] || cat['分類名稱'];
                option.textContent = cat['分類名稱'];
                categorySelect.appendChild(option);
            }
        });
    }
    
    // 動態新增負責人和相關單位欄位到 Modal
    addModalFields();
}

function addModalFields() {
    const modalBody = document.querySelector('.modal-body');
    if (!modalBody || document.getElementById('modalAssignee')) return;
    
    // 新增負責人欄位
    const assigneeGroup = document.createElement('div');
    assigneeGroup.className = 'form-group';
    assigneeGroup.innerHTML = `
        <label for="modalAssignee">負責人</label>
        <select id="modalAssignee" class="form-input">
            <option value="">請選擇負責人</option>
            ${sheetsData.staff.map(s => 
                s['姓名'] && s['啟用'] === '是' 
                    ? `<option value="${s['姓名']}">${s['姓名']} - ${s['職稱'] || ''}</option>` 
                    : ''
            ).join('')}
        </select>
    `;
    
    // 新增相關單位欄位
    const orgGroup = document.createElement('div');
    orgGroup.className = 'form-group';
    orgGroup.innerHTML = `
        <label for="modalOrganization">相關單位</label>
        <select id="modalOrganization" class="form-input">
            <option value="">請選擇單位</option>
            ${sheetsData.organizations.map(o => 
                o['單位全銜'] 
                    ? `<option value="${o['單位全銜']}">${o['單位全銜']}</option>` 
                    : ''
            ).join('')}
        </select>
    `;
    
    // 插入到會議類型後面
    const typeGroup = document.getElementById('modalMeetingType').parentElement;
    typeGroup.after(assigneeGroup, orgGroup);
}

// ========================================
// 更新統計卡片
// ========================================
function updateStatsCards() {
    const meetings = sheetsData.meetings;
    const todos = sheetsData.todos;
    
    // 待處理會議
    const pendingCount = meetings.filter(m => m['狀態'] === '待處理' || m['狀態'] === '進行中').length;
    const pendingEl = document.querySelector('.stat-card:nth-child(1) .stat-value');
    if (pendingEl) pendingEl.textContent = pendingCount || 0;
    
    // 本月公文（會議數）
    const thisMonth = new Date().getMonth();
    const monthlyCount = meetings.filter(m => {
        if (!m['指派日期']) return false;
        const date = new Date(m['指派日期']);
        return date.getMonth() === thisMonth;
    }).length;
    const monthlyEl = document.querySelector('.stat-card:nth-child(2) .stat-value');
    if (monthlyEl) monthlyEl.textContent = monthlyCount || meetings.length;
    
    // 已完成決議
    const completedCount = meetings.filter(m => m['狀態'] === '已完成').length;
    const completedEl = document.querySelector('.stat-card:nth-child(3) .stat-value');
    if (completedEl) completedEl.textContent = completedCount || 0;
    
    // 逾期待辦
    const today = new Date();
    const overdueCount = todos.filter(t => {
        if (t['狀態'] === '已完成' || !t['截止日期']) return false;
        return new Date(t['截止日期']) < today;
    }).length;
    const overdueEl = document.querySelector('.stat-card:nth-child(4) .stat-value');
    if (overdueEl) overdueEl.textContent = overdueCount || 0;
    
    // 初始化統計卡片點擊事件
    initStatCardClicks();
}

// ========================================
// 統計卡片點擊展開詳情
// ========================================
function initStatCardClicks() {
    const statCards = document.querySelectorAll('.stat-card.clickable');
    const detailPanel = document.getElementById('statDetailPanel');
    const detailTitle = document.getElementById('detailTitle');
    const detailList = document.getElementById('detailList');
    const detailClose = document.getElementById('detailClose');
    
    if (!detailPanel) return;
    
    statCards.forEach(card => {
        card.addEventListener('click', () => {
            const type = card.dataset.type;
            showStatDetail(type);
        });
    });
    
    detailClose?.addEventListener('click', () => {
        detailPanel.style.display = 'none';
    });
}

function showStatDetail(type) {
    const detailPanel = document.getElementById('statDetailPanel');
    const detailTitle = document.getElementById('detailTitle');
    const detailList = document.getElementById('detailList');
    
    const meetings = sheetsData.meetings;
    const todos = sheetsData.todos;
    const today = new Date();
    
    let items = [];
    let title = '';
    
    switch(type) {
        case 'pending':
            title = '待處理會議';
            items = meetings.filter(m => m['狀態'] === '待處理' || m['狀態'] === '進行中');
            break;
        case 'monthly':
            title = '本月公文';
            const thisMonth = today.getMonth();
            items = meetings.filter(m => {
                if (!m['指派日期']) return true;
                const date = new Date(m['指派日期']);
                return date.getMonth() === thisMonth;
            });
            break;
        case 'completed':
            title = '已完成決議';
            items = meetings.filter(m => m['狀態'] === '已完成');
            break;
        case 'overdue':
            title = '逾期待辦';
            items = [...todos, ...meetings].filter(item => {
                const status = item['狀態'];
                const dueDate = item['截止日期'];
                if (status === '已完成' || !dueDate) return false;
                return new Date(dueDate) < today;
            });
            break;
    }
    
    detailTitle.textContent = title;
    
    if (items.length === 0) {
        detailList.innerHTML = '<div class="detail-empty">暫無資料</div>';
    } else {
        detailList.innerHTML = items.map(item => {
            const itemTitle = item['主題'] || item['待辦事項'] || '未命名';
            const assignee = item['負責人'] || '未指派';
            const dueDate = item['截止日期'] || '';
            const status = item['狀態'] || '待處理';
            const category = item['工作分類'] || '';
            
            let statusClass = 'pending';
            if (status === '已完成') statusClass = 'completed';
            if (dueDate && new Date(dueDate) < today && status !== '已完成') statusClass = 'overdue';
            
            return `
                <div class="detail-item">
                    <div class="detail-item-info">
                        <div class="detail-item-title">${itemTitle}</div>
                        <div class="detail-item-meta">
                            ${category ? category + ' · ' : ''}${assignee}${dueDate ? ' · 截止: ' + dueDate : ''}
                        </div>
                    </div>
                    <span class="detail-item-status ${statusClass}">${status}</span>
                </div>
            `;
        }).join('');
    }
    
    detailPanel.style.display = 'block';
    detailPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// ========================================
// 更新看板
// ========================================
function updateKanbanBoard() {
    const todos = sheetsData.todos;
    const meetings = sheetsData.meetings;
    
    // 合併待辦和會議到看板
    const allTasks = [
        ...todos.map(t => ({
            id: t['待辦編號'],
            title: t['待辦事項'],
            desc: t['關聯會議編號'] || '',
            assignee: t['負責人'],
            date: t['截止日期'],
            status: t['狀態'],
            priority: t['優先級']
        })),
        ...meetings.map(m => ({
            id: m['編號'],
            title: m['主題'],
            desc: m['工作分類'] || '',
            assignee: m['負責人'],
            date: m['截止日期'],
            status: m['狀態'],
            priority: ''
        }))
    ].filter(t => t.title);
    
    // 分類任務
    const pending = allTasks.filter(t => t.status === '待處理' || t.status === '進行中');
    const review = allTasks.filter(t => t.status === '待核定');
    const completed = allTasks.filter(t => t.status === '已完成');
    
    // 更新各欄
    updateKanbanColumn('pending', pending);
    updateKanbanColumn('review', review);
    updateKanbanColumn('completed', completed);
}

function updateKanbanColumn(status, tasks) {
    const column = document.querySelector(`.kanban-column[data-status="${status}"]`);
    if (!column) return;
    
    const cardsContainer = column.querySelector('.column-cards');
    const countEl = column.querySelector('.column-count');
    
    if (countEl) countEl.textContent = tasks.length;
    
    if (tasks.length === 0) {
        cardsContainer.innerHTML = '<div class="empty-column">暫無項目</div>';
        return;
    }
    
    cardsContainer.innerHTML = tasks.slice(0, 5).map(task => `
        <div class="task-card ${task.status === '已完成' ? 'completed' : ''}">
            <div class="card-header">
                <span class="card-tag ${task.priority === '高' ? 'urgent' : task.status === '已完成' ? 'completed' : 'normal'}">
                    ${task.priority === '高' ? '急件' : task.status === '已完成' ? '完成' : '一般'}
                </span>
                <span class="card-date">${formatDate(task.date)}</span>
            </div>
            <h4 class="card-title">${task.title}</h4>
            <p class="card-desc">${task.desc}</p>
            <div class="card-footer">
                <span class="card-assignee">${task.assignee || '未指派'}</span>
            </div>
        </div>
    `).join('');
}

function formatDate(dateStr) {
    if (!dateStr) return '';
    const date = new Date(dateStr);
    if (isNaN(date)) return dateStr;
    return `${date.getMonth() + 1}/${date.getDate()}`;
}

// ========================================
// 側邊欄控制
// ========================================
function initializeSidebar() {
    const sidebar = document.getElementById('sidebar');
    const sidebarToggle = document.getElementById('sidebarToggle');
    const menuBtn = document.getElementById('menuBtn');

    if (sidebarToggle) {
        sidebarToggle.addEventListener('click', () => {
            sidebar.classList.toggle('collapsed');
        });
    }

    if (menuBtn) {
        menuBtn.addEventListener('click', () => {
            sidebar.classList.toggle('open');
        });
    }

    document.addEventListener('click', (e) => {
        if (window.innerWidth <= 768) {
            if (!sidebar.contains(e.target) && !menuBtn.contains(e.target)) {
                sidebar.classList.remove('open');
            }
        }
    });

    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            navItems.forEach(i => i.classList.remove('active'));
            item.classList.add('active');
            
            if (window.innerWidth <= 768) {
                sidebar.classList.remove('open');
            }
        });
    });
}

// ========================================
// 檔案上傳功能
// ========================================
function initializeUploadZones() {
    const uploadZones = document.querySelectorAll('.upload-zone');
    
    uploadZones.forEach(zone => {
        const type = zone.dataset.type;
        const input = document.getElementById(`${type}Input`);
        const status = document.getElementById(`${type}Status`);
        
        zone.addEventListener('click', () => input.click());
        
        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            zone.classList.add('dragover');
        });
        
        zone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');
        });
        
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                handleFileUpload(zone, e.dataTransfer.files, status);
            }
        });
        
        input.addEventListener('change', () => {
            if (input.files.length > 0) {
                handleFileUpload(zone, input.files, status);
            }
        });
    });
}

function handleFileUpload(zone, files, statusEl) {
    const fileNames = Array.from(files).map(f => f.name).join(', ');
    zone.classList.add('uploaded');
    statusEl.textContent = `✓ ${files.length > 1 ? files.length + ' 個檔案' : fileNames}`;
    statusEl.classList.add('success');
    updateGenerateButtonState();
}

function updateGenerateButtonState() {
    const generateBtn = document.getElementById('generateBtn');
    const whisperUploaded = document.getElementById('whisperZone').classList.contains('uploaded');
    const titleFilled = document.getElementById('meetingTitle').value.trim() !== '';
    
    if (whisperUploaded || titleFilled) {
        generateBtn.disabled = false;
    }
}

// ========================================
// AI 生成功能
// ========================================
function initializeGeneration() {
    const generateBtn = document.getElementById('generateBtn');
    const progressSection = document.getElementById('generationProgress');
    const outputPreview = document.getElementById('outputPreview');
    
    generateBtn.addEventListener('click', async () => {
        progressSection.style.display = 'block';
        outputPreview.style.display = 'none';
        generateBtn.disabled = true;
        generateBtn.innerHTML = `
            <svg class="btn-icon spinning" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
            </svg>
            <span>處理中...</span>
        `;
        
        await simulateGeneration();
        
        progressSection.style.display = 'none';
        outputPreview.style.display = 'block';
        generateBtn.disabled = false;
        generateBtn.innerHTML = `
            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>
            </svg>
            <span>重新生成</span>
        `;
        
        populateSampleResults();
    });
}

async function simulateGeneration() {
    const progressFill = document.getElementById('progressFill');
    const progressPercent = document.querySelector('.progress-percent');
    const progressText = document.querySelector('.progress-text');
    const steps = document.querySelectorAll('.progress-step');
    
    const stages = [
        { step: 1, percent: 25, text: '正在解析上傳資料...' },
        { step: 2, percent: 50, text: '交叉比對校正中...' },
        { step: 3, percent: 75, text: '生成文件內容...' },
        { step: 4, percent: 100, text: '格式校正完成！' }
    ];
    
    for (let i = 0; i < stages.length; i++) {
        const stage = stages[i];
        progressFill.style.width = stage.percent + '%';
        progressPercent.textContent = stage.percent + '%';
        progressText.textContent = stage.text;
        
        steps.forEach((step, idx) => {
            step.classList.remove('active', 'completed');
            if (idx + 1 < stage.step) step.classList.add('completed');
            else if (idx + 1 === stage.step) step.classList.add('active');
        });
        
        await delay(800);
    }
    
    steps[3].classList.remove('active');
    steps[3].classList.add('completed');
}

function populateSampleResults() {
    const meetingTitle = document.getElementById('meetingTitle').value || '113年度數位產業推動策略會議';
    const meetingDate = document.getElementById('meetingDate').value || '2024/12/31';
    
    const summaryText = document.getElementById('summaryText');
    summaryText.innerHTML = `
<h3 style="margin-bottom: 16px; color: var(--primary);">會議摘要</h3>

<p><strong>會議名稱：</strong>${meetingTitle}</p>
<p><strong>會議時間：</strong>${meetingDate}</p>
<p><strong>會議地點：</strong>本署 8 樓會議室</p>
<p><strong>主持人：</strong>○○○署長</p>
<p><strong>出席人員：</strong>各組科室主管、相關業務承辦人</p>

<h4 style="margin: 20px 0 12px; color: var(--primary);">關鍵決策：</h4>
<ul style="list-style: disc; padding-left: 24px; line-height: 2;">
    <li>114年度 AI 產業輔導計畫預算核定通過，總額新台幣 3.2 億元</li>
    <li>數位產業資料平台第二階段開發案，由○○科負責規格確認，預計 1/15 前完成</li>
    <li>跨部會協調會議訂於 1/10 召開，由○○組長代表出席</li>
</ul>

<h4 style="margin: 20px 0 12px; color: var(--primary);">待辦事項：</h4>
<table style="width: 100%; border-collapse: collapse; margin-top: 8px;">
    <tr style="background: var(--bg-tertiary);">
        <th style="padding: 10px; text-align: left; border: 1px solid var(--border-color);">負責人</th>
        <th style="padding: 10px; text-align: left; border: 1px solid var(--border-color);">工作項目</th>
        <th style="padding: 10px; text-align: left; border: 1px solid var(--border-color);">期限</th>
    </tr>
    ${sheetsData.staff.slice(0, 3).map((s, i) => `
    <tr>
        <td style="padding: 10px; border: 1px solid var(--border-color);">${s['姓名'] || '待指派'}</td>
        <td style="padding: 10px; border: 1px solid var(--border-color);">待辦事項 ${i + 1}</td>
        <td style="padding: 10px; border: 1px solid var(--border-color);">1/${10 + i * 5}</td>
    </tr>
    `).join('')}
</table>
    `;
}

// ========================================
// Modal 彈窗
// ========================================
function initializeModal() {
    const newMeetingBtn = document.getElementById('newMeetingBtn');
    const mobileFab = document.getElementById('mobileFab');
    const modalOverlay = document.getElementById('modalOverlay');
    const modalClose = document.getElementById('modalClose');
    const modalCancel = document.getElementById('modalCancel');
    const modalConfirm = document.getElementById('modalConfirm');
    
    // AI 上傳區元素
    const aiUploadZone = document.getElementById('aiUploadZone');
    const aiFileInput = document.getElementById('aiFileInput');
    const aiUploadStatus = document.getElementById('aiUploadStatus');
    
    const openModal = () => {
        modalOverlay.classList.add('active');
        document.body.style.overflow = 'hidden';
        // 重置 AI 上傳狀態
        if (aiUploadStatus) aiUploadStatus.textContent = '';
        if (aiUploadZone) aiUploadZone.classList.remove('processing');
    };
    
    const closeModal = () => {
        modalOverlay.classList.remove('active');
        document.body.style.overflow = '';
    };
    
    // AI 上傳區事件
    if (aiUploadZone && aiFileInput) {
        aiUploadZone.addEventListener('click', () => aiFileInput.click());
        
        aiUploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            aiUploadZone.classList.add('dragover');
        });
        
        aiUploadZone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            aiUploadZone.classList.remove('dragover');
        });
        
        aiUploadZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            aiUploadZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                await handleAiFileUpload(e.dataTransfer.files[0]);
            }
        });
        
        aiFileInput.addEventListener('change', async () => {
            if (aiFileInput.files.length > 0) {
                await handleAiFileUpload(aiFileInput.files[0]);
            }
        });
    }
    
    async function handleAiFileUpload(file) {
        aiUploadZone.classList.add('processing');
        aiUploadStatus.textContent = 'AI 正在辨識中...';
        aiUploadStatus.className = 'ai-upload-status processing';
        
        const result = await analyzeFileWithGemini(file);
        
        aiUploadZone.classList.remove('processing');
        
        if (result) {
            // 自動填入表單
            if (result.title) {
                document.getElementById('modalMeetingTitle').value = result.title;
            }
            if (result.date) {
                document.getElementById('modalMeetingDate').value = result.date;
            }
            if (result.time) {
                document.getElementById('modalMeetingTime').value = result.time;
            }
            if (result.category) {
                const typeSelect = document.getElementById('modalMeetingType');
                const option = Array.from(typeSelect.options).find(o => 
                    o.value === result.category || o.text === result.category
                );
                if (option) typeSelect.value = option.value;
            }
            if (result.dueDate) {
                const dueDateInput = document.getElementById('modalDueDate');
                if (dueDateInput) dueDateInput.value = result.dueDate;
            }
            if (result.organization) {
                const orgSelect = document.getElementById('modalOrganization');
                if (orgSelect) {
                    const option = Array.from(orgSelect.options).find(o => 
                        o.value.includes(result.organization) || result.organization.includes(o.value)
                    );
                    if (option) orgSelect.value = option.value;
                }
            }
            if (result.assignee) {
                const assigneeSelect = document.getElementById('modalAssignee');
                if (assigneeSelect) {
                    const option = Array.from(assigneeSelect.options).find(o => 
                        o.value.includes(result.assignee)
                    );
                    if (option) assigneeSelect.value = option.value;
                }
            }
            
            aiUploadStatus.textContent = '✓ 辨識完成，請確認資訊';
            aiUploadStatus.className = 'ai-upload-status success';
        } else {
            aiUploadStatus.textContent = '辨識失敗，請手動填寫';
            aiUploadStatus.className = 'ai-upload-status error';
        }
    }
    
    newMeetingBtn.addEventListener('click', openModal);
    if (mobileFab) mobileFab.addEventListener('click', (e) => {
        e.preventDefault();
        openModal();
    });
    modalClose.addEventListener('click', closeModal);
    modalCancel.addEventListener('click', closeModal);
    
    modalOverlay.addEventListener('click', (e) => {
        if (e.target === modalOverlay) closeModal();
    });
    
    modalConfirm.addEventListener('click', async () => {
        const title = document.getElementById('modalMeetingTitle').value;
        const date = document.getElementById('modalMeetingDate').value;
        const category = document.getElementById('modalMeetingType')?.value || '';
        const assignee = document.getElementById('modalAssignee')?.value || '';
        const organization = document.getElementById('modalOrganization')?.value || '';
        
        if (title && date) {
            // 填入主表單
            document.getElementById('meetingTitle').value = title;
            document.getElementById('meetingDate').value = date;
            document.getElementById('meetingTime').value = document.getElementById('modalMeetingTime').value;
            
            closeModal();
            
            // 寫入 Google Sheets
            showToast('正在儲存至資料庫...');
            await addMeetingToSheet({
                title: title,
                category: category,
                organization: organization,
                assignee: assignee,
                assignDate: date,
                dueDate: '',
                status: '待處理'
            });
            
            document.querySelector('.upload-section').scrollIntoView({ behavior: 'smooth' });
            updateGenerateButtonState();
        } else {
            alert('請填寫會議主題與日期');
        }
    });
}

// ========================================
// 預覽 Tabs
// ========================================
function initializeTabs() {
    const tabs = document.querySelectorAll('.preview-tab');
    
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
        });
    });
    
    document.querySelectorAll('.toolbar-btn').forEach(btn => {
        btn.addEventListener('click', async () => {
            const action = btn.getAttribute('title');
            
            if (action === '複製') {
                const text = document.getElementById('summaryText').innerText;
                try {
                    await navigator.clipboard.writeText(text);
                    showToast('已複製到剪貼簿');
                } catch (err) {
                    console.error('複製失敗:', err);
                }
            } else if (action === '下載 Word') {
                showToast('Word 檔案準備中...');
            } else if (action === '編輯') {
                showToast('進入編輯模式');
            }
        });
    });
}

// ========================================
// 手機版底部導覽
// ========================================
function initializeMobileNav() {
    const mobileNavItems = document.querySelectorAll('.mobile-nav-item:not(.fab)');
    
    mobileNavItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            mobileNavItems.forEach(i => i.classList.remove('active'));
            item.classList.add('active');
            
            const page = item.dataset.page;
            document.querySelectorAll('.nav-item').forEach(navItem => {
                navItem.classList.toggle('active', navItem.dataset.page === page);
            });
        });
    });
}

// ========================================
// 工具函數
// ========================================
function setDefaultDate() {
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('meetingDate').value = today;
    document.getElementById('modalMeetingDate').value = today;
}

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function showToast(message) {
    const toast = document.createElement('div');
    toast.style.cssText = `
        position: fixed;
        bottom: 100px;
        left: 50%;
        transform: translateX(-50%);
        background: #2D3748;
        color: white;
        padding: 12px 24px;
        border-radius: 8px;
        font-size: 14px;
        z-index: 1000;
        animation: fadeIn 0.3s ease;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    `;
    toast.textContent = message;
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transition = 'opacity 0.3s ease';
        setTimeout(() => toast.remove(), 300);
    }, 2000);
}

// 重新整理資料按鈕
function refreshData() {
    loadAllSheetsData();
}

// CSS 動畫
const spinStyle = document.createElement('style');
spinStyle.textContent = `
    .spinning { animation: spin 1s linear infinite; }
    @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
    .empty-column { text-align: center; color: #adb5bd; padding: 20px; font-size: 14px; }
`;
document.head.appendChild(spinStyle);
