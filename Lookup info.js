// ==========================================
// 1. Data Definitions
// ==========================================
const machineFiles = [
    "1.塗佈機-Coater -BCB F4ACS227.pdf",
    "12~13.乾式蝕刻機-Dry Etch-1.pdf",
    "12~13.乾式蝕刻機-Dry Etch-2.pdf",
    "15~23.晶圓傳送工具-GLA Wafer Transfer Tool & Comet wafer Transfer Tool.pdf",
    "23~29.O2,CF4電漿清潔機-Plasma O2 & Plasma CF4.pdf",
    "2~3.顯影機-Developer - Functional F4ACS641 & F4ACS829.pdf",
    "30~31.氧氣烘箱-Oven O2.pdf",
    "32~34.氮氣烘箱_Oven N2.pdf",
    "35.電子束蒸鍍機-BAK (UBM).pdf",
    "36.電子束蒸鍍機-BAK (UBM).pdf",
    "37.電子束蒸鍍機-BAK(SMD).pdf",
    "38.電子束蒸鍍機-BAK(SMD).pdf",
    "39.X射線度層厚度量測儀-Fischerscoope.pdf",
    "4.曝光機-Mask Aligner F1MAE331.pdf",
    "40.電漿清潔機-Plasma H2.pdf",
    "41.化學式氣相蒸鍍機-PECVD.pdf",
    "42.背面濕蝕刻機-Reverse Side Etch.pdf",
    "43.自動化濕製程設備-SAT.pdf",
    "44.單晶圓多功能濕式處理機-SSEC.pdf",
    "45.單晶圓多功能濕式處理機-SSEC.pdf",
    "46.單晶圓多功能濕式處理機-SSEC.pdf",
    "47~50.-NF3 Trimmer-1.pdf",
    "47~50.-NF3 Trimmer-2.pdf",
    "5.曝光機-Mask Aligner F1MAE723.pdf",
    "51.白光3D表面輪廓儀-M_WLI.pdf",
    "6.曝光機-Mask Aligner F1MAE724.pdf",
    "7.曝光機-Mask Aligner F1MAE725.pdf",
    "8.曝光機-Mask Aligner F1MAE726.pdf",
    "9~11.直立式烘箱-Oven 3TPA F4TPA916,F4TPA917,F4TPA918.pdf"
];

const checkItems = [
    { category: "清點", item: "設備清單核對", target: "逐一核對廠商、型號、Tool ID 是否與清單一致", type: "checkbox+note" },
    { category: "紀錄", item: "機台尺寸", target: "機台主體或拆解後之各部件尺寸", type: "dimensions_multi" },
    { category: "檢查", item: "損傷檢查", target: "檢查有無腐蝕、生鏽、化學噴濺痕跡", type: "radio-yesno" },
    { category: "評估", item: "安裝", target: "評估是否可由公司自行復機", type: "radio-eval" }
];

const docItems = [
    { type: "手冊", nameEn: "Equipment Manual", nameTw: "設備操作說明書" },
    { type: "手冊", nameEn: "Service Manual", nameTw: "維修技術手冊" },
    { type: "圖面", nameEn: "Electrical Schematic", nameTw: "電氣電路圖" },
    { type: "圖面", nameEn: "Pneumatic Diagram", nameTw: "氣壓原理圖" },
    { type: "圖面", nameEn: "Gas Diagram", nameTw: "氣體管路圖" }
];

// ==========================================
// 2. Global State & Auth & Cloud Storage
// ==========================================
let isAdmin = false;
let currentMachine = null;

// Cloud Storage Utilities
const cloudStorage = {
    async get(key) {
        try {
            const res = await fetch(`/api/get-data?key=${encodeURIComponent(key)}`);
            if (!res.ok) throw new Error('Network response was not ok');
            const data = await res.json();
            return data.value;
        } catch (error) {
            console.error('Error fetching data from cloud:', error);
            // Fallback to local storage for offline support/testing
            const localData = localStorage.getItem(key);
            try { return JSON.parse(localData); } catch (e) { return localData; }
        }
    },
    async set(key, value) {
        try {
            const res = await fetch('/api/save-data', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ key, value })
            });
            if (!res.ok) throw new Error('Network response was not ok');
        } catch (error) {
            console.error('Error saving data to cloud:', error);
            // Fallback to local storage
            const dataToSave = typeof value === 'object' ? JSON.stringify(value) : value;
            try {
                localStorage.setItem(key, dataToSave);
            } catch (e) {
                console.error("Local storage error:", e);
            }
        }
    }
};

// ==========================================
// 3. Initialization
// ==========================================
document.addEventListener('DOMContentLoaded', async () => {
    initAuth();
    await initMachineList();
    initTabs();
    initSPA();
    await initDailyReport();
    initExcelViewer();
});

// ==========================================
// 4. SPA Logic
// ==========================================
function initSPA() {
    const navItems = document.querySelectorAll('.nav-item');
    const views = document.querySelectorAll('.view-section');

    navItems.forEach(item => {
        item.addEventListener('click', () => {
            navItems.forEach(nav => nav.classList.remove('active'));
            item.classList.add('active');

            const targetId = item.getAttribute('data-target');
            views.forEach(view => {
                if (view.id === targetId) {
                    view.classList.add('active');
                    view.classList.remove('hidden');
                } else {
                    view.classList.remove('active');
                    view.classList.add('hidden');
                }
            });
        });
    });
}

// ==========================================
// 5. Auth Logic
// ==========================================
function initAuth() {
    const loginModal = document.getElementById('loginModal');
    const loginBtn = document.getElementById('loginBtn');
    const logoutBtn = document.getElementById('logoutBtn');
    const closeModalBtn = document.getElementById('closeModalBtn');
    const cancelLoginBtn = document.getElementById('cancelLoginBtn');
    const submitLoginBtn = document.getElementById('submitLoginBtn');
    const adminIdInput = document.getElementById('adminId');
    const adminPwdInput = document.getElementById('adminPwd');
    const loginError = document.getElementById('loginError');
    const globalUserStatus = document.getElementById('globalUserStatus');
    const userStatus = document.getElementById('userStatus');

    // Load auth state
    const savedAuth = localStorage.getItem('isAdmin');
    if (savedAuth === 'true') {
        setAdminState(true);
    }

    loginBtn.addEventListener('click', () => {
        loginModal.classList.remove('hidden');
        adminIdInput.value = '';
        adminPwdInput.value = '';
        loginError.classList.add('hidden');
        adminIdInput.focus();
    });

    const closeModal = () => {
        loginModal.classList.add('hidden');
    };

    closeModalBtn.addEventListener('click', closeModal);
    cancelLoginBtn.addEventListener('click', closeModal);

    submitLoginBtn.addEventListener('click', handleLogin);
    adminPwdInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleLogin();
    });

    logoutBtn.addEventListener('click', () => {
        setAdminState(false);
        localStorage.removeItem('isAdmin');
    });

    function handleLogin() {
        const id = adminIdInput.value.trim();
        const pwd = adminPwdInput.value;

        if (id === 'Polaris' && pwd === 'Polaris') {
            setAdminState(true);
            localStorage.setItem('isAdmin', 'true');
            closeModal();
        } else {
            loginError.classList.remove('hidden');
        }
    }

    function setAdminState(state) {
        isAdmin = state;
        if (isAdmin) {
            loginBtn.classList.add('hidden');
            logoutBtn.classList.remove('hidden');
            userStatus.textContent = '管理員模式';
            userStatus.className = 'status-badge admin';
            globalUserStatus.textContent = '👑';
            globalUserStatus.className = 'status-badge admin';
            globalUserStatus.title = "目前為管理員模式";
        } else {
            loginBtn.classList.remove('hidden');
            logoutBtn.classList.add('hidden');
            userStatus.textContent = '檢視模式';
            userStatus.className = 'status-badge viewer';
            globalUserStatus.textContent = '👁️';
            globalUserStatus.className = 'status-badge viewer';
            globalUserStatus.title = "目前為檢視模式";
        }

        // Refresh Current View if machine selected
        if (currentMachine) {
            loadMachineData(currentMachine);
        }

        // Toggle Daily Report form features
        toggleDailyReportAuth();
    }
}

function toggleDailyReportAuth() {
    const reportForm = document.getElementById('dailyReportForm');
    if (reportForm) {
        const inputs = reportForm.querySelectorAll('input, textarea, button[type="submit"], #clearFormBtn');
        const uploadWrapper = document.getElementById('uploadBtnWrapper');

        if (isAdmin) {
            inputs.forEach(el => el.disabled = false);
            if (uploadWrapper) uploadWrapper.style.display = 'flex';
        } else {
            inputs.forEach(el => el.disabled = true);
            if (uploadWrapper) uploadWrapper.style.display = 'none';
        }
    }
}

// ==========================================
// 6. Machine List & Search
// ==========================================
async function initMachineList() {
    const searchInput = document.getElementById('searchInput');
    const machineListElement = document.getElementById('machineList');

    async function renderList(files) {
        machineListElement.innerHTML = '<div style="padding: 24px; text-align: center; color: var(--text-secondary);">資料載入中...</div>';

        if (files.length === 0) {
            machineListElement.innerHTML = '<div style="padding: 24px 12px; color: var(--text-secondary); text-align: center; font-size: 0.9rem;">查無符合的機台檔案</div>';
            return;
        }

        // Fetch all status lights at once or sequentially
        const renderedItems = [];
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const displayName = file.replace('.pdf', '');
            const hasData = await window.hasMachineDataUploaded(displayName);
            let savedLight = await cloudStorage.get(`statusLight_${displayName}`);
            if (!savedLight) savedLight = 'red'; // Default

            const itemHtml = `
                <label class="machine-item ${hasData ? 'has-data' : ''}">
                    <div class="checkbox-wrapper">
                        <input type="radio" name="machineSelect" value="${file}" id="machine_${i}">
                    </div>
                    <div class="item-content">
                        <div class="item-title" title="${displayName}">${displayName}</div>
                    </div>
                    <div class="traffic-light-container">
                        <div class="traffic-light red ${savedLight === 'red' ? 'active' : ''}" data-color="red" data-machine="${displayName}" title="尚未紀錄"></div>
                        <div class="traffic-light yellow ${savedLight === 'yellow' ? 'active' : ''}" data-color="yellow" data-machine="${displayName}" title="部分紀錄"></div>
                        <div class="traffic-light green ${savedLight === 'green' ? 'active' : ''}" data-color="green" data-machine="${displayName}" title="完整登錄"></div>
                    </div>
                </label>
            `;
            renderedItems.push({ html: itemHtml, file, displayName });
        }

        machineListElement.innerHTML = renderedItems.map(item => item.html).join('');

        // Attach events
        const items = machineListElement.querySelectorAll('.machine-item');
        items.forEach((item, index) => {
            const { file, displayName } = renderedItems[index];

            item.querySelector('input').addEventListener('change', (e) => {
                document.querySelectorAll('.machine-item').forEach(el => el.classList.remove('active'));
                if (e.target.checked) {
                    item.classList.add('active');
                    showMachineDetail(displayName, file);
                }
            });

            const lights = item.querySelectorAll('.traffic-light');
            lights.forEach(light => {
                light.addEventListener('click', async (e) => {
                    e.preventDefault();
                    if (!isAdmin) {
                        alert("僅管理員可更改機台燈號狀態。");
                        return;
                    }
                    const color = light.getAttribute('data-color');
                    const machine = light.getAttribute('data-machine');

                    lights.forEach(l => l.classList.remove('active'));
                    light.classList.add('active');

                    light.style.opacity = '0.5'; // loading indication
                    await cloudStorage.set(`statusLight_${machine}`, color);
                    light.style.opacity = '1';
                });
            });
        });
    }

    await renderList(machineFiles);

    searchInput.addEventListener('input', async (e) => {
        const term = e.target.value.toLowerCase();
        const filtered = machineFiles.filter(f => f.toLowerCase().includes(term));
        await renderList(filtered);
    });
}

function showMachineDetail(name, filename) {
    const placeholderState = document.getElementById('placeholderState');
    const contentState = document.getElementById('contentState');
    const machineTitle = document.getElementById('machineTitle');
    const pdfContainer = document.getElementById('pdfContainer');

    machineTitle.textContent = name;
    placeholderState.classList.add('hidden');
    contentState.classList.remove('hidden');

    const pdfPath = `機台相片/${filename}`;
    pdfContainer.innerHTML = `<iframe src="${pdfPath}#toolbar=0&navpanes=0&scrollbar=0" title="${name}"></iframe>`;

    currentMachine = name;
    loadMachineData(currentMachine);
}

// ==========================================
// 7. Tabs & Forms Logic
// ==========================================
function initTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabPanes = document.querySelectorAll('.tab-pane');

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            tabBtns.forEach(b => b.classList.remove('active'));
            tabPanes.forEach(p => p.classList.remove('active'));

            btn.classList.add('active');
            const target = btn.getAttribute('data-tab');
            document.getElementById(target).classList.add('active');
        });
    });
}

async function loadMachineData(machineName) {
    const recordsTbody = document.getElementById('recordsTbody');
    const docsTbody = document.getElementById('docsTbody');

    recordsTbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding: 20px;">資料載入中...</td></tr>';
    docsTbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding: 20px;">資料載入中...</td></tr>';

    // Process Records Table
    let savedRecords = await cloudStorage.get(`records_${machineName}`);
    if (typeof savedRecords === 'string') {
        try { savedRecords = JSON.parse(savedRecords); } catch (e) { savedRecords = {}; }
    }
    savedRecords = savedRecords || {};
    recordsTbody.innerHTML = '';

    checkItems.forEach((item, index) => {
        const tr = document.createElement('tr');
        const stateKey = `state_${index}`;
        const photoKey = `photo_${index}`;
        const savedState = savedRecords[stateKey] || '';
        const savedPhoto = savedRecords[photoKey] || null;

        let photoHtml = '';
        if (item.type === 'dimensions_multi') {
            photoHtml = '<span style="color:var(--text-secondary);font-size:0.85rem;">見左側各部件</span>';
        } else if (savedPhoto) {
            photoHtml = `
                <div class="photo-preview-mini" style="background-image: url('${savedPhoto}')">
                    <img src="${savedPhoto}" alt="preview">
                    ${isAdmin ? `<button type="button" class="del-photo-btn" data-index="${index}">&times;</button>` : ''}
                </div>
            `;
        } else {
            if (isAdmin) {
                photoHtml = `
                    <label class="btn-upload">
                        📁 選擇照片
                        <input type="file" accept="image/*" class="photo-input" data-index="${index}" style="display:none;">
                    </label>
                `;
            } else {
                photoHtml = `<span style="color:var(--text-secondary);font-size:0.8rem;">無照片</span>`;
            }
        }

        let stateHtml = '';
        if (item.type === 'checkbox+note') {
            const isChecked = savedRecords[`check_${index}`] || false;
            const noteText = savedRecords[`note_${index}`] || '';
            stateHtml = `
                <div style="display:flex; flex-direction:column; gap:6px;">
                    <label style="display:flex; align-items:center; gap:6px; font-weight: 500;">
                        <input type="checkbox" class="status-check" data-index="${index}" ${isChecked ? 'checked' : ''} ${!isAdmin ? 'disabled' : ''} style="width: 16px; height: 16px; cursor: pointer;">
                        <span>完成 (V)</span>
                    </label>
                    <input type="text" class="status-note w-full" data-index="${index}" value="${noteText}" placeholder="備註..." ${!isAdmin ? 'disabled' : ''}>
                </div>
            `;
        } else if (item.type === 'checkbox') {
            const isChecked = savedRecords[`check_${index}`] || false;
            const labelText = '完成 (V)';
            stateHtml = `
                <div style="display:flex; flex-direction:column; gap:6px;">
                    <label style="display:flex; align-items:center; gap:6px; font-weight: 500; cursor: pointer;">
                        <input type="checkbox" class="status-check" data-index="${index}" ${isChecked ? 'checked' : ''} ${!isAdmin ? 'disabled' : ''} style="width: 16px; height: 16px; accent-color: var(--primary-color);">
                        <span>${labelText}</span>
                    </label>
                </div>
            `;
        } else if (item.type === 'radio-yesno') {
            const savedYesNo = savedRecords[`yesno_${index}`] || '';
            stateHtml = `
                <div style="display:flex; flex-direction:column; gap:8px;">
                    <label style="display:flex; align-items:center; gap:6px; font-weight: 500; cursor: pointer;">
                        <input type="radio" name="yesno_${index}" class="status-yesno" data-index="${index}" value="是 (V)" ${savedYesNo === '是 (V)' ? 'checked' : ''} ${!isAdmin ? 'disabled' : ''} style="width: 16px; height: 16px; accent-color: var(--primary-color);">
                        <span>是 (V)</span>
                    </label>
                    <label style="display:flex; align-items:center; gap:6px; font-weight: 500; cursor: pointer;">
                        <input type="radio" name="yesno_${index}" class="status-yesno" data-index="${index}" value="否" ${savedYesNo === '否' ? 'checked' : ''} ${!isAdmin ? 'disabled' : ''} style="width: 16px; height: 16px; accent-color: var(--primary-color);">
                        <span>否</span>
                    </label>
                </div>
            `;
        } else if (item.type === 'dimensions_multi') {
            let dimHtml = '<div style="display:flex; flex-direction:column; gap:12px; max-height: 450px; overflow-y: auto; padding-right: 8px;">';
            for (let i = 1; i <= 15; i++) {
                const nameKey = `dim_${index}_${i}_name`;
                const lKey = `dim_${index}_${i}_L`;
                const wKey = `dim_${index}_${i}_W`;
                const hKey = `dim_${index}_${i}_H`;
                const holeKey = `dim_${index}_${i}_hole`;
                const partPhotoKey = `dim_${index}_${i}_photo`;

                const savedName = savedRecords[nameKey] || '';
                const savedL = savedRecords[lKey] || '';
                const savedW = savedRecords[wKey] || '';
                const savedH = savedRecords[hKey] || '';
                const savedHole = savedRecords[holeKey] || '';
                const savedPartPhoto = savedRecords[partPhotoKey] || null;

                let partPhotoHtml = '';
                if (savedPartPhoto) {
                    partPhotoHtml = `
                        <div class="photo-preview-mini" style="background-image: url('${savedPartPhoto}'); width: 44px; height: 44px; margin-left: auto; border-radius: 4px; position: relative;">
                            ${isAdmin ? `<button type="button" class="del-part-photo-btn" data-index="${index}" data-sub="${i}" style="position: absolute; top: -6px; right: -6px; background: #ef4444; color: white; border: none; border-radius: 50%; width: 18px; height: 18px; cursor: pointer; display: flex; align-items: center; justify-content: center; font-size: 14px; line-height: 1;">&times;</button>` : ''}
                        </div>
                    `;
                } else {
                    if (isAdmin) {
                        partPhotoHtml = `
                            <label class="btn-upload" style="margin-left: auto; font-size: 0.8rem; padding: 4px 8px; cursor: pointer;">
                                📁 照片
                                <input type="file" accept="image/*" class="part-photo-input" data-index="${index}" data-sub="${i}" style="display:none;">
                            </label>
                        `;
                    } else {
                        partPhotoHtml = `<span style="color:var(--text-secondary);font-size:0.8rem; margin-left: auto;">無照片</span>`;
                    }
                }

                dimHtml += `
                    <div style="background: rgba(255,255,255,0.03); padding: 12px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.1);">
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom: 10px;">
                            <div style="font-size: 0.9rem; font-weight: 600; color: #fff; white-space: nowrap;">部件 ${i}</div>
                            <input type="text" class="status-dim name-input form-control" data-index="${index}" data-sub="${i}" value="${savedName}" placeholder="在此輸入部件名稱..." style="flex: 1; padding: 4px 8px; font-size: 0.85rem;" ${!isAdmin ? 'disabled' : ''}>
                            ${partPhotoHtml}
                        </div>
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom: 10px;">
                            <input type="number" class="status-dim L-input form-control" data-index="${index}" data-sub="${i}" value="${savedL}" placeholder="長" style="width: 70px; text-align: center; padding: 6px; flex: 1;" ${!isAdmin ? 'disabled' : ''}>
                            <span style="color:var(--text-secondary); font-weight: bold;">x</span>
                            <input type="number" class="status-dim W-input form-control" data-index="${index}" data-sub="${i}" value="${savedW}" placeholder="寬" style="width: 70px; text-align: center; padding: 6px; flex: 1;" ${!isAdmin ? 'disabled' : ''}>
                            <span style="color:var(--text-secondary); font-weight: bold;">x</span>
                            <input type="number" class="status-dim H-input form-control" data-index="${index}" data-sub="${i}" value="${savedH}" placeholder="高" style="width: 70px; text-align: center; padding: 6px; flex: 1;" ${!isAdmin ? 'disabled' : ''}>
                        </div>
                        <div style="display:flex; align-items:center; gap:8px; font-size: 0.95rem;">
                            <span style="font-size: 0.85rem; color: #fff; white-space: nowrap;">水氣電管孔資訊：</span>
                            <div style="display:flex; gap: 12px; margin-left: 4px;">
                                <label style="display:flex; align-items:center; gap:4px; cursor: pointer;">
                                    <input type="radio" name="hole_${index}_${i}" class="status-dim hole-radio form-control" data-index="${index}" data-sub="${i}" value="有" ${savedHole === '有' ? 'checked' : ''} ${!isAdmin ? 'disabled' : ''} style="width: 14px; height: 14px; accent-color: var(--primary-color);">
                                    <span style="font-size: 0.85rem;">有</span>
                                </label>
                                <label style="display:flex; align-items:center; gap:4px; cursor: pointer;">
                                    <input type="radio" name="hole_${index}_${i}" class="status-dim hole-radio form-control" data-index="${index}" data-sub="${i}" value="無" ${savedHole === '無' ? 'checked' : ''} ${!isAdmin ? 'disabled' : ''} style="width: 14px; height: 14px; accent-color: var(--primary-color);">
                                    <span style="font-size: 0.85rem;">無</span>
                                </label>
                            </div>
                        </div>
                    </div>
                `;
            }
            dimHtml += '</div>';
            stateHtml = dimHtml;
        } else if (item.type === 'radio-eval') {
            const savedEval = savedRecords[`eval_${index}`] || '請選擇';
            stateHtml = `
                <select class="status-eval form-control w-full" data-index="${index}" ${!isAdmin ? 'disabled' : ''}>
                    <option value="請選擇" ${savedEval === '請選擇' ? 'selected' : ''}>請選擇</option>
                    <option value="是" ${savedEval === '是' ? 'selected' : ''}>是</option>
                    <option value="否" ${savedEval === '否' ? 'selected' : ''}>否</option>
                    <option value="不確定" ${savedEval === '不確定' ? 'selected' : ''}>不確定</option>
                </select>
            `;
        } else {
            stateHtml = `<input type="text" class="status-input w-full" data-index="${index}" value="${savedState}" placeholder="輸入紀錄..." ${!isAdmin ? 'disabled' : ''}>`;
        }

        tr.innerHTML = `
            <td><strong>${item.category}</strong></td>
            <td>${item.item}</td>
            <td>${item.target}</td>
            <td>${stateHtml}</td>
            <td>${photoHtml}</td>
        `;

        // Bind input event
        if (isAdmin) {
            if (item.type === 'checkbox+note') {
                const checkInput = tr.querySelector('.status-check');
                const noteInput = tr.querySelector('.status-note');
                checkInput.addEventListener('change', (e) => saveRecord(machineName, `check_${index}`, e.target.checked));
                noteInput.addEventListener('change', (e) => saveRecord(machineName, `note_${index}`, e.target.value));
            } else if (item.type === 'checkbox') {
                const checkInput = tr.querySelector('.status-check');
                checkInput.addEventListener('change', (e) => saveRecord(machineName, `check_${index}`, e.target.checked));
            } else if (item.type === 'radio-yesno') {
                const radios = tr.querySelectorAll('.status-yesno');
                radios.forEach(radio => {
                    radio.addEventListener('change', (e) => {
                        if (e.target.checked) {
                            saveRecord(machineName, `yesno_${index}`, e.target.value);
                        }
                    });
                });
            } else if (item.type === 'dimensions_multi') {
                const dims = tr.querySelectorAll('.status-dim');
                dims.forEach(inp => {
                    inp.addEventListener('change', (e) => {
                        let dl = '';
                        if (e.target.classList.contains('L-input')) dl = 'L';
                        else if (e.target.classList.contains('W-input')) dl = 'W';
                        else if (e.target.classList.contains('H-input')) dl = 'H';
                        else if (e.target.classList.contains('name-input')) dl = 'name';
                        else if (e.target.classList.contains('hole-radio') && e.target.checked) dl = 'hole';

                        if (dl) {
                            saveRecord(machineName, `dim_${index}_${e.target.dataset.sub}_${dl}`, e.target.value);
                        }
                    });
                });
                const partPhotoInputs = tr.querySelectorAll('.part-photo-input');
                partPhotoInputs.forEach(pInput => {
                    pInput.addEventListener('change', (e) => {
                        handlePhotoUpload(e, machineName, `dim_${index}_${e.target.dataset.sub}_photo`);
                    });
                });
                const delPartBtns = tr.querySelectorAll('.del-part-photo-btn');
                delPartBtns.forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        saveRecord(machineName, `dim_${index}_${e.currentTarget.dataset.sub}_photo`, null);
                        loadMachineData(machineName); // re-render
                        if (window.updateMachineListStatus) window.updateMachineListStatus();
                    });
                });
            } else if (item.type === 'radio-eval') {
                const evalSelect = tr.querySelector('.status-eval');
                evalSelect.addEventListener('change', (e) => {
                    saveRecord(machineName, `eval_${index}`, e.target.value);
                });
            } else {
                const textInput = tr.querySelector('.status-input');
                if (textInput) {
                    textInput.addEventListener('change', (e) => {
                        saveRecord(machineName, stateKey, e.target.value);
                    });
                }
            }

            const pInput = tr.querySelector('.photo-input');
            if (pInput) {
                pInput.addEventListener('change', (e) => handlePhotoUpload(e, machineName, photoKey));
            }

            const delBtn = tr.querySelector('.del-photo-btn');
            if (delBtn) {
                delBtn.addEventListener('click', () => {
                    saveRecord(machineName, photoKey, null);
                    loadMachineData(machineName); // re-render
                    if (window.updateMachineListStatus) window.updateMachineListStatus();
                });
            }
        }

        recordsTbody.appendChild(tr);
    });

    // Process Docs Table
    docsTbody.innerHTML = '';
    let savedDocs = await cloudStorage.get(`docs_${machineName}`);
    if (typeof savedDocs === 'string') {
        try { savedDocs = JSON.parse(savedDocs); } catch (e) { savedDocs = {}; }
    }
    savedDocs = savedDocs || {};

    docItems.forEach((item, index) => {
        const tr = document.createElement('tr');
        const receiveKey = `receive_${index}`;
        const noteKey = `note_${index}`;

        const savedReceive = savedDocs[receiveKey] || '無';
        const savedNote = savedDocs[noteKey] || '';

        tr.innerHTML = `
            <td><strong>${item.type}</strong></td>
            <td>${item.nameEn}</td>
            <td>${item.nameTw}</td>
            <td>
                <select class="docs-select w-full" data-key="${receiveKey}" ${!isAdmin ? 'disabled' : ''}>
                    <option value="無" ${savedReceive === '無' ? 'selected' : ''}>無</option>
                    <option value="紙本" ${savedReceive === '紙本' ? 'selected' : ''}>紙本</option>
                    <option value="電子檔" ${savedReceive === '電子檔' ? 'selected' : ''}>電子檔</option>
                    <option value="皆有" ${savedReceive === '皆有' ? 'selected' : ''}>皆有</option>
                </select>
            </td>
            <td>
                <input type="text" class="status-input w-full docs-note" data-key="${noteKey}" value="${savedNote}" placeholder="備註事項..." ${!isAdmin ? 'disabled' : ''}>
            </td>
        `;

        if (isAdmin) {
            const selectEl = tr.querySelector('.docs-select');
            const noteEl = tr.querySelector('.docs-note');

            selectEl.addEventListener('change', (e) => saveDocState(machineName, e.target.dataset.key, e.target.value));
            noteEl.addEventListener('change', (e) => saveDocState(machineName, e.target.dataset.key, e.target.value));
        }

        docsTbody.appendChild(tr);
    });
}

async function saveRecord(machine, key, value) {
    let data = await cloudStorage.get(`records_${machine}`) || {};
    if (typeof data === 'string') {
        try { data = JSON.parse(data); } catch (e) { data = {}; }
    }
    data[key] = value;
    await cloudStorage.set(`records_${machine}`, data);
}

async function saveDocState(machine, key, value) {
    let data = await cloudStorage.get(`docs_${machine}`) || {};
    if (typeof data === 'string') {
        try { data = JSON.parse(data); } catch (e) { data = {}; }
    }
    data[key] = value;
    await cloudStorage.set(`docs_${machine}`, data);
}

function handlePhotoUpload(e, machine, key) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async function (evt) {
        try {
            const compressed = await window.compressImage(evt.target.result);
            await saveRecord(machine, key, compressed);
            await loadMachineData(machine);
            if (window.updateMachineListStatus) window.updateMachineListStatus();
        } catch (err) {
            console.error(err);
        }
    };
    reader.readAsDataURL(file);
    e.target.value = ''; // Reset input so the same file could be selected again if needed
}

window.hasMachineDataUploaded = async function (machineName) {
    let savedRecords = await cloudStorage.get(`records_${machineName}`);
    if (typeof savedRecords === 'string') {
        try { savedRecords = JSON.parse(savedRecords); } catch (e) { savedRecords = {}; }
    }
    savedRecords = savedRecords || {};

    for (const key in savedRecords) {
        if (key.includes('photo') && savedRecords[key] !== null) {
            return true;
        }
    }
    return false;
};

window.updateMachineListStatus = async function () {
    const items = document.querySelectorAll('.machine-item');
    for (const item of items) {
        const input = item.querySelector('input[type="radio"]');
        if (!input) continue;
        const displayName = input.value.replace('.pdf', '');
        const hasData = await window.hasMachineDataUploaded(displayName);
        if (hasData) {
            item.classList.add('has-data');
        } else {
            item.classList.remove('has-data');
        }
    }
};

// ==========================================
// 8. Daily Report Logic
// ==========================================
async function initDailyReport() {
    const form = document.getElementById('dailyReportForm');
    if (!form) return;

    const reportDateInput = document.getElementById('reportDate');
    const photoInput = document.getElementById('photoInput');
    const photoGrid = document.getElementById('photoGrid');
    const uploadBtnWrapper = document.getElementById('uploadBtnWrapper');
    const photoCountSpan = document.getElementById('photoCount');
    const historyList = document.getElementById('historyList');
    const historyPlaceholder = document.getElementById('historyPlaceholder');
    const clearFormBtn = document.getElementById('clearFormBtn');
    const submitBtn = form.querySelector('button[type="submit"]');

    const MAX_PHOTOS = 5;
    let photosBase64 = [];

    // Date default
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    if (reportDateInput) reportDateInput.value = `${yyyy}-${mm}-${dd}`;

    await loadHistory();
    setTimeout(toggleDailyReportAuth, 100);

    if (photoInput) {
        photoInput.addEventListener('change', async (e) => {
            const files = Array.from(e.target.files);
            if (photosBase64.length + files.length > MAX_PHOTOS) {
                alert(`最多只能上傳 ${MAX_PHOTOS} 張照片，您目前已選 ${photosBase64.length} 張。`);
                photoInput.value = '';
                return;
            }

            for (const file of files) {
                try {
                    const base64 = await getBase64Report(file);
                    const compressed = await window.compressImage(base64);
                    photosBase64.push(compressed);
                } catch (err) {
                    console.error(err);
                }
            }
            photoInput.value = '';
            renderPhotoPreviews();
        });
    }

    function getBase64Report(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => resolve(reader.result);
            reader.onerror = error => reject(error);
        });
    }

    function renderPhotoPreviews() {
        if (!photoGrid) return;
        const existingPreviews = photoGrid.querySelectorAll('.photo-preview-box');
        existingPreviews.forEach(el => el.remove());

        photosBase64.forEach((base64, index) => {
            const box = document.createElement('div');
            box.className = 'photo-preview-box';
            box.innerHTML = `
                <img src="${base64}" alt="Preview">
                <button type="button" class="remove-photo-btn" data-index="${index}">&times;</button>
            `;
            photoGrid.insertBefore(box, uploadBtnWrapper);
        });

        photoGrid.querySelectorAll('.remove-photo-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                photosBase64.splice(parseInt(e.target.getAttribute('data-index')), 1);
                renderPhotoPreviews();
            });
        });

        if (photoCountSpan) photoCountSpan.textContent = photosBase64.length;
        if (uploadBtnWrapper) {
            uploadBtnWrapper.style.display = (isAdmin && photosBase64.length < MAX_PHOTOS) ? 'flex' : 'none';
        }
    }

    if (form) {
        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!isAdmin) return;

            if (photosBase64.length === 0) {
                if (!confirm("您尚未上傳任何照片。確定要送出嗎？")) return;
            }

            submitBtn.textContent = '傳送中...';
            submitBtn.disabled = true;

            const newReport = {
                id: Date.now().toString(),
                date: document.getElementById('reportDate').value,
                reporter: document.getElementById('reporterName').value.trim(),
                site: document.getElementById('siteName').value.trim(),
                progress: document.getElementById('progressDesc').value.trim(),
                issue: document.getElementById('issueDesc').value.trim() || '無',
                photos: [...photosBase64],
                timestamp: new Date().toISOString()
            };

            let history = await cloudStorage.get('dailyReports');
            if (typeof history === 'string') {
                try { history = JSON.parse(history); } catch (e) { history = []; }
            }
            history = history || [];
            history.unshift(newReport);

            try {
                await cloudStorage.set('dailyReports', history);
                alert('回報已成功送出！');
                clearForm();
                await loadHistory();
            } catch (err) {
                console.error(err);
                alert('進度回報儲存失敗，請重試。');
            } finally {
                submitBtn.textContent = '送出回報紀錄';
                submitBtn.disabled = false;
            }
        });
    }

    if (clearFormBtn) {
        clearFormBtn.addEventListener('click', () => {
            if (isAdmin && confirm('確定要清除所有填寫的內容與照片嗎？')) clearForm();
        });
    }

    function clearForm() {
        if (form) form.reset();
        if (reportDateInput) reportDateInput.value = `${yyyy}-${mm}-${dd}`;
        photosBase64 = [];
        renderPhotoPreviews();
    }

    async function loadHistory() {
        if (!historyList) return;
        historyList.innerHTML = '<div style="padding: 24px; text-align: center; color: var(--text-secondary);">資料載入中...</div>';

        let history = await cloudStorage.get('dailyReports');
        if (typeof history === 'string') {
            try { history = JSON.parse(history); } catch (e) { history = []; }
        }
        history = history || [];

        if (history.length === 0) {
            if (historyPlaceholder) historyPlaceholder.style.display = 'flex';
            historyList.innerHTML = '';
            // Just leaving the placeholder visible is fine
            return;
        }

        if (historyPlaceholder) historyPlaceholder.style.display = 'none';
        historyList.innerHTML = '';

        history.forEach(report => {
            const card = document.createElement('div');
            card.className = 'history-card';

            let issueHtml = '';
            if (report.issue && report.issue !== '無') {
                issueHtml = `
                    <div class="history-issue">
                        <div class="history-issue-title">異常 / 需協助事項</div>
                        <div>${report.issue}</div>
                    </div>
                `;
            }

            let photosHtml = '';
            if (report.photos && report.photos.length > 0) {
                const imgTags = report.photos.map(p => `<img src="${p}" class="history-photo-thumb" onclick="window.open('${p}', '_blank')" title="點擊放大預覽">`).join('');
                photosHtml = `<div class="history-photos">${imgTags}</div>`;
            }

            card.innerHTML = `
                <div class="history-header">
                    <span class="history-date">${report.date}</span>
                    <span class="history-reporter">${report.reporter}</span>
                </div>
                <div class="history-site">${report.site}</div>
                <div class="history-desc">${report.progress}</div>
                ${issueHtml}
                ${photosHtml}
            `;
            historyList.appendChild(card);
        });
    }
}

window.compressImage = function (base64Str, maxWidth = 1200, quality = 0.7) {
    return new Promise((resolve) => {
        const img = new Image();
        img.src = base64Str;
        img.onload = () => {
            const canvas = document.createElement('canvas');
            let width = img.width;
            let height = img.height;

            if (width > maxWidth) {
                height = Math.round((height * maxWidth) / width);
                width = maxWidth;
            }

            canvas.width = width;
            canvas.height = height;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, width, height);

            resolve(canvas.toDataURL('image/jpeg', quality));
        };
        img.onerror = () => resolve(base64Str); // Fallback to original if load fails
    });
};

// ==========================================
// 9. Excel Viewer Logic
// ==========================================
async function initExcelViewer() {
    const tableContainer = document.querySelector('#excelTable tbody');
    if (!tableContainer) return;

    try {
        const response = await fetch('機台清單文字/RF360_For_Sale_List_compiled.xlsx');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Grab the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert to JSON (array of arrays to capture header style easily)
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length === 0) {
            tableContainer.innerHTML = '<tr><td style="text-align: center; padding: 40px;">Excel 檔案為空</td></tr>';
            return;
        }

        let html = '';

        // We will render manually to add custom classes based on content
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            // Skip completely empty rows
            if (row.length === 0 || !row.some(cell => cell !== undefined && cell !== null && cell !== '')) {
                continue;
            }

            const isHeader = i === 0 || (row[0] && row[0].toString().includes('Check Date'));

            if (isHeader) {
                html += '<thead><tr>';
                for (let j = 0; j < row.length; j++) {
                    // Usually we only care about columns up to 'H' (index 7) based on the image provided
                    if (j > 7) break;
                    html += `<th>${row[j] || ''}</th>`;
                }
                html += '</tr></thead><tbody>';
            } else {
                html += '<tr>';
                for (let j = 0; j < row.length; j++) {
                    if (j > 7) break;

                    let cellValue = row[j] || '';
                    let cellClass = '';

                    // Apply special styling for the "狀態" column (usually index 7 / col H)
                    if (j === 7) {
                        cellClass = 'status-cell';
                        if (cellValue.includes('拆機中')) cellClass += ' 拆機中';
                        else if (cellValue.includes('已拆機')) {
                            cellClass += ' 已拆機';
                            if (cellValue.includes('警示') || cellValue.includes('管線')) cellClass += '-警示';
                        }
                    }

                    // Format dates if it's a number (Excel serial date)
                    if (typeof cellValue === 'number' && j === 0) {
                        // Very rough excel date parsing (usually not perfect without formatting info, but good enough for generic display)
                        const execDate = new Date(Math.round((cellValue - 25569) * 86400 * 1000));
                        if (!isNaN(execDate.getTime())) {
                            const month = execDate.getMonth() + 1;
                            const dec = execDate.getDate() + 1; // timezone offset adj
                            cellValue = `${month}月${dec}日`;
                        }
                    }

                    html += `<td class="${cellClass}">${cellValue}</td>`;
                }

                // Pad missing cells if the row is shorter
                for (let j = row.length; j <= 7; j++) {
                    html += `<td></td>`;
                }

                html += '</tr>';
            }
        }

        if (html.endsWith('<tbody>')) {
            html += '</tbody>';
        }

        // Remove the existing '<tbody>' wrapper, replace entire inner HTML of table
        document.getElementById('excelTable').innerHTML = html;

    } catch (error) {
        console.error("Error loading Excel file:", error);
        tableContainer.innerHTML = '<tr><td style="text-align: center; padding: 40px; color: #ef4444;">無法讀取 Excel 檔案，請確認檔案「機台清單文字/RF360_For_Sale_List_compiled.xlsx」存在。</td></tr>';
    }
}
