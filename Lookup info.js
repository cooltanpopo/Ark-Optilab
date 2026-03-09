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
    { category: "紀錄", item: "整體尺寸", target: "量測設備整體：長(cm) x 寬(cm) x 高(cm)", type: "checkbox" },
    { category: "紀錄", item: "組件尺寸", target: "各獨立組件拆解後的尺寸紀錄", type: "checkbox" },
    { category: "檢查", item: "損傷檢查", target: "檢查有無腐蝕、生鏽、化學噴濺痕跡", type: "checkbox-yesno" },
    { category: "紀錄", item: "電力資訊", target: "電壓(V)、相數(Ph)、瓦數(W) 或電流(A)", type: "checkbox" },
    { category: "紀錄", item: "水/氣介面", target: "紀錄管徑尺寸與介面規格 (如 VCR, Swagelok)", type: "checkbox" }
];

const docItems = [
    { type: "手冊", nameEn: "Equipment Manual", nameTw: "設備操作說明書" },
    { type: "手冊", nameEn: "Service Manual", nameTw: "維修技術手冊" },
    { type: "圖面", nameEn: "Electrical Schematic", nameTw: "電氣電路圖" },
    { type: "圖面", nameEn: "Pneumatic Diagram", nameTw: "氣壓原理圖" },
    { type: "圖面", nameEn: "Gas Diagram", nameTw: "氣體管路圖" }
];

// ==========================================
// 2. Global State & Auth
// ==========================================
let isAdmin = false;
let currentMachine = null;

// ==========================================
// 3. Initialization
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
    initAuth();
    initMachineList();
    initTabs();
    initSPA();
    initDailyReport();
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
function initMachineList() {
    const searchInput = document.getElementById('searchInput');
    const machineListElement = document.getElementById('machineList');

    function renderList(files) {
        machineListElement.innerHTML = '';

        if (files.length === 0) {
            machineListElement.innerHTML = '<div style="padding: 24px 12px; color: var(--text-secondary); text-align: center; font-size: 0.9rem;">查無符合的機台檔案</div>';
            return;
        }

        files.forEach((file, index) => {
            const displayName = file.replace('.pdf', '');
            const item = document.createElement('label');
            item.className = 'machine-item';
            if (window.hasMachineDataUploaded && window.hasMachineDataUploaded(displayName)) {
                item.classList.add('has-data');
            }

            const savedLight = localStorage.getItem(`statusLight_${displayName}`) || 'red';

            const radioHtml = `
                <div class="checkbox-wrapper">
                    <input type="radio" name="machineSelect" value="${file}" id="machine_${index}">
                </div>
                <div class="item-content">
                    <div class="item-title" title="${displayName}">${displayName}</div>
                </div>
                <div class="traffic-light-container">
                    <div class="traffic-light red ${savedLight === 'red' ? 'active' : ''}" data-color="red" data-machine="${displayName}" title="尚未紀錄"></div>
                    <div class="traffic-light yellow ${savedLight === 'yellow' ? 'active' : ''}" data-color="yellow" data-machine="${displayName}" title="部分紀錄"></div>
                    <div class="traffic-light green ${savedLight === 'green' ? 'active' : ''}" data-color="green" data-machine="${displayName}" title="完整登錄"></div>
                </div>
            `;
            item.innerHTML = radioHtml;

            // Handle machine selection
            item.querySelector('input').addEventListener('change', (e) => {
                document.querySelectorAll('.machine-item').forEach(el => el.classList.remove('active'));
                if (e.target.checked) {
                    item.classList.add('active');
                    showMachineDetail(displayName, file);
                }
            });

            // Handle traffic light clicks
            const lights = item.querySelectorAll('.traffic-light');
            lights.forEach(light => {
                light.addEventListener('click', (e) => {
                    e.preventDefault(); // prevent triggering the label radio
                    if (!isAdmin) {
                        alert("僅管理員可更改機台燈號狀態。");
                        return;
                    }
                    const color = light.getAttribute('data-color');
                    const machine = light.getAttribute('data-machine');

                    // Update UI
                    lights.forEach(l => l.classList.remove('active'));
                    light.classList.add('active');

                    // Save State
                    localStorage.setItem(`statusLight_${machine}`, color);
                });
            });

            machineListElement.appendChild(item);
        });
    }

    renderList(machineFiles);

    searchInput.addEventListener('input', (e) => {
        const term = e.target.value.toLowerCase();
        const filtered = machineFiles.filter(f => f.toLowerCase().includes(term));
        renderList(filtered);
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

function loadMachineData(machineName) {
    const recordsTbody = document.getElementById('recordsTbody');
    const docsTbody = document.getElementById('docsTbody');

    // Process Records Table
    recordsTbody.innerHTML = '';
    const savedRecords = JSON.parse(localStorage.getItem(`records_${machineName}`)) || {};

    checkItems.forEach((item, index) => {
        const tr = document.createElement('tr');
        const stateKey = `state_${index}`;
        const photoKey = `photo_${index}`;
        const savedState = savedRecords[stateKey] || '';
        const savedPhoto = savedRecords[photoKey] || null;

        let photoHtml = '';
        if (savedPhoto) {
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
        } else if (item.type === 'checkbox' || item.type === 'checkbox-yesno') {
            const isChecked = savedRecords[`check_${index}`] || false;
            const labelText = item.type === 'checkbox-yesno' ? '是 (V)' : '完成 (V)';
            stateHtml = `
                <div style="display:flex; flex-direction:column; gap:6px;">
                    <label style="display:flex; align-items:center; gap:6px; font-weight: 500;">
                        <input type="checkbox" class="status-check" data-index="${index}" ${isChecked ? 'checked' : ''} ${!isAdmin ? 'disabled' : ''} style="width: 16px; height: 16px; cursor: pointer;">
                        <span>${labelText}</span>
                    </label>
                </div>
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
            } else if (item.type === 'checkbox' || item.type === 'checkbox-yesno') {
                const checkInput = tr.querySelector('.status-check');
                checkInput.addEventListener('change', (e) => saveRecord(machineName, `check_${index}`, e.target.checked));
            } else {
                const textInput = tr.querySelector('.status-input');
                textInput.addEventListener('change', (e) => {
                    saveRecord(machineName, stateKey, e.target.value);
                });
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
    const savedDocs = JSON.parse(localStorage.getItem(`docs_${machineName}`)) || {};

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

            selectEl.addEventListener('change', (e) => {
                saveDocState(machineName, e.target.dataset.key, e.target.value);
            });
            noteEl.addEventListener('change', (e) => {
                saveDocState(machineName, e.target.dataset.key, e.target.value);
            });
        }

        docsTbody.appendChild(tr);
    });
}

function saveRecord(machine, key, value) {
    let data = JSON.parse(localStorage.getItem(`records_${machine}`)) || {};
    data[key] = value;
    try {
        localStorage.setItem(`records_${machine}`, JSON.stringify(data));
    } catch (e) {
        alert('照片儲存失敗！可能是因為您的裝置/瀏覽器儲存空間已滿，請嘗試刪除其他機台的照片後再試。');
        console.error(e);
    }
}

function saveDocState(machine, key, value) {
    let data = JSON.parse(localStorage.getItem(`docs_${machine}`)) || {};
    data[key] = value;
    localStorage.setItem(`docs_${machine}`, JSON.stringify(data));
}

function handlePhotoUpload(e, machine, key) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async function (evt) {
        try {
            const compressed = await window.compressImage(evt.target.result);
            saveRecord(machine, key, compressed);
            loadMachineData(machine);
            if (window.updateMachineListStatus) window.updateMachineListStatus();
        } catch (err) {
            console.error(err);
        }
    };
    reader.readAsDataURL(file);
    e.target.value = ''; // Reset input so the same file could be selected again if needed
}

window.hasMachineDataUploaded = function (machineName) {
    const savedRecords = JSON.parse(localStorage.getItem(`records_${machineName}`)) || {};
    for (const key in savedRecords) {
        if (key.startsWith('photo_') && savedRecords[key] !== null) {
            return true;
        }
    }
    return false;
};

window.updateMachineListStatus = function () {
    const items = document.querySelectorAll('.machine-item');
    items.forEach(item => {
        const input = item.querySelector('input[type="radio"]');
        if (!input) return;
        const displayName = input.value.replace('.pdf', '');
        if (window.hasMachineDataUploaded(displayName)) {
            item.classList.add('has-data');
        } else {
            item.classList.remove('has-data');
        }
    });
};

// ==========================================
// 8. Daily Report Logic
// ==========================================
function initDailyReport() {
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

    const MAX_PHOTOS = 5;
    let photosBase64 = [];

    // Date default
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    if (reportDateInput) reportDateInput.value = `${yyyy}-${mm}-${dd}`;

    loadHistory();
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
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            if (!isAdmin) return;

            if (photosBase64.length === 0) {
                if (!confirm("您尚未上傳任何照片。確定要送出嗎？")) return;
            }

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

            const history = JSON.parse(localStorage.getItem('dailyReports')) || [];
            history.unshift(newReport);

            try {
                localStorage.setItem('dailyReports', JSON.stringify(history));
                alert('回報已成功送出！');
                clearForm();
                loadHistory();
            } catch (err) {
                console.error(err);
                alert('儲存失敗，可能是照片總容量超過瀏覽器限制 (10MB)。');
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

    function loadHistory() {
        if (!historyList) return;
        const history = JSON.parse(localStorage.getItem('dailyReports')) || [];

        if (history.length === 0) {
            if (historyPlaceholder) historyPlaceholder.style.display = 'flex';
            historyList.querySelectorAll('.history-card').forEach(c => c.remove());
            return;
        }

        if (historyPlaceholder) historyPlaceholder.style.display = 'none';
        historyList.querySelectorAll('.history-card').forEach(c => c.remove());

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
