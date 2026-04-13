/**
 * app.js - Main application logic for Easy Statistic Analysis Tools
 * Depends on: window.Stats (stats.js), XLSX (SheetJS CDN), jStat (CDN)
 */
(function () {
    "use strict";

    // =========================================================================
    // Global State
    // =========================================================================

    var state = {
        data: null,
        columns: [],
        numericCols: [],
        categoricalCols: [],
        currentPage: 'home',
        results: {},
        aiResults: {},
        aiSettings: { apiKey: '', model: 'gemini-2.5-flash-lite' },
        chatHistory: []
    };

    // =========================================================================
    // Utility: Format Number
    // =========================================================================

    function fmt(v, decimals) {
        if (v === null || v === undefined || (typeof v === 'number' && isNaN(v))) return '';
        if (typeof v !== 'number') return String(v);
        if (decimals === undefined) decimals = 4;
        return v.toFixed(decimals);
    }

    // =========================================================================
    // Login / Logout
    // =========================================================================

    function handleLogin(event) {
        if (event) event.preventDefault();
        var username = document.getElementById('username').value.trim();
        var password = document.getElementById('password').value.trim();
        if (username === 'thankyou' && password === '1234') {
            document.getElementById('login-page').style.display = 'none';
            document.getElementById('app-container').style.display = 'flex';
            navigateTo('home');
        } else {
            alert('Invalid username or password. Please try again.');
        }
    }

    function handleLogout() {
        state.data = null;
        state.columns = [];
        state.numericCols = [];
        state.categoricalCols = [];
        state.currentPage = 'home';
        state.results = {};
        state.aiResults = {};
        state.chatHistory = [];
        document.getElementById('app-container').style.display = 'none';
        document.getElementById('login-page').style.display = 'flex';
        document.getElementById('username').value = '';
        document.getElementById('password').value = '';
        var uploadStatus = document.getElementById('upload-status');
        if (uploadStatus) uploadStatus.innerHTML = '';
        var fileInput = document.getElementById('file-upload');
        if (fileInput) fileInput.value = '';
    }

    // =========================================================================
    // File Upload
    // =========================================================================

    function handleFileUpload(event) {
        var file = event.target.files[0];
        if (!file) return;

        var reader = new FileReader();
        reader.onload = function (e) {
            try {
                var data = new Uint8Array(e.target.result);
                var workbook = XLSX.read(data, { type: 'array' });
                var sheetName = workbook.SheetNames[0];
                var sheet = workbook.Sheets[sheetName];
                var jsonData = XLSX.utils.sheet_to_json(sheet, { defval: null });

                if (!jsonData || jsonData.length === 0) {
                    alert('The uploaded file is empty or has no valid data.');
                    return;
                }

                state.data = jsonData;
                state.columns = Object.keys(jsonData[0]);
                classifyColumns();

                var statusEl = document.getElementById('upload-status');
                if (statusEl) {
                    statusEl.innerHTML = '<span class="status-success">&#10004; ' + jsonData.length + ' rows &times; ' + state.columns.length + ' cols</span>';
                }

                populateSelects();
                updateHomePage();

            } catch (err) {
                alert('Error reading file: ' + err.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function classifyColumns() {
        state.numericCols = [];
        state.categoricalCols = [];
        if (!state.data || state.data.length === 0) return;

        state.columns.forEach(function (col) {
            var isNumeric = true;
            for (var i = 0; i < state.data.length; i++) {
                var v = state.data[i][col];
                if (v === null || v === undefined || v === '') continue;
                if (typeof v === 'number') continue;
                var parsed = Number(v);
                if (isNaN(parsed)) { isNumeric = false; break; }
            }
            if (isNumeric) {
                state.numericCols.push(col);
            } else {
                state.categoricalCols.push(col);
            }
        });
    }

    function updateHomePage() {
        if (!state.data) return;
        var rows = state.data.length;
        var cols = state.columns.length;
        setText('summary-rows', rows);
        setText('summary-cols', cols);
        setText('summary-numeric', state.numericCols.length);
        setText('summary-categorical', state.categoricalCols.length);

        // Data preview (first 20 rows)
        var preview = state.data.slice(0, 20);
        var previewEl = document.getElementById('data-preview');
        if (previewEl) {
            previewEl.innerHTML = buildTable(preview, state.columns);
        }
    }

    function setText(id, text) {
        var el = document.getElementById(id);
        if (el) el.textContent = text;
    }

    // =========================================================================
    // Navigation
    // =========================================================================

    function navigateTo(pageName) {
        var pages = document.querySelectorAll('.page');
        pages.forEach(function (p) { p.classList.remove('active'); });

        var target = document.getElementById('page-' + pageName);
        if (target) target.classList.add('active');

        state.currentPage = pageName;

        // Close sidebar on mobile after navigation
        var sidebar = document.getElementById('sidebar');
        var overlay = document.getElementById('sidebar-overlay');
        if (sidebar) sidebar.classList.remove('open');
        if (overlay) overlay.classList.remove('active');

        var menuItems = document.querySelectorAll('.menu-item');
        menuItems.forEach(function (item) { item.classList.remove('active'); });
        menuItems.forEach(function (item) {
            if (item.getAttribute('onclick') && item.getAttribute('onclick').indexOf("'" + pageName + "'") !== -1) {
                item.classList.add('active');
            }
        });

        if (state.data) {
            populateSelects();
            // Initialize dual-list pickers if they exist on the page
            initDualListsForPage(pageName);
        }

        // Restore results if they exist for current page prefix
        var prefixMap = getPagePrefixMap();
        var prefix = prefixMap[pageName];
        if (prefix && state.results[prefix]) {
            displayResults(prefix);
        }

        // AI chat: check key warning
        if (pageName === 'ai-chat') {
            var warn = document.getElementById('ai-no-key-warning');
            if (warn) {
                warn.style.display = state.aiSettings.apiKey ? 'none' : '';
            }
        }
    }

    function getPagePrefixMap() {
        return {
            'demographics': 'demo',
            'descriptive': 'desc',
            'numeric': 'num',
            'nominal': 'nom',
            'likert': 'lk',
            'interval': 'intv',
            'outlier': 'out',
            'normality': 'norm',
            'independent-ttest': 'itt',
            'paired-ttest': 'ptt',
            'oneway-anova': 'ow',
            'twoway-anova': 'tw',
            'rm-anova': 'rm',
            'ancova': 'anc',
            'mann-whitney': 'mw',
            'wilcoxon': 'wx',
            'kruskal-wallis': 'kw',
            'friedman': 'fr',
            'chi-square': 'chi',
            'correlation': 'cor',
            'linear-regression': 'lr',
            'logistic-regression': 'logr',
            'assumption': 'asn',
            'reliability': 'rel',
            'effect-size': 'cd',
            'factor-analysis': 'efa',
            'crosstab': 'ct',
            'posthoc': 'ph',
            'bootstrap': 'boot',
            'zscore': 'zs',
            'multicollinearity': 'vifp',
            'charts': 'chart',
            'partial-corr': 'pcor',
            'hierarchical-reg': 'hreg',
            'roc': 'roc',
            'icc': 'icc',
            'split-half': 'sh',
            'mcnemar': 'mcn',
            'fisher-exact': 'fe',
            'cochran-q': 'cq',
            'power-analysis':'pwr','welch-anova':'wa','dunnett':'dnt','median-test':'mdt',
            'runs-test':'run','ks2':'ks2','cluster':'cls','discriminant':'da','missing':'miss','qqplot':'qq',
            'survival':'surv','time-series':'ts'
        };
    }

    // =========================================================================
    // Menu Group Toggle
    // =========================================================================

    function toggleSidebar() {
        var sidebar = document.getElementById('sidebar');
        var overlay = document.getElementById('sidebar-overlay');
        if (!sidebar) return;
        var isOpen = sidebar.classList.contains('open');
        if (isOpen) {
            sidebar.classList.remove('open');
            if (overlay) overlay.classList.remove('active');
        } else {
            sidebar.classList.add('open');
            if (overlay) overlay.classList.add('active');
        }
    }

    function toggleMenuGroup(header) {
        var group = header.parentElement;
        var items = header.nextElementSibling;
        var arrow = header.querySelector('.menu-arrow');
        if (!items) return;

        var isOpen = group.classList.contains('open');
        if (isOpen) {
            // Close
            group.classList.remove('open');
            items.style.display = 'none';
            if (arrow) arrow.textContent = '\u25B8'; // ▸
        } else {
            // Open
            group.classList.add('open');
            items.style.display = 'block';
            if (arrow) arrow.textContent = '\u25BE'; // ▾
        }
    }

    // =========================================================================
    // Tab Switching
    // =========================================================================

    function switchTab(prefix, tabId) {
        // Hide all tab contents in the same page
        var page = document.getElementById('page-' + getPageFromPrefix(prefix));
        if (!page) return;
        var contents = page.querySelectorAll('.tab-content');
        contents.forEach(function (tc) { tc.style.display = 'none'; });

        var target = document.getElementById(prefix + '-' + tabId.replace('-tab', '-tab'));
        if (target) target.style.display = '';

        // Toggle active class on tab buttons
        var tabs = page.querySelectorAll('.tab-btn');
        tabs.forEach(function (tb) { tb.classList.remove('active'); });
        // Find the clicked button
        tabs.forEach(function (tb) {
            if (tb.getAttribute('onclick') && tb.getAttribute('onclick').indexOf("'" + tabId + "'") !== -1) {
                tb.classList.add('active');
            }
        });
    }

    function getPageFromPrefix(prefix) {
        var map = getPagePrefixMap();
        for (var page in map) {
            if (map[page] === prefix) return page;
        }
        return '';
    }

    // =========================================================================
    // Populate Variable Selects
    // =========================================================================

    function populateSelects() {
        if (!state.data) return;
        var selects = document.querySelectorAll('.var-select');
        selects.forEach(function (sel) {
            // Skip non-variable selects (those with existing option values like alpha, scale, method)
            if (sel.id === 'lk-scale' || sel.id === 'ai-model' ||
                sel.id.indexOf('-alpha') !== -1 || sel.id === 'cor-method') {
                return;
            }

            var dataType = sel.getAttribute('data-type');
            var currentValues = getSelected(sel.id);
            sel.innerHTML = '';

            // Add placeholder for single selects
            if (!sel.multiple) {
                var placeholder = document.createElement('option');
                placeholder.value = '';
                placeholder.textContent = '-- Select Variable --';
                sel.appendChild(placeholder);
            }

            var cols;
            if (dataType === 'numeric') {
                cols = state.numericCols;
            } else {
                cols = state.columns;
            }

            cols.forEach(function (col) {
                var opt = document.createElement('option');
                opt.value = col;
                opt.textContent = col;
                if (currentValues.indexOf(col) !== -1) opt.selected = true;
                sel.appendChild(opt);
            });
        });

        // Populate checkbox pickers
        var pickerConfigs = [
            // Basic Analysis
            {id: 'desc-picker', filter: 'numeric', label: 'เลือกตัวแปร'},
            {id: 'num-picker', filter: 'numeric', label: 'เลือกตัวแปร'},
            {id: 'nom-picker', filter: 'all', label: 'เลือกตัวแปร'},
            {id: 'lk-picker', filter: 'numeric', label: 'เลือกตัวแปร Likert'},
            {id: 'intv-picker', filter: 'numeric', label: 'เลือกตัวแปร'},
            {id: 'out-picker', filter: 'numeric', label: 'เลือกตัวแปร'},
            {id: 'norm-picker', filter: 'numeric', label: 'เลือกตัวแปร'},
            // Parametric — DV + IV
            {id: 'itt-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'itt-iv-picker', filter: 'all', label: 'ตัวแปรอิสระ / กลุ่ม (IV)'},
            {id: 'ptt-before-picker', filter: 'numeric', label: 'ตัวแปร Before'},
            {id: 'ptt-after-picker', filter: 'numeric', label: 'ตัวแปร After'},
            {id: 'ow-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'ow-iv-picker', filter: 'all', label: 'ตัวแปรอิสระ / Factor'},
            {id: 'tw-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'tw-iv1-picker', filter: 'all', label: 'Factor A'},
            {id: 'tw-iv2-picker', filter: 'all', label: 'Factor B'},
            {id: 'rm-picker', filter: 'numeric', label: 'ตัวแปรวัดซ้ำ (3+)'},
            {id: 'anc-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'anc-iv-picker', filter: 'all', label: 'Factor'},
            {id: 'anc-cov-picker', filter: 'numeric', label: 'Covariate'},
            // Non-Parametric — DV + IV
            {id: 'mw-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'mw-iv-picker', filter: 'all', label: 'ตัวแปรอิสระ / กลุ่ม (IV)'},
            {id: 'wx-before-picker', filter: 'numeric', label: 'ตัวแปร Before'},
            {id: 'wx-after-picker', filter: 'numeric', label: 'ตัวแปร After'},
            {id: 'kw-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'kw-iv-picker', filter: 'all', label: 'ตัวแปรอิสระ / กลุ่ม (IV)'},
            {id: 'fr-picker', filter: 'numeric', label: 'ตัวแปรวัดซ้ำ (3+)'},
            {id: 'chi-var1-picker', filter: 'all', label: 'ตัวแปร 1 (แถว)'},
            {id: 'chi-var2-picker', filter: 'all', label: 'ตัวแปร 2 (คอลัมน์)'},
            // Advanced
            {id: 'cor-picker', filter: 'numeric', label: 'ตัวแปร (2 ตัวขึ้นไป)'},
            {id: 'lr-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'lr-iv-picker', filter: 'numeric', label: 'ตัวแปรอิสระ (IVs)'},
            {id: 'logr-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV) — 0/1'},
            {id: 'logr-iv-picker', filter: 'numeric', label: 'ตัวแปรอิสระ (IVs)'},
            {id: 'rel-picker', filter: 'numeric', label: 'เลือกข้อ (Items)'},
            // Assumption & Effect Size
            {id: 'asn-lev-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'asn-lev-iv-picker', filter: 'all', label: 'ตัวแปรกลุ่ม'},
            {id: 'cd-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'cd-iv-picker', filter: 'all', label: 'ตัวแปรอิสระ (2 กลุ่ม)'},
            {id: 'cd-or-var1-picker', filter: 'all', label: 'ตัวแปร Exposure'},
            {id: 'cd-or-var2-picker', filter: 'all', label: 'ตัวแปร Outcome'},
            {id: 'asn-norm-vars-picker', filter: 'numeric', label: 'เลือกตัวแปร'},
            {id: 'asn-vif-vars-picker', filter: 'numeric', label: 'เลือกตัวแปร (2+)'},
            // SPSS/Stata Advanced
            {id: 'efa-picker', filter: 'numeric', label: 'เลือกตัวแปร (3+)'},
            {id: 'ct-var1-picker', filter: 'all', label: 'ตัวแปร 1 (แถว)'},
            {id: 'ct-var2-picker', filter: 'all', label: 'ตัวแปร 2 (คอลัมน์)'},
            {id: 'ph-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'ph-iv-picker', filter: 'all', label: 'Factor'},
            {id: 'boot-picker', filter: 'numeric', label: 'เลือกตัวแปร'},
            {id: 'zs-picker', filter: 'numeric', label: 'เลือกตัวแปร'},
            {id: 'vifp-picker', filter: 'numeric', label: 'เลือกตัวแปร (2+)'},
            // Charts & New SPSS features
            {id: 'chart-picker', filter: 'numeric', label: 'เลือกตัวแปร'},
            {id: 'pcor-x-picker', filter: 'numeric', label: 'Variable X'},
            {id: 'pcor-y-picker', filter: 'numeric', label: 'Variable Y'},
            {id: 'pcor-ctrl-picker', filter: 'numeric', label: 'Control Variables'},
            {id: 'hreg-dv-picker', filter: 'numeric', label: 'ตัวแปรตาม (DV)'},
            {id: 'hreg-b1-picker', filter: 'numeric', label: 'Block 1 IVs'},
            {id: 'hreg-b2-picker', filter: 'numeric', label: 'Block 2 IVs'},
            {id: 'roc-actual-picker', filter: 'numeric', label: 'Actual (0/1)'},
            {id: 'roc-pred-picker', filter: 'numeric', label: 'Predicted Prob'},
            {id: 'icc-picker', filter: 'numeric', label: 'Rater/Measure Variables'},
            {id: 'sh-picker', filter: 'numeric', label: 'เลือก Items'},
            {id: 'mcn-before-picker', filter: 'all', label: 'Before (0/1)'},
            {id: 'mcn-after-picker', filter: 'all', label: 'After (0/1)'},
            {id: 'fe-var1-picker', filter: 'all', label: 'Variable 1'},
            {id: 'fe-var2-picker', filter: 'all', label: 'Variable 2'},
            {id: 'cq-picker', filter: 'numeric', label: 'Conditions (0/1, 3+)'},
            // Demographics
            {id: 'demo-picker', filter: 'all', label: 'เลือกตัวแปรคุณลักษณะส่วนบุคคล'},
            {id: 'demo-numeric-picker', filter: 'numeric', label: 'ตัวแปรที่เป็นตัวเลข (แสดง Mean, S.D.)'},
            {id:'wa-dv-picker',filter:'numeric',label:'DV'},{id:'wa-iv-picker',filter:'all',label:'Factor'},
            {id:'dnt-dv-picker',filter:'numeric',label:'DV'},{id:'dnt-iv-picker',filter:'all',label:'Factor'},
            {id:'mdt-dv-picker',filter:'numeric',label:'DV'},{id:'mdt-iv-picker',filter:'all',label:'Factor'},
            {id:'run-picker',filter:'numeric',label:'เลือกตัวแปร'},
            {id:'ks2-dv-picker',filter:'numeric',label:'DV'},{id:'ks2-iv-picker',filter:'all',label:'Grouping'},
            {id:'cls-picker',filter:'numeric',label:'ตัวแปร (2+)'},
            {id:'da-group-picker',filter:'all',label:'Grouping (2 กลุ่ม)'},{id:'da-vars-picker',filter:'numeric',label:'Predictors'},
            {id:'qq-var-picker',filter:'numeric',label:'เลือกตัวแปร'},
            {id:'surv-time-picker',filter:'numeric',label:'Time'},
            {id:'surv-event-picker',filter:'numeric',label:'Event (0/1)'},
            {id:'surv-group-picker',filter:'all',label:'Group (optional)'},
            {id:'surv-cov-picker',filter:'numeric',label:'Covariates'},
            {id:'ts-var-picker',filter:'numeric',label:'Value Variable'},
        ];
        pickerConfigs.forEach(function(cfg) {
            var el = document.getElementById(cfg.id);
            if (el) createCheckboxPicker(cfg.id, cfg.filter, cfg.label);
        });
    }

    // =========================================================================
    // Checkbox-based Variable Picker
    // =========================================================================

    // Store selected vars per picker: { pickerId: [var1, var2, ...] }
    var pickerSelections = {};

    // Store picker configs for modal reuse
    var pickerRegistry = {};

    function createCheckboxPicker(containerId, filterType, label) {
        var container = document.getElementById(containerId);
        if (!container || !state.data) return;

        // Register this picker
        pickerRegistry[containerId] = { filter: filterType, label: label };
        if (!pickerSelections[containerId]) pickerSelections[containerId] = [];

        // Render button + tags display
        renderPickerButton(containerId, label);
    }

    function renderPickerButton(containerId, label) {
        var container = document.getElementById(containerId);
        if (!container) return;
        var selected = pickerSelections[containerId] || [];

        var html = '<div class="picker-box">';
        html += '<button type="button" class="btn-picker" onclick="openVarModal(\'' + containerId + '\')">';
        html += '📋 ' + (label || 'เลือกตัวแปร');
        if (selected.length > 0) html += ' <span class="picker-badge">' + selected.length + '</span>';
        html += '</button>';

        if (selected.length > 0) {
            html += '<div class="picker-tags">';
            selected.forEach(function(v) {
                html += '<span class="picker-tag">' + v + ' <span class="picker-tag-x" onclick="removePickerVar(\'' + containerId + '\',\'' + v + '\')">&times;</span></span>';
            });
            html += '</div>';
        }
        html += '</div>';
        container.innerHTML = html;
    }

    function openVarModal(pickerId) {
        var reg = pickerRegistry[pickerId];
        if (!reg || !state.data) return;

        var cols = reg.filter === 'numeric' ? state.numericCols :
                   reg.filter === 'categorical' ? state.categoricalCols : state.columns;
        var selected = pickerSelections[pickerId] || [];

        // Build modal content
        var modal = document.getElementById('var-picker-modal');
        if (!modal) {
            // Create modal if not exists
            modal = document.createElement('div');
            modal.id = 'var-picker-modal';
            modal.className = 'modal-overlay';
            modal.innerHTML = '<div class="modal-content">' +
                '<div class="modal-header"><h3 id="var-modal-title"></h3><button class="modal-close" onclick="closeVarModal()">&times;</button></div>' +
                '<div class="modal-body">' +
                '<div class="var-modal-actions"><button class="btn-sm" onclick="varModalSelectAll()">เลือกทั้งหมด</button> <button class="btn-sm" onclick="varModalDeselectAll()">ยกเลิกทั้งหมด</button></div>' +
                '<div id="var-modal-grid" class="var-modal-grid"></div>' +
                '<div id="var-modal-selected" class="var-modal-selected-area"></div>' +
                '</div>' +
                '<div class="modal-footer"><button class="btn btn-primary" onclick="confirmVarModal()">✅ ตกลง</button><button class="btn btn-export" onclick="closeVarModal()">ยกเลิก</button></div>' +
                '</div>';
            document.body.appendChild(modal);
        }

        // Store current picker ID
        modal.setAttribute('data-picker-id', pickerId);

        // Set title
        document.getElementById('var-modal-title').textContent = '📋 ' + (reg.label || 'เลือกตัวแปร');

        // Build pills grid
        var grid = document.getElementById('var-modal-grid');
        var html = '';
        cols.forEach(function(col) {
            var isSelected = selected.indexOf(col) !== -1;
            html += '<button type="button" class="var-pill ' + (isSelected ? 'var-pill-active' : '') + '" onclick="toggleVarPill(this)" data-var="' + col + '">' + col + '</button>';
        });
        grid.innerHTML = html;

        // Update selected count
        updateVarModalCount();

        modal.style.display = 'flex';
    }

    function toggleVarPill(btn) {
        btn.classList.toggle('var-pill-active');
        updateVarModalCount();
    }

    function varModalSelectAll() {
        var pills = document.querySelectorAll('#var-modal-grid .var-pill');
        pills.forEach(function(p) { p.classList.add('var-pill-active'); });
        updateVarModalCount();
    }

    function varModalDeselectAll() {
        var pills = document.querySelectorAll('#var-modal-grid .var-pill');
        pills.forEach(function(p) { p.classList.remove('var-pill-active'); });
        updateVarModalCount();
    }

    function updateVarModalCount() {
        var active = document.querySelectorAll('#var-modal-grid .var-pill-active');
        var area = document.getElementById('var-modal-selected');
        if (area) {
            area.textContent = 'เลือกแล้ว ' + active.length + ' ตัวแปร';
        }
    }

    function confirmVarModal() {
        var modal = document.getElementById('var-picker-modal');
        var pickerId = modal.getAttribute('data-picker-id');
        var active = document.querySelectorAll('#var-modal-grid .var-pill-active');
        var selected = Array.from(active).map(function(p) { return p.getAttribute('data-var'); });

        pickerSelections[pickerId] = selected;

        var reg = pickerRegistry[pickerId];
        renderPickerButton(pickerId, reg ? reg.label : '');

        modal.style.display = 'none';
    }

    function closeVarModal() {
        var modal = document.getElementById('var-picker-modal');
        if (modal) modal.style.display = 'none';
    }

    function removePickerVar(pickerId, varName) {
        var arr = pickerSelections[pickerId] || [];
        pickerSelections[pickerId] = arr.filter(function(v) { return v !== varName; });
        var reg = pickerRegistry[pickerId];
        renderPickerButton(pickerId, reg ? reg.label : '');
    }

    function getCheckedVars(containerId) {
        return pickerSelections[containerId] || [];
    }

    function getPickerValue(pickerId, fallbackSelectId) {
        // Get first checked value from picker, fallback to select dropdown
        var vals = getCheckedVars(pickerId);
        if (vals.length > 0) return vals[0];
        if (fallbackSelectId) return getSelectValue(fallbackSelectId);
        return '';
    }

    function getPickerValues(pickerId, fallbackSelectId) {
        // Get all checked values from picker, fallback to select
        var vals = getCheckedVars(pickerId);
        if (vals.length > 0) return vals;
        if (fallbackSelectId) {
            var sel = getSelected(fallbackSelectId);
            if (sel.length > 0) return sel;
            var single = getSelectValue(fallbackSelectId);
            return single ? [single] : [];
        }
        return [];
    }

    // =========================================================================
    // Dual-List Variable Picker
    // =========================================================================

    function createDualListPicker(containerId, sourceLabel, targetLabel, filterType) {
        // filterType: 'numeric', 'categorical', 'all'
        var container = document.getElementById(containerId);
        if (!container || !state.data) return;

        var cols = filterType === 'numeric' ? state.numericCols :
                   filterType === 'categorical' ? state.categoricalCols : state.columns;

        container.innerHTML =
            '<div class="dual-list">' +
            '  <div class="dual-list-panel">' +
            '    <div class="dual-list-title">' + (sourceLabel || 'ตัวแปรทั้งหมด') + '</div>' +
            '    <select class="dual-list-select" id="' + containerId + '-source" multiple>' +
            cols.map(function(c) { return '<option value="' + c + '">' + c + '</option>'; }).join('') +
            '    </select>' +
            '  </div>' +
            '  <div class="dual-list-buttons">' +
            '    <button type="button" onclick="dualListMove(\'' + containerId + '\', \'right\')" class="dual-btn">▶</button>' +
            '    <button type="button" onclick="dualListMove(\'' + containerId + '\', \'left\')" class="dual-btn">◀</button>' +
            '    <button type="button" onclick="dualListMove(\'' + containerId + '\', \'all-right\')" class="dual-btn">▶▶</button>' +
            '    <button type="button" onclick="dualListMove(\'' + containerId + '\', \'all-left\')" class="dual-btn">◀◀</button>' +
            '  </div>' +
            '  <div class="dual-list-panel">' +
            '    <div class="dual-list-title">' + (targetLabel || 'ตัวแปรที่เลือก') + '</div>' +
            '    <select class="dual-list-select" id="' + containerId + '-target" multiple></select>' +
            '  </div>' +
            '</div>';
    }

    function dualListMove(containerId, direction) {
        var source = document.getElementById(containerId + '-source');
        var target = document.getElementById(containerId + '-target');
        if (!source || !target) return;

        if (direction === 'right') {
            Array.from(source.selectedOptions).forEach(function(opt) { target.appendChild(opt); });
        } else if (direction === 'left') {
            Array.from(target.selectedOptions).forEach(function(opt) { source.appendChild(opt); });
        } else if (direction === 'all-right') {
            Array.from(source.options).forEach(function(opt) { target.appendChild(opt); });
        } else if (direction === 'all-left') {
            Array.from(target.options).forEach(function(opt) { source.appendChild(opt); });
        }
    }

    function getDualListSelected(containerId) {
        var target = document.getElementById(containerId + '-target');
        if (!target) return [];
        return Array.from(target.options).map(function(opt) { return opt.value; });
    }

    // =========================================================================
    // Assumption Check + AI Recommendation
    // =========================================================================

    function runAssumptionCheck(prefix, testType, data) {
        var el = document.getElementById(prefix + '-assumptions');
        if (!el) return null;

        var result = Stats.checkAssumptions(testType, data);
        if (!result || !result.checks || result.checks.length === 0) {
            el.style.display = 'none';
            return result;
        }

        var html = '<div class="assumption-panel">';
        html += '<h4>🔍 Assumption Check</h4>';
        html += '<table class="result-table"><thead><tr><th>Test</th><th>Result</th><th>Status</th><th>Detail</th></tr></thead><tbody>';
        result.checks.forEach(function(c) {
            var icon = c.passed ? '✅' : '⚠️';
            html += '<tr><td>' + escapeHtml(c.name) + '</td><td>' + escapeHtml(c.result) + '</td><td>' + icon + '</td><td>' + escapeHtml(c.detail) + '</td></tr>';
        });
        html += '</tbody></table>';

        if (result.recommendation) {
            var cls = result.passed ? 'assumption-pass' : 'assumption-warn';
            html += '<div class="' + cls + '">';
            html += '<strong>💡 คำแนะนำ:</strong> ' + escapeHtml(result.recommendation);
            html += '</div>';
        }
        html += '</div>';

        el.innerHTML = html;
        el.style.display = 'block';
        return result;
    }

    // =========================================================================
    // Initialize Dual-Lists for Page
    // =========================================================================

    function initDualListsForPage(pageName) {
        if (!state.data) return;
        // Map page names to their dual-list containers
        var dualLists = {
            'descriptive': [{id: 'desc-dual', filter: 'numeric', source: 'ตัวแปรทั้งหมด', target: 'ตัวแปรที่เลือกวิเคราะห์'}],
            'normality': [{id: 'norm-dual', filter: 'numeric', source: 'ตัวแปร Numeric', target: 'ตัวแปรที่เลือกทดสอบ'}],
            'correlation': [{id: 'cor-dual', filter: 'numeric', source: 'ตัวแปร Numeric', target: 'ตัวแปรที่เลือก (2+)'}],
            'reliability': [{id: 'rel-dual', filter: 'numeric', source: 'รายการข้อ', target: 'ข้อที่เลือกวิเคราะห์'}],
            'friedman': [{id: 'fr-dual', filter: 'numeric', source: 'ตัวแปร Numeric', target: 'ตัวแปรวัดซ้ำ (3+)'}],
            'linear-regression': [{id: 'lr-dual', filter: 'numeric', source: 'ตัวแปร Numeric', target: 'ตัวแปรอิสระ (IVs)'}],
            'logistic-regression': [{id: 'logr-dual', filter: 'numeric', source: 'ตัวแปร Numeric', target: 'ตัวแปรอิสระ (IVs)'}],
        };

        var lists = dualLists[pageName];
        if (lists) {
            lists.forEach(function(cfg) {
                var el = document.getElementById(cfg.id);
                if (el && el.innerHTML.trim() === '') {
                    createDualListPicker(cfg.id, cfg.source, cfg.target, cfg.filter);
                }
            });
        }
    }

    // =========================================================================
    // Utility: Reorder Columns (p-value before 95% CI)
    // =========================================================================

    function reorderColumns(cols) {
        var ordered = cols.slice();

        // Define desired relative order for known columns
        var orderMap = {
            'df': 1, 'Mean': 2, 'S.D.': 3, 'Mean Diff': 4, 'S.D. Diff': 5,
            't': 6, 'F': 6, 'U': 6, 'W': 6, 'Z': 6, 'H': 6, 'Chi-Square': 6,
            'p-value': 7, 'p': 7, 'p (2-tailed)': 7, 'Sig.': 8,
            '95% CI': 9, 'CI 95%': 9,
            "Cohen's d": 10, 'Effect': 11, 'r': 10, 'eta2': 10,
        };

        // Sort columns that appear in orderMap, keep others in original position
        var knownCols = [];
        var knownPositions = [];

        ordered.forEach(function(col, idx) {
            if (orderMap[col] !== undefined) {
                knownCols.push({ col: col, order: orderMap[col], origIdx: idx });
                knownPositions.push(idx);
            }
        });

        if (knownCols.length > 1) {
            knownCols.sort(function(a, b) { return a.order - b.order; });
            knownPositions.sort(function(a, b) { return a - b; });
            knownCols.forEach(function(item, i) {
                ordered[knownPositions[i]] = item.col;
            });
        }

        return ordered;
    }

    // =========================================================================
    // Utility: Build HTML Table
    // =========================================================================

    function buildTable(data, columns) {
        if (!data || data.length === 0) return '<p>No data to display.</p>';
        var explicitCols = !!columns;
        if (!columns) columns = Object.keys(data[0]);
        if (!explicitCols) columns = reorderColumns(columns);

        var html = '<table class="result-table"><thead><tr>';
        columns.forEach(function (col) {
            html += '<th>' + escapeHtml(String(col)) + '</th>';
        });
        html += '</tr></thead><tbody>';

        data.forEach(function (row) {
            html += '<tr>';
            columns.forEach(function (col) {
                var val = row[col];
                if (val === null || val === undefined || (typeof val === 'number' && isNaN(val))) {
                    html += '<td></td>';
                } else {
                    html += '<td>' + escapeHtml(String(val)) + '</td>';
                }
            });
            html += '</tr>';
        });

        html += '</tbody></table>';
        return html;
    }

    function escapeHtml(str) {
        return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                  .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
    }

    // =========================================================================
    // Utility: Get Selected Values
    // =========================================================================

    function getSelected(selectId) {
        var sel = document.getElementById(selectId);
        if (!sel) return [];
        var result = [];
        for (var i = 0; i < sel.options.length; i++) {
            if (sel.options[i].selected && sel.options[i].value) {
                result.push(sel.options[i].value);
            }
        }
        return result;
    }

    function getSelectValue(selectId) {
        var sel = document.getElementById(selectId);
        return sel ? sel.value : '';
    }

    // =========================================================================
    // Utility: Get Column Data
    // =========================================================================

    function getColumnData(colName, numericOnly) {
        if (!state.data || !colName) return [];
        var values = [];
        state.data.forEach(function (row) {
            var v = row[colName];
            if (numericOnly) {
                var num = parseFloat(v);
                if (!isNaN(num) && isFinite(num)) values.push(num);
            } else {
                values.push(v);
            }
        });
        return values;
    }

    // =========================================================================
    // Utility: Get Checked Checkboxes
    // =========================================================================

    function getChecked(name) {
        var boxes = document.querySelectorAll('input[name="' + name + '"]:checked');
        return Array.prototype.map.call(boxes, function (cb) { return cb.value; });
    }

    // =========================================================================
    // Utility: Split Data by Group
    // =========================================================================

    function splitByGroup(dvCol, ivCol) {
        if (!state.data) return { groups: {}, groupNames: [] };
        var groups = {};
        state.data.forEach(function (row) {
            var gVal = String(row[ivCol]);
            var dVal = parseFloat(row[dvCol]);
            if (isNaN(dVal)) return;
            if (!groups[gVal]) groups[gVal] = [];
            groups[gVal].push(dVal);
        });
        var groupNames = Object.keys(groups);
        return { groups: groups, groupNames: groupNames };
    }

    // =========================================================================
    // Display Results
    // =========================================================================

    function displayResults(prefix) {
        var result = state.results[prefix];
        var resultsArea = document.getElementById(prefix + '-results');

        if (!result) {
            if (resultsArea) resultsArea.classList.remove('visible');
            return;
        }

        if (resultsArea) {
            resultsArea.classList.add('visible', 'fade-in');
        }

        // Render extras
        var extrasEl = document.getElementById(prefix + '-extras');
        if (extrasEl && result.extras && result.extras.length > 0) {
            var extrasHtml = '';
            result.extras.forEach(function (extra) {
                extrasHtml += '<div class="card"><h4>' + escapeHtml(extra.title) + '</h4>';
                if (extra.html) {
                    extrasHtml += extra.html;
                } else if (extra.data && extra.data.length > 0) {
                    extrasHtml += buildTable(extra.data);
                }
                extrasHtml += '</div>';
            });
            extrasEl.innerHTML = extrasHtml;
        } else if (extrasEl) {
            extrasEl.innerHTML = '';
        }

        // Render main table
        var tableEl = document.getElementById(prefix + '-table');
        if (tableEl) {
            var mainHtml = '';
            if (result.title) {
                mainHtml += '<h4>' + escapeHtml(result.title) + '</h4>';
            }
            if (result.html) {
                mainHtml += result.html;
            } else if (result.data && result.data.length > 0) {
                mainHtml += buildTable(result.data, result.columns || null);
            }
            tableEl.innerHTML = mainHtml;
        }

        // Restore AI result if present
        var aiEl = document.getElementById(prefix + '-ai-result');
        if (aiEl && state.aiResults[prefix]) {
            aiEl.innerHTML = state.aiResults[prefix];
            aiEl.style.display = '';
        }
    }

    // =========================================================================
    // Analysis Runner
    // =========================================================================

    function runAnalysis(type) {
        var noDataTypes = ['effect-', 'assumption-', 'sample-size', 'power-analysis'];
        var needsData = !noDataTypes.some(function(t) { return type.indexOf(t) === 0 || type === t; });
        if (needsData && !state.data) {
            alert('Please upload a data file first.');
            return;
        }

        try {
            switch (type) {
                case 'demographics': runDemographics(); break;
                case 'descriptive': runDescriptive(); break;
                case 'numeric': runNumeric(); break;
                case 'nominal': runNominal(); break;
                case 'likert': runLikert(); break;
                case 'interval': runInterval(); break;
                case 'outlier': runOutlier(); break;
                case 'normality': runNormality(); break;
                case 'independent-ttest': runIndependentTTest(); break;
                case 'paired-ttest': runPairedTTest(); break;
                case 'oneway-anova': runOnewayAnova(); break;
                case 'twoway-anova': runTwowayAnova(); break;
                case 'rm-anova': runRmAnova(); break;
                case 'ancova': runAncova(); break;
                case 'mann-whitney': runMannWhitney(); break;
                case 'wilcoxon': runWilcoxon(); break;
                case 'kruskal-wallis': runKruskalWallis(); break;
                case 'friedman': runFriedman(); break;
                case 'chi-square': runChiSquare(); break;
                case 'correlation': runCorrelation(); break;
                case 'linear-regression': runLinearRegression(); break;
                case 'logistic-regression': runLogisticRegression(); break;
                case 'assumption-normality': runAssumptionNormality(); break;
                case 'assumption-levene': runAssumptionLevene(); break;
                case 'assumption-vif': runAssumptionVif(); break;
                case 'reliability': runReliability(); break;
                case 'effect-cohen': runEffectCohen(); break;
                case 'effect-odds': runEffectOdds(); break;
                case 'factor-analysis': runFactorAnalysis(); break;
                case 'crosstab': runCrossTab(); break;
                case 'posthoc': runPostHoc(); break;
                case 'bootstrap': runBootstrap(); break;
                case 'zscore': runZScore(); break;
                case 'multicollinearity': runVIF(); break;
                case 'charts': runCharts(); break;
                case 'partial-corr': runPartialCorr(); break;
                case 'hierarchical-reg': runHierarchicalReg(); break;
                case 'roc': runROC(); break;
                case 'icc': runICC(); break;
                case 'split-half': runSplitHalf(); break;
                case 'mcnemar': runMcNemar(); break;
                case 'fisher-exact': runFisherExact(); break;
                case 'cochran-q': runCochranQ(); break;
                case 'power-analysis': runPowerAnalysis(); break;
                case 'welch-anova': runWelchAnova(); break;
                case 'dunnett': runDunnett(); break;
                case 'median-test': runMedianTest(); break;
                case 'runs-test': runRunsTest(); break;
                case 'ks2': runKS2(); break;
                case 'cluster': runCluster(); break;
                case 'discriminant': runDiscriminant(); break;
                case 'missing': runMissing(); break;
                case 'qqplot': runQQPlot(); break;
                case 'survival': runSurvival(); break;
                case 'time-series': runTimeSeries(); break;
                case 'sample-size': runSampleSize(); break;
                case 'one-sample-ttest': runOneSampleTTest(); break;
                case 'sign-test': runSignTest(); break;
                case 'binomial-test': runBinomialTest(); break;
                case 'homogeneity': runHomogeneity(); break;
                case 'manova': runManova(); break;
                case 'games-howell': runGamesHowell(); break;
                case 'multiple-regression': runMultipleRegression(); break;
                case 'kappa': runKappa(); break;
                case 'cfa': runCFA(); break;
                case 'item-analysis': runItemAnalysis(); break;
                case 'path-analysis': runPathAnalysis(); break;
                default:
                    alert('Analysis type "' + type + '" is not yet implemented.');
            }
        } catch (err) {
            alert('Error running analysis: ' + err.message);
            console.error(err);
        }
    }

    // =========================================================================
    // Demographics (ข้อมูลทั่วไป / คุณลักษณะส่วนบุคคล)
    // =========================================================================

    function runDemographics() {
        var vars = getCheckedVars('demo-picker');
        if (vars.length === 0) { alert('กรุณาเลือกตัวแปรอย่างน้อย 1 ตัว'); return; }
        var numericVars = getCheckedVars('demo-numeric-picker');
        var tableTitle = (document.getElementById('demo-table-title') || {}).value || 'คุณลักษณะส่วนบุคคลของกลุ่มตัวอย่าง';
        var totalN = state.data.length;

        // Build the demographics table: Variable | Category | n | %
        var rows = [];
        var rowNum = 1;

        vars.forEach(function(varName) {
            var isNumeric = numericVars.indexOf(varName) !== -1;

            if (isNumeric) {
                var values = getColumnData(varName, true);
                if (values.length === 0) return;
                var desc = Stats.descriptive(values);

                // Check if custom intervals defined
                var intervals = demoIntervalConfigs[varName];
                if (intervals && intervals.length > 1) {
                    // Group by intervals
                    var isFirst = true;
                    for (var b = 0; b < intervals.length - 1; b++) {
                        var lo = intervals[b], hi = intervals[b + 1];
                        var count = 0;
                        values.forEach(function(v) {
                            if (b < intervals.length - 2) {
                                if (v >= lo && v < hi) count++;
                            } else {
                                if (v >= lo && v <= hi) count++;
                            }
                        });
                        var pct = values.length > 0 ? (count / values.length * 100) : 0;
                        var isInt = Number.isInteger(lo) && Number.isInteger(hi);
                        var label;
                        if (b < intervals.length - 2) {
                            label = isInt ? (lo + ' - ' + (hi - 1)) : (lo + ' - ' + (hi - 0.01).toFixed(2));
                        } else {
                            label = isInt ? (lo + ' - ' + hi) : (lo + ' - ' + hi);
                        }
                        rows.push({
                            'ลำดับ': isFirst ? rowNum : '',
                            'คุณลักษณะ': isFirst ? varName : '',
                            'รายการ': label,
                            'จำนวน (n)': count,
                            'ร้อยละ': fmt(pct, 1),
                        });
                        isFirst = false;
                    }
                    rowNum++;
                } else {
                    // Default: show Mean, S.D.
                    rows.push({
                        'ลำดับ': rowNum++, 'คุณลักษณะ': varName,
                        'รายการ': 'Mean = ' + fmt(desc.mean, 2) + ', S.D. = ' + fmt(desc.sd, 2),
                        'จำนวน (n)': desc.n, 'ร้อยละ': '-',
                    });
                    rows.push({
                        'ลำดับ': '', 'คุณลักษณะ': '',
                        'รายการ': 'Min = ' + fmt(desc.min, 2) + ', Max = ' + fmt(desc.max, 2),
                        'จำนวน (n)': '', 'ร้อยละ': '',
                    });
                }
            } else {
                // Categorical variable: frequency table
                var valueCounts = {};
                var total = 0;
                state.data.forEach(function(row) {
                    var val = row[varName];
                    if (val === null || val === undefined || val === '') return;
                    var key = String(val);
                    valueCounts[key] = (valueCounts[key] || 0) + 1;
                    total++;
                });

                // Sort by frequency descending
                var sorted = Object.keys(valueCounts).sort(function(a, b) {
                    return valueCounts[b] - valueCounts[a];
                });

                var isFirst = true;
                sorted.forEach(function(cat) {
                    var count = valueCounts[cat];
                    var pct = total > 0 ? (count / total * 100) : 0;
                    rows.push({
                        'ลำดับ': isFirst ? rowNum : '',
                        'คุณลักษณะ': isFirst ? varName : '',
                        'รายการ': cat,
                        'จำนวน (n)': count,
                        'ร้อยละ': fmt(pct, 1),
                    });
                    isFirst = false;
                });
                rowNum++;
            }
        });

        // Add total row
        rows.push({
            'ลำดับ': '',
            'คุณลักษณะ': '',
            'รายการ': 'รวม (Total)',
            'จำนวน (n)': totalN,
            'ร้อยละ': '100.0',
        });

        // Explicit column order
        var columns = ['ลำดับ', 'คุณลักษณะ', 'รายการ', 'จำนวน (n)', 'ร้อยละ'];

        state.results['demo'] = {
            data: rows,
            columns: columns,
            title: tableTitle + ' (N = ' + totalN + ')',
            extras: []
        };
        displayResults('demo');
    }

    // =========================================================================
    // Descriptive Analysis
    // =========================================================================

    function runDescriptive() {
        var vars = getCheckedVars('desc-picker');
        if (vars.length === 0) vars = getDualListSelected('desc-dual');
        if (!vars || vars.length === 0) vars = getSelected('desc-vars');
        if (vars.length === 0) { alert('Please select at least one variable.'); return; }
        var stats = getChecked('desc-stat');
        if (stats.length === 0) { alert('Please select at least one statistic.'); return; }

        var statLabelMap = {
            count: 'N', mean: 'Mean', se: 'S.E.', sd: 'S.D.',
            min: 'Min', max: 'Max', skewness: 'Skewness', kurtosis: 'Kurtosis'
        };

        var rows = [];
        vars.forEach(function (v) {
            var values = getColumnData(v, true);
            var d = Stats.descriptive(values);
            if (!d) return;
            var row = { Variable: v };
            stats.forEach(function (s) {
                var label = statLabelMap[s] || s;
                switch (s) {
                    case 'count': row[label] = d.n; break;
                    case 'mean': row[label] = fmt(d.mean); break;
                    case 'se': row[label] = fmt(d.se); break;
                    case 'sd': row[label] = fmt(d.sd); break;
                    case 'min': row[label] = fmt(d.min, 2); break;
                    case 'max': row[label] = fmt(d.max, 2); break;
                    case 'skewness': row[label] = fmt(d.skewness); break;
                    case 'kurtosis': row[label] = fmt(d.kurtosis); break;
                }
            });
            rows.push(row);
        });

        state.results['desc'] = {
            data: rows,
            title: 'Descriptive Statistics',
            extras: []
        };
        displayResults('desc');
    }

    // =========================================================================
    // Numeric Variable Analysis
    // =========================================================================

    function runNumeric() {
        var vars = getCheckedVars('num-picker');
        if (vars.length === 0) vars = getSelected('num-vars');
        if (vars.length === 0) { alert('Please select at least one variable.'); return; }

        var rows = [];
        vars.forEach(function (v) {
            var values = getColumnData(v, true);
            var d = Stats.descriptive(values);
            if (!d) return;
            rows.push({
                Variable: v, N: d.n, Mean: fmt(d.mean), 'S.E.': fmt(d.se), 'S.D.': fmt(d.sd),
                Min: fmt(d.min, 2), P25: fmt(d.p25), Median: fmt(d.median),
                P75: fmt(d.p75), Max: fmt(d.max, 2),
                '95% CI': Stats.formatCI ? Stats.formatCI(d.ci95_lo, d.ci95_hi) : '[' + fmt(d.ci95_lo) + ', ' + fmt(d.ci95_hi) + ']',
                Skewness: fmt(d.skewness), Kurtosis: fmt(d.kurtosis)
            });
        });

        state.results['num'] = { data: rows, title: 'Numeric Variable Summary', extras: [] };
        displayResults('num');
    }

    // =========================================================================
    // Nominal Variable Analysis
    // =========================================================================

    function runNominal() {
        var vars = getCheckedVars('nom-picker');
        if (vars.length === 0) vars = getSelected('nom-vars');
        if (vars.length === 0) { alert('Please select at least one variable.'); return; }

        var extras = [];
        vars.forEach(function (v) {
            var values = getColumnData(v, false);
            var freq = {};
            var total = 0;
            values.forEach(function (val) {
                if (val === null || val === undefined || val === '') return;
                var key = String(val);
                freq[key] = (freq[key] || 0) + 1;
                total++;
            });

            var sorted = Object.keys(freq).sort(function (a, b) { return freq[b] - freq[a]; });
            var rows = [];
            var cumPct = 0;
            sorted.forEach(function (key) {
                var pct = total > 0 ? (freq[key] / total * 100) : 0;
                cumPct += pct;
                rows.push({
                    Value: key,
                    Frequency: freq[key],
                    'Percent (%)': fmt(pct, 2),
                    'Cumulative (%)': fmt(cumPct, 2)
                });
            });
            rows.push({ Value: 'Total', Frequency: total, 'Percent (%)': '100.00', 'Cumulative (%)': '' });

            extras.push({ title: 'Frequency Table: ' + v, data: rows });
        });

        // Use first variable's data as main if only one, otherwise empty main
        var mainData = extras.length === 1 ? extras[0].data : [];
        var mainTitle = extras.length === 1 ? extras[0].title : 'Nominal Variable Analysis';
        if (extras.length === 1) extras = [];

        state.results['nom'] = { data: mainData, title: mainTitle, extras: extras };
        displayResults('nom');
    }

    // =========================================================================
    // Likert Scale Analysis
    // =========================================================================

    function getLikertCriteria() {
        var container = document.getElementById('lk-criteria-rows');
        if (!container) return null;
        var rows = container.querySelectorAll('.criteria-row');
        var criteria = [];
        rows.forEach(function(row) {
            var lo = parseFloat(row.querySelector('.criteria-lo').value);
            var hi = parseFloat(row.querySelector('.criteria-hi').value);
            var label = row.querySelector('.criteria-label').value.trim();
            if (!isNaN(lo) && !isNaN(hi) && label) {
                criteria.push({ lo: lo, hi: hi, label: label });
            }
        });
        return criteria.length > 0 ? criteria : null;
    }

    function updateLikertCriteria() {
        var scale = parseInt(getSelectValue('lk-scale')) || 5;
        var container = document.getElementById('lk-criteria-rows');
        if (!container) return;

        var defaults5 = [
            { lo: 4.21, hi: 5.00, label: 'มากที่สุด' },
            { lo: 3.41, hi: 4.20, label: 'มาก' },
            { lo: 2.61, hi: 3.40, label: 'ปานกลาง' },
            { lo: 1.81, hi: 2.60, label: 'น้อย' },
            { lo: 1.00, hi: 1.80, label: 'น้อยที่สุด' }
        ];
        var defaults3 = [
            { lo: 2.34, hi: 3.00, label: 'มาก' },
            { lo: 1.67, hi: 2.33, label: 'ปานกลาง' },
            { lo: 1.00, hi: 1.66, label: 'น้อย' }
        ];
        var defaults = scale === 3 ? defaults3 : defaults5;

        var html = '';
        defaults.forEach(function(d) {
            html += '<div class="criteria-row">' +
                '<input type="number" class="criteria-lo" value="' + d.lo + '" step="0.01">' +
                '<span>—</span>' +
                '<input type="number" class="criteria-hi" value="' + d.hi + '" step="0.01">' +
                '<input type="text" class="criteria-label" value="' + d.label + '" placeholder="ชื่อระดับ">' +
                '</div>';
        });
        container.innerHTML = html;
    }

    function runLikert() {
        var vars = getCheckedVars('lk-picker');
        if (vars.length === 0) vars = getSelected('lk-vars');
        if (vars.length === 0) { alert('กรุณาเลือกตัวแปรอย่างน้อย 1 ตัว'); return; }
        var scale = parseInt(getSelectValue('lk-scale')) || 5;

        // Read custom criteria from UI
        var criteria = getLikertCriteria();

        var dataArrays = vars.map(function (v) { return getColumnData(v, true); });
        var result = Stats.likertAnalysis(dataArrays, vars, scale, criteria);
        if (!result) { alert('ไม่สามารถวิเคราะห์ได้ ตรวจสอบข้อมูล'); return; }

        // Main table: No. → Variable → 5 → 4 → 3 → 2 → 1 → Mean → S.D. → Interpretation
        var rows = result.items.map(function (item) {
            var row = { 'No.': item.no, 'Variable': item.variable };
            for (var lv = scale; lv >= 1; lv--) {
                var f = item.frequencies[lv];
                row[String(lv)] = f ? f.count + ' (' + fmt(f.pct, 1) + '%)' : '0';
            }
            row['Mean'] = fmt(item.mean);
            row['S.D.'] = fmt(item.sd);
            row['Interpretation'] = item.interpretation;
            return row;
        });

        // Overall row
        var overallRow = { 'No.': '', 'Variable': 'รวม (Overall)' };
        for (var lv = scale; lv >= 1; lv--) overallRow[String(lv)] = '';
        overallRow['Mean'] = fmt(result.overall.mean);
        overallRow['S.D.'] = fmt(result.overall.sd);
        overallRow['Interpretation'] = result.overall.interpretation;
        rows.push(overallRow);

        // Criteria reference table
        var usedCriteria = criteria || [];
        if (usedCriteria.length === 0 && result.items.length > 0) {
            // Get defaults from Stats
            if (scale === 5) {
                usedCriteria = [{lo:4.21,hi:5.00,label:'มากที่สุด'},{lo:3.41,hi:4.20,label:'มาก'},{lo:2.61,hi:3.40,label:'ปานกลาง'},{lo:1.81,hi:2.60,label:'น้อย'},{lo:1.00,hi:1.80,label:'น้อยที่สุด'}];
            } else {
                usedCriteria = [{lo:2.34,hi:3.00,label:'มาก'},{lo:1.67,hi:2.33,label:'ปานกลาง'},{lo:1.00,hi:1.66,label:'น้อย'}];
            }
        }
        var criteriaRows = usedCriteria.map(function(c) {
            return { 'ช่วงค่าเฉลี่ย': fmt(c.lo, 2) + ' - ' + fmt(c.hi, 2), 'ระดับการแปลผล': c.label };
        });

        // Ranking table
        var rankRows = result.ranking.map(function (r) {
            return {
                'อันดับ': r.rank, 'Variable': r.variable,
                'Mean': fmt(r.mean), 'S.D.': fmt(r.sd),
                'Interpretation': r.interpretation
            };
        });

        // Explicit column order to prevent JS numeric key sorting
        var likertColumns = ['No.', 'Variable'];
        for (var lv = scale; lv >= 1; lv--) likertColumns.push(String(lv));
        likertColumns.push('Mean', 'S.D.', 'Interpretation');

        state.results['lk'] = {
            data: rows,
            columns: likertColumns,
            title: 'Likert Scale Analysis (' + scale + '-point)',
            extras: [
                { title: 'เกณฑ์การแปลผล (Interpretation Criteria)', data: criteriaRows },
                { title: 'การจัดลำดับ (Ranking)', data: rankRows }
            ]
        };
        displayResults('lk');
    }

    // =========================================================================
    // Interval Analysis
    // =========================================================================

    // Interval config storage: { varName: { mode: 'auto'|'custom', bins: 5, breaks: [0,18,30,40,50] } }
    var intervalConfigs = {};
    var demoIntervalConfigs = {};

    function openIntervalConfig() {
        var vars = getCheckedVars('intv-picker');
        if (vars.length === 0) { alert('กรุณาเลือกตัวแปรก่อน'); return; }

        var body = document.getElementById('intv-modal-body');
        var html = '';

        vars.forEach(function(varName) {
            var values = getColumnData(varName, true);
            var min = values.length > 0 ? Math.min.apply(null, values) : 0;
            var max = values.length > 0 ? Math.max.apply(null, values) : 0;
            var cfg = intervalConfigs[varName] || { mode: 'auto', bins: 5, breaks: '' };

            html += '<div class="intv-var-config" data-var="' + varName + '">';
            html += '<h4>📊 ' + varName + '</h4>';
            html += '<div class="var-stats">N = ' + values.length + ' | Min = ' + fmt(min, 2) + ' | Max = ' + fmt(max, 2) + '</div>';

            html += '<div class="intv-mode-row">';
            html += '<label><input type="radio" name="intv-mode-' + varName + '" value="auto" ' + (cfg.mode === 'auto' ? 'checked' : '') + ' onchange="toggleIntvMode(this)"> แบ่งอัตโนมัติ จำนวน ';
            html += '<input type="number" class="intv-bins-input" value="' + cfg.bins + '" min="2" max="50"> ช่วง</label>';
            html += '</div>';

            html += '<div class="intv-mode-row">';
            html += '<label><input type="radio" name="intv-mode-' + varName + '" value="custom" ' + (cfg.mode === 'custom' ? 'checked' : '') + ' onchange="toggleIntvMode(this)"> กำหนดจุดแบ่งเอง</label>';
            html += '</div>';

            var defaultBreaks = cfg.breaks || (fmt(min, 0) + ',' + fmt((min + max) / 2, 0) + ',' + fmt(max, 0));
            html += '<div class="intv-custom-area" style="' + (cfg.mode === 'custom' ? '' : 'display:none') + '">';
            html += '<input type="text" class="intv-custom-breaks" value="' + defaultBreaks + '" placeholder="เช่น 0,18,25,35,45,60">';
            html += '<div style="font-size:0.75rem;color:#64748b;margin-top:4px">ใส่ตัวเลขคั่นด้วยจุลภาค เช่น 0,18,25,35,45,60</div>';
            html += '</div>';

            html += '</div>';
        });

        body.innerHTML = html;
        document.getElementById('intv-modal').style.display = 'flex';
    }

    function toggleIntvMode(radio) {
        var configCard = radio.closest('.intv-var-config');
        var customArea = configCard.querySelector('.intv-custom-area');
        if (radio.value === 'custom') {
            customArea.style.display = '';
        } else {
            customArea.style.display = 'none';
        }
    }

    function applyIntervalConfig() {
        var cards = document.querySelectorAll('.intv-var-config');
        intervalConfigs = {};

        var previewHtml = '<div class="intv-preview-pills">';

        cards.forEach(function(card) {
            var varName = card.getAttribute('data-var');
            var modeRadio = card.querySelector('input[type="radio"]:checked');
            var mode = modeRadio ? modeRadio.value : 'auto';
            var binsInput = card.querySelector('.intv-bins-input');
            var bins = binsInput ? parseInt(binsInput.value) || 5 : 5;
            var breaksInput = card.querySelector('.intv-custom-breaks');
            var breaks = breaksInput ? breaksInput.value.trim() : '';

            intervalConfigs[varName] = { mode: mode, bins: bins, breaks: breaks };

            if (mode === 'custom' && breaks) {
                previewHtml += '<span class="intv-preview-pill">' + varName + ': ' + breaks + '</span>';
            } else {
                previewHtml += '<span class="intv-preview-pill">' + varName + ': ' + bins + ' ช่วง</span>';
            }
        });

        previewHtml += '</div>';
        var previewEl = document.getElementById('intv-config-preview');
        if (previewEl) previewEl.innerHTML = previewHtml;

        closeIntervalConfig();
    }

    function closeIntervalConfig() {
        document.getElementById('intv-modal').style.display = 'none';
    }

    // =========================================================================
    // Demographics Interval Config
    // =========================================================================

    function openDemoIntervalConfig() {
        var numVars = getCheckedVars('demo-numeric-picker');
        if (numVars.length === 0) { alert('กรุณาเลือกตัวแปรเชิงปริมาณก่อน'); return; }

        var body = document.getElementById('demo-interval-body');
        var html = '';
        numVars.forEach(function(varName) {
            var values = getColumnData(varName, true);
            var min = values.length > 0 ? Math.min.apply(null, values) : 0;
            var max = values.length > 0 ? Math.max.apply(null, values) : 100;
            var cfg = demoIntervalConfigs[varName];
            var defaultBreaks = cfg ? cfg.join(',') : '';
            if (!defaultBreaks) {
                // Auto suggest nice breaks
                var step = Math.ceil((max - min) / 5);
                var start = Math.floor(min / step) * step;
                var breaks = [];
                for (var v = start; v <= max + step; v += step) breaks.push(v);
                defaultBreaks = breaks.join(',');
            }
            html += '<div class="intv-var-config" data-var="' + varName + '">';
            html += '<h4>📊 ' + varName + '</h4>';
            html += '<div class="var-stats">N = ' + values.length + ' | Min = ' + fmt(min, 1) + ' | Max = ' + fmt(max, 1) + '</div>';
            html += '<div style="margin-top:8px"><label style="font-size:0.82rem;font-weight:600;color:#475569">จุดแบ่ง (คั่นด้วยจุลภาค เช่น 20,30,40,50,60)</label>';
            html += '<input type="text" class="intv-custom-breaks" value="' + defaultBreaks + '" style="width:100%;margin-top:4px">';
            html += '<div style="font-size:0.75rem;color:#64748b;margin-top:4px">ตัวอย่าง: ถ้าใส่ 20,30,40,50,60 จะได้ช่วง 20-29, 30-39, 40-49, 50-60</div>';
            html += '</div></div>';
        });
        body.innerHTML = html;
        document.getElementById('demo-interval-modal').style.display = 'flex';
    }

    function closeDemoIntervalModal() {
        document.getElementById('demo-interval-modal').style.display = 'none';
    }

    function applyDemoIntervals() {
        var cards = document.querySelectorAll('#demo-interval-body .intv-var-config');
        demoIntervalConfigs = {};
        var previewHtml = '<div class="intv-preview-pills">';
        cards.forEach(function(card) {
            var varName = card.getAttribute('data-var');
            var breaksInput = card.querySelector('.intv-custom-breaks');
            var breaks = breaksInput ? breaksInput.value.trim() : '';
            if (breaks) {
                var arr = breaks.split(',').map(function(s){return parseFloat(s.trim());}).filter(function(n){return !isNaN(n);});
                arr.sort(function(a,b){return a-b;});
                if (arr.length >= 2) {
                    demoIntervalConfigs[varName] = arr;
                    previewHtml += '<span class="intv-preview-pill">' + varName + ': ' + arr.join(', ') + '</span>';
                }
            }
        });
        previewHtml += '</div>';
        var preview = document.getElementById('demo-interval-preview');
        if (preview) preview.innerHTML = previewHtml;
        closeDemoIntervalModal();
    }

    function runInterval() {
        var vars = getCheckedVars('intv-picker');
        if (vars.length === 0) {
            var singleVar = getSelectValue('intv-var');
            if (singleVar) vars = [singleVar];
        }
        if (vars.length === 0) { alert('กรุณาเลือกตัวแปร'); return; }

        // If no config set yet, use default 5 bins for all
        var allRows = [];
        vars.forEach(function(varName) {
            var values = getColumnData(varName, true);
            if (values.length === 0) return;

            var cfg = intervalConfigs[varName] || { mode: 'auto', bins: 5, breaks: '' };
            var breakpoints = [];

            if (cfg.mode === 'custom' && cfg.breaks) {
                // Parse custom breakpoints
                breakpoints = cfg.breaks.split(',').map(function(s) { return parseFloat(s.trim()); }).filter(function(n) { return !isNaN(n); });
                breakpoints.sort(function(a, b) { return a - b; });
            } else {
                // Auto: equal-width bins
                var min = Math.min.apply(null, values);
                var max = Math.max.apply(null, values);
                var binWidth = (max - min) / cfg.bins;
                if (binWidth === 0) binWidth = 1;
                for (var i = 0; i <= cfg.bins; i++) {
                    breakpoints.push(min + i * binWidth);
                }
                // Ensure last breakpoint covers max
                breakpoints[breakpoints.length - 1] = max;
            }

            if (breakpoints.length < 2) return;

            // Build bins from breakpoints
            var total = values.length;
            var cumFreq = 0;
            for (var b = 0; b < breakpoints.length - 1; b++) {
                var lo = breakpoints[b];
                var hi = breakpoints[b + 1];
                var count = 0;
                values.forEach(function(v) {
                    if (b === breakpoints.length - 2) {
                        if (v >= lo && v <= hi) count++;
                    } else {
                        if (v >= lo && v < hi) count++;
                    }
                });
                var pct = total > 0 ? (count / total * 100) : 0;
                cumFreq += count;
                // For auto mode with integers, show as 20-29, 30-39
                // For custom mode or decimals, show as 20.00-29.99, 30.00-39.99
                var isInt = Number.isInteger(lo) && Number.isInteger(hi);
                var loStr = isInt ? String(Math.round(lo)) : fmt(lo, 2);
                var hiStr;
                if (b < breakpoints.length - 2) {
                    // Not the last bin: hi is exclusive, so show hi-1 (int) or hi-0.01 (decimal)
                    hiStr = isInt ? String(Math.round(hi) - 1) : fmt(hi - 0.01, 2);
                } else {
                    // Last bin: hi is inclusive
                    hiStr = isInt ? String(Math.round(hi)) : fmt(hi, 2);
                }
                allRows.push({
                    Variable: varName,
                    Interval: loStr + ' - ' + hiStr,
                    Frequency: count,
                    '%': fmt(pct, 1),
                    'Cum Freq': cumFreq,
                    'Cum %': fmt(cumFreq / total * 100, 1)
                });
            }
        });

        state.results['intv'] = { data: allRows, title: 'Interval Analysis' };
        displayResults('intv');
    }

    // =========================================================================
    // Outlier Detection
    // =========================================================================

    function runOutlier() {
        var vars = getCheckedVars('out-picker');
        if (vars.length === 0) vars = getSelected('out-vars');
        if (vars.length === 0) { alert('Please select at least one variable.'); return; }

        var rows = [];
        vars.forEach(function (v) {
            var values = getColumnData(v, true);
            if (values.length === 0) return;
            var sorted = values.slice().sort(function (a, b) { return a - b; });
            var n = sorted.length;
            var q1Idx = (n - 1) * 0.25;
            var q3Idx = (n - 1) * 0.75;
            var q1 = sorted[Math.floor(q1Idx)] + (sorted[Math.ceil(q1Idx)] - sorted[Math.floor(q1Idx)]) * (q1Idx % 1);
            var q3 = sorted[Math.floor(q3Idx)] + (sorted[Math.ceil(q3Idx)] - sorted[Math.floor(q3Idx)]) * (q3Idx % 1);
            var iqr = q3 - q1;
            var lowerBound = q1 - 1.5 * iqr;
            var upperBound = q3 + 1.5 * iqr;

            var outliers = values.filter(function (val) { return val < lowerBound || val > upperBound; });

            rows.push({
                Variable: v,
                N: values.length,
                Q1: fmt(q1),
                Q3: fmt(q3),
                IQR: fmt(iqr),
                'Lower Bound': fmt(lowerBound),
                'Upper Bound': fmt(upperBound),
                'Outlier Count': outliers.length,
                'Outlier Values': outliers.length > 0 ? outliers.slice(0, 10).map(function (o) { return fmt(o, 2); }).join(', ') + (outliers.length > 10 ? '...' : '') : 'None'
            });
        });

        state.results['out'] = { data: rows, title: 'Outlier Detection (IQR Method)', extras: [] };
        displayResults('out');
    }

    // =========================================================================
    // Normality Test
    // =========================================================================

    function runNormality() {
        var vars = getCheckedVars('norm-picker');
        if (vars.length === 0) vars = getDualListSelected('norm-dual');
        if (!vars || vars.length === 0) vars = getSelected('norm-vars');
        if (vars.length === 0) { alert('Please select at least one variable.'); return; }
        var alpha = parseFloat(getSelectValue('norm-alpha')) || 0.05;

        var rows = [];
        vars.forEach(function (v) {
            var values = getColumnData(v, true);
            if (values.length < 3) return;
            var sw = Stats.shapiroWilk(values);
            var ks = Stats.ksTest(values);
            var row = { Variable: v, N: values.length };
            row['Shapiro-Wilk W'] = sw ? fmt(sw.W) : 'N/A';
            row['S-W p-value'] = sw ? Stats.formatPValue(sw.p) : 'N/A';
            row['K-S D'] = ks ? fmt(ks.D) : 'N/A';
            row['K-S p-value'] = ks ? Stats.formatPValue(ks.p) : 'N/A';

            var swNormal = sw ? (sw.p >= alpha) : null;
            var ksNormal = ks ? (ks.p >= alpha) : null;
            var conclusion = 'N/A';
            if (swNormal !== null && ksNormal !== null) {
                conclusion = (swNormal && ksNormal) ? 'Normal' : 'Not Normal';
            } else if (swNormal !== null) {
                conclusion = swNormal ? 'Normal' : 'Not Normal';
            } else if (ksNormal !== null) {
                conclusion = ksNormal ? 'Normal' : 'Not Normal';
            }
            row['Conclusion (alpha=' + alpha + ')'] = conclusion;
            rows.push(row);
        });

        state.results['norm'] = { data: rows, title: 'Normality Test Results', extras: [] };
        displayResults('norm');
    }

    // =========================================================================
    // Independent Samples t-Test
    // =========================================================================

    function runIndependentTTest() {
        var dvList = getCheckedVars('itt-dv-picker');
        if (dvList.length === 0) {
            var singleDv = getSelectValue('itt-dv');
            if (singleDv) dvList = [singleDv];
        }
        var iv = getPickerValue('itt-iv-picker', 'itt-iv');
        if (dvList.length === 0 || !iv) { alert('Please select DV and IV.'); return; }

        var opts = getChecked('itt-opt');
        var alpha = parseFloat(getSelectValue('itt-alpha')) || 0.05;

        var allMainRows = [];
        var allExtras = [];

        dvList.forEach(function(dv) {
            var split = splitByGroup(dv, iv);
            var gNames = split.groupNames;
            if (gNames.length !== 2) {
                alert('Independent t-test requires exactly 2 groups. Found ' + gNames.length + ' group(s) in "' + iv + '" for DV "' + dv + '".');
                return;
            }

            var g1 = split.groups[gNames[0]];
            var g2 = split.groups[gNames[1]];

            // Run assumption check (only for first/single DV)
            if (dvList.length === 1) {
                runAssumptionCheck('itt', 'independent-ttest', { group1: g1, group2: g2 });
            }

            var result = Stats.independentTTest(g1, g2);
            if (!result) { alert('Could not compute t-test for "' + dv + '". Check your data.'); return; }

            // Build comprehensive result table with ALL info in one row
            var extras = [];

            // Descriptive table (keep as extra for detail)
            if (opts.indexOf('descriptive') !== -1) {
                extras.push({
                    title: 'Group Descriptive Statistics' + (dvList.length > 1 ? ' (' + dv + ')' : ''),
                    data: [
                        { Group: gNames[0], N: result.desc1.n, Mean: fmt(result.desc1.mean), 'S.D.': fmt(result.desc1.sd), 'S.E.': fmt(result.desc1.se) },
                        { Group: gNames[1], N: result.desc2.n, Mean: fmt(result.desc2.mean), 'S.D.': fmt(result.desc2.sd), 'S.E.': fmt(result.desc2.se) }
                    ]
                });
            }

            // Levene's test (keep as extra)
            if (opts.indexOf('levene') !== -1) {
                extras.push({
                    title: "Levene's Test for Equality of Variances" + (dvList.length > 1 ? ' (' + dv + ')' : ''),
                    data: [{
                        'F': fmt(result.leveneF),
                        'p-value': Stats.formatPValue(result.leveneP),
                        'Conclusion': result.leveneP >= alpha ? 'Equal variances assumed' : 'Equal variances not assumed'
                    }]
                });
            }

            // Main comprehensive summary table
            var equalVarAssumed = result.leveneP >= alpha;
            var mainRows = [];

            if (opts.indexOf('both') !== -1) {
                mainRows.push({
                    'DV': dv, 'IV': iv, 'Method': 'Equal var assumed',
                    ['n(' + gNames[0] + ')']: result.desc1.n, ['M(' + gNames[0] + ')']: fmt(result.desc1.mean), ['SD(' + gNames[0] + ')']: fmt(result.desc1.sd),
                    ['n(' + gNames[1] + ')']: result.desc2.n, ['M(' + gNames[1] + ')']: fmt(result.desc2.mean), ['SD(' + gNames[1] + ')']: fmt(result.desc2.sd),
                    'df': fmt(result.df, 0), 'Mean Diff': fmt(result.meanDiff),
                    't': fmt(result.t), 'p-value': Stats.formatPValue(result.p),
                    '95% CI': Stats.formatCI ? Stats.formatCI(result.ci95_lo, result.ci95_hi) : '',
                    "Cohen's d": fmt(result.cohensD), 'Effect': result.effectSize || ''
                });
                mainRows.push({
                    'DV': dv, 'IV': iv, 'Method': "Welch's",
                    ['n(' + gNames[0] + ')']: result.desc1.n, ['M(' + gNames[0] + ')']: fmt(result.desc1.mean), ['SD(' + gNames[0] + ')']: fmt(result.desc1.sd),
                    ['n(' + gNames[1] + ')']: result.desc2.n, ['M(' + gNames[1] + ')']: fmt(result.desc2.mean), ['SD(' + gNames[1] + ')']: fmt(result.desc2.sd),
                    'df': fmt(result.welchDf, 2), 'Mean Diff': fmt(result.meanDiff),
                    't': fmt(result.welchT), 'p-value': Stats.formatPValue(result.welchP),
                    '95% CI': '',
                    "Cohen's d": fmt(result.cohensD), 'Effect': result.effectSize || ''
                });
            } else {
                var useT = equalVarAssumed ? result.t : result.welchT;
                var useP = equalVarAssumed ? result.p : result.welchP;
                var useDf = equalVarAssumed ? result.df : result.welchDf;
                var method = equalVarAssumed ? 'Equal var assumed' : "Welch's";
                mainRows.push({
                    'DV': dv, 'IV': iv, 'Method': method,
                    ['n(' + gNames[0] + ')']: result.desc1.n, ['M(' + gNames[0] + ')']: fmt(result.desc1.mean), ['SD(' + gNames[0] + ')']: fmt(result.desc1.sd),
                    ['n(' + gNames[1] + ')']: result.desc2.n, ['M(' + gNames[1] + ')']: fmt(result.desc2.mean), ['SD(' + gNames[1] + ')']: fmt(result.desc2.sd),
                    'df': fmt(useDf, equalVarAssumed ? 0 : 2), 'Mean Diff': fmt(result.meanDiff),
                    't': fmt(useT), 'p-value': Stats.formatPValue(useP),
                    '95% CI': Stats.formatCI ? Stats.formatCI(result.ci95_lo, result.ci95_hi) : '',
                    "Cohen's d": fmt(result.cohensD), 'Effect': result.effectSize || ''
                });
            }

            // Summary box
            var usedP = equalVarAssumed ? result.p : result.welchP;
            var sig = usedP < alpha;
            var summaryHtml = '<div class="detail-box">';
            summaryHtml += '<strong>\u0E2A\u0E23\u0E38\u0E1B:</strong> ' + gNames[0] + ' (M=' + fmt(result.desc1.mean) + ', SD=' + fmt(result.desc1.sd) + ') vs ' +
                           gNames[1] + ' (M=' + fmt(result.desc2.mean) + ', SD=' + fmt(result.desc2.sd) + ') \u2014 ';
            summaryHtml += (sig ? '\u0E41\u0E15\u0E01\u0E15\u0E48\u0E32\u0E07\u0E2D\u0E22\u0E48\u0E32\u0E07\u0E21\u0E35\u0E19\u0E31\u0E22\u0E2A\u0E33\u0E04\u0E31\u0E0D' : '\u0E44\u0E21\u0E48\u0E41\u0E15\u0E01\u0E15\u0E48\u0E32\u0E07') + ' (p=' + Stats.formatPValue(usedP) + ', d=' + fmt(result.cohensD) + ')';
            summaryHtml += '</div>';
            extras.push({ title: '', html: summaryHtml });

            allMainRows = allMainRows.concat(mainRows);
            allExtras = allExtras.concat(extras);
        });

        state.results['itt'] = { data: allMainRows, title: 'Independent Samples t-Test', extras: allExtras };
        displayResults('itt');
    }

    // =========================================================================
    // Paired Samples t-Test
    // =========================================================================

    function runPairedTTest() {
        var before = getPickerValue('ptt-before-picker', 'ptt-before');
        var after = getPickerValue('ptt-after-picker', 'ptt-after');
        if (!before || !after) { alert('Please select both Before and After variables.'); return; }

        var opts = getChecked('ptt-opt');
        var alpha = parseFloat(getSelectValue('ptt-alpha')) || 0.05;

        var bData = getColumnData(before, true);
        var aData = getColumnData(after, true);
        var minLen = Math.min(bData.length, aData.length);
        bData = bData.slice(0, minLen);
        aData = aData.slice(0, minLen);

        // Run assumption check
        runAssumptionCheck('ptt', 'paired-ttest', { before: bData, after: aData });

        var result = Stats.pairedTTest(bData, aData);
        if (!result) { alert('Could not compute paired t-test. Check your data.'); return; }

        var extras = [];

        if (opts.indexOf('descriptive') !== -1) {
            extras.push({
                title: 'Paired Samples Descriptive Statistics',
                data: [
                    { Variable: before, N: result.descBefore.n, Mean: fmt(result.descBefore.mean), 'S.D.': fmt(result.descBefore.sd), 'S.E.': fmt(result.descBefore.se) },
                    { Variable: after, N: result.descAfter.n, Mean: fmt(result.descAfter.mean), 'S.D.': fmt(result.descAfter.sd), 'S.E.': fmt(result.descAfter.se) }
                ]
            });
        }

        if (opts.indexOf('correlation') !== -1) {
            extras.push({
                title: 'Paired Samples Correlation',
                data: [{ Pair: before + ' & ' + after, r: fmt(result.r), 'p-value': Stats.formatPValue(result.rP) }]
            });
        }

        var mainRows = [];
        mainRows.push({
            'Before': before, 'After': after, 'N': result.descBefore.n,
            'M (Before)': fmt(result.descBefore.mean), 'SD (Before)': fmt(result.descBefore.sd),
            'M (After)': fmt(result.descAfter.mean), 'SD (After)': fmt(result.descAfter.sd),
            'df': fmt(result.df, 0), 'Mean Diff': fmt(result.meanDiff),
            't': fmt(result.t), 'p-value': Stats.formatPValue(result.p),
            '95% CI': result.ci95 || '',
            "Cohen's d": fmt(result.cohensD), 'Effect': result.effectSize || ''
        });

        // Summary box
        var sig = result.p < alpha;
        var summaryHtml = '<div class="detail-box">';
        summaryHtml += '<strong>\u0E2A\u0E23\u0E38\u0E1B:</strong> ' + before + ' (M=' + fmt(result.descBefore.mean) + ', SD=' + fmt(result.descBefore.sd) + ') vs ' +
                       after + ' (M=' + fmt(result.descAfter.mean) + ', SD=' + fmt(result.descAfter.sd) + ') \u2014 ';
        summaryHtml += (sig ? '\u0E41\u0E15\u0E01\u0E15\u0E48\u0E32\u0E07\u0E2D\u0E22\u0E48\u0E32\u0E07\u0E21\u0E35\u0E19\u0E31\u0E22\u0E2A\u0E33\u0E04\u0E31\u0E0D' : '\u0E44\u0E21\u0E48\u0E41\u0E15\u0E01\u0E15\u0E48\u0E32\u0E07') + ' (p=' + Stats.formatPValue(result.p) + ', d=' + fmt(result.cohensD) + ')';
        summaryHtml += '</div>';
        extras.push({ title: '', html: summaryHtml });

        state.results['ptt'] = { data: mainRows, title: 'Paired Samples t-Test', extras: extras };
        displayResults('ptt');
    }

    // =========================================================================
    // One-way ANOVA
    // =========================================================================

    function runOnewayAnova() {
        var dvList = getCheckedVars('ow-dv-picker');
        if (dvList.length === 0) {
            var singleDv = getSelectValue('ow-dv');
            if (singleDv) dvList = [singleDv];
        }
        var iv = getPickerValue('ow-iv-picker', 'ow-iv');
        if (dvList.length === 0 || !iv) { alert('Please select both DV and Factor.'); return; }
        var alpha = parseFloat(getSelectValue('ow-alpha')) || 0.05;

        var allMainRows = [];
        var allExtras = [];

        dvList.forEach(function(dv) {
            var split = splitByGroup(dv, iv);
            var gNames = split.groupNames;
            if (gNames.length < 2) { alert('Need at least 2 groups for "' + dv + '". Found ' + gNames.length + '.'); return; }

            var groups = gNames.map(function (g) { return split.groups[g]; });

            // Run assumption check (only for first/single DV)
            if (dvList.length === 1) {
                runAssumptionCheck('ow', 'oneway-anova', { groups: groups, names: gNames });
            }

            var result = Stats.onewayAnova(groups, gNames);
            if (!result) { alert('Could not compute ANOVA for "' + dv + '". Check your data.'); return; }

            var extras = [];

            // Descriptive
            extras.push({
                title: 'Group Descriptive Statistics' + (dvList.length > 1 ? ' (' + dv + ')' : ''),
                data: result.descriptives.map(function (d) {
                    return { Group: d.group, N: d.n, Mean: fmt(d.mean), 'S.D.': fmt(d.sd), 'S.E.': fmt(d.se) };
                })
            });

            // Levene
            extras.push({
                title: "Levene's Test" + (dvList.length > 1 ? ' (' + dv + ')' : ''),
                data: [{ F: fmt(result.leveneF), 'p-value': Stats.formatPValue(result.leveneP), Conclusion: result.leveneP >= alpha ? 'Homogeneous' : 'Not Homogeneous' }]
            });

            // ANOVA Table with eta-squared and effect in main rows
            var dvLabel = dvList.length > 1 ? dv : '';
            var mainRows = [
                { DV: dvLabel, Source: 'Between Groups', SS: fmt(result.ssBetween), df: result.dfBetween, MS: fmt(result.msBetween), F: fmt(result.f), 'p-value': Stats.formatPValue(result.p), '\u03B7\u00B2': fmt(result.etaSquared), 'Effect': result.effectSize || '' },
                { DV: dvLabel, Source: 'Within Groups', SS: fmt(result.ssWithin), df: result.dfWithin, MS: fmt(result.msWithin), F: '', 'p-value': '', '\u03B7\u00B2': '', 'Effect': '' },
                { DV: dvLabel, Source: 'Total', SS: fmt(result.ssTotal), df: result.dfBetween + result.dfWithin, MS: '', F: '', 'p-value': '', '\u03B7\u00B2': '', 'Effect': '' }
            ];

            // If single DV, remove DV column
            if (dvList.length === 1) {
                mainRows.forEach(function(row) { delete row.DV; });
            }

            // Post-hoc
            if (result.posthoc && result.posthoc.length > 0) {
                extras.push({
                    title: 'Post-Hoc Pairwise Comparisons (Bonferroni)' + (dvList.length > 1 ? ' (' + dv + ')' : ''),
                    data: result.posthoc.map(function (ph) {
                        return {
                            'Group A': ph.groupA, 'Group B': ph.groupB,
                            'Mean Diff': fmt(ph.meanDiff), t: fmt(ph.t),
                            'Adj. p-value': Stats.formatPValue(ph.p),
                            "Cohen's d": fmt(ph.cohensD),
                            Sig: ph.p < alpha ? '*' : ''
                        };
                    })
                });
            }

            allMainRows = allMainRows.concat(mainRows);
            allExtras = allExtras.concat(extras);
        });

        state.results['ow'] = { data: allMainRows, title: 'One-way ANOVA', extras: allExtras };
        displayResults('ow');
    }

    // =========================================================================
    // Two-way ANOVA (simplified - using cell means approach)
    // =========================================================================

    function runTwowayAnova() {
        var dv = getPickerValue('tw-dv-picker', 'tw-dv');
        var iv1 = getPickerValue('tw-iv1-picker', 'tw-iv1');
        var iv2 = getPickerValue('tw-iv2-picker', 'tw-iv2');
        if (!dv || !iv1 || !iv2) { alert('Please select DV and both Factors.'); return; }

        // Group data by factor combinations
        var cells = {};
        var levels1 = [], levels2 = [];
        var l1Set = {}, l2Set = {};

        state.data.forEach(function (row) {
            var val = parseFloat(row[dv]);
            if (isNaN(val)) return;
            var f1 = String(row[iv1]);
            var f2 = String(row[iv2]);
            if (!l1Set[f1]) { l1Set[f1] = true; levels1.push(f1); }
            if (!l2Set[f2]) { l2Set[f2] = true; levels2.push(f2); }
            var key = f1 + '|' + f2;
            if (!cells[key]) cells[key] = [];
            cells[key].push(val);
        });

        // For a simplified two-way, just run one-way on each factor separately
        // Full factorial ANOVA is complex; provide main effects
        var split1 = splitByGroup(dv, iv1);
        var groups1 = split1.groupNames.map(function (g) { return split1.groups[g]; });
        var result1 = Stats.onewayAnova(groups1, split1.groupNames);

        var split2 = splitByGroup(dv, iv2);
        var groups2 = split2.groupNames.map(function (g) { return split2.groups[g]; });
        var result2 = Stats.onewayAnova(groups2, split2.groupNames);

        var mainRows = [];
        if (result1) {
            mainRows.push({ Source: iv1 + ' (Main Effect)', SS: fmt(result1.ssBetween), df: result1.dfBetween, MS: fmt(result1.msBetween), F: fmt(result1.f), 'p-value': Stats.formatPValue(result1.p) });
        }
        if (result2) {
            mainRows.push({ Source: iv2 + ' (Main Effect)', SS: fmt(result2.ssBetween), df: result2.dfBetween, MS: fmt(result2.msBetween), F: fmt(result2.f), 'p-value': Stats.formatPValue(result2.p) });
        }

        var extras = [];
        if (result1 && result1.descriptives) {
            extras.push({
                title: 'Descriptives by ' + iv1,
                data: result1.descriptives.map(function (d) {
                    return { Group: d.group, N: d.n, Mean: fmt(d.mean), 'S.D.': fmt(d.sd) };
                })
            });
        }
        if (result2 && result2.descriptives) {
            extras.push({
                title: 'Descriptives by ' + iv2,
                data: result2.descriptives.map(function (d) {
                    return { Group: d.group, N: d.n, Mean: fmt(d.mean), 'S.D.': fmt(d.sd) };
                })
            });
        }

        state.results['tw'] = { data: mainRows, title: 'Two-way ANOVA (Main Effects)', extras: extras };
        displayResults('tw');
    }

    // =========================================================================
    // Repeated Measures ANOVA (simplified via Friedman approach)
    // =========================================================================

    function runRmAnova() {
        var vars = getCheckedVars('rm-picker');
        if (vars.length === 0) vars = getSelected('rm-vars');
        if (vars.length < 2) { alert('Please select at least 2 variables.'); return; }

        // Treat as one-way with repeated measures
        // Use each variable as a "condition"
        var groups = vars.map(function (v) { return getColumnData(v, true); });
        var n = Math.min.apply(null, groups.map(function (g) { return g.length; }));
        groups = groups.map(function (g) { return g.slice(0, n); });

        // Use Friedman as non-parametric RM equivalent
        var fr = Stats.friedmanTest(groups);

        // Also provide descriptives
        var extras = [];
        var descRows = vars.map(function (v, i) {
            var d = Stats.descriptive(groups[i]);
            return d ? { Variable: v, N: d.n, Mean: fmt(d.mean), 'S.D.': fmt(d.sd), 'S.E.': fmt(d.se) } : null;
        }).filter(Boolean);
        extras.push({ title: 'Descriptive Statistics', data: descRows });

        var mainRows = [];
        if (fr) {
            mainRows.push({
                'Chi-Square': fmt(fr.chiSquare),
                df: fr.df,
                'p-value': Stats.formatPValue(fr.p),
                "Kendall's W": fmt(fr.kendallW)
            });
        }

        state.results['rm'] = { data: mainRows, title: 'Repeated Measures Analysis', extras: extras };
        displayResults('rm');
    }

    // =========================================================================
    // ANCOVA (simplified)
    // =========================================================================

    function runAncova() {
        var dv = getPickerValue('anc-dv-picker', 'anc-dv');
        var iv = getPickerValue('anc-iv-picker', 'anc-iv');
        var cov = getPickerValue('anc-cov-picker', 'anc-cov');
        if (!dv || !iv || !cov) { alert('Please select DV, Factor, and Covariate.'); return; }

        // Run linear regression with covariate and dummy-coded IV
        var split = splitByGroup(dv, iv);
        var gNames = split.groupNames;

        // Create dummy variables for IV
        var yVals = [];
        var covVals = [];
        var dummies = {};
        for (var i = 1; i < gNames.length; i++) dummies[gNames[i]] = [];

        state.data.forEach(function (row) {
            var y = parseFloat(row[dv]);
            var c = parseFloat(row[cov]);
            if (isNaN(y) || isNaN(c)) return;
            var g = String(row[iv]);
            yVals.push(y);
            covVals.push(c);
            for (var i = 1; i < gNames.length; i++) {
                dummies[gNames[i]].push(g === gNames[i] ? 1 : 0);
            }
        });

        var xs = [covVals];
        var names = ['(Intercept)', cov];
        for (var i = 1; i < gNames.length; i++) {
            xs.push(dummies[gNames[i]]);
            names.push(iv + '=' + gNames[i]);
        }

        var result = Stats.linearRegression(yVals, xs, names);
        if (!result) { alert('Could not compute ANCOVA. Check your data.'); return; }

        var mainRows = result.coefficients.map(function (c) {
            return {
                Variable: c.variable,
                B: fmt(c.b), 'S.E.': fmt(c.se), t: fmt(c.t),
                'p-value': Stats.formatPValue(c.p)
            };
        });

        var extras = [{
            title: 'Model Summary',
            data: [{ R: fmt(result.r), 'R-Squared': fmt(result.rSquared), 'Adj. R-Squared': fmt(result.adjRSquared), F: fmt(result.f), 'p-value': Stats.formatPValue(result.fP) }]
        }];

        state.results['anc'] = { data: mainRows, title: 'ANCOVA Coefficients', extras: extras };
        displayResults('anc');
    }

    // =========================================================================
    // Mann-Whitney U
    // =========================================================================

    function runMannWhitney() {
        var dvList = getCheckedVars('mw-dv-picker');
        if (dvList.length === 0) {
            var singleDv = getSelectValue('mw-dv');
            if (singleDv) dvList = [singleDv];
        }
        var iv = getPickerValue('mw-iv-picker', 'mw-iv');
        if (dvList.length === 0 || !iv) { alert('Please select both DV and IV.'); return; }
        var alpha = parseFloat(getSelectValue('mw-alpha')) || 0.05;

        var allMainRows = [];
        var allExtras = [];

        dvList.forEach(function(dv) {
            var split = splitByGroup(dv, iv);
            var gNames = split.groupNames;
            if (gNames.length !== 2) { alert('Mann-Whitney U requires exactly 2 groups. Found ' + gNames.length + ' for "' + dv + '".'); return; }

            var g1 = split.groups[gNames[0]];
            var g2 = split.groups[gNames[1]];

            // Run assumption check (only for first/single DV)
            if (dvList.length === 1) {
                runAssumptionCheck('mw', 'independent-ttest', { group1: g1, group2: g2 });
            }

            var result = Stats.mannWhitneyU(g1, g2);
            if (!result) { alert('Could not compute Mann-Whitney U for "' + dv + '". Check your data.'); return; }

            var extras = [{
                title: 'Group Descriptive Statistics' + (dvList.length > 1 ? ' (' + dv + ')' : ''),
                data: [
                    { Group: gNames[0], N: result.desc1.n, Median: fmt(result.desc1.median), Mean: fmt(result.desc1.mean), 'S.D.': fmt(result.desc1.sd), 'Mean Rank': fmt(result.desc1.meanRank), 'Sum of Ranks': fmt(result.desc1.sumRank, 1) },
                    { Group: gNames[1], N: result.desc2.n, Median: fmt(result.desc2.median), Mean: fmt(result.desc2.mean), 'S.D.': fmt(result.desc2.sd), 'Mean Rank': fmt(result.desc2.meanRank), 'Sum of Ranks': fmt(result.desc2.sumRank, 1) }
                ]
            }];

            var mainRows = [{
                'DV': dv,
                'Mann-Whitney U': fmt(result.U, 1),
                Z: fmt(result.Z),
                'p-value': Stats.formatPValue(result.p),
                'Effect Size (r)': fmt(result.r),
                Interpretation: result.effectSize
            }];

            // If single DV, remove DV column
            if (dvList.length === 1) {
                mainRows.forEach(function(row) { delete row.DV; });
            }

            var sig = result.p < alpha;
            var summaryHtml = '<div class="detail-box"><h4>Summary' + (dvList.length > 1 ? ' (' + dv + ')' : '') + '</h4>';
            summaryHtml += '<p>Mann-Whitney U = ' + fmt(result.U, 1) + ', Z = ' + fmt(result.Z) + ', p = ' + Stats.formatPValue(result.p) + '</p>';
            summaryHtml += '<p>Result: ' + (sig ? 'Statistically significant difference (p < ' + alpha + ')' : 'No significant difference (p >= ' + alpha + ')') + '</p>';
            summaryHtml += '</div>';
            extras.push({ title: '', html: summaryHtml });

            allMainRows = allMainRows.concat(mainRows);
            allExtras = allExtras.concat(extras);
        });

        state.results['mw'] = { data: allMainRows, title: 'Mann-Whitney U Test', extras: allExtras };
        displayResults('mw');
    }

    // =========================================================================
    // Wilcoxon Signed-Rank
    // =========================================================================

    function runWilcoxon() {
        var before = getPickerValue('wx-before-picker', 'wx-before');
        var after = getPickerValue('wx-after-picker', 'wx-after');
        if (!before || !after) { alert('Please select both variables.'); return; }
        var alpha = parseFloat(getSelectValue('wx-alpha')) || 0.05;

        var bData = getColumnData(before, true);
        var aData = getColumnData(after, true);
        var minLen = Math.min(bData.length, aData.length);
        bData = bData.slice(0, minLen);
        aData = aData.slice(0, minLen);

        // Run assumption check
        runAssumptionCheck('wx', 'paired-ttest', { before: bData, after: aData });

        var result = Stats.wilcoxonSignedRank(bData, aData);
        if (!result) { alert('Could not compute Wilcoxon test. Check your data.'); return; }

        var extras = [{
            title: 'Descriptive Statistics',
            data: [
                { Variable: before, N: result.descBefore.n, Median: fmt(result.descBefore.median), Mean: fmt(result.descBefore.mean), 'S.D.': fmt(result.descBefore.sd) },
                { Variable: after, N: result.descAfter.n, Median: fmt(result.descAfter.median), Mean: fmt(result.descAfter.mean), 'S.D.': fmt(result.descAfter.sd) }
            ]
        }, {
            title: 'Ranks',
            data: [{ 'Positive Ranks': result.nPos, 'Negative Ranks': result.nNeg, Ties: result.nTies }]
        }];

        var mainRows = [{
            W: fmt(result.W, 1),
            Z: fmt(result.Z),
            'p-value': Stats.formatPValue(result.p),
            'Effect Size (r)': fmt(result.r),
            Interpretation: result.effectSize
        }];

        var sig = result.p < alpha;
        var summaryHtml = '<div class="detail-box"><h4>Summary</h4>';
        summaryHtml += '<p>Wilcoxon W = ' + fmt(result.W, 1) + ', Z = ' + fmt(result.Z) + ', p = ' + Stats.formatPValue(result.p) + '</p>';
        summaryHtml += '<p>Result: ' + (sig ? 'Statistically significant difference' : 'No significant difference') + '</p></div>';
        extras.push({ title: '', html: summaryHtml });

        state.results['wx'] = { data: mainRows, title: 'Wilcoxon Signed-Rank Test', extras: extras };
        displayResults('wx');
    }

    // =========================================================================
    // Kruskal-Wallis
    // =========================================================================

    function runKruskalWallis() {
        var dvList = getCheckedVars('kw-dv-picker');
        if (dvList.length === 0) {
            var singleDv = getSelectValue('kw-dv');
            if (singleDv) dvList = [singleDv];
        }
        var iv = getPickerValue('kw-iv-picker', 'kw-iv');
        if (dvList.length === 0 || !iv) { alert('Please select both DV and IV.'); return; }
        var alpha = parseFloat(getSelectValue('kw-alpha')) || 0.05;

        var allMainRows = [];
        var allExtras = [];

        dvList.forEach(function(dv) {
            var split = splitByGroup(dv, iv);
            var gNames = split.groupNames;
            if (gNames.length < 2) { alert('Need at least 2 groups for "' + dv + '".'); return; }

            var groups = gNames.map(function (g) { return split.groups[g]; });

            // Run assumption check (only for first/single DV)
            if (dvList.length === 1) {
                runAssumptionCheck('kw', 'oneway-anova', { groups: groups, names: gNames });
            }

            var result = Stats.kruskalWallis(groups, gNames);
            if (!result) { alert('Could not compute Kruskal-Wallis for "' + dv + '". Check your data.'); return; }

            var extras = [{
                title: 'Group Descriptive Statistics' + (dvList.length > 1 ? ' (' + dv + ')' : ''),
                data: result.descriptives.map(function (d) {
                    return { Group: d.group, N: d.n, Median: fmt(d.median), Mean: fmt(d.mean), 'S.D.': fmt(d.sd), 'Mean Rank': fmt(d.meanRank) };
                })
            }];

            var mainRows = [{
                'DV': dv,
                'H (Chi-Square)': fmt(result.H),
                df: result.df,
                'p-value': Stats.formatPValue(result.p),
                'Eta Squared H': fmt(result.etaSquaredH),
                Interpretation: result.effectSize
            }];

            // If single DV, remove DV column
            if (dvList.length === 1) {
                mainRows.forEach(function(row) { delete row.DV; });
            }

            if (result.posthoc && result.posthoc.length > 0) {
                extras.push({
                    title: 'Post-Hoc Pairwise Comparisons (Bonferroni)' + (dvList.length > 1 ? ' (' + dv + ')' : ''),
                    data: result.posthoc.map(function (ph) {
                        return {
                            'Group A': ph.groupA, 'Group B': ph.groupB,
                            U: fmt(ph.U, 1), Z: fmt(ph.Z),
                            'Adj. p-value': Stats.formatPValue(ph.p),
                            r: fmt(ph.r),
                            Sig: ph.p < alpha ? '*' : ''
                        };
                    })
                });
            }

            allMainRows = allMainRows.concat(mainRows);
            allExtras = allExtras.concat(extras);
        });

        state.results['kw'] = { data: allMainRows, title: 'Kruskal-Wallis H Test', extras: allExtras };
        displayResults('kw');
    }

    // =========================================================================
    // Friedman Test
    // =========================================================================

    function runFriedman() {
        var vars = getCheckedVars('fr-picker');
        if (vars.length === 0) vars = getSelected('fr-vars');
        if (vars.length < 2) { alert('Please select at least 2 variables.'); return; }
        var alpha = parseFloat(getSelectValue('fr-alpha')) || 0.05;

        var dataArrays = vars.map(function (v) { return getColumnData(v, true); });
        var minLen = Math.min.apply(null, dataArrays.map(function (a) { return a.length; }));
        dataArrays = dataArrays.map(function (a) { return a.slice(0, minLen); });

        var result = Stats.friedmanTest(dataArrays);
        if (!result) { alert('Could not compute Friedman test. Check your data.'); return; }

        // Check normality of each variable
        var fAssumptionEl = document.getElementById('fr-assumptions');
        if (fAssumptionEl) {
            var checksHtml = '<div class="assumption-panel"><h4>\ud83d\udd0d Assumption Check</h4>';
            checksHtml += '<table class="result-table"><thead><tr><th>Variable</th><th>Shapiro-Wilk p</th><th>Normal?</th></tr></thead><tbody>';
            var allNormal = true;
            vars.forEach(function(v, i) {
                var sw = Stats.shapiroWilk(dataArrays[i]);
                var normal = sw && sw.p > 0.05;
                if (!normal) allNormal = false;
                checksHtml += '<tr><td>' + v + '</td><td>' + (sw ? Stats.formatPValue(sw.p) : 'N/A') + '</td><td>' + (normal ? '\u2705' : '\u26a0\ufe0f') + '</td></tr>';
            });
            checksHtml += '</tbody></table>';
            checksHtml += '<div class="' + (allNormal ? 'assumption-warn' : 'assumption-pass') + '">';
            checksHtml += '<strong>\ud83d\udca1</strong> ' + (allNormal ? '\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25\u0e41\u0e08\u0e01\u0e41\u0e08\u0e07\u0e1b\u0e01\u0e15\u0e34 \u0e2d\u0e32\u0e08\u0e1e\u0e34\u0e08\u0e32\u0e23\u0e13\u0e32\u0e43\u0e0a\u0e49 RM-ANOVA \u0e41\u0e17\u0e19' : '\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25\u0e44\u0e21\u0e48\u0e41\u0e08\u0e01\u0e41\u0e08\u0e07\u0e1b\u0e01\u0e15\u0e34 Friedman test \u0e40\u0e2b\u0e21\u0e32\u0e30\u0e2a\u0e21');
            checksHtml += '</div></div>';
            fAssumptionEl.innerHTML = checksHtml;
            fAssumptionEl.style.display = 'block';
        }

        // Patch variable names into descriptives
        result.descriptives.forEach(function (d, i) { d.variable = vars[i]; });

        var extras = [{
            title: 'Descriptive Statistics',
            data: result.descriptives.map(function (d) {
                return { Variable: d.variable, N: d.n, Median: fmt(d.median), Mean: fmt(d.mean), 'S.D.': fmt(d.sd), 'Mean Rank': fmt(d.meanRank) };
            })
        }];

        var mainRows = [{
            'Chi-Square': fmt(result.chiSquare),
            df: result.df,
            'p-value': Stats.formatPValue(result.p),
            "Kendall's W": fmt(result.kendallW)
        }];

        if (result.posthoc && result.posthoc.length > 0) {
            // Patch variable names into post-hoc
            result.posthoc.forEach(function (ph, idx) {
                var pairIdx = 0;
                for (var i = 0; i < vars.length; i++) {
                    for (var j = i + 1; j < vars.length; j++) {
                        if (pairIdx === idx) {
                            ph.variableA = vars[i];
                            ph.variableB = vars[j];
                        }
                        pairIdx++;
                    }
                }
            });
            extras.push({
                title: 'Post-Hoc Pairwise Comparisons (Bonferroni)',
                data: result.posthoc.map(function (ph) {
                    return {
                        'Variable A': ph.variableA, 'Variable B': ph.variableB,
                        W: fmt(ph.W, 1), Z: fmt(ph.Z),
                        'Adj. p-value': Stats.formatPValue(ph.p),
                        r: fmt(ph.r),
                        Sig: ph.p < alpha ? '*' : ''
                    };
                })
            });
        }

        state.results['fr'] = { data: mainRows, title: 'Friedman Test', extras: extras };
        displayResults('fr');
    }

    // =========================================================================
    // Chi-Square Test
    // =========================================================================

    function runChiSquare() {
        var var1 = getPickerValue('chi-var1-picker', 'chi-var1');
        var var2 = getPickerValue('chi-var2-picker', 'chi-var2');
        if (!var1 || !var2) { alert('Please select both variables.'); return; }
        var alpha = parseFloat(getSelectValue('chi-alpha')) || 0.05;

        var v1 = getColumnData(var1, false);
        var v2 = getColumnData(var2, false);

        // Filter to rows where both are present
        var pairs = [];
        for (var i = 0; i < Math.min(v1.length, v2.length); i++) {
            if (v1[i] !== null && v1[i] !== undefined && v1[i] !== '' &&
                v2[i] !== null && v2[i] !== undefined && v2[i] !== '') {
                pairs.push({ a: v1[i], b: v2[i] });
            }
        }
        var filteredV1 = pairs.map(function (p) { return p.a; });
        var filteredV2 = pairs.map(function (p) { return p.b; });

        var result = Stats.chiSquare(filteredV1, filteredV2);
        if (!result) { alert('Could not compute Chi-Square. Need at least 2 categories for each variable.'); return; }

        // Assumption check: expected frequencies >= 5
        var chiAssumptionEl = document.getElementById('chi-assumptions');
        if (chiAssumptionEl && result) {
            var expected = result.expected;
            var totalCells = 0, cellsBelow5 = 0;
            if (expected) {
                for (var r = 0; r < expected.length; r++) {
                    for (var c = 0; c < expected[r].length; c++) {
                        totalCells++;
                        if (expected[r][c] < 5) cellsBelow5++;
                    }
                }
            }
            var pctBelow = totalCells > 0 ? (cellsBelow5 / totalCells * 100) : 0;
            var chiOk = pctBelow <= 20;
            var checksHtml = '<div class="assumption-panel"><h4>\ud83d\udd0d Assumption Check</h4>';
            checksHtml += '<table class="result-table"><thead><tr><th>Check</th><th>Result</th><th>Status</th></tr></thead><tbody>';
            checksHtml += '<tr><td>Expected Frequency >= 5</td><td>' + (100 - pctBelow).toFixed(0) + '% \u0E02\u0E2D\u0E07 cells</td><td>' + (chiOk ? '\u2705' : '\u26a0\ufe0f') + '</td></tr>';
            checksHtml += '<tr><td>Sample Size</td><td>N = ' + result.n + '</td><td>' + (result.n >= 20 ? '\u2705' : '\u26a0\ufe0f') + '</td></tr>';
            checksHtml += '</tbody></table>';
            checksHtml += '<div class="' + (chiOk ? 'assumption-pass' : 'assumption-warn') + '">';
            checksHtml += '<strong>\ud83d\udca1</strong> ' + (chiOk ? 'Assumptions met. Chi-Square test \u0E40\u0E2B\u0E21\u0E32\u0E30\u0E2A\u0E21' : 'Expected < 5 \u0E21\u0E32\u0E01\u0E01\u0E27\u0E48\u0E32 20% \u2014 \u0E1E\u0E34\u0E08\u0E32\u0E23\u0E13\u0E32\u0E43\u0E0A\u0E49 Fisher\'s Exact Test');
            checksHtml += '</div></div>';
            chiAssumptionEl.innerHTML = checksHtml;
            chiAssumptionEl.style.display = 'block';
        }

        var extras = [];

        // Observed frequency table
        var obsRows = [];
        for (var i = 0; i < result.rowLabels.length; i++) {
            var row = {};
            row[var1 + ' / ' + var2] = result.rowLabels[i];
            for (var j = 0; j < result.colLabels.length; j++) {
                row[result.colLabels[j]] = result.observed[i][j];
            }
            obsRows.push(row);
        }
        extras.push({ title: 'Observed Frequencies', data: obsRows });

        // Expected frequency table
        var expRows = [];
        for (var i = 0; i < result.rowLabels.length; i++) {
            var row = {};
            row[var1 + ' / ' + var2] = result.rowLabels[i];
            for (var j = 0; j < result.colLabels.length; j++) {
                row[result.colLabels[j]] = fmt(result.expected[i][j], 2);
            }
            expRows.push(row);
        }
        extras.push({ title: 'Expected Frequencies', data: expRows });

        var mainRows = [{
            'Chi-Square': fmt(result.chiSquare),
            df: result.df,
            'p-value': Stats.formatPValue(result.p),
            N: result.n,
            "Cramer's V": fmt(result.cramersV),
            Interpretation: result.effectSize
        }];

        var sig = result.p < alpha;
        var summaryHtml = '<div class="detail-box"><h4>Summary</h4>';
        summaryHtml += '<p>Chi-Square(' + result.df + ') = ' + fmt(result.chiSquare) + ', p = ' + Stats.formatPValue(result.p) + '</p>';
        summaryHtml += '<p>Result: ' + (sig ? 'Significant association between variables' : 'No significant association') + '</p></div>';
        extras.push({ title: '', html: summaryHtml });

        state.results['chi'] = { data: mainRows, title: 'Chi-Square Test of Independence', extras: extras };
        displayResults('chi');
    }

    // =========================================================================
    // Correlation
    // =========================================================================

    function runCorrelation() {
        var vars = getCheckedVars('cor-picker');
        if (vars.length === 0) vars = getDualListSelected('cor-dual');
        if (!vars || vars.length === 0) vars = getSelected('cor-vars');
        if (vars.length < 2) { alert('Please select at least 2 variables.'); return; }
        var method = getSelectValue('cor-method') || 'pearson';

        var dataArrays = vars.map(function (v) { return getColumnData(v, true); });

        // Run assumption check
        runAssumptionCheck('cor', 'correlation', { vars: dataArrays });

        var result = Stats.correlation(dataArrays, vars, method);
        if (!result) { alert('Could not compute correlation. Check your data.'); return; }

        // Pairs table
        var pairsRows = result.pairs.map(function (p) {
            return {
                'Variable 1': p.var1, 'Variable 2': p.var2,
                r: fmt(p.r), 'p-value': Stats.formatPValue(p.p),
                'R-Squared': fmt(p.rSquared),
                Direction: p.direction, Strength: p.strength,
                Sig: p.sig ? '*' : ''
            };
        });

        // Correlation matrix as extra
        var matrixRows = [];
        for (var i = 0; i < vars.length; i++) {
            var row = { Variable: vars[i] };
            for (var j = 0; j < vars.length; j++) {
                row[vars[j]] = fmt(result.matrix[i][j]);
            }
            matrixRows.push(row);
        }

        var extras = [{ title: 'Correlation Matrix (' + method.charAt(0).toUpperCase() + method.slice(1) + ')', data: matrixRows }];

        state.results['cor'] = { data: pairsRows, title: 'Correlation Pairs', extras: extras };
        displayResults('cor');
    }

    // =========================================================================
    // Linear Regression
    // =========================================================================

    function runLinearRegression() {
        var dv = getPickerValue('lr-dv-picker', 'lr-dv');
        var ivs = getCheckedVars('lr-iv-picker');
        if (ivs.length === 0) ivs = getSelected('lr-ivs');
        if (!dv || ivs.length === 0) { alert('Please select DV and at least one IV.'); return; }

        var method = getSelectValue('lr-method') || 'enter';
        var pIn = parseFloat(document.getElementById('lr-pin') ? document.getElementById('lr-pin').value : 0.05) || 0.05;
        var pOut = parseFloat(document.getElementById('lr-pout') ? document.getElementById('lr-pout').value : 0.10) || 0.10;

        // Get display options
        var lrOpts = {};
        document.querySelectorAll('input[name="lr-opt"]').forEach(function(cb) { lrOpts[cb.value] = cb.checked; });

        var yData = getColumnData(dv, true);
        var xData = ivs.map(function (iv) { return getColumnData(iv, true); });
        var minLen = Math.min(yData.length, Math.min.apply(null, xData.map(function (x) { return x.length; })));
        yData = yData.slice(0, minLen);
        xData = xData.map(function (x) { return x.slice(0, minLen); });

        // --- Stepwise / Forward / Backward variable selection ---
        var selectedIVs = ivs.slice();
        var selectedXData = xData.slice();
        var stepLog = [];
        var methodLabel = 'Enter';

        if (method === 'stepwise' || method === 'forward' || method === 'backward') {
            methodLabel = method.charAt(0).toUpperCase() + method.slice(1);
            var result_steps = _stepwiseRegression(yData, xData, ivs, method, pIn, pOut);
            selectedIVs = result_steps.selectedVars;
            selectedXData = result_steps.selectedXData;
            stepLog = result_steps.steps;
            if (selectedIVs.length === 0) {
                alert('ไม่มีตัวแปรใดผ่านเกณฑ์ p-in = ' + pIn + ' ลองปรับเกณฑ์หรือเปลี่ยน Method');
                return;
            }
        }

        var names = ['(Intercept)'].concat(selectedIVs);

        // Run assumption check
        runAssumptionCheck('lr', 'correlation', { vars: selectedXData });

        var result = Stats.linearRegression(yData, selectedXData, names);
        if (!result) { alert('Could not compute regression. Check your data (possible multicollinearity).'); return; }

        // --- Build extras ---
        var extras = [];

        // 1. Method Information
        var methodDesc = {
            'enter': 'Enter — ใส่ตัวแปรอิสระทั้งหมดเข้าสมการพร้อมกัน ไม่มีการคัดเลือก',
            'stepwise': 'Stepwise — เพิ่ม/ลดตัวแปรอัตโนมัติตามเกณฑ์ p-in=' + pIn + ', p-out=' + pOut,
            'forward': 'Forward Selection — เพิ่มตัวแปรทีละตัวตามเกณฑ์ p-in=' + pIn,
            'backward': 'Backward Elimination — เริ่มจากทุกตัว แล้วตัดตัวที่ไม่มีนัยสำคัญ p-out=' + pOut
        };
        extras.push({
            title: 'Method / กระบวนการสร้างโมเดล',
            data: [{
                'Method': methodLabel,
                'Description': methodDesc[method] || method,
                'DV': dv,
                'Candidate IVs': ivs.join(', '),
                'Selected IVs': selectedIVs.join(', '),
                'Variables Entered': selectedIVs.length,
                'Variables Excluded': ivs.length - selectedIVs.length,
                'p-in Criteria': pIn,
                'p-out Criteria': pOut,
                'N': yData.length
            }]
        });

        // 2. Step-by-step log (for Stepwise/Forward/Backward)
        if (stepLog.length > 0) {
            extras.push({ title: 'Variable Selection Steps / ขั้นตอนการคัดเลือกตัวแปร', data: stepLog });
        }

        // 3. Model Summary
        var modelSummary = {
            'Model': '1', 'Method': methodLabel,
            'R': fmt(result.r), 'R²': fmt(result.rSquared),
            'Adjusted R²': fmt(result.adjRSquared),
            'Std. Error of Estimate': fmt(Math.sqrt(result.anova.msRes)),
            'F': fmt(result.f), 'Sig.': Stats.formatPValue(result.fP)
        };
        if (lrOpts.durbin && result.durbinWatson !== undefined) {
            modelSummary['Durbin-Watson'] = fmt(result.durbinWatson);
        }
        extras.push({ title: 'Model Summary', data: [modelSummary] });

        // 4. ANOVA Table
        extras.push({
            title: 'ANOVA',
            data: [
                { Source: 'Regression', SS: fmt(result.anova.ssReg), df: result.anova.dfReg, MS: fmt(result.anova.msReg), F: fmt(result.f), 'Sig.': Stats.formatPValue(result.fP) },
                { Source: 'Residual', SS: fmt(result.anova.ssRes), df: result.anova.dfRes, MS: fmt(result.anova.msRes), F: '', 'Sig.': '' },
                { Source: 'Total', SS: fmt(result.anova.ssTotal), df: result.anova.dfReg + result.anova.dfRes, MS: '', F: '', 'Sig.': '' }
            ]
        });

        // 5. Collinearity Diagnostics
        if (lrOpts.collinearity && selectedIVs.length > 1) {
            var collinRows = [];
            selectedIVs.forEach(function(ivName, idx) {
                var xi = selectedXData[idx];
                // Regress xi on all other IVs to get R²_i for VIF
                var otherX = selectedXData.filter(function(_,j){return j!==idx;});
                var otherNames = ['(Intercept)'].concat(selectedIVs.filter(function(_,j){return j!==idx;}));
                var regI = Stats.linearRegression(xi, otherX, otherNames);
                var r2i = regI ? regI.rSquared : 0;
                var tol = 1 - r2i;
                var vif = tol > 0 ? 1 / tol : 999;
                collinRows.push({ 'Variable': ivName, 'Tolerance': fmt(tol), 'VIF': fmt(vif), 'Status': vif > 10 ? 'Multicollinearity!' : (vif > 5 ? 'Warning' : 'OK') });
            });
            extras.push({ title: 'Collinearity Diagnostics / ค่าสหสัมพันธ์ร่วมเชิงเส้น', data: collinRows });
        }

        // 6. Correlation Matrix
        if (lrOpts.correlation) {
            var allVars = [dv].concat(selectedIVs);
            var allData = [yData].concat(selectedXData);
            var corrRows = [];
            allVars.forEach(function(v1, i) {
                var row = { 'Variable': v1 };
                allVars.forEach(function(v2, j) {
                    row[v2] = fmt(jStat.corrcoeff(allData[i], allData[j]));
                });
                corrRows.push(row);
            });
            extras.push({ title: 'Correlation Matrix (Pearson)', data: corrRows });
        }

        // --- Main: Coefficients Table ---
        var mainRows = result.coefficients.map(function (c, idx) {
            var row = {
                Variable: c.variable,
                B: fmt(c.b), 'S.E.': fmt(c.se),
                'Beta (Std.)': idx === 0 ? '—' : fmt(c.beta !== undefined ? c.beta : ''),
                t: fmt(c.t),
                'Sig.': Stats.formatPValue(c.p),
                '95% CI': Stats.formatCI ? Stats.formatCI(c.ci95Lo, c.ci95Hi) : '[' + fmt(c.ci95Lo) + ', ' + fmt(c.ci95Hi) + ']'
            };
            // Add collinearity to coefficient table
            if (lrOpts.collinearity && idx > 0 && selectedIVs.length > 1) {
                var xi = selectedXData[idx-1];
                var otherX = selectedXData.filter(function(_,j){return j!==idx-1;});
                var otherNames = ['(Intercept)'].concat(selectedIVs.filter(function(_,j){return j!==idx-1;}));
                var regI = Stats.linearRegression(xi, otherX, otherNames);
                var r2i = regI ? regI.rSquared : 0;
                row['Tolerance'] = fmt(1 - r2i);
                row['VIF'] = fmt((1-r2i) > 0 ? 1/(1-r2i) : 999);
            }
            return row;
        });

        // 7. Residual Statistics
        if (lrOpts.residuals) {
            var residuals = yData.map(function(y, i) {
                var pred = result.coefficients[0].b;
                selectedXData.forEach(function(xd, j) { pred += result.coefficients[j+1].b * xd[i]; });
                return y - pred;
            });
            var predicted = yData.map(function(y, i) {
                var pred = result.coefficients[0].b;
                selectedXData.forEach(function(xd, j) { pred += result.coefficients[j+1].b * xd[i]; });
                return pred;
            });
            var sdRes = jStat.stdev(residuals, true);
            var stdResiduals = residuals.map(function(r) { return r / sdRes; });
            extras.push({
                title: 'Residual Statistics',
                data: [{
                    'Statistic': 'Predicted Value', 'Min': fmt(jStat.min(predicted)), 'Max': fmt(jStat.max(predicted)), 'Mean': fmt(jStat.mean(predicted)), 'S.D.': fmt(jStat.stdev(predicted, true))
                }, {
                    'Statistic': 'Residual', 'Min': fmt(jStat.min(residuals)), 'Max': fmt(jStat.max(residuals)), 'Mean': fmt(jStat.mean(residuals)), 'S.D.': fmt(sdRes)
                }, {
                    'Statistic': 'Std. Residual', 'Min': fmt(jStat.min(stdResiduals)), 'Max': fmt(jStat.max(stdResiduals)), 'Mean': fmt(jStat.mean(stdResiduals)), 'S.D.': fmt(jStat.stdev(stdResiduals, true))
                }]
            });
        }

        // 8. Excluded Variables (for Stepwise/Forward)
        if (method !== 'enter') {
            var excludedVars = ivs.filter(function(v) { return selectedIVs.indexOf(v) === -1; });
            if (excludedVars.length > 0) {
                var exclRows = excludedVars.map(function(v) {
                    var idx = ivs.indexOf(v);
                    var xi = xData[idx];
                    var r = jStat.corrcoeff(xi, yData);
                    return { 'Variable': v, 'Partial Corr. (r)': fmt(r), 'Status': 'Excluded (p > ' + pIn + ')' };
                });
                extras.push({ title: 'Excluded Variables / ตัวแปรที่ไม่ผ่านเกณฑ์', data: exclRows });
            }
        }

        // 9. Regression Equation
        var eqParts = [dv + ' = ' + fmt(result.coefficients[0].b)];
        result.coefficients.forEach(function(c, idx) {
            if (idx === 0) return;
            var sign = c.b >= 0 ? ' + ' : ' - ';
            eqParts.push(sign + fmt(Math.abs(c.b)) + '(' + c.variable + ')');
        });
        var unstdEq = eqParts.join('');
        var stdParts = [dv + '(Z) ='];
        var hasBeta = false;
        result.coefficients.forEach(function(c, idx) {
            if (idx === 0) return;
            var betaVal = c.beta !== undefined ? c.beta : 0;
            hasBeta = true;
            var sign = betaVal >= 0 ? (idx === 1 ? ' ' : ' + ') : ' - ';
            stdParts.push(sign + fmt(Math.abs(betaVal)) + '(Z_' + c.variable + ')');
        });
        var stdEq = hasBeta ? stdParts.join('') : '';

        var eqData = [{'สมการ Unstandardized (ค่าดิบ)': unstdEq}];
        if (stdEq) eqData[0]['สมการ Standardized (ค่ามาตรฐาน)'] = stdEq;
        extras.push({ title: 'สมการถดถอย (Regression Equation)', data: eqData });

        // 10. Diagnostic Warnings
        var warnings = _buildRegressionWarnings(result, selectedIVs, selectedXData, yData, lrOpts, mainRows);
        if (warnings.length > 0) {
            extras.push({ title: '⚠️ จุดสังเกต / ข้อควรระวัง (Diagnostics)', data: warnings });
        }

        state.results['lr'] = { data: mainRows, title: 'Regression Coefficients (Method: ' + methodLabel + ')', extras: extras };
        displayResults('lr');
    }

    // --- Stepwise / Forward / Backward selection helper ---
    function _stepwiseRegression(yData, allXData, allIVNames, method, pIn, pOut) {
        var n = yData.length;
        var steps = [];
        var selected = [];
        var selectedIdx = [];
        var remaining = allIVNames.map(function(_,i){return i;});

        function runOLS(yD, xIdxArr) {
            if (xIdxArr.length === 0) return null;
            var xD = xIdxArr.map(function(i){return allXData[i];});
            var nms = ['(Intercept)'].concat(xIdxArr.map(function(i){return allIVNames[i];}));
            return Stats.linearRegression(yD, xD, nms);
        }

        if (method === 'backward') {
            // Start with all variables
            selectedIdx = allIVNames.map(function(_,i){return i;});
            remaining = [];
            var step = 0;
            while (true) {
                step++;
                var model = runOLS(yData, selectedIdx);
                if (!model) break;
                // Find worst p-value among IVs (skip intercept at index 0)
                var worstP = 0, worstCoefIdx = -1;
                model.coefficients.forEach(function(c, ci) {
                    if (ci === 0) return; // skip intercept
                    if (c.p > worstP) { worstP = c.p; worstCoefIdx = ci; }
                });
                if (worstP > pOut && worstCoefIdx > 0) {
                    var removedVarName = model.coefficients[worstCoefIdx].variable;
                    var removedOrigIdx = allIVNames.indexOf(removedVarName);
                    selectedIdx = selectedIdx.filter(function(i){return i !== removedOrigIdx;});
                    steps.push({
                        'Step': step, 'Action': 'Removed', 'Variable': removedVarName,
                        'p-value': fmt(worstP), 'Criterion': 'p > ' + pOut,
                        'R²': model.rSquared ? fmt(model.rSquared) : '',
                        'Remaining Vars': selectedIdx.map(function(i){return allIVNames[i];}).join(', ')
                    });
                } else {
                    steps.push({
                        'Step': step, 'Action': 'Final Model', 'Variable': '—',
                        'p-value': '—', 'Criterion': 'All p <= ' + pOut,
                        'R²': model.rSquared ? fmt(model.rSquared) : '',
                        'Remaining Vars': selectedIdx.map(function(i){return allIVNames[i];}).join(', ')
                    });
                    break;
                }
                if (selectedIdx.length === 0) break;
            }
        } else {
            // Forward or Stepwise
            var step = 0;
            while (remaining.length > 0) {
                step++;
                var bestP = 1, bestIdx = -1;
                // Try adding each remaining variable
                remaining.forEach(function(ri) {
                    var tryIdx = selectedIdx.concat([ri]);
                    var model = runOLS(yData, tryIdx);
                    if (model) {
                        // Find p-value of newly added variable (last coefficient)
                        var lastCoef = model.coefficients[model.coefficients.length - 1];
                        if (lastCoef && lastCoef.p < bestP) {
                            bestP = lastCoef.p; bestIdx = ri;
                        }
                    }
                });
                if (bestP <= pIn && bestIdx >= 0) {
                    selectedIdx.push(bestIdx);
                    remaining = remaining.filter(function(i){return i !== bestIdx;});
                    var model = runOLS(yData, selectedIdx);
                    steps.push({
                        'Step': step, 'Action': 'Entered', 'Variable': allIVNames[bestIdx],
                        'p-value': fmt(bestP), 'Criterion': 'p <= ' + pIn,
                        'R²': model ? fmt(model.rSquared) : '',
                        'Selected Vars': selectedIdx.map(function(i){return allIVNames[i];}).join(', ')
                    });

                    // Stepwise: after adding, check if any should be removed
                    if (method === 'stepwise' && selectedIdx.length > 1) {
                        var model2 = runOLS(yData, selectedIdx);
                        if (model2) {
                            var worstP2 = 0, worstCI2 = -1;
                            model2.coefficients.forEach(function(c, ci) {
                                if (ci === 0) return;
                                if (c.p > worstP2) { worstP2 = c.p; worstCI2 = ci; }
                            });
                            if (worstP2 > pOut && worstCI2 > 0) {
                                var removedName = model2.coefficients[worstCI2].variable;
                                var removedOI = allIVNames.indexOf(removedName);
                                selectedIdx = selectedIdx.filter(function(i){return i !== removedOI;});
                                remaining.push(removedOI);
                                step++;
                                steps.push({
                                    'Step': step, 'Action': 'Removed', 'Variable': removedName,
                                    'p-value': fmt(worstP2), 'Criterion': 'p > ' + pOut,
                                    'R²': '', 'Selected Vars': selectedIdx.map(function(i){return allIVNames[i];}).join(', ')
                                });
                            }
                        }
                    }
                } else {
                    steps.push({
                        'Step': step, 'Action': 'Stopped', 'Variable': '—',
                        'p-value': fmt(bestP), 'Criterion': 'No more p <= ' + pIn,
                        'R²': '', 'Selected Vars': selectedIdx.map(function(i){return allIVNames[i];}).join(', ')
                    });
                    break;
                }
            }
        }

        return {
            selectedVars: selectedIdx.map(function(i){return allIVNames[i];}),
            selectedXData: selectedIdx.map(function(i){return allXData[i];}),
            steps: steps
        };
    }

    // --- Regression diagnostic warnings builder (Linear Regression) ---
    function _buildRegressionWarnings(result, selectedIVs, selectedXData, yData, lrOpts, mainRows) {
        var warnings = [];
        var n = yData.length;
        var p = selectedIVs.length;

        // 1. R² interpretation
        if (result.rSquared < 0.1) warnings.push({'ประเด็น': 'R² ต่ำมาก', 'รายละเอียด': 'R² = ' + fmt(result.rSquared) + ' — ตัวแปรอิสระอธิบายตัวแปรตามได้ < 10%', 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'โมเดลอธิบายได้น้อย ควรพิจารณาเพิ่มตัวแปร หรือตรวจสอบความสัมพันธ์เชิงเส้น'});
        if (result.adjRSquared < 0) warnings.push({'ประเด็น': 'Adjusted R² ติดลบ', 'รายละเอียด': 'Adj R² = ' + fmt(result.adjRSquared), 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'โมเดลแย่กว่าค่าเฉลี่ย — ตัวแปรที่ใส่อาจไม่เหมาะสม หรือ overfitting'});

        // 2. Model significance
        if (result.fP >= 0.05) warnings.push({'ประเด็น': 'โมเดลไม่มีนัยสำคัญ', 'รายละเอียด': 'F-test p = ' + Stats.formatPValue(result.fP), 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'ตัวแปรอิสระไม่สามารถพยากรณ์ตัวแปรตามได้ ผลทั้งหมดไม่น่าเชื่อถือ'});

        // 3. Durbin-Watson
        if (result.durbinWatson !== undefined) {
            if (result.durbinWatson < 1.5) warnings.push({'ประเด็น': 'Positive Autocorrelation (DW ต่ำ)', 'รายละเอียด': 'Durbin-Watson = ' + fmt(result.durbinWatson) + ' (ค่าปกติ 1.5-2.5)', 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'Residuals อาจมีความสัมพันธ์กัน ผิด Assumption ของ OLS'});
            if (result.durbinWatson > 2.5) warnings.push({'ประเด็น': 'Negative Autocorrelation (DW สูง)', 'รายละเอียด': 'Durbin-Watson = ' + fmt(result.durbinWatson) + ' (ค่าปกติ 1.5-2.5)', 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'Residuals อาจมีความสัมพันธ์เชิงลบ'});
        }

        // 4. Multicollinearity (VIF)
        if (selectedIVs.length > 1) {
            selectedIVs.forEach(function(ivName, idx) {
                var xi = selectedXData[idx];
                var otherX = selectedXData.filter(function(_,j){return j!==idx;});
                var otherNames = ['(Intercept)'].concat(selectedIVs.filter(function(_,j){return j!==idx;}));
                var regI = Stats.linearRegression(xi, otherX, otherNames);
                var r2i = regI ? regI.rSquared : 0;
                var vif = (1-r2i) > 0 ? 1/(1-r2i) : 999;
                if (vif > 10) warnings.push({'ประเด็น': 'Multicollinearity สูงมาก: ' + ivName, 'รายละเอียด': 'VIF = ' + fmt(vif) + ' (เกณฑ์ < 10)', 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'ตัวแปรนี้ซ้ำซ้อนกับตัวอื่นมาก ควรตัดออก หรือรวมเป็นองค์ประกอบ'});
                else if (vif > 5) warnings.push({'ประเด็น': 'Multicollinearity ปานกลาง: ' + ivName, 'รายละเอียด': 'VIF = ' + fmt(vif) + ' (เกณฑ์ < 5 ดี)', 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'S.E. อาจพองตัว ค่า t อาจไม่แม่นยำ'});
            });
        }

        // 5. Sample size adequacy
        var nPerIV = n / p;
        if (nPerIV < 10) warnings.push({'ประเด็น': 'ขนาดตัวอย่างน้อยเกินไป', 'รายละเอียด': 'N/IV = ' + fmt(nPerIV,1) + ' (แนะนำ >= 15-20)', 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'Tabachnick & Fidell แนะนำ N >= 50 + 8*IV (' + (50+8*p) + ' คน) สำหรับ R² และ N >= 104 + IV (' + (104+p) + ' คน) สำหรับ coefficients'});
        else if (nPerIV < 20) warnings.push({'ประเด็น': 'ขนาดตัวอย่างพอใช้', 'รายละเอียด': 'N/IV = ' + fmt(nPerIV,1), 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'แนะนำ N >= ' + (50+8*p) + ' (Tabachnick & Fidell) ปัจจุบัน N = ' + n});

        // 6. Non-significant predictors
        var nonSigCount = 0;
        mainRows.forEach(function(r, idx) {
            if (idx === 0) return; // skip intercept
            var pVal = parseFloat(r['p-value'] || r['Sig.']);
            if (!isNaN(pVal) && pVal >= 0.05) nonSigCount++;
        });
        if (nonSigCount > 0 && nonSigCount === selectedIVs.length) warnings.push({'ประเด็น': 'ไม่มีตัวแปรอิสระใดมีนัยสำคัญ', 'รายละเอียด': 'ตัวแปรอิสระทุกตัว p >= 0.05', 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'ไม่มีตัวแปรใดพยากรณ์ได้ ตรวจสอบ Multicollinearity หรือความเหมาะสมของตัวแปร'});
        else if (nonSigCount > 0) warnings.push({'ประเด็น': 'มีตัวแปรไม่มีนัยสำคัญ ' + nonSigCount + ' ตัว', 'รายละเอียด': 'จาก ' + selectedIVs.length + ' ตัว', 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'พิจารณาใช้ Stepwise/Forward เพื่อคัดเลือกเฉพาะตัวที่มีนัยสำคัญ'});

        if (warnings.length === 0) warnings.push({'ประเด็น': 'ไม่พบปัญหาเบื้องต้น', 'รายละเอียด': '—', 'ระดับ': '✅ ปกติ', 'คำแนะนำ': 'ตรวจสอบ Normality ของ Residual, Linearity, Homoscedasticity เพิ่มเติม'});
        return warnings;
    }

    // --- Regression diagnostic warnings builder (Multiple Regression) ---
    function _buildRegressionWarnings2(R2, adjR2, pF, dw, n, p, coeffRows, selectedIVs, selectedXData, Y) {
        var warnings = [];

        if (R2 < 0.1) warnings.push({'ประเด็น': 'R² ต่ำมาก', 'รายละเอียด': 'R² = ' + fmt(R2), 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'ตัวแปรอิสระอธิบายตัวแปรตามได้น้อยมาก'});
        if (adjR2 < 0) warnings.push({'ประเด็น': 'Adjusted R² ติดลบ', 'รายละเอียด': fmt(adjR2), 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'โมเดลไม่เหมาะสม'});
        if (pF >= 0.05) warnings.push({'ประเด็น': 'โมเดลไม่มีนัยสำคัญ', 'รายละเอียด': 'F-test p = ' + fmt(pF), 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'ตัวแปรอิสระไม่สามารถพยากรณ์ตัวแปรตามได้'});

        // Durbin-Watson
        if (dw < 1.5) warnings.push({'ประเด็น': 'Positive Autocorrelation', 'รายละเอียด': 'DW = ' + fmt(dw) + ' (ปกติ 1.5-2.5)', 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'Residual อาจมีความสัมพันธ์กัน'});
        if (dw > 2.5) warnings.push({'ประเด็น': 'Negative Autocorrelation', 'รายละเอียด': 'DW = ' + fmt(dw), 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'Residual อาจมีความสัมพันธ์เชิงลบ'});

        // VIF
        if (selectedIVs.length > 1) {
            selectedIVs.forEach(function(ivN, idx) {
                var xi = selectedXData[idx];
                var otherX = selectedXData.filter(function(_,j){return j!==idx;});
                var otherNms = ['(Intercept)'].concat(selectedIVs.filter(function(_,j){return j!==idx;}));
                var regI = Stats.linearRegression(xi, otherX, otherNms);
                var r2i = regI ? regI.rSquared : 0;
                var vif = (1-r2i)>0?1/(1-r2i):999;
                if (vif>10) warnings.push({'ประเด็น':'Multicollinearity สูงมาก: '+ivN,'รายละเอียด':'VIF = '+fmt(vif),'ระดับ':'🔴 สำคัญ','คำแนะนำ':'ควรตัดตัวแปรนี้ออก'});
                else if (vif>5) warnings.push({'ประเด็น':'Multicollinearity ปานกลาง: '+ivN,'รายละเอียด':'VIF = '+fmt(vif),'ระดับ':'🟡 ระวัง','คำแนะนำ':'S.E. อาจพองตัว'});
            });
        }

        // Sample size
        var nPerIV = n/p;
        if (nPerIV < 10) warnings.push({'ประเด็น':'ขนาดตัวอย่างน้อย','รายละเอียด':'N/IV = '+fmt(nPerIV,1),'ระดับ':'🔴 สำคัญ','คำแนะนำ':'แนะนำ N >= '+(50+8*p)+' (Tabachnick & Fidell)'});
        else if (nPerIV < 20) warnings.push({'ประเด็น':'ขนาดตัวอย่างพอใช้','รายละเอียด':'N/IV = '+fmt(nPerIV,1),'ระดับ':'🟡 ระวัง','คำแนะนำ':'N = '+n+', แนะนำ >= '+(50+8*p)});

        // Non-significant predictors
        var nonSig = 0;
        coeffRows.forEach(function(r,idx){if(idx===0)return;var pVal=parseFloat(r['Sig.']);if(!isNaN(pVal)&&pVal>=0.05)nonSig++;});
        if (nonSig>0 && nonSig===p) warnings.push({'ประเด็น':'ไม่มี IV ใดมีนัยสำคัญ','รายละเอียด':'ทุกตัว p >= 0.05','ระดับ':'🔴 สำคัญ','คำแนะนำ':'ลองใช้ Stepwise หรือตรวจ Multicollinearity'});
        else if (nonSig>0) warnings.push({'ประเด็น':'มี IV ไม่มีนัยสำคัญ '+nonSig+' ตัว','รายละเอียด':'จาก '+p+' ตัว','ระดับ':'🟡 ระวัง','คำแนะนำ':'พิจารณาตัดออก หรือใช้ Stepwise'});

        if (warnings.length===0) warnings.push({'ประเด็น':'ไม่พบปัญหาเบื้องต้น','รายละเอียด':'—','ระดับ':'✅ ปกติ','คำแนะนำ':'ตรวจ Normality/Linearity/Homoscedasticity เพิ่ม'});
        return warnings;
    }

    // =========================================================================
    // Logistic Regression
    // =========================================================================

    function runLogisticRegression() {
        var dv = getPickerValue('logr-dv-picker', 'logr-dv');
        var ivs = getCheckedVars('logr-iv-picker');
        if (ivs.length === 0) ivs = getSelected('logr-ivs');
        if (!dv || ivs.length === 0) { alert('Please select DV and at least one IV.'); return; }

        var method = getSelectValue('logr-method') || 'enter';
        var logrOpts = {};
        document.querySelectorAll('input[name="logr-opt"]').forEach(function(cb) { logrOpts[cb.value] = cb.checked; });

        var yData = getColumnData(dv, true);
        var xData = ivs.map(function (iv) { return getColumnData(iv, true); });
        var minLen = Math.min(yData.length, Math.min.apply(null, xData.map(function (x) { return x.length; })));
        yData = yData.slice(0, minLen);
        xData = xData.map(function (x) { return x.slice(0, minLen); });

        var names = ['(Intercept)'].concat(ivs);
        var result = Stats.logisticRegression(yData, xData, names);
        if (!result) { alert('Could not compute logistic regression. Check your data.'); return; }

        // Method label
        var methodLabels = {
            'enter': 'Enter', 'forward-wald': 'Forward: Wald', 'forward-lr': 'Forward: LR',
            'backward-wald': 'Backward: Wald', 'backward-lr': 'Backward: LR'
        };
        var methodLabel = methodLabels[method] || 'Enter';
        var methodDescs = {
            'enter': 'ใส่ตัวแปรอิสระทั้งหมดเข้าสมการพร้อมกัน',
            'forward-wald': 'เพิ่มตัวแปรทีละตัว โดยใช้ Wald statistic เป็นเกณฑ์',
            'forward-lr': 'เพิ่มตัวแปรทีละตัว โดยใช้ Likelihood Ratio เป็นเกณฑ์',
            'backward-wald': 'เริ่มจากทุกตัวแปร แล้วตัดทีละตัว โดยใช้ Wald statistic',
            'backward-lr': 'เริ่มจากทุกตัวแปร แล้วตัดทีละตัว โดยใช้ Likelihood Ratio'
        };

        // Simple assumption check inline
        var logrAssumptionEl = document.getElementById('logr-assumptions');
        if (logrAssumptionEl) {
            var n = yData.length;
            var nOk = n >= 10 * (ivs.length + 1);
            var checksHtml = '<div class="assumption-panel"><h4>🔍 Assumption Check</h4>';
            checksHtml += '<table class="result-table"><thead><tr><th>Check</th><th>Result</th><th>Status</th></tr></thead><tbody>';
            checksHtml += '<tr><td>Sample Size (10 per predictor)</td><td>N=' + n + ', needed>=' + (10*(ivs.length+1)) + '</td><td>' + (nOk ? '✅' : '⚠️') + '</td></tr>';
            checksHtml += '<tr><td>DV is Binary (0/1)</td><td>Check data</td><td>ℹ️</td></tr>';
            checksHtml += '</tbody></table></div>';
            logrAssumptionEl.innerHTML = checksHtml;
            logrAssumptionEl.style.display = 'block';
        }

        var extras = [];

        // 1. Method Information
        extras.push({
            title: 'Method / กระบวนการสร้างโมเดล',
            data: [{
                'Method': methodLabel,
                'Description': methodDescs[method] || method,
                'DV': dv,
                'IVs': ivs.join(', '),
                'N': yData.length,
                'N (Event=1)': yData.filter(function(v){return v>=0.5;}).length,
                'N (Event=0)': yData.filter(function(v){return v<0.5;}).length,
                'Convergence': result.converged !== false ? 'Yes' : 'No',
                'Iterations': result.iterations || 'N/A'
            }]
        });

        // 2. Model Summary
        var n = yData.length;
        var n1 = yData.filter(function(v){return v>=0.5;}).length;
        var n0 = n - n1;
        var nullLL = n1 * Math.log(n1/n) + n0 * Math.log(n0/n);
        var modelLL = result.logLikelihood || (result.aic ? -(result.aic/2 - ivs.length - 1) : nullLL);
        var chiSq = -2 * (nullLL - modelLL);
        var dfChi = ivs.length;
        var pChi = chiSq > 0 ? (1 - jStat.chisquare.cdf(chiSq, dfChi)) : 1;
        var coxSnell = 1 - Math.exp(-chiSq / n);
        var nagelkerke = coxSnell > 0 ? coxSnell / (1 - Math.exp(2 * nullLL / n)) : 0;

        extras.push({
            title: 'Model Summary',
            data: [{
                'Method': methodLabel,
                '-2 Log Likelihood': fmt(-2 * modelLL),
                'Chi-square': fmt(chiSq),
                'df': dfChi,
                'Sig.': fmt(pChi),
                "Cox & Snell R²": fmt(coxSnell),
                "Nagelkerke R²": fmt(nagelkerke),
                'Accuracy': fmt(result.accuracy * 100, 1) + '%',
                'AIC': fmt(result.aic)
            }]
        });

        // 3. Omnibus Test of Model Coefficients
        extras.push({
            title: 'Omnibus Tests of Model Coefficients',
            data: [
                { 'Test': 'Step', 'Chi-square': fmt(chiSq), 'df': dfChi, 'Sig.': fmt(pChi) },
                { 'Test': 'Block', 'Chi-square': fmt(chiSq), 'df': dfChi, 'Sig.': fmt(pChi) },
                { 'Test': 'Model', 'Chi-square': fmt(chiSq), 'df': dfChi, 'Sig.': fmt(pChi) }
            ]
        });

        // 4. Classification Table
        if (logrOpts.classification) {
            var tp=0,tn=0,fp=0,fn=0;
            yData.forEach(function(y,i) {
                var pred = result.coefficients[0].b;
                xData.forEach(function(xd,j) { pred += result.coefficients[j+1].b * xd[i]; });
                var prob = 1 / (1 + Math.exp(-pred));
                var predClass = prob >= 0.5 ? 1 : 0;
                var actual = y >= 0.5 ? 1 : 0;
                if (actual===1 && predClass===1) tp++;
                else if (actual===0 && predClass===0) tn++;
                else if (actual===0 && predClass===1) fp++;
                else fn++;
            });
            extras.push({
                title: 'Classification Table (Cut-off = 0.50)',
                data: [
                    { 'Observed': 'Event=0', 'Predicted 0': tn, 'Predicted 1': fp, '% Correct': fmt(tn/(tn+fp)*100,1)+'%' },
                    { 'Observed': 'Event=1', 'Predicted 0': fn, 'Predicted 1': tp, '% Correct': fmt(tp/(tp+fn)*100,1)+'%' },
                    { 'Observed': 'Overall', 'Predicted 0': '', 'Predicted 1': '', '% Correct': fmt((tp+tn)/n*100,1)+'%' }
                ]
            });
        }

        // 5. Coefficients (Variables in the Equation)
        var mainRows = result.coefficients.map(function (c) {
            var row = {
                Variable: c.variable,
                B: fmt(c.b), 'S.E.': fmt(c.se),
                Wald: fmt(c.wald), df: 1,
                'Sig.': Stats.formatPValue(c.p),
                'Exp(B) / OR': fmt(c.or)
            };
            if (logrOpts.ci) {
                var orLo = Math.exp(c.b - 1.96 * c.se);
                var orHi = Math.exp(c.b + 1.96 * c.se);
                row['95% CI Lower'] = fmt(orLo);
                row['95% CI Upper'] = fmt(orHi);
            }
            return row;
        });

        // 6. Logistic Regression Equation
        var logitParts = ['ln(P/(1-P)) = ' + fmt(result.coefficients[0].b)];
        result.coefficients.forEach(function(c, idx) {
            if (idx === 0) return;
            var sign = c.b >= 0 ? ' + ' : ' - ';
            logitParts.push(sign + fmt(Math.abs(c.b)) + '(' + c.variable + ')');
        });
        var logitEq = logitParts.join('');
        // Probability equation
        var probParts = ['P(Y=1) = 1 / (1 + e^-(' + fmt(result.coefficients[0].b)];
        result.coefficients.forEach(function(c, idx) {
            if (idx === 0) return;
            var sign = c.b >= 0 ? ' + ' : ' - ';
            probParts.push(sign + fmt(Math.abs(c.b)) + '*' + c.variable);
        });
        var probEq = probParts.join('') + '))';
        // OR interpretation
        var orInterpRows = [];
        result.coefficients.forEach(function(c, idx) {
            if (idx === 0) return;
            var orVal = c.or;
            var direction = orVal > 1 ? 'เพิ่มโอกาส' : (orVal < 1 ? 'ลดโอกาส' : 'ไม่มีผล');
            var pctChange = orVal > 1 ? fmt((orVal - 1) * 100, 1) + '% เพิ่มขึ้น' : fmt((1 - orVal) * 100, 1) + '% ลดลง';
            orInterpRows.push({
                'Variable': c.variable,
                'Exp(B) / OR': fmt(orVal),
                'ทิศทาง': direction,
                'แปลผล': 'เมื่อ ' + c.variable + ' เพิ่มขึ้น 1 หน่วย โอกาสเกิดเหตุการณ์ ' + pctChange,
                'Sig.': Stats.formatPValue(c.p),
                'มีนัยสำคัญ': c.p < 0.05 ? 'ใช่' : 'ไม่'
            });
        });

        extras.push({ title: 'สมการ Logistic Regression', data: [
            {'สมการ Logit': logitEq},
            {'สมการ Probability': probEq}
        ]});
        if (orInterpRows.length > 0) {
            extras.push({ title: 'การแปลผล Odds Ratio (OR) แต่ละตัวแปร', data: orInterpRows });
        }

        // 7. Diagnostic Warnings for Logistic Regression
        var logrWarnings = [];
        var nEvents = yData.filter(function(v){return v>=0.5;}).length;
        var nNonEvents = n - nEvents;
        var epp = Math.min(nEvents, nNonEvents) / ivs.length; // events per predictor
        if (epp < 10) logrWarnings.push({'ประเด็น': 'Events Per Predictor (EPP) ต่ำ', 'รายละเอียด': 'EPP = ' + fmt(epp,1) + ' (แนะนำ >= 10)', 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'ผลอาจไม่เสถียร ควรเพิ่มขนาดตัวอย่าง หรือลดจำนวนตัวแปรอิสระ'});
        else if (epp < 20) logrWarnings.push({'ประเด็น': 'Events Per Predictor พอใช้', 'รายละเอียด': 'EPP = ' + fmt(epp,1) + ' (เหมาะสม >= 20)', 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'ใช้ได้แต่ผลอาจมี bias เล็กน้อย'});
        if (result.accuracy < 0.6) logrWarnings.push({'ประเด็น': 'Model Accuracy ต่ำ', 'รายละเอียด': 'Accuracy = ' + fmt(result.accuracy*100,1) + '%', 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'โมเดลพยากรณ์ได้ไม่ดี ควรพิจารณาเพิ่มตัวแปร หรือตรวจสอบข้อมูล'});
        if (nagelkerke < 0.1) logrWarnings.push({'ประเด็น': 'Nagelkerke R² ต่ำมาก', 'รายละเอียด': 'R² = ' + fmt(nagelkerke), 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'ตัวแปรอิสระอธิบายตัวแปรตามได้น้อย'});
        if (pChi >= 0.05) logrWarnings.push({'ประเด็น': 'Omnibus Test ไม่มีนัยสำคัญ', 'รายละเอียด': 'Chi-square p = ' + fmt(pChi), 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'โมเดลโดยรวมไม่ดีกว่า Null Model ผลอาจไม่น่าเชื่อถือ'});
        // Check for extreme OR
        result.coefficients.forEach(function(c, idx) {
            if (idx === 0) return;
            if (c.or > 50) logrWarnings.push({'ประเด็น': 'OR สูงมาก: ' + c.variable, 'รายละเอียด': 'OR = ' + fmt(c.or), 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'อาจเกิดจาก quasi-complete separation หรือ sample size น้อย ตรวจสอบข้อมูล'});
            if (c.se > 5) logrWarnings.push({'ประเด็น': 'S.E. สูงมาก: ' + c.variable, 'รายละเอียด': 'S.E. = ' + fmt(c.se), 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'อาจเกิด complete separation ค่า B และ OR ไม่น่าเชื่อถือ'});
        });
        if (logrWarnings.length === 0) logrWarnings.push({'ประเด็น': 'ไม่พบปัญหาเบื้องต้น', 'รายละเอียด': '—', 'ระดับ': '✅ ปกติ', 'คำแนะนำ': 'ตรวจสอบ Assumption เพิ่มเติมตามบริบทงานวิจัย'});
        extras.push({ title: '⚠️ จุดสังเกต / ข้อควรระวัง (Diagnostics)', data: logrWarnings });

        state.results['logr'] = { data: mainRows, title: 'Variables in the Equation (Method: ' + methodLabel + ')', extras: extras };
        displayResults('logr');
    }

    // =========================================================================
    // Assumption Tests
    // =========================================================================

    function runAssumptionNormality() {
        var vars = getSelected('asn-norm-vars');
        if (vars.length === 0) { alert('Please select at least one variable.'); return; }

        var rows = [];
        vars.forEach(function (v) {
            var values = getColumnData(v, true);
            if (values.length < 3) return;
            var sw = Stats.shapiroWilk(values);
            var ks = Stats.ksTest(values);
            rows.push({
                Variable: v, N: values.length,
                'Shapiro-Wilk W': sw ? fmt(sw.W) : 'N/A',
                'S-W p-value': sw ? Stats.formatPValue(sw.p) : 'N/A',
                'K-S D': ks ? fmt(ks.D) : 'N/A',
                'K-S p-value': ks ? Stats.formatPValue(ks.p) : 'N/A',
                Conclusion: (sw && sw.p >= 0.05) ? 'Normal' : 'Not Normal'
            });
        });

        state.results['asn'] = { data: rows, title: 'Normality Test (Assumption Check)', extras: [] };
        displayResults('asn');
    }

    function runAssumptionLevene() {
        var dv = getPickerValue('asn-lev-dv-picker', 'asn-lev-dv');
        var iv = getPickerValue('asn-lev-iv-picker', 'asn-lev-iv');
        if (!dv || !iv) { alert('Please select both variables.'); return; }

        var split = splitByGroup(dv, iv);
        var gNames = split.groupNames;
        if (gNames.length < 2) { alert('Need at least 2 groups.'); return; }

        var groups = gNames.map(function (g) { return split.groups[g]; });
        // Use Stats' internals via independentTTest for 2 groups, or onewayAnova
        var result = Stats.onewayAnova(groups, gNames);
        if (!result) { alert('Could not compute Levene test.'); return; }

        var rows = [{
            'Levene F': fmt(result.leveneF),
            'p-value': Stats.formatPValue(result.leveneP),
            Conclusion: result.leveneP >= 0.05 ? 'Homogeneous variances (equal)' : 'Non-homogeneous variances (unequal)'
        }];

        state.results['asn'] = { data: rows, title: "Levene's Test for Equality of Variances", extras: [] };
        displayResults('asn');
    }

    function runAssumptionVif() {
        var vars = getSelected('asn-vif-vars');
        if (vars.length < 2) { alert('Please select at least 2 variables.'); return; }

        var dataArrays = vars.map(function (v) { return getColumnData(v, true); });
        var minLen = Math.min.apply(null, dataArrays.map(function (a) { return a.length; }));
        dataArrays = dataArrays.map(function (a) { return a.slice(0, minLen); });

        var rows = [];
        vars.forEach(function (targetVar, idx) {
            // Regress targetVar on all other vars
            var y = dataArrays[idx];
            var xs = [];
            var xNames = ['(Intercept)'];
            for (var j = 0; j < vars.length; j++) {
                if (j === idx) continue;
                xs.push(dataArrays[j]);
                xNames.push(vars[j]);
            }

            var result = Stats.linearRegression(y, xs, xNames);
            var vif = 'N/A';
            var tolerance = 'N/A';
            if (result && !isNaN(result.rSquared) && result.rSquared < 1) {
                var tol = 1 - result.rSquared;
                tolerance = fmt(tol);
                vif = fmt(1 / tol);
            }

            rows.push({ Variable: targetVar, Tolerance: tolerance, VIF: vif });
        });

        state.results['asn'] = { data: rows, title: 'Variance Inflation Factor (VIF)', extras: [] };
        displayResults('asn');
    }

    // =========================================================================
    // Reliability (Cronbach's Alpha)
    // =========================================================================

    function runReliability() {
        var vars = getCheckedVars('rel-picker');
        if (vars.length === 0) vars = getDualListSelected('rel-dual');
        if (!vars || vars.length === 0) vars = getSelected('rel-vars');
        if (vars.length < 2) { alert('Please select at least 2 variables.'); return; }

        var dataArrays = vars.map(function (v) { return getColumnData(v, true); });
        var minLen = Math.min.apply(null, dataArrays.map(function (a) { return a.length; }));
        dataArrays = dataArrays.map(function (a) { return a.slice(0, minLen); });

        var result = Stats.cronbachAlpha(dataArrays);
        if (!result) { alert('Could not compute Cronbach Alpha. Check your data.'); return; }

        // Patch item names
        result.itemStats.forEach(function (item, i) { item.item = vars[i]; });

        var extras = [{
            title: 'Reliability Summary',
            html: '<div class="metric-cards">' +
                  '<div class="metric-card"><div class="metric-value">' + fmt(result.alpha) + '</div><div class="metric-label">Cronbach\'s Alpha</div></div>' +
                  '<div class="metric-card"><div class="metric-value">' + result.interpretation + '</div><div class="metric-label">Interpretation</div></div>' +
                  '<div class="metric-card"><div class="metric-value">' + result.nItems + '</div><div class="metric-label">Number of Items</div></div>' +
                  '<div class="metric-card"><div class="metric-value">' + result.nCases + '</div><div class="metric-label">Number of Cases</div></div>' +
                  '</div>'
        }];

        var mainRows = result.itemStats.map(function (item) {
            return {
                Item: item.item,
                Mean: fmt(item.mean),
                'S.D.': fmt(item.sd),
                'Item-Total r': fmt(item.itemTotalR),
                'Alpha if Deleted': fmt(item.alphaIfDeleted),
                'Suggest Delete?': item.shouldDelete ? 'Yes' : 'No'
            };
        });

        state.results['rel'] = { data: mainRows, title: 'Item-Total Statistics', extras: extras };
        displayResults('rel');
    }

    // =========================================================================
    // Effect Size
    // =========================================================================

    function runEffectCohen() {
        var dv = getPickerValue('cd-dv-picker', 'cd-dv');
        var iv = getPickerValue('cd-iv-picker', 'cd-iv');
        if (!dv || !iv) { alert('Please select both variables.'); return; }

        var split = splitByGroup(dv, iv);
        var gNames = split.groupNames;
        if (gNames.length !== 2) { alert("Cohen's d requires exactly 2 groups. Found " + gNames.length + '.'); return; }

        var g1 = split.groups[gNames[0]];
        var g2 = split.groups[gNames[1]];
        var result = Stats.independentTTest(g1, g2);
        if (!result) { alert('Could not compute effect size.'); return; }

        var mainRows = [{
            'Group 1': gNames[0] + ' (N=' + result.desc1.n + ', M=' + fmt(result.desc1.mean) + ')',
            'Group 2': gNames[1] + ' (N=' + result.desc2.n + ', M=' + fmt(result.desc2.mean) + ')',
            'Mean Diff': fmt(result.meanDiff),
            "Cohen's d": fmt(result.cohensD),
            Interpretation: result.effectSize
        }];

        state.results['cd'] = { data: mainRows, title: "Cohen's d Effect Size", extras: [] };
        displayResults('cd');
    }

    function runEffectOdds() {
        var var1 = getPickerValue('cd-or-var1-picker', 'cd-or-var1');
        var var2 = getPickerValue('cd-or-var2-picker', 'cd-or-var2');
        if (!var1 || !var2) { alert('Please select both variables.'); return; }

        var v1 = getColumnData(var1, false);
        var v2 = getColumnData(var2, false);

        // Build 2x2 contingency table
        var cats1 = [], cats2 = [];
        var c1Set = {}, c2Set = {};
        for (var i = 0; i < Math.min(v1.length, v2.length); i++) {
            var a = String(v1[i]), b = String(v2[i]);
            if (!c1Set[a]) { c1Set[a] = true; cats1.push(a); }
            if (!c2Set[b]) { c2Set[b] = true; cats2.push(b); }
        }

        if (cats1.length !== 2 || cats2.length !== 2) {
            alert('Odds Ratio requires exactly 2 categories per variable. Got ' + cats1.length + ' and ' + cats2.length + '.');
            return;
        }

        var table = [[0, 0], [0, 0]];
        for (var i = 0; i < Math.min(v1.length, v2.length); i++) {
            var r = cats1.indexOf(String(v1[i]));
            var c = cats2.indexOf(String(v2[i]));
            if (r >= 0 && c >= 0) table[r][c]++;
        }

        var a = table[0][0], b = table[0][1], c = table[1][0], d = table[1][1];
        var or = (b * c) > 0 ? (a * d) / (b * c) : NaN;
        var logOR = !isNaN(or) && or > 0 ? Math.log(or) : NaN;
        var seLogOR = NaN;
        if (a > 0 && b > 0 && c > 0 && d > 0) {
            seLogOR = Math.sqrt(1 / a + 1 / b + 1 / c + 1 / d);
        }
        var ci95Lo = !isNaN(logOR) && !isNaN(seLogOR) ? Math.exp(logOR - 1.96 * seLogOR) : NaN;
        var ci95Hi = !isNaN(logOR) && !isNaN(seLogOR) ? Math.exp(logOR + 1.96 * seLogOR) : NaN;

        var extras = [{
            title: '2x2 Contingency Table',
            data: [
                { '': cats1[0], [cats2[0]]: a, [cats2[1]]: b },
                { '': cats1[1], [cats2[0]]: c, [cats2[1]]: d }
            ]
        }];

        var mainRows = [{
            'Odds Ratio': fmt(or),
            'ln(OR)': fmt(logOR),
            'SE(ln OR)': fmt(seLogOR),
            '95% CI': Stats.formatCI ? Stats.formatCI(ci95Lo, ci95Hi) : '[' + fmt(ci95Lo) + ', ' + fmt(ci95Hi) + ']'
        }];

        state.results['cd'] = { data: mainRows, title: 'Odds Ratio', extras: extras };
        displayResults('cd');
    }

    // =========================================================================
    // Factor Analysis (EFA)
    // =========================================================================

    function runFactorAnalysis() {
        var vars = getCheckedVars('efa-picker');
        if (vars.length < 3) { alert('กรุณาเลือกตัวแปรอย่างน้อย 3 ตัว'); return; }
        var dataArrays = vars.map(function(v) { return getColumnData(v, true); });
        var result = Stats.factorAnalysis(dataArrays, vars);
        if (!result) { alert('ไม่สามารถวิเคราะห์ได้'); return; }

        // KMO & Bartlett's
        var extras = [];
        extras.push({
            title: 'KMO & Sampling Adequacy',
            data: [{ 'KMO': fmt(result.kmo), 'Interpretation': result.kmoInterpretation,
                     'Suitable for FA': result.kmo >= 0.5 ? '✅ Yes' : '⚠️ No' }]
        });

        // Total Variance Explained
        var varianceRows = result.components.map(function(c) {
            return {
                'Component': c.component, 'Eigenvalue': fmt(c.eigenvalue),
                '% of Variance': fmt(c.pctVariance, 1), 'Cumulative %': fmt(c.cumPctVariance, 1),
                'Extract': c.eigenvalue >= 1 ? '✓' : ''
            };
        });
        extras.push({ title: 'Total Variance Explained', data: varianceRows });

        // Summary
        var summaryRows = [{
            'Variables': vars.length, 'Factors Extracted (Kaiser)': result.nFactors,
            'Total Variance Explained': fmt(result.totalVarianceExplained, 1) + '%',
            'KMO': fmt(result.kmo)
        }];

        state.results['efa'] = { data: summaryRows, title: 'Factor Analysis Summary', extras: extras };
        displayResults('efa');
    }

    // =========================================================================
    // Cross Tabulation
    // =========================================================================

    function runCrossTab() {
        var var1 = getPickerValue('ct-var1-picker', 'ct-var1');
        var var2 = getPickerValue('ct-var2-picker', 'ct-var2');
        if (!var1 || !var2) { alert('กรุณาเลือกตัวแปร 2 ตัว'); return; }

        var data1 = state.data.map(function(row) { return row[var1]; });
        var data2 = state.data.map(function(row) { return row[var2]; });
        var result = Stats.crossTab(data1, data2);
        if (!result) { alert('ไม่สามารถวิเคราะห์ได้'); return; }

        var opts = getChecked('ct-opt');
        var extras = [];

        // Observed frequencies table
        var obsRows = [];
        for (var i = 0; i < result.rowLabels.length; i++) {
            var row = {};
            row[var1] = result.rowLabels[i];
            for (var j = 0; j < result.colLabels.length; j++) {
                var cell = String(result.observed[i][j]);
                if (opts.indexOf('rowpct') !== -1) cell += ' (' + fmt(result.rowPct[i][j], 1) + '%)';
                row[result.colLabels[j]] = cell;
            }
            row['Total'] = result.rowTotals[i];
            obsRows.push(row);
        }
        // Total row
        var totalRow = {};
        totalRow[var1] = 'Total';
        for (var j = 0; j < result.colLabels.length; j++) totalRow[result.colLabels[j]] = result.colTotals[j];
        totalRow['Total'] = result.n;
        obsRows.push(totalRow);
        extras.push({ title: 'Observed Frequencies', data: obsRows });

        // Expected
        if (opts.indexOf('expected') !== -1) {
            var expRows = [];
            for (var i = 0; i < result.rowLabels.length; i++) {
                var row = {};
                row[var1] = result.rowLabels[i];
                for (var j = 0; j < result.colLabels.length; j++) {
                    row[result.colLabels[j]] = fmt(result.expected[i][j], 1);
                }
                expRows.push(row);
            }
            extras.push({ title: 'Expected Frequencies', data: expRows });
        }

        // Std Residuals
        if (opts.indexOf('residuals') !== -1) {
            var resRows = [];
            for (var i = 0; i < result.rowLabels.length; i++) {
                var row = {};
                row[var1] = result.rowLabels[i];
                for (var j = 0; j < result.colLabels.length; j++) {
                    row[result.colLabels[j]] = fmt(result.stdResiduals[i][j]);
                }
                resRows.push(row);
            }
            extras.push({ title: 'Standardized Residuals', data: resRows });
        }

        // Chi-square test result
        var mainRows = [{
            'Chi-Square (χ²)': fmt(result.chi2), 'df': result.df,
            'p-value': Stats.formatPValue(result.p), 'Phi (φ)': fmt(result.phi),
            "Cramer's V": fmt(result.cramersV), 'Effect': result.interpretation,
            'N': result.n, 'Sig.': result.p < 0.05 ? '✓' : ''
        }];

        state.results['ct'] = { data: mainRows, title: 'Cross Tabulation — Chi-Square Test', extras: extras };
        displayResults('ct');
    }

    // =========================================================================
    // Multiple Comparisons (Post-hoc)
    // =========================================================================

    function runPostHoc() {
        var dv = getPickerValue('ph-dv-picker', 'ph-dv');
        var iv = getPickerValue('ph-iv-picker', 'ph-iv');
        if (!dv || !iv) { alert('กรุณาเลือก DV และ IV'); return; }
        var method = getSelectValue('ph-method') || 'tukey';

        var split = splitByGroup(dv, iv);
        var gNames = split.groupNames;
        if (gNames.length < 3) { alert('ต้องมีอย่างน้อย 3 กลุ่ม'); return; }

        var groups = gNames.map(function(name) { return split.groups[name]; });

        var result;
        if (method === 'tukey') result = Stats.tukeyHSD(groups, gNames);
        else if (method === 'bonferroni') result = Stats.bonferroni(groups, gNames);
        else result = Stats.scheffeTest(groups, gNames);

        if (!result) { alert('ไม่สามารถวิเคราะห์ได้'); return; }

        var rows = result.map(function(r) {
            return {
                'Group A': r.groupA, 'Group B': r.groupB,
                'Mean A': fmt(r.meanA), 'Mean B': fmt(r.meanB),
                'Mean Diff': fmt(r.meanDiff),
                't / F': fmt(r.t || r.F),
                'p-value': Stats.formatPValue(r.p),
                'p (adjusted)': Stats.formatPValue(r.pAdjusted || r.p),
                'Sig.': r.significant ? '✓' : ''
            };
        });

        var methodName = method === 'tukey' ? 'Tukey HSD' : method === 'bonferroni' ? 'Bonferroni' : 'Scheffe';
        state.results['ph'] = { data: rows, title: 'Multiple Comparisons — ' + methodName };
        displayResults('ph');
    }

    // =========================================================================
    // Bootstrap CI
    // =========================================================================

    function runBootstrap() {
        var vars = getCheckedVars('boot-picker');
        if (vars.length === 0) { alert('กรุณาเลือกตัวแปร'); return; }
        var statType = getSelectValue('boot-stat') || 'mean';
        var nBoot = parseInt(document.getElementById('boot-n').value) || 1000;

        var statFn;
        var statLabel;
        if (statType === 'mean') {
            statFn = function(arr) { return arr.reduce(function(a,b){return a+b;},0) / arr.length; };
            statLabel = 'Mean';
        } else if (statType === 'median') {
            statFn = function(arr) { var s = arr.slice().sort(function(a,b){return a-b;}); return s.length%2===0 ? (s[s.length/2-1]+s[s.length/2])/2 : s[Math.floor(s.length/2)]; };
            statLabel = 'Median';
        } else {
            statFn = function(arr) { var m = arr.reduce(function(a,b){return a+b;},0)/arr.length; return Math.sqrt(arr.reduce(function(a,b){return a+(b-m)*(b-m);},0)/(arr.length-1)); };
            statLabel = 'S.D.';
        }

        var rows = [];
        vars.forEach(function(v) {
            var values = getColumnData(v, true);
            if (values.length < 2) return;
            var result = Stats.bootstrapCI(values, statFn, nBoot);
            if (!result) return;
            rows.push({
                'Variable': v, 'Statistic': statLabel,
                'Estimate': fmt(result.estimate),
                'Bootstrap S.E.': fmt(result.se),
                'Bias': fmt(result.bias),
                '95% CI': result.ci95,
                'N Bootstrap': result.nBoot
            });
        });

        state.results['boot'] = { data: rows, title: 'Bootstrap Confidence Interval (' + statLabel + ', B=' + nBoot + ')' };
        displayResults('boot');
    }

    // =========================================================================
    // Z-Score & Percentile
    // =========================================================================

    function runZScore() {
        var vars = getCheckedVars('zs-picker');
        if (vars.length === 0) { alert('กรุณาเลือกตัวแปร'); return; }

        var extras = [];
        var summaryRows = [];

        vars.forEach(function(v) {
            var values = getColumnData(v, true);
            if (values.length < 2) return;
            var zScores = Stats.zScores(values);
            var desc = Stats.descriptive(values);

            summaryRows.push({
                'Variable': v, 'N': desc.n, 'Mean': fmt(desc.mean), 'S.D.': fmt(desc.sd),
                'Min Z': fmt(Math.min.apply(null, zScores)),
                'Max Z': fmt(Math.max.apply(null, zScores)),
                'P25': fmt(desc.p25), 'P50 (Median)': fmt(desc.median), 'P75': fmt(desc.p75)
            });

            // Detailed z-scores (first 50 rows)
            if (vars.length === 1) {
                var detailRows = values.slice(0, 50).map(function(val, i) {
                    return {
                        'Row': i + 1, 'Value': fmt(val, 2), 'Z-Score': fmt(zScores[i]),
                        'Percentile': fmt(Stats.percentileRank(values, val), 1) + '%'
                    };
                });
                extras.push({ title: 'Z-Scores (first 50 rows)', data: detailRows });
            }
        });

        state.results['zs'] = { data: summaryRows, title: 'Z-Score & Percentile Summary', extras: extras };
        displayResults('zs');
    }

    // =========================================================================
    // Multicollinearity (VIF)
    // =========================================================================

    function runVIF() {
        var vars = getCheckedVars('vifp-picker');
        if (vars.length < 2) { alert('กรุณาเลือกตัวแปรอย่างน้อย 2 ตัว'); return; }

        var dataArrays = vars.map(function(v) { return getColumnData(v, true); });
        var result = Stats.vif(dataArrays, vars);
        if (!result) { alert('ไม่สามารถคำนวณ VIF ได้'); return; }

        var rows = result.map(function(r) {
            return {
                'Variable': r.variable,
                'R²': fmt(r.rSquared),
                'Tolerance': fmt(r.tolerance),
                'VIF': fmt(r.vif),
                'Status': r.status
            };
        });

        var hasIssue = result.some(function(r) { return r.vif > 5; });
        var extras = [];
        extras.push({
            title: '', html: '<div class="detail-box"><strong>สรุป:</strong> ' +
                (hasIssue ? '⚠️ พบ Multicollinearity — พิจารณาลบตัวแปรที่ VIF > 10 ออก' : '✅ ไม่พบปัญหา Multicollinearity (VIF ทุกตัว < 5)') +
                '<br>เกณฑ์: VIF < 5 = OK, 5-10 = Moderate, > 10 = Severe</div>'
        });

        state.results['vifp'] = { data: rows, title: 'Multicollinearity Analysis (VIF)', extras: extras };
        displayResults('vifp');
    }

    // =========================================================================
    // Charts & Visualization
    // =========================================================================

    function runCharts() {
        var vars = getCheckedVars('chart-picker');
        if (vars.length === 0) { alert('กรุณาเลือกตัวแปร'); return; }
        var chartType = getSelectValue('chart-type') || 'histogram';
        var canvas = document.getElementById('chart-canvas');
        if (!canvas) return;

        // Destroy previous chart
        if (window._currentChart) window._currentChart.destroy();

        var values = getColumnData(vars[0], true);

        if (chartType === 'histogram') {
            var nBins = Math.min(Math.ceil(Math.sqrt(values.length)), 30);
            var min = Math.min.apply(null, values), max = Math.max.apply(null, values);
            var binW = (max - min) / nBins || 1;
            var bins = []; for(var i=0;i<nBins;i++) bins.push({lo:min+i*binW,hi:min+(i+1)*binW,count:0});
            values.forEach(function(v){for(var i=0;i<bins.length;i++){if(v>=bins[i].lo&&(i===bins.length-1?v<=bins[i].hi:v<bins[i].hi)){bins[i].count++;break;}}});
            window._currentChart = new Chart(canvas, {
                type: 'bar',
                data: { labels: bins.map(function(b){return fmt(b.lo,1)+'-'+fmt(b.hi,1);}), datasets: [{label:vars[0],data:bins.map(function(b){return b.count;}),backgroundColor:'rgba(37,99,235,0.6)',borderColor:'#2563eb',borderWidth:1}] },
                options: { responsive:true, plugins:{title:{display:true,text:'Histogram: '+vars[0]}} }
            });
        } else if (chartType === 'boxplot') {
            var datasets = vars.map(function(v,i){
                var d = getColumnData(v,true).sort(function(a,b){return a-b;});
                var q1=d[Math.floor(d.length*0.25)],med=d[Math.floor(d.length*0.5)],q3=d[Math.floor(d.length*0.75)];
                var colors = ['rgba(37,99,235,0.6)','rgba(5,150,105,0.6)','rgba(220,38,38,0.6)','rgba(124,58,237,0.6)'];
                return {label:v,data:[{min:d[0],q1:q1,median:med,q3:q3,max:d[d.length-1]}],backgroundColor:colors[i%4]};
            });
            // Fallback: show as bar chart with min/q1/median/q3/max
            var labels = ['Min','Q1','Median','Q3','Max'];
            var ds = vars.map(function(v,i){
                var d = getColumnData(v,true).sort(function(a,b){return a-b;});
                var colors = ['rgba(37,99,235,0.6)','rgba(5,150,105,0.6)','rgba(220,38,38,0.6)'];
                return {label:v,data:[d[0],d[Math.floor(d.length*0.25)],d[Math.floor(d.length*0.5)],d[Math.floor(d.length*0.75)],d[d.length-1]],backgroundColor:colors[i%3]};
            });
            window._currentChart = new Chart(canvas, {
                type:'bar', data:{labels:labels,datasets:ds},
                options:{responsive:true,plugins:{title:{display:true,text:'Box Plot Summary'}}}
            });
        } else if (chartType === 'scatter' && vars.length >= 2) {
            var xVals = getColumnData(vars[0],true), yVals = getColumnData(vars[1],true);
            var n = Math.min(xVals.length,yVals.length);
            var pts = []; for(var i=0;i<n;i++) pts.push({x:xVals[i],y:yVals[i]});
            window._currentChart = new Chart(canvas, {
                type:'scatter', data:{datasets:[{label:vars[0]+' vs '+vars[1],data:pts,backgroundColor:'rgba(37,99,235,0.5)',pointRadius:4}]},
                options:{responsive:true,plugins:{title:{display:true,text:'Scatter: '+vars[0]+' vs '+vars[1]}},scales:{x:{title:{display:true,text:vars[0]}},y:{title:{display:true,text:vars[1]}}}}
            });
        } else if (chartType === 'bar') {
            var ds = vars.map(function(v,i){
                var d = Stats.descriptive(getColumnData(v,true));
                var colors = ['rgba(37,99,235,0.7)','rgba(5,150,105,0.7)','rgba(220,38,38,0.7)','rgba(124,58,237,0.7)'];
                return {label:v,data:[d.mean],backgroundColor:colors[i%4],borderWidth:1};
            });
            window._currentChart = new Chart(canvas, {
                type:'bar', data:{labels:['Mean'],datasets:ds},
                options:{responsive:true,plugins:{title:{display:true,text:'Mean Comparison'}}}
            });
        }
    }

    // =========================================================================
    // Partial Correlation
    // =========================================================================

    function runPartialCorr() {
        var xVar = getPickerValue('pcor-x-picker','pcor-x');
        var yVar = getPickerValue('pcor-y-picker','pcor-y');
        var ctrlVars = getCheckedVars('pcor-ctrl-picker');
        if (!xVar || !yVar) { alert('กรุณาเลือก X และ Y'); return; }

        var x = getColumnData(xVar,true), y = getColumnData(yVar,true);
        var controls = ctrlVars.map(function(v){return getColumnData(v,true);});

        var zeroOrder = Stats.correlation([x,y],[xVar,yVar],'pearson');
        var result = Stats.partialCorrelation(x,y,controls.length>0?controls:null);
        if (!result) { alert('ไม่สามารถคำนวณได้'); return; }

        var extras = [];
        if (zeroOrder && zeroOrder.pairs) {
            extras.push({title:'Zero-Order Correlation',data:[{
                Variables:xVar+' & '+yVar, r:fmt(zeroOrder.pairs[0].r), 'p-value':Stats.formatPValue(zeroOrder.pairs[0].p)
            }]});
        }

        var rows = [{
            'Variable X':xVar, 'Variable Y':yVar,
            'Control':ctrlVars.join(', ')||'None',
            'Partial r':fmt(result.r), 'r²':fmt(result.rSquared),
            'df':result.df, 't':fmt(result.t),
            'p-value':Stats.formatPValue(result.p),
            'Sig.':result.p<0.05?'✓':''
        }];

        state.results['pcor'] = {data:rows, title:'Partial Correlation', extras:extras};
        displayResults('pcor');
    }

    // =========================================================================
    // Hierarchical Regression
    // =========================================================================

    function runHierarchicalReg() {
        var dv = getPickerValue('hreg-dv-picker','hreg-dv');
        var b1Vars = getCheckedVars('hreg-b1-picker');
        var b2Vars = getCheckedVars('hreg-b2-picker');
        if (!dv || b1Vars.length===0) { alert('กรุณาเลือก DV และ Block 1'); return; }

        var y = getColumnData(dv,true);
        var blocks = [b1Vars.map(function(v){return getColumnData(v,true);})];
        var names = [b1Vars];
        if (b2Vars.length > 0) {
            blocks.push(b2Vars.map(function(v){return getColumnData(v,true);}));
            names.push(b2Vars);
        }

        var steps = Stats.hierarchicalRegression(y, blocks, names);
        if (!steps || steps.length === 0) { alert('ไม่สามารถวิเคราะห์ได้'); return; }

        var extras = [];

        // 1. Method Information
        var blockInfo = 'Block 1: ' + b1Vars.join(', ');
        if (b2Vars.length > 0) blockInfo += ' → Block 2: ' + b2Vars.join(', ');
        extras.push({
            title: 'Method / กระบวนการสร้างโมเดล',
            data: [{
                'Method': 'Hierarchical (Enter per Block)',
                'Description': 'ใส่ตัวแปรเป็นลำดับขั้น (Block) เพื่อดูว่าแต่ละ Block เพิ่มอำนาจพยากรณ์ (R² Change) ได้เท่าไหร่',
                'DV': dv,
                'Block Structure': blockInfo,
                'Number of Blocks': blocks.length,
                'Total IVs': b1Vars.length + b2Vars.length,
                'N': y.length
            }]
        });

        // 2. Model Comparison Summary
        var modelRows = steps.map(function(s) {
            return {
                'Model/Step':s.step, 'R':fmt(s.r), 'R²':fmt(s.rSquared), 'Adj R²':fmt(s.adjRSquared),
                'R² Change':fmt(s.r2Change), 'F Change':fmt(s.fChange),
                'df1':s.df1Change, 'df2':s.df2Change,
                'Sig. F Change':Stats.formatPValue(s.pChange),
                'Variables in Model':s.varsAdded.join(', '),
                'Interpretation': s.pChange < 0.05 ? 'Block นี้เพิ่ม R² อย่างมีนัยสำคัญ' : 'Block นี้ไม่เพิ่ม R² อย่างมีนัยสำคัญ'
            };
        });

        // 3. Coefficients per step + Equations
        steps.forEach(function(s) {
            if (s.coefficients) {
                var coefRows = s.coefficients.map(function(c){
                    return {Variable:c.variable,B:fmt(c.b),'S.E.':fmt(c.se),'Beta (Std.)':c.beta!==undefined?fmt(c.beta):'',t:fmt(c.t),'Sig.':Stats.formatPValue(c.p),'95% CI':c.ci95||''};
                });
                extras.push({title:'Model '+s.step+' — Coefficients (Variables: '+s.varsAdded.join(', ')+')',data:coefRows});

                // Equation for this step
                var eqParts = [dv + ' = ' + fmt(s.coefficients[0].b)];
                s.coefficients.forEach(function(c, idx) {
                    if (idx === 0) return;
                    var sign = c.b >= 0 ? ' + ' : ' - ';
                    eqParts.push(sign + fmt(Math.abs(c.b)) + '(' + c.variable + ')');
                });
                var stdParts = [dv + '(Z) ='];
                s.coefficients.forEach(function(c, idx) {
                    if (idx === 0) return;
                    var betaVal = c.beta !== undefined ? c.beta : 0;
                    var sign = betaVal >= 0 ? (idx === 1 ? ' ' : ' + ') : ' - ';
                    stdParts.push(sign + fmt(Math.abs(betaVal)) + '(Z_' + c.variable + ')');
                });
                extras.push({title:'Model '+s.step+' — สมการถดถอย',data:[
                    {'Unstandardized': eqParts.join('')},
                    {'Standardized': stdParts.join('')}
                ]});
            }
        });

        // 4. Diagnostic Warnings
        var hregWarnings = [];
        var lastStep = steps[steps.length - 1];
        if (lastStep) {
            if (lastStep.rSquared < 0.1) hregWarnings.push({'ประเด็น': 'R² รวมต่ำมาก', 'รายละเอียด': 'R² = ' + fmt(lastStep.rSquared), 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'ตัวแปรอิสระทั้งหมดอธิบายตัวแปรตามได้น้อย ควรพิจารณาเพิ่มตัวแปร'});
            if (lastStep.adjRSquared < 0) hregWarnings.push({'ประเด็น': 'Adjusted R² ติดลบ', 'รายละเอียด': 'Adj R² = ' + fmt(lastStep.adjRSquared), 'ระดับ': '🔴 สำคัญ', 'คำแนะนำ': 'โมเดลแย่กว่าค่าเฉลี่ย ตัวแปรที่ใส่อาจไม่เหมาะสม'});
        }
        steps.forEach(function(s, si) {
            if (si > 0 && s.pChange >= 0.05) hregWarnings.push({'ประเด็น': 'Block ' + s.step + ' ไม่มีนัยสำคัญ', 'รายละเอียด': 'R² Change = ' + fmt(s.r2Change) + ', p = ' + Stats.formatPValue(s.pChange), 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'ตัวแปรใน Block นี้ไม่เพิ่มอำนาจพยากรณ์ อาจตัดออกได้'});
        });
        var nTotal = y.length;
        var pTotal = b1Vars.length + b2Vars.length;
        if (nTotal / pTotal < 15) hregWarnings.push({'ประเด็น': 'อัตราส่วน N/IV ต่ำ', 'รายละเอียด': 'N=' + nTotal + ', IVs=' + pTotal + ' (ratio=' + fmt(nTotal/pTotal,1) + ')', 'ระดับ': '🟡 ระวัง', 'คำแนะนำ': 'แนะนำ N/IV >= 15 สำหรับผลที่เสถียร (Tabachnick & Fidell)'});
        if (hregWarnings.length === 0) hregWarnings.push({'ประเด็น': 'ไม่พบปัญหาเบื้องต้น', 'รายละเอียด': '—', 'ระดับ': '✅ ปกติ', 'คำแนะนำ': 'ตรวจสอบ Normality ของ Residual และ Linearity เพิ่มเติม'});
        extras.push({ title: '⚠️ จุดสังเกต / ข้อควรระวัง (Diagnostics)', data: hregWarnings });

        state.results['hreg'] = {data:modelRows, title:'Hierarchical Regression — Model Comparison (Method: Hierarchical Enter)', extras:extras};
        displayResults('hreg');
    }

    // =========================================================================
    // ROC Curve / AUC
    // =========================================================================

    function runROC() {
        var actualVar = getPickerValue('roc-actual-picker','roc-actual');
        var predVar = getPickerValue('roc-pred-picker','roc-pred');
        if (!actualVar || !predVar) { alert('กรุณาเลือก Actual และ Predicted'); return; }

        var actual = getColumnData(actualVar,true).map(function(v){return v>=0.5?1:0;});
        var predicted = getColumnData(predVar,true);
        var n = Math.min(actual.length,predicted.length);
        actual=actual.slice(0,n); predicted=predicted.slice(0,n);

        var result = Stats.roc(actual,predicted);
        if (!result) { alert('ไม่สามารถวิเคราะห์ได้'); return; }

        // Draw ROC curve
        var canvas = document.getElementById('roc-canvas');
        if (canvas && window.Chart) {
            if (window._rocChart) window._rocChart.destroy();
            var pts = result.points.map(function(p){return {x:p.fpr,y:p.tpr};});
            window._rocChart = new Chart(canvas, {
                type:'scatter',
                data:{datasets:[
                    {label:'ROC Curve (AUC='+fmt(result.auc,3)+')',data:pts,showLine:true,fill:false,borderColor:'#2563eb',backgroundColor:'rgba(37,99,235,0.1)',pointRadius:0,borderWidth:2},
                    {label:'Reference',data:[{x:0,y:0},{x:1,y:1}],showLine:true,borderColor:'#dc2626',borderDash:[5,5],pointRadius:0,borderWidth:1}
                ]},
                options:{responsive:true,scales:{x:{title:{display:true,text:'1 - Specificity (FPR)'},min:0,max:1},y:{title:{display:true,text:'Sensitivity (TPR)'},min:0,max:1}},plugins:{title:{display:true,text:'ROC Curve'}}}
            });
        }

        var rows = [{
            'AUC':fmt(result.auc,3), 'Interpretation':result.interpretation,
            'Optimal Threshold':fmt(result.optimalThreshold,3),
            'Sensitivity':fmt(result.sensitivity,3), 'Specificity':fmt(result.specificity,3),
            'PPV':fmt(result.ppv,3), 'NPV':fmt(result.npv,3),
            'Accuracy':fmt(result.accuracy*100,1)+'%'
        }];

        state.results['roc'] = {data:rows, title:'ROC Analysis'};
        displayResults('roc');
    }

    // =========================================================================
    // ICC
    // =========================================================================

    function runICC() {
        var vars = getCheckedVars('icc-picker');
        if (vars.length < 2) { alert('กรุณาเลือกตัวแปรอย่างน้อย 2 ตัว'); return; }
        var dataArrays = vars.map(function(v){return getColumnData(v,true);});
        var result = Stats.icc(dataArrays);
        if (!result) { alert('ไม่สามารถคำนวณ ICC ได้'); return; }

        var rows = [
            {Type:result.icc1.type, ICC:fmt(result.icc1.value), Interpretation:result.icc1.interpretation},
            {Type:result.icc2.type, ICC:fmt(result.icc2.value), Interpretation:result.icc2.interpretation},
            {Type:result.icc3.type, ICC:fmt(result.icc3.value), Interpretation:result.icc3.interpretation}
        ];

        state.results['icc'] = {data:rows, title:'ICC (Intraclass Correlation Coefficient)', extras:[{title:'ANOVA Components',data:[{
            Subjects:result.n, Raters:result.k, MSB:fmt(result.MSB), MSW:fmt(result.MSW), MSR:fmt(result.MSR), MSE:fmt(result.MSE)
        }]}]};
        displayResults('icc');
    }

    // =========================================================================
    // Split-Half Reliability
    // =========================================================================

    function runSplitHalf() {
        var vars = getCheckedVars('sh-picker');
        if (vars.length < 2) { alert('กรุณาเลือกอย่างน้อย 2 Items'); return; }
        var dataArrays = vars.map(function(v){return getColumnData(v,true);});
        var result = Stats.splitHalf(dataArrays);
        if (!result) { alert('ไม่สามารถคำนวณได้'); return; }

        var rows = [{
            'Split-Half r':fmt(result.rHalf), 'Spearman-Brown':fmt(result.spearmanBrown),
            'Guttman Split-Half':fmt(result.guttman), 'N Items':result.nItems, 'N Cases':result.nCases
        }];

        state.results['sh'] = {data:rows, title:'Split-Half Reliability'};
        displayResults('sh');
    }

    // =========================================================================
    // McNemar Test
    // =========================================================================

    function runMcNemar() {
        var beforeVar = getPickerValue('mcn-before-picker','mcn-before');
        var afterVar = getPickerValue('mcn-after-picker','mcn-after');
        if (!beforeVar || !afterVar) { alert('กรุณาเลือกตัวแปร Before และ After'); return; }

        var before = getColumnData(beforeVar,true).map(function(v){return v>=0.5?1:0;});
        var after = getColumnData(afterVar,true).map(function(v){return v>=0.5?1:0;});
        var n = Math.min(before.length,after.length);

        var result = Stats.mcnemar(before.slice(0,n), after.slice(0,n));
        if (!result) { alert('ไม่สามารถวิเคราะห์ได้'); return; }

        var extras = [{title:'2x2 Contingency Table',data:[
            {'':afterVar+'=1', [beforeVar+'=1']:result.a, [beforeVar+'=0']:result.c},
            {'':afterVar+'=0', [beforeVar+'=1']:result.b, [beforeVar+'=0']:result.d}
        ]}];

        var rows = [{
            'Chi-Square':fmt(result.chi2), 'df':result.df,
            'p-value':Stats.formatPValue(result.p),
            'Discordant Pairs':result.discordant, 'N':result.n,
            'Sig.':result.significant?'✓':''
        }];

        state.results['mcn'] = {data:rows, title:'McNemar Test', extras:extras};
        displayResults('mcn');
    }

    // =========================================================================
    // Fisher's Exact Test
    // =========================================================================

    function runFisherExact() {
        var var1 = getPickerValue('fe-var1-picker','fe-var1');
        var var2 = getPickerValue('fe-var2-picker','fe-var2');
        if (!var1 || !var2) { alert('กรุณาเลือกตัวแปร 2 ตัว'); return; }

        var d1 = state.data.map(function(r){return r[var1];}), d2 = state.data.map(function(r){return r[var2];});
        var ct = Stats.crossTab(d1,d2);
        if (!ct || ct.rowLabels.length!==2 || ct.colLabels.length!==2) { alert('Fisher\'s Exact ต้องการตาราง 2x2'); return; }

        var result = Stats.fisherExact(ct.observed[0][0],ct.observed[0][1],ct.observed[1][0],ct.observed[1][1]);
        if (!result) { alert('ไม่สามารถคำนวณได้'); return; }

        var rows = [{
            'p (exact)':Stats.formatPValue(result.pExact),
            'p (two-tailed)':Stats.formatPValue(result.pTwoTail),
            'Odds Ratio':fmt(result.oddsRatio), 'N':result.n,
            'Sig.':result.pTwoTail<0.05?'✓':''
        }];

        state.results['fe'] = {data:rows, title:"Fisher's Exact Test"};
        displayResults('fe');
    }

    // =========================================================================
    // Cochran's Q Test
    // =========================================================================

    function runCochranQ() {
        var vars = getCheckedVars('cq-picker');
        if (vars.length < 3) { alert('กรุณาเลือกอย่างน้อย 3 ตัวแปร'); return; }
        var dataArrays = vars.map(function(v){return getColumnData(v,true).map(function(x){return x>=0.5?1:0;});});
        var result = Stats.cochranQ(dataArrays);
        if (!result) { alert('ไม่สามารถคำนวณได้'); return; }

        var extras = [{title:'Proportions',data:vars.map(function(v,i){
            var sum = dataArrays[i].reduce(function(a,b){return a+b;},0);
            return {Variable:v, 'Success (1)':sum, 'Total':dataArrays[i].length, 'Proportion':fmt(sum/dataArrays[i].length,3)};
        })}];

        var rows = [{"Cochran's Q":fmt(result.Q), df:result.df, 'p-value':Stats.formatPValue(result.p),
                     'k (conditions)':result.k, 'N':result.n, 'Sig.':result.significant?'✓':''}];

        state.results['cq'] = {data:rows, title:"Cochran's Q Test", extras:extras};
        displayResults('cq');
    }

    // =========================================================================
    // New Analysis Functions
    // =========================================================================

    function runPowerAnalysis() {
        var test = getSelectValue('pwr-test')||'ttest';
        var es = parseFloat(document.getElementById('pwr-es').value)||0.5;
        var alpha = parseFloat(document.getElementById('pwr-alpha').value)||0.05;
        var power = parseFloat(document.getElementById('pwr-power').value)||0.80;
        var groups = parseInt(document.getElementById('pwr-groups').value)||3;
        var result = Stats.powerAnalysis(test,{effectSize:es,alpha:alpha,power:power,groups:groups});
        if(!result){alert('ไม่สามารถคำนวณได้');return;}
        var rows=[{'Test':result.test,'Effect Size':fmt(es),'Alpha':alpha,'Power':power,
            'N per Group':result.nPerGroup||result.n,'Total N':result.totalN}];
        state.results['pwr']={data:rows,title:'Power Analysis — Required Sample Size'};
        displayResults('pwr');
    }

    function runWelchAnova() {
        var dv=getPickerValue('wa-dv-picker','wa-dv'),iv=getPickerValue('wa-iv-picker','wa-iv');
        if(!dv||!iv){alert('กรุณาเลือก DV และ Factor');return;}
        var split=splitByGroup(dv,iv),gNames=split.groupNames;
        if(gNames.length<3){alert('ต้องมีอย่างน้อย 3 กลุ่ม');return;}
        var groups=gNames.map(function(name){return split.groups[name];});
        var result=Stats.welchAnova(groups,gNames);
        if(!result){alert('ไม่สามารถวิเคราะห์ได้');return;}
        var extras=[{title:'Descriptive',data:result.descriptives.map(function(d){
            return{Group:d.group,N:d.n,Mean:fmt(d.mean),'S.D.':fmt(Math.sqrt(d.variance)),Variance:fmt(d.variance)};})}];
        var rows=[{'Welch F':fmt(result.F),'df1':fmt(result.df1,0),'df2':fmt(result.df2,2),
            'p-value':Stats.formatPValue(result.p),'Sig.':result.significant?'✓':''}];
        state.results['wa']={data:rows,title:"Welch's ANOVA",extras:extras};
        displayResults('wa');
    }

    function runDunnett() {
        var dv=getPickerValue('dnt-dv-picker','dnt-dv'),iv=getPickerValue('dnt-iv-picker','dnt-iv');
        if(!dv||!iv){alert('กรุณาเลือก DV และ Factor');return;}
        var split=splitByGroup(dv,iv),gNames=split.groupNames;
        if(gNames.length<2){alert('ต้องมีอย่างน้อย 2 กลุ่ม');return;}
        var groups=gNames.map(function(name){return split.groups[name];});
        var result=Stats.dunnettTest(groups,gNames,0);
        if(!result){alert('ไม่สามารถวิเคราะห์ได้');return;}
        var rows=result.map(function(r){return{
            'Group':r.group,'Control':r.control,'Mean(Group)':fmt(r.meanGroup),'Mean(Control)':fmt(r.meanControl),
            'Mean Diff':fmt(r.meanDiff),'t':fmt(r.t),'df':r.df,'p-value':Stats.formatPValue(r.p),
            'p(adjusted)':Stats.formatPValue(r.pAdjusted),'Sig.':r.significant?'✓':''};});
        state.results['dnt']={data:rows,title:"Dunnett's Test (Control: "+gNames[0]+")"};
        displayResults('dnt');
    }

    function runMedianTest() {
        var dv=getPickerValue('mdt-dv-picker','mdt-dv'),iv=getPickerValue('mdt-iv-picker','mdt-iv');
        if(!dv||!iv){alert('กรุณาเลือก DV และ Factor');return;}
        var split=splitByGroup(dv,iv),gNames=split.groupNames;
        var groups=gNames.map(function(name){return split.groups[name];});
        var result=Stats.medianTest(groups,gNames);
        if(!result){alert('ไม่สามารถวิเคราะห์ได้');return;}
        var extras=[{title:'Group Counts',data:result.groupStats.map(function(g){
            return{Group:g.group,N:g.n,'Above Median':g.above,'At/Below Median':g.below};})}];
        var rows=[{'Grand Median':fmt(result.grandMedian),'Chi-Square':fmt(result.chi2),'df':result.df,
            'p-value':Stats.formatPValue(result.p),'Sig.':result.significant?'✓':''}];
        state.results['mdt']={data:rows,title:'Median Test',extras:extras};
        displayResults('mdt');
    }

    function runRunsTest() {
        var vars=getCheckedVars('run-picker');
        if(vars.length===0){alert('กรุณาเลือกตัวแปร');return;}
        var rows=[];
        vars.forEach(function(v){
            var values=getColumnData(v,true);
            if(values.length<10)return;
            var result=Stats.runsTest(values);
            if(!result)return;
            rows.push({Variable:v,N:result.n,'Runs':result.runs,'Expected Runs':fmt(result.expectedRuns),
                'Z':fmt(result.z),'p-value':Stats.formatPValue(result.p),'Random?':result.random?'✅ Yes':'⚠️ No'});
        });
        state.results['run']={data:rows,title:'Runs Test (Randomness)'};
        displayResults('run');
    }

    function runKS2() {
        var dv=getPickerValue('ks2-dv-picker','ks2-dv'),iv=getPickerValue('ks2-iv-picker','ks2-iv');
        if(!dv||!iv){alert('กรุณาเลือก DV และ Grouping');return;}
        var split=splitByGroup(dv,iv),gNames=split.groupNames;
        if(gNames.length!==2){alert('ต้องมี 2 กลุ่มเท่านั้น');return;}
        var g1=split.groups[gNames[0]],g2=split.groups[gNames[1]];
        var result=Stats.ks2Sample(g1,g2);
        if(!result){alert('ไม่สามารถวิเคราะห์ได้');return;}
        var rows=[{'Group 1':gNames[0]+' (n='+result.n1+')','Group 2':gNames[1]+' (n='+result.n2+')',
            'D':fmt(result.D),'p-value':Stats.formatPValue(result.p),
            'Same Distribution?':result.significant?'⚠️ No (Different)':'✅ Yes (Same)'}];
        state.results['ks2']={data:rows,title:'Kolmogorov-Smirnov 2-Sample Test'};
        displayResults('ks2');
    }

    function runCluster() {
        var vars=getCheckedVars('cls-picker');
        if(vars.length<2){alert('กรุณาเลือกตัวแปรอย่างน้อย 2 ตัว');return;}
        var k=parseInt(document.getElementById('cls-k').value)||3;
        var dataArrays=vars.map(function(v){return getColumnData(v,true);});
        var result=Stats.kMeans(dataArrays,vars,k);
        if(!result){alert('ไม่สามารถวิเคราะห์ได้');return;}
        var rows=result.clusters.map(function(c){
            var row={'Cluster':c.cluster,'N':c.n};
            vars.forEach(function(v){row['Mean('+v+')']=fmt(c.means[v]);});
            return row;
        });
        var extras=[{title:'',html:'<div class="detail-box"><strong>K-Means Clustering:</strong> '+result.k+' clusters, '+result.n+' observations, '+result.iterations+' iterations</div>'}];
        state.results['cls']={data:rows,title:'Cluster Analysis (K='+k+')',extras:extras};
        displayResults('cls');
    }

    function runDiscriminant() {
        var groupVar=getPickerValue('da-group-picker','da-group');
        var predVars=getCheckedVars('da-vars-picker');
        if(!groupVar||predVars.length<1){alert('กรุณาเลือก Grouping Variable และ Predictors');return;}
        var groupData=state.data.map(function(r){return r[groupVar];});
        var dataArrays=predVars.map(function(v){return getColumnData(v,true);});
        var result=Stats.discriminantAnalysis(dataArrays,predVars,groupData);
        if(!result||result.error){alert(result?result.error:'ไม่สามารถวิเคราะห์ได้');return;}
        var extras=[];
        if(result.groupMeans)extras.push({title:'Group Means',data:result.groupMeans});
        var rows=[{"Wilks' Lambda":fmt(result.wilksLambda),'F':fmt(result.F),'df1':result.df1,'df2':result.df2,
            'p-value':Stats.formatPValue(result.p),'Accuracy':fmt(result.accuracy,1)+'%','Correct':result.correct+'/'+result.n}];
        state.results['da']={data:rows,title:'Discriminant Analysis',extras:extras};
        displayResults('da');
    }

    function runMissing() {
        if(!state.data){alert('กรุณา Upload ข้อมูลก่อน');return;}
        var result=Stats.missingValueAnalysis(state.data,state.columns);
        if(!result){alert('ไม่สามารถวิเคราะห์ได้');return;}
        var rows=result.variables.map(function(v){return{
            Variable:v.variable,N:v.n,Valid:v.valid,Missing:v.missing,
            'Missing %':fmt(v.missingPct,1)+'%',Status:v.status};});
        var extras=[{title:'',html:'<div class="detail-box"><strong>Overall:</strong> '+result.totalCells+' cells, '+result.totalMissing+' missing ('+fmt(result.overallPct,1)+'%)</div>'}];
        state.results['miss']={data:rows,title:'Missing Value Analysis',extras:extras};
        displayResults('miss');
    }

    function runQQPlot() {
        var varName=getPickerValue('qq-var-picker','qq-var');
        if(!varName){alert('กรุณาเลือกตัวแปร');return;}
        var values=getColumnData(varName,true);
        var result=Stats.qqPlotData(values);
        if(!result){alert('ไม่สามารถสร้าง Q-Q Plot ได้');return;}
        var canvas=document.getElementById('qq-canvas');
        if(canvas&&window.Chart){
            if(window._qqChart)window._qqChart.destroy();
            var pts=result.points.map(function(p){return{x:p.theoretical,y:p.observed};});
            var minV=Math.min(pts[0].x,pts[0].y)-0.5,maxV=Math.max(pts[pts.length-1].x,pts[pts.length-1].y)+0.5;
            window._qqChart=new Chart(canvas,{
                type:'scatter',
                data:{datasets:[
                    {label:'Data Points',data:pts,backgroundColor:'rgba(37,99,235,0.6)',pointRadius:4,showLine:false},
                    {label:'Reference Line',data:[{x:minV,y:minV},{x:maxV,y:maxV}],showLine:true,borderColor:'#dc2626',borderDash:[5,5],pointRadius:0,borderWidth:2}
                ]},
                options:{responsive:true,scales:{x:{title:{display:true,text:'Theoretical Quantiles'}},y:{title:{display:true,text:'Sample Quantiles'}}},
                    plugins:{title:{display:true,text:'Q-Q Plot: '+varName+' (N='+result.n+')'}}}
            });
        }
        var sw=Stats.shapiroWilk(values);
        var rows=[{Variable:varName,N:result.n,Mean:fmt(result.mean),'S.D.':fmt(result.sd),
            'Shapiro-Wilk W':sw?fmt(sw.W):'N/A','S-W p':sw?Stats.formatPValue(sw.p):'N/A',
            'Normal?':sw?(sw.p>0.05?'✅ Yes':'⚠️ No'):'N/A'}];
        state.results['qq']={data:rows,title:'Q-Q Plot Analysis'};
        displayResults('qq');
    }

    // =========================================================================
    // AI Integration
    // =========================================================================

    async function aiAnalyze(prefix) {
        var result = state.results[prefix];
        if (!result) {
            var aiEl = document.getElementById(prefix + '-ai-result');
            if (aiEl) { aiEl.innerHTML = '<p class="error-text">⚠️ กรุณารันการวิเคราะห์ก่อน</p>'; aiEl.style.display = ''; }
            return;
        }
        if (!state.aiSettings.apiKey) {
            var aiEl2 = document.getElementById(prefix + '-ai-result');
            if (aiEl2) { aiEl2.innerHTML = '<p class="error-text">⚠️ กรุณาตั้งค่า API Key ใน AI Settings ก่อน</p>'; aiEl2.style.display = ''; }
            return;
        }

        var aiResultEl = document.getElementById(prefix + '-ai-result');
        if (aiResultEl) {
            aiResultEl.innerHTML = '<p class="loading">Analyzing with AI... Please wait.</p>';
            aiResultEl.style.display = '';
        }

        // Build text representation of results
        var textData = '';
        if (result.title) textData += result.title + '\n\n';
        if (result.extras) {
            result.extras.forEach(function (extra) {
                if (extra.title) textData += extra.title + '\n';
                if (extra.data) {
                    textData += tableToText(extra.data) + '\n\n';
                }
            });
        }
        if (result.data) {
            textData += tableToText(result.data);
        }

        // Get interpretation format
        var formatEl = document.getElementById(prefix + '-interp-format');
        var interpFormat = formatEl ? formatEl.value : 'chapter4';

        try {
            var response = await fetch('/api/ai/summarize', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    apiKey: state.aiSettings.apiKey,
                    model: state.aiSettings.model,
                    data: textData,
                    title: result.title || 'Analysis Results',
                    format: interpFormat
                })
            });

            if (!response.ok) {
                var errText = await response.text();
                throw new Error('AI request failed: ' + errText);
            }

            var json = await response.json();
            var aiText = json.summary || json.result || json.text || '';
            if (!aiText && json.success === false) aiText = json.error || 'No response';

            // Clean AI text: remove markdown symbols and JSON artifacts
            aiText = cleanAIText(aiText);

            state.aiResults[prefix] = aiText;

            if (aiResultEl) {
                aiResultEl.innerHTML = '<h4>🤖 ผลการวิเคราะห์ — พร้อมใช้ในบทที่ 4</h4><p>' + aiText.replace(/\n/g, '<br>') + '</p>';
                aiResultEl.style.display = '';
                // Inject follow-up chat UI
                injectFollowupChat(prefix, aiResultEl);
            }
        } catch (err) {
            if (aiResultEl) {
                aiResultEl.innerHTML = '<p class="error-text">Error: ' + escapeHtml(err.message) + '</p>';
                aiResultEl.style.display = '';
            }
        }
    }

    // =========================================================================
    // AI Follow-up Chat in Results
    // =========================================================================
    // Stores per-prefix chat histories
    state.followupChats = state.followupChats || {};

    function injectFollowupChat(prefix, parentEl) {
        // Remove existing followup chat if any
        var existingChat = document.getElementById('followup-chat-' + prefix);
        if (existingChat) existingChat.remove();

        var chatDiv = document.createElement('div');
        chatDiv.id = 'followup-chat-' + prefix;
        chatDiv.className = 'followup-chat-container';
        chatDiv.innerHTML =
            '<div class="followup-chat-header" onclick="toggleFollowupChat(\'' + prefix + '\')">' +
                '<span>💬 สอบถาม / ขอคำแนะนำเพิ่มเติมจาก AI</span>' +
                '<span class="followup-arrow" id="followup-arrow-' + prefix + '">▸</span>' +
            '</div>' +
            '<div class="followup-chat-body" id="followup-body-' + prefix + '" style="display:none">' +
                '<div class="followup-messages" id="followup-msgs-' + prefix + '">' +
                    '<div class="chat-msg-ai"><div class="chat-avatar">🤖</div><div class="chat-bubble">สอบถามเพิ่มเติมเกี่ยวกับผลวิเคราะห์นี้ได้เลยครับ เช่น อธิบายเพิ่ม, แนะนำการเขียนบทที่ 4, แปลผลอย่างอื่น</div></div>' +
                '</div>' +
                '<div class="followup-suggestions">' +
                    '<button class="suggestion-btn-sm" onclick="sendFollowup(\'' + prefix + '\',\'อธิบายผลวิเคราะห์ให้ละเอียดขึ้น\')">📝 อธิบายเพิ่มเติม</button>' +
                    '<button class="suggestion-btn-sm" onclick="sendFollowup(\'' + prefix + '\',\'ช่วยเขียนเนื้อหาบทที่ 4 จากผลวิเคราะห์นี้\')">📖 เขียนบทที่ 4</button>' +
                    '<button class="suggestion-btn-sm" onclick="sendFollowup(\'' + prefix + '\',\'อธิบายให้เข้าใจง่ายๆ แบบชาวบ้าน\')">💡 อธิบายง่ายๆ</button>' +
                    '<button class="suggestion-btn-sm" onclick="sendFollowup(\'' + prefix + '\',\'มีข้อจำกัดหรือข้อควรระวังอะไรบ้าง\')">⚠️ ข้อควรระวัง</button>' +
                '</div>' +
                '<div class="followup-input-row">' +
                    '<input type="text" class="followup-input" id="followup-input-' + prefix + '" placeholder="พิมพ์คำถามเกี่ยวกับผลวิเคราะห์..." onkeypress="if(event.key===\'Enter\') sendFollowupFromInput(\'' + prefix + '\')">' +
                    '<button class="btn btn-ai btn-sm" onclick="sendFollowupFromInput(\'' + prefix + '\')">📤 ส่ง</button>' +
                '</div>' +
            '</div>';

        parentEl.parentNode.insertBefore(chatDiv, parentEl.nextSibling);
        // Initialize chat history for this prefix
        if (!state.followupChats[prefix]) state.followupChats[prefix] = [];
    }

    function toggleFollowupChat(prefix) {
        var body = document.getElementById('followup-body-' + prefix);
        var arrow = document.getElementById('followup-arrow-' + prefix);
        if (body) {
            var isHidden = body.style.display === 'none';
            body.style.display = isHidden ? '' : 'none';
            if (arrow) arrow.textContent = isHidden ? '▾' : '▸';
        }
    }

    function sendFollowupFromInput(prefix) {
        var input = document.getElementById('followup-input-' + prefix);
        if (!input) return;
        var msg = input.value.trim();
        if (!msg) return;
        input.value = '';
        sendFollowup(prefix, msg);
    }

    async function sendFollowup(prefix, message) {
        if (!state.aiSettings.apiKey) {
            appendFollowupMsg(prefix, 'bot', '⚠️ กรุณาตั้งค่า API Key ใน AI Settings ก่อน');
            return;
        }

        // Build context from the analysis results
        var result = state.results[prefix];
        var textData = '';
        if (result) {
            if (result.title) textData += result.title + '\n\n';
            if (result.extras) {
                result.extras.forEach(function(extra) {
                    if (extra.title) textData += extra.title + '\n';
                    if (extra.data) textData += tableToText(extra.data) + '\n\n';
                });
            }
            if (result.data) textData += tableToText(result.data);
        }

        var aiSummary = state.aiResults[prefix] || '';

        // Add user message
        if (!state.followupChats[prefix]) state.followupChats[prefix] = [];
        state.followupChats[prefix].push({ role: 'user', content: message });
        appendFollowupMsg(prefix, 'user', message);

        // Show loading
        var loadingId = 'followup-loading-' + prefix;
        appendFollowupMsg(prefix, 'bot', '<span id="' + loadingId + '" class="loading-dots">กำลังคิด...</span>', true);

        try {
            var systemContext = 'คุณคือผู้เชี่ยวชาญด้านสถิติวิจัย กำลังให้คำปรึกษาเกี่ยวกับผลวิเคราะห์ต่อไปนี้:\n\n' +
                'ข้อมูลผลวิเคราะห์:\n' + textData + '\n\n' +
                'ผลสรุปจาก AI ก่อนหน้า:\n' + aiSummary + '\n\n' +
                'ตอบเป็นภาษาไทย กระชับ ชัดเจน เหมาะสำหรับใช้ในงานวิจัย/วิทยานิพนธ์';

            var history = state.followupChats[prefix].map(function(m) {
                return { role: m.role === 'bot' ? 'assistant' : m.role, content: m.content };
            });

            var response = await fetch('/api/ai/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    apiKey: state.aiSettings.apiKey,
                    model: state.aiSettings.model,
                    history: history,
                    context: systemContext
                })
            });

            // Remove loading message
            var loadingEl = document.getElementById(loadingId);
            if (loadingEl) loadingEl.closest('.chat-msg-ai').remove();

            if (!response.ok) {
                var errText = await response.text();
                throw new Error('AI request failed: ' + errText);
            }

            var json = await response.json();
            var reply = json.reply || json.result || json.text || 'ไม่ได้รับคำตอบจาก AI';
            reply = cleanAIText(reply);

            state.followupChats[prefix].push({ role: 'bot', content: reply });
            appendFollowupMsg(prefix, 'bot', reply);
        } catch (err) {
            var loadingEl2 = document.getElementById(loadingId);
            if (loadingEl2) loadingEl2.closest('.chat-msg-ai').remove();
            appendFollowupMsg(prefix, 'bot', 'Error: ' + err.message);
        }
    }

    function appendFollowupMsg(prefix, role, text, isHtml) {
        var container = document.getElementById('followup-msgs-' + prefix);
        if (!container) return;
        var div = document.createElement('div');
        var avatar = role === 'user' ? '👤' : '🤖';
        div.className = role === 'user' ? 'chat-msg-user' : 'chat-msg-ai';
        var displayText = isHtml ? text : (role === 'user' ? escapeHtml(text) : formatAIResponse(text));
        div.innerHTML = '<div class="chat-avatar">' + avatar + '</div><div class="chat-bubble">' + displayText + '</div>';
        container.appendChild(div);
        container.scrollTop = container.scrollHeight;

        // Auto-open the chat body if it's hidden
        var body = document.getElementById('followup-body-' + prefix);
        var arrow = document.getElementById('followup-arrow-' + prefix);
        if (body && body.style.display === 'none') {
            body.style.display = '';
            if (arrow) arrow.textContent = '▾';
        }
    }

    function tableToText(data) {
        if (!data || data.length === 0) return '';
        var cols = Object.keys(data[0]);
        var lines = [cols.join('\t')];
        data.forEach(function (row) {
            lines.push(cols.map(function (c) { return row[c] !== undefined && row[c] !== null ? String(row[c]) : ''; }).join('\t'));
        });
        return lines.join('\n');
    }

    function cleanAIText(text) {
        if (!text) return '';
        // Remove JSON wrappers
        text = text.replace(/^\s*\{[\s\S]*?"(?:summary|result|text)"\s*:\s*"/, '');
        text = text.replace(/"\s*(?:,\s*"success"\s*:\s*true)?\s*\}\s*$/, '');
        // Remove markdown symbols
        text = text.replace(/\*\*/g, '');
        text = text.replace(/\*/g, '');
        text = text.replace(/^#+\s*/gm, '');
        text = text.replace(/^[-•]\s*/gm, '');
        text = text.replace(/^>\s*/gm, '');
        text = text.replace(/```[\s\S]*?```/g, '');
        text = text.replace(/`/g, '');
        // Clean extra whitespace
        text = text.replace(/\n{3,}/g, '\n\n');
        return text.trim();
    }

    function formatAIResponse(text) {
        if (!text) return '';
        var clean = cleanAIText(text);
        var html = escapeHtml(clean);
        html = html.replace(/\n/g, '<br>');
        return '<div class="ai-text">' + html + '</div>';
    }

    // =========================================================================
    // AI Chat
    // =========================================================================

    function sendChat() {
        var input = document.getElementById('chat-input');
        if (!input) return;
        var message = input.value.trim();
        if (!message) return;
        input.value = '';
        aiChat(message);
    }

    function sendQuickChat(message) {
        aiChat(message);
    }

    async function aiChat(message) {
        if (!state.aiSettings.apiKey) {
            appendChatMessage('bot', '⚠️ กรุณาตั้งค่า API Key ใน AI Settings ก่อนใช้งาน');
            return;
        }

        // Add user message
        state.chatHistory.push({ role: 'user', content: message });
        appendChatMessage('user', message);

        // Provide data context if available
        var context = '';
        if (state.data) {
            context = 'The user has loaded a dataset with ' + state.data.length + ' rows and ' + state.columns.length + ' columns. ';
            context += 'Columns: ' + state.columns.join(', ') + '. ';
            context += 'Numeric columns: ' + state.numericCols.join(', ') + '. ';
            context += 'Categorical columns: ' + (state.categoricalCols.length > 0 ? state.categoricalCols.join(', ') : 'None') + '. ';
        }

        try {
            var response = await fetch('/api/ai/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    apiKey: state.aiSettings.apiKey,
                    model: state.aiSettings.model,
                    history: state.chatHistory,
                    context: context
                })
            });

            if (!response.ok) {
                var errText = await response.text();
                throw new Error('AI chat failed: ' + errText);
            }

            var json = await response.json();
            var reply = json.reply || json.result || json.text || 'No response from AI.';
            reply = cleanAIText(reply);

            state.chatHistory.push({ role: 'assistant', content: reply });
            appendChatMessage('bot', reply);
        } catch (err) {
            appendChatMessage('bot', 'Error: ' + err.message);
        }
    }

    function appendChatMessage(role, text) {
        var container = document.getElementById('chat-messages');
        if (!container) return;
        var div = document.createElement('div');
        var avatar = role === 'user' ? '👤' : '🤖';
        div.className = role === 'user' ? 'chat-msg-user' : 'chat-msg-ai';
        var cleanText = role === 'user' ? escapeHtml(text) : formatAIResponse(text);
        div.innerHTML = '<div class="chat-avatar">' + avatar + '</div><div class="chat-bubble">' + cleanText + '</div>';
        container.appendChild(div);
        container.scrollTop = container.scrollHeight;
    }

    function clearChatHistory() {
        state.chatHistory = [];
        var container = document.getElementById('chat-messages');
        if (container) {
            container.innerHTML = '<div class="chat-msg-ai"><div class="chat-avatar">🤖</div><div class="chat-bubble">สวัสดีครับ! ผมคือ Stat Advisor ผู้ช่วยด้านสถิติวิจัย พร้อมให้คำปรึกษาเรื่องการเลือกใช้สถิติ การแปลผล และการเขียนผลการวิจัย ถามมาได้เลยครับ</div></div>';
        }
        showSettingsStatus('ai-settings-status', '✅ ล้างประวัติแชทเรียบร้อย', 'success');
    }

    // =========================================================================
    // AI Settings
    // =========================================================================

    function saveAISettings() {
        var apiKey = document.getElementById('ai-api-key').value.trim();
        var model = getSelectValue('ai-model');
        state.aiSettings.apiKey = apiKey;
        state.aiSettings.model = model || 'gemini-2.5-flash-lite';

        try {
            localStorage.setItem('eesy_ai_apiKey', apiKey);
            localStorage.setItem('eesy_ai_model', state.aiSettings.model);
        } catch (e) { /* localStorage may not be available */ }

        showSettingsStatus('ai-settings-status', '✅ บันทึกสำเร็จ! Model: ' + state.aiSettings.model, 'success');
    }

    function loadAISettings() {
        try {
            var apiKey = localStorage.getItem('eesy_ai_apiKey');
            var model = localStorage.getItem('eesy_ai_model');
            if (apiKey) state.aiSettings.apiKey = apiKey;
            if (model) state.aiSettings.model = model;
        } catch (e) { /* ignore */ }

        var apiKeyInput = document.getElementById('ai-api-key');
        if (apiKeyInput && state.aiSettings.apiKey) apiKeyInput.value = state.aiSettings.apiKey;

        var modelSelect = document.getElementById('ai-model');
        if (modelSelect && state.aiSettings.model) modelSelect.value = state.aiSettings.model;
    }

    async function testAIConnection() {
        if (!state.aiSettings.apiKey) {
            showSettingsStatus('ai-settings-status', '❌ กรุณาใส่ API Key และกด Save ก่อน', 'error');
            return;
        }

        showSettingsStatus('ai-settings-status', '🔄 กำลังทดสอบการเชื่อมต่อ...', 'loading');

        try {
            var response = await fetch('/api/ai/test', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    apiKey: state.aiSettings.apiKey,
                    model: state.aiSettings.model
                })
            });

            if (response.ok) {
                var json = await response.json();
                showSettingsStatus('ai-settings-status', '✅ เชื่อมต่อสำเร็จ! AI พร้อมใช้งาน — ' + (json.message || ''), 'success');
            } else {
                var errText = await response.text();
                showSettingsStatus('ai-settings-status', '❌ เชื่อมต่อล้มเหลว: ' + errText, 'error');
            }
        } catch (err) {
            showSettingsStatus('ai-settings-status', '❌ เชื่อมต่อล้มเหลว: ' + err.message, 'error');
        }
    }

    function showSettingsStatus(elementId, message, type) {
        var el = document.getElementById(elementId);
        if (!el) return;
        el.style.display = 'block';
        el.className = 'settings-status settings-status-' + type;
        el.textContent = message;
    }

    // =========================================================================
    // Export Functions
    // =========================================================================

    function exportExcel(prefix) {
        var result = state.results[prefix];
        if (!result) { alert('No results to export. Please run analysis first.'); return; }

        try {
            var wb = XLSX.utils.book_new();

            // Main result sheet
            if (result.data && result.data.length > 0) {
                var ws = XLSX.utils.json_to_sheet(result.data);
                XLSX.utils.book_append_sheet(wb, ws, 'Results');
            }

            // Extra sheets
            if (result.extras) {
                result.extras.forEach(function (extra, idx) {
                    if (extra.data && extra.data.length > 0) {
                        var sheetName = (extra.title || 'Extra ' + (idx + 1)).substring(0, 31).replace(/[\/\\?*\[\]]/g, '');
                        var ws = XLSX.utils.json_to_sheet(extra.data);
                        XLSX.utils.book_append_sheet(wb, ws, sheetName);
                    }
                });
            }

            var filename = (result.title || 'analysis').replace(/[^a-zA-Z0-9]/g, '_') + '.xlsx';
            XLSX.writeFile(wb, filename);
        } catch (err) {
            alert('Error exporting Excel: ' + err.message);
        }
    }

    function exportWord(prefix) {
        var result = state.results[prefix];
        if (!result) { alert('No results to export. Please run analysis first.'); return; }

        try {
            var html = '<html><head><meta charset="UTF-8"><style>';
            html += 'body{font-family:TH Sarabun New,Sarabun,sans-serif;font-size:14pt;margin:2cm;}';
            html += 'table{border-collapse:collapse;width:100%;margin:10px 0;}';
            html += 'th,td{border:1px solid #000;padding:4px 8px;text-align:center;font-size:12pt;}';
            html += 'th{background:#f0f0f0;font-weight:bold;}';
            html += 'h2,h3,h4{margin:10px 0;}';
            html += '.ai-text{margin:10px 0;padding:10px;background:#f9f9f9;border-left:3px solid #4a90d9;}';
            html += '</style></head><body>';

            if (result.title) html += '<h2>' + escapeHtml(result.title) + '</h2>';

            // Extras tables
            if (result.extras) {
                result.extras.forEach(function (extra) {
                    if (extra.title) html += '<h3>' + escapeHtml(extra.title) + '</h3>';
                    if (extra.data && extra.data.length > 0) {
                        html += buildTable(extra.data);
                    }
                    if (extra.html) html += extra.html;
                });
            }

            // Main table
            if (result.data && result.data.length > 0) {
                html += buildTable(result.data);
            }

            // AI result
            if (state.aiResults[prefix]) {
                html += '<h3>AI Analysis</h3>';
                html += state.aiResults[prefix];
            }

            html += '</body></html>';

            var blob = new Blob(['\ufeff' + html], { type: 'application/msword' });
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url;
            a.download = (result.title || 'analysis').replace(/[^a-zA-Z0-9]/g, '_') + '.doc';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (err) {
            alert('Error exporting Word: ' + err.message);
        }
    }

    // =========================================================================
    // Template Download
    // =========================================================================

    function downloadTemplate(type) {
        window.location.href = '/api/templates/' + encodeURIComponent(type) + '/download';
    }

    // =========================================================================
    // Initialization
    // =========================================================================

    document.addEventListener('DOMContentLoaded', function () {
        // File upload listener
        var fileInput = document.getElementById('file-upload');
        if (fileInput) {
            fileInput.addEventListener('change', handleFileUpload);
        }

        // Login form
        var loginForm = document.getElementById('login-form');
        if (loginForm) {
            loginForm.addEventListener('submit', handleLogin);
        }

        // Load AI settings
        loadAISettings();

        // Navigate to home
        navigateTo('home');
    });

    // =========================================================================
    // =========================================================================
    // Survival Analysis
    // =========================================================================

    function runSurvival() {
        var timeVar = getPickerValue('surv-time-picker','surv-time');
        var eventVar = getPickerValue('surv-event-picker','surv-event');
        if (!timeVar || !eventVar) { alert('กรุณาเลือก Time และ Event variables'); return; }

        var time = getColumnData(timeVar, true);
        var event = getColumnData(eventVar, true).map(function(v){return v>=0.5?1:0;});
        var n = Math.min(time.length, event.length);
        time = time.slice(0,n); event = event.slice(0,n);

        var km = Stats.kaplanMeier(time, event);
        if (!km) { alert('ไม่สามารถวิเคราะห์ได้'); return; }

        var extras = [];

        // KM Summary
        var summaryRows = [{'N':km.n, 'Events':km.totalEvents, 'Censored':km.totalCensored,
            'Median Survival':km.medianSurvival!==null?fmt(km.medianSurvival):'Not reached',
            'Mean Survival':fmt(km.meanSurvival)}];
        extras.push({title:'Kaplan-Meier Summary', data:summaryRows});

        // Survival curve chart
        var canvas = document.getElementById('surv-canvas');
        if (canvas && window.Chart) {
            if (window._survChart) window._survChart.destroy();
            var pts = km.curve.map(function(p){return {x:p.time, y:p.survival};});
            window._survChart = new Chart(canvas, {
                type:'line',
                data:{datasets:[{label:'Survival Probability',data:pts,borderColor:'#2563eb',backgroundColor:'rgba(37,99,235,0.1)',fill:true,stepped:'before',pointRadius:2}]},
                options:{responsive:true,scales:{x:{title:{display:true,text:'Time'}},y:{title:{display:true,text:'Survival Probability'},min:0,max:1}},plugins:{title:{display:true,text:'Kaplan-Meier Survival Curve'}}}
            });
        }

        // Log-Rank test if group variable selected
        var groupVar = getPickerValue('surv-group-picker','surv-group');
        var mainRows = [];
        if (groupVar) {
            var groupData = state.data.map(function(r){return r[groupVar];});
            var groups = {};
            groupData.forEach(function(g,i){if(i<n){if(!groups[g])groups[g]={t:[],e:[]};groups[g].t.push(time[i]);groups[g].e.push(event[i]);}});
            var gNames = Object.keys(groups);
            if (gNames.length === 2) {
                var lr = Stats.logRankTest(groups[gNames[0]].t, groups[gNames[0]].e, groups[gNames[1]].t, groups[gNames[1]].e);
                if (lr) {
                    mainRows.push({'Test':'Log-Rank', 'Chi-Square':fmt(lr.chi2), 'df':lr.df, 'p-value':Stats.formatPValue(lr.p), 'Sig.':lr.significant?'✓':''});
                    extras.push({title:'Group Comparison',data:[
                        {Group:gNames[0],N:lr.group1.n,Events:lr.group1.events,Median:lr.group1.median!==null?fmt(lr.group1.median):'NR'},
                        {Group:gNames[1],N:lr.group2.n,Events:lr.group2.events,Median:lr.group2.median!==null?fmt(lr.group2.median):'NR'}
                    ]});
                }
            }
        }

        // Cox regression if covariates selected
        var covVars = getCheckedVars('surv-cov-picker');
        if (covVars.length > 0) {
            var covData = covVars.map(function(v){return getColumnData(v,true).slice(0,n);});
            var cox = Stats.coxRegression(time, event, covData, covVars);
            if (cox && cox.coefficients) {
                var coxRows = cox.coefficients.map(function(c){
                    return {Variable:c.variable, B:fmt(c.b), 'S.E.':fmt(c.se), 'HR':fmt(c.hr), '95% CI (HR)':c.hrCI, Wald:fmt(c.wald), 'p-value':Stats.formatPValue(c.p), 'Sig.':c.significant?'✓':''};
                });
                extras.push({title:'Cox Regression', data:coxRows});
            }
        }

        if (mainRows.length === 0) mainRows = summaryRows;
        state.results['surv'] = {data:mainRows, title:'Survival Analysis', extras:extras};
        displayResults('surv');
    }

    // =========================================================================
    // Time Series Analysis
    // =========================================================================

    function runTimeSeries() {
        var varName = getPickerValue('ts-var-picker','ts-var');
        if (!varName) { alert('กรุณาเลือกตัวแปร'); return; }
        var period = parseInt(document.getElementById('ts-period').value) || 12;
        var values = getColumnData(varName, true);
        if (values.length < period * 2) { alert('ข้อมูลไม่เพียงพอ (ต้องมีอย่างน้อย ' + (period*2) + ' ค่า)'); return; }

        var result = Stats.timeSeries(values, period);
        if (!result) { alert('ไม่สามารถวิเคราะห์ได้'); return; }

        // Chart
        var canvas = document.getElementById('ts-canvas');
        if (canvas && window.Chart) {
            if (window._tsChart) window._tsChart.destroy();
            var labels = result.decomposition.map(function(d){return d.t;});
            window._tsChart = new Chart(canvas, {
                type:'line',
                data:{labels:labels,datasets:[
                    {label:'Original',data:result.decomposition.map(function(d){return d.original;}),borderColor:'#2563eb',borderWidth:2,pointRadius:1,fill:false},
                    {label:'Trend',data:result.decomposition.map(function(d){return d.trend;}),borderColor:'#dc2626',borderWidth:2,borderDash:[5,5],pointRadius:0,fill:false},
                    {label:'Forecast',data:new Array(values.length).fill(null).concat(result.forecast.map(function(f){return f.forecast;})),borderColor:'#059669',borderWidth:2,borderDash:[3,3],pointRadius:0,fill:false}
                ]},
                options:{responsive:true,plugins:{title:{display:true,text:'Time Series: '+varName}}}
            });
        }

        var extras = [];
        // Trend info
        extras.push({title:'Trend & Model Summary',data:[{
            'Trend Slope':fmt(result.trend.slope),'Intercept':fmt(result.trend.intercept),
            'Period':result.period,'RMSE':fmt(result.rmse),'ACF(1)':fmt(result.acf1),'N':result.n
        }]});
        // Seasonal pattern
        var seasonRows = result.seasonal.pattern.map(function(s,i){return {'Period Position':i+1,'Seasonal Effect':fmt(s)};});
        extras.push({title:'Seasonal Component',data:seasonRows});
        // Forecast
        var forecastRows = result.forecast.map(function(f){return {'Period':f.period,'Trend':fmt(f.trend),'Seasonal':fmt(f.seasonal),'Forecast':fmt(f.forecast)};});

        state.results['ts'] = {data:forecastRows, title:'Time Series Forecast (Next '+period+' periods)', extras:extras};
        displayResults('ts');
    }

    // =========================================================================
    // SAMPLE SIZE CALCULATOR
    // =========================================================================
    function runSampleSize() {
        var formula = getSelectValue('ss-formula') || 'yamane';
        var result = {};
        var extras = [];

        if (formula === 'yamane') {
            var N = parseFloat(document.getElementById('ss-pop').value) || 1000;
            var e = parseFloat(getSelectValue('ss-error')) || 0.05;
            var n = Math.ceil(N / (1 + N * e * e));
            result = {'สูตร':'Taro Yamane','จำนวนประชากร (N)':N,'ค่าความคลาดเคลื่อน (e)':e,'ขนาดตัวอย่างที่คำนวณได้ (n)':n};
            extras.push({title:'สูตร Yamane: n = N / (1 + Ne²)',data:[result]});
        } else if (formula === 'cochran') {
            var z = parseFloat(getSelectValue('ss-z')) || 1.96;
            var e2 = parseFloat(document.getElementById('ss-coch-e').value) || 0.05;
            var p = parseFloat(document.getElementById('ss-coch-p').value) || 0.50;
            var q = 1 - p;
            var n2 = Math.ceil((z * z * p * q) / (e2 * e2));
            result = {'สูตร':'Cochran','Z':z,'p':p,'q':q,'e':e2,'ขนาดตัวอย่าง (n)':n2};
            extras.push({title:'สูตร Cochran: n = Z²pq / e²',data:[result]});
        } else if (formula === 'krejcie') {
            var Nk = parseFloat(document.getElementById('ss-krej-pop').value) || 1000;
            // Krejcie & Morgan formula: S = X²NP(1-P) / (d²(N-1) + X²P(1-P))
            var X2 = 3.841; // chi-square at .05 with df=1
            var Pk = 0.5; var dk = 0.05;
            var nk = Math.ceil((X2 * Nk * Pk * (1-Pk)) / (dk*dk*(Nk-1) + X2*Pk*(1-Pk)));
            result = {'สูตร':'Krejcie & Morgan','จำนวนประชากร (N)':Nk,'ขนาดตัวอย่าง (S)':nk};
            extras.push({title:'Krejcie & Morgan Table Formula (α=.05, P=.50)',data:[result]});
            // Add common table values
            var tableVals = [[10,10],[50,44],[100,80],[200,132],[300,169],[400,196],[500,217],[750,254],[1000,278],[1500,306],[2000,322],[3000,341],[5000,357],[10000,370],[50000,381],[100000,384]];
            var tableRows = tableVals.map(function(v){return {'Population (N)':v[0],'Sample Size (S)':v[1]};});
            extras.push({title:'ตาราง Krejcie & Morgan (อ้างอิง)',data:tableRows});
        } else if (formula === 'gpower') {
            var test = getSelectValue('ss-gp-test') || 'ttest';
            var es = parseFloat(document.getElementById('ss-gp-es').value) || 0.5;
            var alpha = parseFloat(document.getElementById('ss-gp-alpha').value) || 0.05;
            var power = parseFloat(document.getElementById('ss-gp-power').value) || 0.80;
            var groups = parseInt(document.getElementById('ss-gp-groups').value) || 2;
            // Approximate using Cohen's formulas
            var za = jStat.normal.inv(1 - alpha/2, 0, 1);
            var zb = jStat.normal.inv(power, 0, 1);
            var nGP = 0;
            if (test === 'ttest') { nGP = Math.ceil(2 * Math.pow((za + zb) / es, 2)); }
            else if (test === 'paired') { nGP = Math.ceil(Math.pow((za + zb) / es, 2)); }
            else if (test === 'anova') { nGP = Math.ceil(Math.pow((za + zb), 2) / (es * es) * groups); }
            else if (test === 'correlation') { nGP = Math.ceil(Math.pow((za + zb), 2) / (es * es) + 3); }
            else if (test === 'chi-square') { nGP = Math.ceil(Math.pow((za + zb), 2) / (es * es)); }
            else if (test === 'regression') { var f2 = es; nGP = Math.ceil((Math.pow(za+zb,2) * (1+groups*f2)) / f2 + groups + 1); }
            result = {'Test':test,'Effect Size':es,'Alpha':alpha,'Power':power,'Groups/Predictors':groups,'Sample Size (per group)':nGP,'Total Sample':test==='ttest'?nGP*2:(test==='anova'?nGP:nGP)};
            extras.push({title:'G*Power Approximation',data:[result]});
        } else if (formula === 'proportion') {
            var zp = parseFloat(getSelectValue('ss-prop-z')) || 1.96;
            var ep = parseFloat(document.getElementById('ss-prop-e').value) || 0.05;
            var pp = parseFloat(document.getElementById('ss-prop-p').value) || 0.50;
            var Np = document.getElementById('ss-prop-pop').value ? parseFloat(document.getElementById('ss-prop-pop').value) : null;
            var n0 = Math.ceil((zp * zp * pp * (1 - pp)) / (ep * ep));
            var nFinal = n0;
            if (Np && Np > 0) { nFinal = Math.ceil(n0 / (1 + (n0 - 1) / Np)); }
            result = {'สูตร':'Cochran Proportion','Z':zp,'p':pp,'e':ep,'n₀ (ไม่จำกัดประชากร)':n0};
            if (Np) result['Population (N)'] = Np;
            result['ขนาดตัวอย่าง (n)'] = nFinal;
            extras.push({title:'Cochran Proportion: n₀ = Z²p(1-p)/e²' + (Np ? ', adjusted for finite population' : ''),data:[result]});
        }

        state.results['ss'] = {data:[result], title:'ผลการคำนวณขนาดตัวอย่าง — ' + formula.toUpperCase(), extras:extras};
        displayResults('ss');
    }

    // =========================================================================
    // ONE-SAMPLE t-TEST
    // =========================================================================
    function runOneSampleTTest() {
        var varName = getCheckedVars('ost-var-picker')[0];
        if (!varName) { alert('กรุณาเลือกตัวแปร'); return; }
        var testVal = parseFloat(document.getElementById('ost-testval').value) || 0;
        var alpha = parseFloat(getSelectValue('ost-alpha')) || 0.05;
        var values = state.data.map(function(r){return parseFloat(r[varName]);}).filter(function(v){return !isNaN(v);});
        var n = values.length;
        if (n < 2) { alert('ข้อมูลไม่เพียงพอ'); return; }
        var mean = jStat.mean(values);
        var sd = jStat.stdev(values, true);
        var se = sd / Math.sqrt(n);
        var t = (mean - testVal) / se;
        var df = n - 1;
        var p = 2 * (1 - jStat.studentt.cdf(Math.abs(t), df));
        var d = (mean - testVal) / sd;
        var ci95lo = mean - jStat.studentt.inv(1-alpha/2,df)*se;
        var ci95hi = mean + jStat.studentt.inv(1-alpha/2,df)*se;
        var sig = p < alpha ? 'Sig.' : 'Not Sig.';

        var extras = [{title:'Descriptive Statistics',data:[{Variable:varName,N:n,Mean:fmt(mean),'S.D.':fmt(sd),'S.E.':fmt(se),'Test Value':testVal}]}];
        var data = [{Variable:varName,t:fmt(t),df:df,'Sig. (2-tailed)':fmt(p),'Mean Diff.':fmt(mean-testVal),'95% CI Lower':fmt(ci95lo),'95% CI Upper':fmt(ci95hi),"Cohen's d":fmt(d),Result:sig}];

        state.results['ost'] = {data:data, title:'One-Sample t-test: '+varName+' vs '+testVal, extras:extras};
        displayResults('ost');
    }

    // =========================================================================
    // SIGN TEST
    // =========================================================================
    function runSignTest() {
        var before = getCheckedVars('sign-before-picker')[0];
        var after = getCheckedVars('sign-after-picker')[0];
        if (!before || !after) { alert('กรุณาเลือกตัวแปร Before และ After'); return; }
        var alpha = parseFloat(getSelectValue('sign-alpha')) || 0.05;
        var pairs = [];
        state.data.forEach(function(r){
            var b = parseFloat(r[before]), a = parseFloat(r[after]);
            if (!isNaN(b) && !isNaN(a)) pairs.push({b:b,a:a,diff:a-b});
        });
        var positive = pairs.filter(function(p){return p.diff>0;}).length;
        var negative = pairs.filter(function(p){return p.diff<0;}).length;
        var ties = pairs.filter(function(p){return p.diff===0;}).length;
        var n = positive + negative; // exclude ties
        var k = Math.min(positive, negative);
        // Binomial test: p = 2 * sum(C(n,i) * 0.5^n) for i=0..k
        var p = 0;
        for (var i = 0; i <= k; i++) { p += jStat.combination(n, i) * Math.pow(0.5, n); }
        p = Math.min(p * 2, 1);
        var sig = p < alpha ? 'Sig.' : 'Not Sig.';

        var data = [{'Before':before,'After':after,'Positive Diffs':positive,'Negative Diffs':negative,'Ties':ties,'N (excl. ties)':n,'p-value':fmt(p),'Alpha':alpha,'Result':sig}];
        state.results['sign'] = {data:data, title:'Sign Test: '+before+' vs '+after, extras:[]};
        displayResults('sign');
    }

    // =========================================================================
    // BINOMIAL TEST
    // =========================================================================
    function runBinomialTest() {
        var varName = getCheckedVars('binom-var-picker')[0];
        if (!varName) { alert('กรุณาเลือกตัวแปร'); return; }
        var expectedP = parseFloat(document.getElementById('binom-p').value) || 0.50;
        var alpha = parseFloat(getSelectValue('binom-alpha')) || 0.05;
        var values = state.data.map(function(r){return parseFloat(r[varName]);}).filter(function(v){return v===0||v===1;});
        var n = values.length;
        if (n < 1) { alert('ข้อมูลต้องเป็น 0/1 (Binary)'); return; }
        var successes = values.filter(function(v){return v===1;}).length;
        var obsProp = successes / n;
        // Two-sided binomial test
        var p = 0;
        for (var i = 0; i <= n; i++) {
            var prob_i = jStat.combination(n, i) * Math.pow(expectedP, i) * Math.pow(1-expectedP, n-i);
            var prob_obs = jStat.combination(n, successes) * Math.pow(expectedP, successes) * Math.pow(1-expectedP, n-successes);
            if (prob_i <= prob_obs + 1e-10) p += prob_i;
        }
        p = Math.min(p, 1);
        var sig = p < alpha ? 'Sig.' : 'Not Sig.';

        var data = [{Variable:varName,N:n,'Successes (1)':successes,'Observed Proportion':fmt(obsProp),'Expected Proportion':expectedP,'p-value':fmt(p),'Alpha':alpha,'Result':sig}];
        state.results['binom'] = {data:data, title:'Binomial Test: '+varName, extras:[]};
        displayResults('binom');
    }

    // =========================================================================
    // HOMOGENEITY OF VARIANCE (Levene's Test)
    // =========================================================================
    function runHomogeneity() {
        var dvName = getCheckedVars('hov-dv-picker')[0];
        var ivName = getCheckedVars('hov-iv-picker')[0];
        if (!dvName || !ivName) { alert('กรุณาเลือก DV และ Grouping Variable'); return; }
        var groups = {};
        state.data.forEach(function(r){
            var g = String(r[ivName]);
            var v = parseFloat(r[dvName]);
            if (!isNaN(v) && g) { if (!groups[g]) groups[g] = []; groups[g].push(v); }
        });
        var gNames = Object.keys(groups);
        if (gNames.length < 2) { alert('ต้องมีอย่างน้อย 2 กลุ่ม'); return; }
        // Levene's Test (using mean)
        var allMedians = {};
        gNames.forEach(function(g){ allMedians[g] = jStat.mean(groups[g]); });
        var zScores = {};
        gNames.forEach(function(g){ zScores[g] = groups[g].map(function(v){ return Math.abs(v - allMedians[g]); }); });
        var allZ = []; gNames.forEach(function(g){ allZ = allZ.concat(zScores[g]); });
        var grandMeanZ = jStat.mean(allZ);
        var groupMeansZ = {};
        gNames.forEach(function(g){ groupMeansZ[g] = jStat.mean(zScores[g]); });
        var N = allZ.length; var k = gNames.length;
        var SSB = 0; gNames.forEach(function(g){ SSB += zScores[g].length * Math.pow(groupMeansZ[g] - grandMeanZ, 2); });
        var SSW = 0; gNames.forEach(function(g){ zScores[g].forEach(function(z){ SSW += Math.pow(z - groupMeansZ[g], 2); }); });
        var df1 = k - 1; var df2 = N - k;
        var F = (SSB / df1) / (SSW / df2);
        var p = 1 - jStat.centralF.cdf(F, df1, df2);

        var descRows = gNames.map(function(g){ return {Group:g, N:groups[g].length, Mean:fmt(jStat.mean(groups[g])), 'S.D.':fmt(jStat.stdev(groups[g],true)), Variance:fmt(jStat.variance(groups[g],true))}; });
        var extras = [{title:'Descriptive by Group',data:descRows}];
        var data = [{"Test":"Levene's Test (based on Mean)",'F':fmt(F),'df1':df1,'df2':df2,'Sig.':fmt(p),'Result':p<0.05?'Variance NOT equal':'Variance Equal'}];

        state.results['hov'] = {data:data, title:"Homogeneity of Variance: "+dvName+" by "+ivName, extras:extras};
        displayResults('hov');
    }

    // =========================================================================
    // MANOVA (simplified — Wilks' Lambda approximation)
    // =========================================================================
    function runManova() {
        var dvNames = getCheckedVars('manova-dv-picker');
        var ivName = getCheckedVars('manova-iv-picker')[0];
        if (!dvNames || dvNames.length < 2 || !ivName) { alert('กรุณาเลือก DV อย่างน้อย 2 ตัว และ Factor 1 ตัว'); return; }
        var alpha = parseFloat(getSelectValue('manova-alpha')) || 0.05;
        // Group data
        var groups = {};
        state.data.forEach(function(r) {
            var g = String(r[ivName]);
            if (!g) return;
            if (!groups[g]) groups[g] = [];
            var row = {};
            dvNames.forEach(function(dv) { row[dv] = parseFloat(r[dv]); });
            if (dvNames.every(function(dv){return !isNaN(row[dv]);})) groups[g].push(row);
        });
        var gNames = Object.keys(groups);
        if (gNames.length < 2) { alert('Factor ต้องมีอย่างน้อย 2 กลุ่ม'); return; }

        // Run individual ANOVAs for each DV and combine
        var anovaResults = [];
        dvNames.forEach(function(dv) {
            var grpData = {};
            gNames.forEach(function(g) { grpData[g] = groups[g].map(function(r){return r[dv];}); });
            var allVals = []; gNames.forEach(function(g){allVals=allVals.concat(grpData[g]);});
            var grandMean = jStat.mean(allVals);
            var N = allVals.length; var k = gNames.length;
            var SSB=0,SSW=0;
            gNames.forEach(function(g){
                var gm=jStat.mean(grpData[g]);
                SSB+=grpData[g].length*Math.pow(gm-grandMean,2);
                grpData[g].forEach(function(v){SSW+=Math.pow(v-gm,2);});
            });
            var df1=k-1,df2=N-k;
            var F=(SSB/df1)/(SSW/df2);
            var p=1-jStat.centralF.cdf(F,df1,df2);
            anovaResults.push({DV:dv,F:fmt(F),df1:df1,df2:df2,'Sig.':fmt(p),'Result':p<alpha?'Sig.':'Not Sig.'});
        });

        // Descriptive
        var descRows = [];
        gNames.forEach(function(g) {
            var row = {Group:g,N:groups[g].length};
            dvNames.forEach(function(dv){
                var vals = groups[g].map(function(r){return r[dv];});
                row['Mean('+dv+')'] = fmt(jStat.mean(vals));
                row['SD('+dv+')'] = fmt(jStat.stdev(vals,true));
            });
            descRows.push(row);
        });
        var extras = [{title:'Descriptive Statistics by Group',data:descRows},{title:'Univariate ANOVA Results (per DV)',data:anovaResults}];

        // Approximate Wilks' Lambda (product of 1/(1+eigenvalue))
        // Simplified: use product of (SSW/(SSW+SSB)) per DV
        var wilks = 1;
        dvNames.forEach(function(dv) {
            var grpData = {};
            gNames.forEach(function(g) { grpData[g] = groups[g].map(function(r){return r[dv];}); });
            var allVals = []; gNames.forEach(function(g){allVals=allVals.concat(grpData[g]);});
            var grandMean = jStat.mean(allVals);
            var SSB=0,SSW=0;
            gNames.forEach(function(g){
                var gm=jStat.mean(grpData[g]);
                SSB+=grpData[g].length*Math.pow(gm-grandMean,2);
                grpData[g].forEach(function(v){SSW+=Math.pow(v-gm,2);});
            });
            wilks *= SSW/(SSW+SSB);
        });

        var data = [{"Wilks' Lambda":fmt(wilks),'DVs':dvNames.join(', '),'Factor':ivName,'Groups':gNames.length,'Note':'ดู Univariate ANOVA ด้านบนสำหรับผลรายตัวแปร'}];
        state.results['manova'] = {data:data, title:'MANOVA: '+dvNames.join(', ')+' by '+ivName, extras:extras};
        displayResults('manova');
    }

    // =========================================================================
    // GAMES-HOWELL POST-HOC
    // =========================================================================
    function runGamesHowell() {
        var dvName = getCheckedVars('gh-dv-picker')[0];
        var ivName = getCheckedVars('gh-iv-picker')[0];
        if (!dvName || !ivName) { alert('กรุณาเลือก DV และ Factor'); return; }
        var groups = {};
        state.data.forEach(function(r){
            var g=String(r[ivName]),v=parseFloat(r[dvName]);
            if(!isNaN(v)&&g){if(!groups[g])groups[g]=[];groups[g].push(v);}
        });
        var gNames = Object.keys(groups);
        if (gNames.length < 3) { alert('ต้องมีอย่างน้อย 3 กลุ่ม'); return; }

        var results = [];
        for (var i=0;i<gNames.length;i++){
            for (var j=i+1;j<gNames.length;j++){
                var g1=groups[gNames[i]],g2=groups[gNames[j]];
                var m1=jStat.mean(g1),m2=jStat.mean(g2);
                var v1=jStat.variance(g1,true),v2=jStat.variance(g2,true);
                var n1=g1.length,n2=g2.length;
                var se=Math.sqrt(v1/n1+v2/n2);
                var t=(m1-m2)/se;
                var df_num=Math.pow(v1/n1+v2/n2,2);
                var df_den=Math.pow(v1/n1,2)/(n1-1)+Math.pow(v2/n2,2)/(n2-1);
                var df=df_num/df_den;
                var p=2*(1-jStat.studentt.cdf(Math.abs(t),df));
                results.push({'Group (I)':gNames[i],'Group (J)':gNames[j],'Mean Diff (I-J)':fmt(m1-m2),'S.E.':fmt(se),'t':fmt(t),'df':fmt(df),'Sig.':fmt(p),'Result':p<0.05?'Sig.':'Not Sig.'});
            }
        }
        state.results['gh'] = {data:results, title:'Games-Howell Post-hoc: '+dvName+' by '+ivName, extras:[]};
        displayResults('gh');
    }

    // =========================================================================
    // MULTIPLE REGRESSION
    // =========================================================================
    function runMultipleRegression() {
        var dvName = getCheckedVars('mreg-dv-picker')[0];
        var ivNames = getCheckedVars('mreg-iv-picker');
        if (!dvName || !ivNames || ivNames.length < 2) { alert('กรุณาเลือก DV 1 ตัว และ IV อย่างน้อย 2 ตัว'); return; }

        var method = getSelectValue('mreg-method') || 'enter';
        var methodLabels = {'enter':'Enter','stepwise':'Stepwise','forward':'Forward','backward':'Backward'};
        var methodLabel = methodLabels[method] || 'Enter';

        // Build data matrix
        var dataRows = [];
        state.data.forEach(function(r){
            var y = parseFloat(r[dvName]);
            var xs = ivNames.map(function(iv){return parseFloat(r[iv]);});
            if (!isNaN(y) && xs.every(function(x){return !isNaN(x);})) dataRows.push({y:y,xs:xs});
        });
        var n = dataRows.length; var origP = ivNames.length;
        if (n <= origP+1) { alert('ข้อมูลไม่เพียงพอสำหรับจำนวนตัวแปรอิสระ'); return; }

        var Y = dataRows.map(function(r){return r.y;});
        var allXData = ivNames.map(function(_,idx){ return dataRows.map(function(r){return r.xs[idx];}); });

        // Variable selection
        var selectedIVs = ivNames.slice();
        var selectedXData = allXData.slice();
        var stepLog = [];

        if (method !== 'enter') {
            var pIn = 0.05, pOut = 0.10;
            var stepResult = _stepwiseRegression(Y, allXData, ivNames, method, pIn, pOut);
            selectedIVs = stepResult.selectedVars;
            selectedXData = stepResult.selectedXData;
            stepLog = stepResult.steps;
            if (selectedIVs.length === 0) { alert('ไม่มีตัวแปรใดผ่านเกณฑ์'); return; }
        }

        var p = selectedIVs.length;

        // OLS
        var X = dataRows.map(function(r){
            var row = [1];
            selectedIVs.forEach(function(iv){
                var idx = ivNames.indexOf(iv);
                row.push(r.xs[idx]);
            });
            return row;
        });
        var Xt = jStat.transpose(X);
        var XtX = jStat.multiply(Xt,X);
        var XtXinv;
        try { XtXinv = jStat.inv(XtX); } catch(e) { alert('Matrix singular — Multicollinearity สูง'); return; }
        var XtY = jStat.multiply(Xt,[Y]);
        var beta = jStat.multiply(XtXinv,jStat.transpose(XtY));
        var betas = beta.map(function(b){return b[0];});

        var yPred = dataRows.map(function(r,i){
            var pred = betas[0];
            selectedIVs.forEach(function(iv,j){
                var idx = ivNames.indexOf(iv);
                pred += betas[j+1]*r.xs[idx];
            });
            return pred;
        });
        var yMean = jStat.mean(Y);
        var SST=0,SSR=0,SSE=0;
        Y.forEach(function(y,i){SST+=Math.pow(y-yMean,2);SSR+=Math.pow(yPred[i]-yMean,2);SSE+=Math.pow(y-yPred[i],2);});
        var R2 = SSR/SST;
        var adjR2 = 1-(1-R2)*(n-1)/(n-p-1);
        var MSR = SSR/p; var MSE = SSE/(n-p-1);
        var F = MSR/MSE;
        var pF = 1-jStat.centralF.cdf(F,p,n-p-1);

        // Durbin-Watson
        var residuals = Y.map(function(y,i){return y - yPred[i];});
        var dwNum = 0;
        for (var di=1;di<residuals.length;di++) dwNum+=Math.pow(residuals[di]-residuals[di-1],2);
        var dwDen = 0; residuals.forEach(function(r){dwDen+=r*r;});
        var dw = dwDen > 0 ? dwNum/dwDen : 0;

        var extras = [];

        // 1. Method
        var methodDescs = {
            'enter':'ใส่ตัวแปรอิสระทั้งหมดเข้าสมการพร้อมกัน',
            'stepwise':'เพิ่ม/ลดตัวแปรอัตโนมัติ (p-in=0.05, p-out=0.10)',
            'forward':'เพิ่มตัวแปรทีละตัว (p-in=0.05)',
            'backward':'เริ่มจากทุกตัว แล้วตัดทีละตัว (p-out=0.10)'
        };
        extras.push({title:'Method / กระบวนการสร้างโมเดล',data:[{
            'Method':methodLabel,'Description':methodDescs[method]||'','DV':dvName,
            'Candidate IVs':ivNames.join(', '),'Selected IVs':selectedIVs.join(', '),
            'Entered':p,'Excluded':origP-p,'N':n
        }]});

        // 2. Step log
        if (stepLog.length > 0) extras.push({title:'Variable Selection Steps',data:stepLog});

        // 3. Model Summary
        extras.push({title:'Model Summary',data:[{
            'Method':methodLabel,R:fmt(Math.sqrt(R2)),'R²':fmt(R2),'Adjusted R²':fmt(adjR2),
            'Std. Error':fmt(Math.sqrt(MSE)),'Durbin-Watson':fmt(dw),N:n
        }]});

        // 4. ANOVA
        extras.push({title:'ANOVA',data:[
            {Source:'Regression',SS:fmt(SSR),df:p,MS:fmt(MSR),F:fmt(F),'Sig.':fmt(pF)},
            {Source:'Residual',SS:fmt(SSE),df:n-p-1,MS:fmt(MSE),F:'','Sig.':''},
            {Source:'Total',SS:fmt(SST),df:n-1,MS:'',F:'','Sig.':''}
        ]});

        // 5. Collinearity
        if (p > 1) {
            var collinRows = [];
            selectedIVs.forEach(function(ivN,idx){
                var xi = selectedXData[idx];
                var otherX = selectedXData.filter(function(_,j){return j!==idx;});
                var otherNms = ['(Intercept)'].concat(selectedIVs.filter(function(_,j){return j!==idx;}));
                var regI = Stats.linearRegression(xi, otherX, otherNms);
                var r2i = regI ? regI.rSquared : 0;
                var tol = 1-r2i; var vif = tol>0 ? 1/tol : 999;
                collinRows.push({Variable:ivN,Tolerance:fmt(tol),VIF:fmt(vif),Status:vif>10?'Multicollinearity!':(vif>5?'Warning':'OK')});
            });
            extras.push({title:'Collinearity Diagnostics',data:collinRows});
        }

        // Coefficients
        var se_beta = [];
        for (var i=0;i<=p;i++) se_beta.push(Math.sqrt(XtXinv[i][i]*MSE));
        var coeffRows = [{Variable:'(Constant)',B:fmt(betas[0]),'S.E.':fmt(se_beta[0]),'Beta (Std.)':'—',t:fmt(betas[0]/se_beta[0]),'Sig.':fmt(2*(1-jStat.studentt.cdf(Math.abs(betas[0]/se_beta[0]),n-p-1)))}];
        var sdY = jStat.stdev(Y,true);
        selectedIVs.forEach(function(iv,idx){
            var origIdx = ivNames.indexOf(iv);
            var sdX = jStat.stdev(dataRows.map(function(r){return r.xs[origIdx];}),true);
            var stdBeta = betas[idx+1]*(sdX/sdY);
            var tVal = betas[idx+1]/se_beta[idx+1];
            var pVal = 2*(1-jStat.studentt.cdf(Math.abs(tVal),n-p-1));
            var ci95Lo = betas[idx+1] - jStat.studentt.inv(0.975,n-p-1)*se_beta[idx+1];
            var ci95Hi = betas[idx+1] + jStat.studentt.inv(0.975,n-p-1)*se_beta[idx+1];
            // VIF inline
            var xi = selectedXData[idx];
            var otherX2 = selectedXData.filter(function(_,j){return j!==idx;});
            var otherNms2 = ['(Intercept)'].concat(selectedIVs.filter(function(_,j){return j!==idx;}));
            var regI2 = p>1 ? Stats.linearRegression(xi, otherX2, otherNms2) : null;
            var tol2 = regI2 ? 1-regI2.rSquared : 1;
            var vif2 = tol2>0 ? 1/tol2 : 999;
            coeffRows.push({Variable:iv,B:fmt(betas[idx+1]),'S.E.':fmt(se_beta[idx+1]),'Beta (Std.)':fmt(stdBeta),t:fmt(tVal),'Sig.':fmt(pVal),'95% CI':'['+fmt(ci95Lo)+', '+fmt(ci95Hi)+']',Tolerance:fmt(tol2),VIF:fmt(vif2)});
        });

        // 6. Residual Statistics
        var sdRes = jStat.stdev(residuals,true);
        var stdRes = residuals.map(function(r){return r/sdRes;});
        extras.push({title:'Residual Statistics',data:[
            {Statistic:'Predicted Value',Min:fmt(jStat.min(yPred)),Max:fmt(jStat.max(yPred)),Mean:fmt(jStat.mean(yPred)),'S.D.':fmt(jStat.stdev(yPred,true))},
            {Statistic:'Residual',Min:fmt(jStat.min(residuals)),Max:fmt(jStat.max(residuals)),Mean:fmt(jStat.mean(residuals)),'S.D.':fmt(sdRes)},
            {Statistic:'Std. Residual',Min:fmt(jStat.min(stdRes)),Max:fmt(jStat.max(stdRes)),Mean:fmt(jStat.mean(stdRes)),'S.D.':fmt(jStat.stdev(stdRes,true))}
        ]});

        // 7. Excluded Variables
        if (method !== 'enter') {
            var excl = ivNames.filter(function(v){return selectedIVs.indexOf(v)===-1;});
            if (excl.length > 0) {
                var exclRows = excl.map(function(v){
                    var idx = ivNames.indexOf(v);
                    var r = jStat.corrcoeff(allXData[idx], Y);
                    return {Variable:v,'Bivariate Corr.':fmt(r),Status:'Excluded'};
                });
                extras.push({title:'Excluded Variables',data:exclRows});
            }
        }

        // 8. Regression Equation
        var eqParts2 = [dvName + ' = ' + fmt(betas[0])];
        selectedIVs.forEach(function(iv, idx) {
            var sign = betas[idx+1] >= 0 ? ' + ' : ' - ';
            eqParts2.push(sign + fmt(Math.abs(betas[idx+1])) + '(' + iv + ')');
        });
        var unstdEq2 = eqParts2.join('');
        var stdParts2 = [dvName + '(Z) ='];
        coeffRows.forEach(function(c, idx) {
            if (idx === 0) return;
            var betaVal = parseFloat(c['Beta (Std.)']) || 0;
            var sign = betaVal >= 0 ? (idx === 1 ? ' ' : ' + ') : ' - ';
            stdParts2.push(sign + fmt(Math.abs(betaVal)) + '(Z_' + c.Variable + ')');
        });
        extras.push({title:'สมการถดถอย (Regression Equation)',data:[
            {'สมการ Unstandardized (ค่าดิบ)': unstdEq2},
            {'สมการ Standardized (ค่ามาตรฐาน)': stdParts2.join('')}
        ]});

        // 9. Diagnostic Warnings
        var mregWarnings = _buildRegressionWarnings2(R2, adjR2, pF, dw, n, p, coeffRows, selectedIVs, selectedXData, Y);
        if (mregWarnings.length > 0) {
            extras.push({ title: '⚠️ จุดสังเกต / ข้อควรระวัง (Diagnostics)', data: mregWarnings });
        }

        state.results['mreg'] = {data:coeffRows, title:'Multiple Regression Coefficients (Method: '+methodLabel+')', extras:extras};
        displayResults('mreg');
    }

    // =========================================================================
    // COHEN'S KAPPA
    // =========================================================================
    function runKappa() {
        var r1Name = getCheckedVars('kap-r1-picker')[0];
        var r2Name = getCheckedVars('kap-r2-picker')[0];
        if (!r1Name || !r2Name) { alert('กรุณาเลือกตัวแปรผู้ประเมิน 2 คน'); return; }
        var pairs = [];
        state.data.forEach(function(r){
            var a=String(r[r1Name]).trim(),b=String(r[r2Name]).trim();
            if(a&&b) pairs.push({a:a,b:b});
        });
        var n=pairs.length;
        if(n<2){alert('ข้อมูลไม่เพียงพอ');return;}
        var cats=[]; pairs.forEach(function(p){if(cats.indexOf(p.a)===-1)cats.push(p.a);if(cats.indexOf(p.b)===-1)cats.push(p.b);});
        cats.sort();
        // Build confusion matrix
        var matrix = {};
        cats.forEach(function(c1){matrix[c1]={};cats.forEach(function(c2){matrix[c1][c2]=0;});});
        pairs.forEach(function(p){matrix[p.a][p.b]++;});
        var po=0; cats.forEach(function(c){po+=matrix[c][c];}); po/=n;
        var pe=0;
        cats.forEach(function(c){
            var row=0,col=0;
            cats.forEach(function(c2){row+=matrix[c][c2];col+=matrix[c2][c];});
            pe+=(row/n)*(col/n);
        });
        var kappa = (po-pe)/(1-pe);
        var interp = kappa<0.20?'Poor':kappa<0.40?'Fair':kappa<0.60?'Moderate':kappa<0.80?'Good':'Very Good';

        var data = [{Rater1:r1Name,Rater2:r2Name,N:n,'Observed Agreement (Po)':fmt(po),'Expected Agreement (Pe)':fmt(pe),"Cohen's Kappa":fmt(kappa),'Interpretation':interp}];
        state.results['kap'] = {data:data, title:"Cohen's Kappa: "+r1Name+' vs '+r2Name, extras:[]};
        displayResults('kap');
    }

    // =========================================================================
    // CFA, ITEM ANALYSIS, PATH ANALYSIS (placeholder with message)
    // =========================================================================
    function runCFA() {
        alert('CFA requires advanced matrix operations. กรุณาใช้ AI Chat เพื่อขอคำแนะนำเรื่อง CFA หรือใช้ EFA เป็นทางเลือก');
    }
    function runItemAnalysis() {
        var itemNames = getCheckedVars('item-vars-picker');
        if (!itemNames || itemNames.length < 2) { alert('กรุณาเลือกข้อสอบอย่างน้อย 2 ข้อ'); return; }
        var rows = [];
        state.data.forEach(function(r){
            var scores = {};
            itemNames.forEach(function(item){ scores[item] = parseFloat(r[item]); });
            if (itemNames.every(function(item){return !isNaN(scores[item]);})) rows.push(scores);
        });
        var n = rows.length;
        if (n < 5) { alert('ข้อมูลไม่เพียงพอ (ต้องมีอย่างน้อย 5 คน)'); return; }
        // Total scores
        var totals = rows.map(function(r){var s=0;itemNames.forEach(function(item){s+=r[item];});return s;});
        var results = [];
        itemNames.forEach(function(item){
            var itemScores = rows.map(function(r){return r[item];});
            var difficulty = jStat.mean(itemScores);
            // Point-biserial or Pearson correlation with total
            var totalMinusItem = totals.map(function(t,i){return t-itemScores[i];});
            var corr = jStat.corrcoeff(itemScores, totalMinusItem);
            results.push({Item:item,N:n,'Mean (Difficulty)':fmt(difficulty),'S.D.':fmt(jStat.stdev(itemScores,true)),'Item-Total Correlation':fmt(corr),'Quality':corr>=0.3?'Good':(corr>=0.2?'Acceptable':'Poor')});
        });
        state.results['item'] = {data:results, title:'Item Analysis: '+itemNames.length+' items, N='+n, extras:[]};
        displayResults('item');
    }
    function runPathAnalysis() {
        alert('Path Analysis requires SEM engine. กรุณาใช้ AI Chat เพื่อขอคำแนะนำหรือใช้ Multiple Regression เป็นทางเลือก');
    }

    // Expose global functions for onclick handlers in HTML
    // =========================================================================

    window.handleLogin = handleLogin;
    window.handleLogout = handleLogout;
    window.navigateTo = navigateTo;
    window.toggleSidebar = toggleSidebar;
    window.toggleMenuGroup = toggleMenuGroup;
    window.runAnalysis = runAnalysis;
    window.aiAnalyze = aiAnalyze;
    window.sendChat = sendChat;
    window.sendQuickChat = sendQuickChat;
    window.exportExcel = exportExcel;
    window.exportWord = exportWord;
    window.downloadTemplate = downloadTemplate;
    window.saveAISettings = saveAISettings;
    window.testAIConnection = testAIConnection;
    window.clearChatHistory = clearChatHistory;
    window.switchTab = switchTab;
    window.createDualListPicker = createDualListPicker;
    window.dualListMove = dualListMove;
    window.getDualListSelected = getDualListSelected;
    window.runAssumptionCheck = runAssumptionCheck;
    window.createCheckboxPicker = createCheckboxPicker;
    window.getCheckedVars = getCheckedVars;
    window.openVarModal = openVarModal;
    window.closeVarModal = closeVarModal;
    window.confirmVarModal = confirmVarModal;
    window.toggleVarPill = toggleVarPill;
    window.varModalSelectAll = varModalSelectAll;
    window.varModalDeselectAll = varModalDeselectAll;
    window.removePickerVar = removePickerVar;
    window.updateLikertCriteria = updateLikertCriteria;
    window.openIntervalConfig = openIntervalConfig;
    window.closeIntervalConfig = closeIntervalConfig;
    window.applyIntervalConfig = applyIntervalConfig;
    window.toggleIntvMode = toggleIntvMode;
    window.openDemoIntervalConfig = openDemoIntervalConfig;
    window.closeDemoIntervalModal = closeDemoIntervalModal;
    window.applyDemoIntervals = applyDemoIntervals;
    window.toggleFollowupChat = toggleFollowupChat;
    window.sendFollowup = sendFollowup;
    window.sendFollowupFromInput = sendFollowupFromInput;
    window.toggleSampleSizeInputs = toggleSampleSizeInputs;

    // =========================================================================
    // Sample Size Calculator — toggle inputs based on formula
    // =========================================================================
    function toggleSampleSizeInputs() {
        var formula = getSelectValue('ss-formula') || 'yamane';
        var sections = ['yamane','cochran','krejcie','gpower','proportion'];
        sections.forEach(function(s) {
            var el = document.getElementById('ss-' + s + '-inputs');
            if (el) el.style.display = (s === formula) ? '' : 'none';
        });
    }

})();
