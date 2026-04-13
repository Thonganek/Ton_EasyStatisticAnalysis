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
            'effect-size': 'cd'
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
        if (!state.data && type.indexOf('effect-') !== 0 && type.indexOf('assumption-') !== 0) {
            if (!state.data) {
                alert('Please upload a data file first.');
                return;
            }
        }
        if (!state.data) {
            alert('Please upload a data file first.');
            return;
        }

        try {
            switch (type) {
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
                default:
                    alert('Analysis type "' + type + '" is not yet implemented.');
            }
        } catch (err) {
            alert('Error running analysis: ' + err.message);
            console.error(err);
        }
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

        var yData = getColumnData(dv, true);
        var xData = ivs.map(function (iv) { return getColumnData(iv, true); });
        var minLen = Math.min(yData.length, Math.min.apply(null, xData.map(function (x) { return x.length; })));
        yData = yData.slice(0, minLen);
        xData = xData.map(function (x) { return x.slice(0, minLen); });

        var names = ['(Intercept)'].concat(ivs);

        // Run assumption check
        runAssumptionCheck('lr', 'correlation', { vars: xData });

        var result = Stats.linearRegression(yData, xData, names);
        if (!result) { alert('Could not compute regression. Check your data (possible multicollinearity).'); return; }

        var extras = [{
            title: 'Model Summary',
            data: [{
                R: fmt(result.r), 'R-Squared': fmt(result.rSquared),
                'Adj. R-Squared': fmt(result.adjRSquared),
                'Durbin-Watson': fmt(result.durbinWatson),
                F: fmt(result.f), 'p-value': Stats.formatPValue(result.fP)
            }]
        }, {
            title: 'ANOVA',
            data: [
                { Source: 'Regression', SS: fmt(result.anova.ssReg), df: result.anova.dfReg, MS: fmt(result.anova.msReg), F: fmt(result.f), 'p-value': Stats.formatPValue(result.fP) },
                { Source: 'Residual', SS: fmt(result.anova.ssRes), df: result.anova.dfRes, MS: fmt(result.anova.msRes), F: '', 'p-value': '' },
                { Source: 'Total', SS: fmt(result.anova.ssTotal), df: result.anova.dfReg + result.anova.dfRes, MS: '', F: '', 'p-value': '' }
            ]
        }];

        var mainRows = result.coefficients.map(function (c) {
            return {
                Variable: c.variable,
                B: fmt(c.b), 'S.E.': fmt(c.se), t: fmt(c.t),
                'p-value': Stats.formatPValue(c.p),
                '95% CI': Stats.formatCI ? Stats.formatCI(c.ci95Lo, c.ci95Hi) : '[' + fmt(c.ci95Lo) + ', ' + fmt(c.ci95Hi) + ']'
            };
        });

        state.results['lr'] = { data: mainRows, title: 'Regression Coefficients', extras: extras };
        displayResults('lr');
    }

    // =========================================================================
    // Logistic Regression
    // =========================================================================

    function runLogisticRegression() {
        var dv = getPickerValue('logr-dv-picker', 'logr-dv');
        var ivs = getCheckedVars('logr-iv-picker');
        if (ivs.length === 0) ivs = getSelected('logr-ivs');
        if (!dv || ivs.length === 0) { alert('Please select DV and at least one IV.'); return; }

        var yData = getColumnData(dv, true);
        var xData = ivs.map(function (iv) { return getColumnData(iv, true); });
        var minLen = Math.min(yData.length, Math.min.apply(null, xData.map(function (x) { return x.length; })));
        yData = yData.slice(0, minLen);
        xData = xData.map(function (x) { return x.slice(0, minLen); });

        var names = ['(Intercept)'].concat(ivs);
        var result = Stats.logisticRegression(yData, xData, names);
        if (!result) { alert('Could not compute logistic regression. Check your data.'); return; }

        // Simple assumption check inline
        var logrAssumptionEl = document.getElementById('logr-assumptions');
        if (logrAssumptionEl) {
            var n = yData.length;
            var nOk = n >= 10 * (ivs.length + 1);
            var checksHtml = '<div class="assumption-panel"><h4>\ud83d\udd0d Assumption Check</h4>';
            checksHtml += '<table class="result-table"><thead><tr><th>Check</th><th>Result</th><th>Status</th></tr></thead><tbody>';
            checksHtml += '<tr><td>Sample Size (10 per predictor)</td><td>N=' + n + ', needed>=' + (10*(ivs.length+1)) + '</td><td>' + (nOk ? '\u2705' : '\u26a0\ufe0f') + '</td></tr>';
            checksHtml += '<tr><td>DV is Binary (0/1)</td><td>Check data</td><td>\u2139\ufe0f</td></tr>';
            checksHtml += '</tbody></table></div>';
            logrAssumptionEl.innerHTML = checksHtml;
            logrAssumptionEl.style.display = 'block';
        }

        var extras = [{
            title: 'Model Summary',
            data: [{ 'Accuracy': fmt(result.accuracy * 100, 2) + '%', 'AIC': fmt(result.aic) }]
        }];

        var mainRows = result.coefficients.map(function (c) {
            return {
                Variable: c.variable,
                B: fmt(c.b), 'S.E.': fmt(c.se),
                Wald: fmt(c.wald),
                'p-value': Stats.formatPValue(c.p),
                'Odds Ratio': fmt(c.or)
            };
        });

        state.results['logr'] = { data: mainRows, title: 'Logistic Regression Coefficients', extras: extras };
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
            }
        } catch (err) {
            if (aiResultEl) {
                aiResultEl.innerHTML = '<p class="error-text">Error: ' + escapeHtml(err.message) + '</p>';
                aiResultEl.style.display = '';
            }
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

})();
