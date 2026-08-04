/* ===== 飞凡AI - 超管后台 (v3.4.0 批次4：档位定制 + 参数统一) ===== */

const Admin = (function () {

    /* ---------- 模块级状态 ---------- */
    let _curTab = 'users';
    let _presetData = null;
    let _curPresetIdx = -1;
    let _openSteps = {};
    let _presetDirty = false;
    let _usersCache = [];
    let _stepSel = {};
    let _curVar = {};
    let _dragStepIdx = null;

    let _tplCache = [];
    let _tplSel = {};
    let _deployUserSel = {};
    let _healthRows = [];

    /* 批次4 */
    let _slotTplCache = [];
    let _slotTplSel = '';
    let _slotDeployUserSel = {};
    let _engCacheAll = [];        // 全部引擎（供档位下拉选引擎）

    const DEFAULT_VID = '__default__';

    /* 参数一键预设 */
    const PARAM_PRESETS = {
        strict:  { label:'🎯 严谨', useTemp:true,  temperature:0.2, useMax:false, max_tokens:4096,  useTopP:false, top_p:1, useFreq:false, frequency_penalty:0 },
        balance: { label:'⚖️ 平衡', useTemp:true,  temperature:0.7, useMax:false, max_tokens:4096,  useTopP:false, top_p:1, useFreq:false, frequency_penalty:0 },
        creative:{ label:'🎨 创意', useTemp:true,  temperature:1.0, useMax:false, max_tokens:8192,  useTopP:false, top_p:1, useFreq:false, frequency_penalty:0 },
        off:     { label:'🚫 全关', useTemp:false, temperature:0.7, useMax:false, max_tokens:4096,  useTopP:false, top_p:1, useFreq:false, frequency_penalty:0 },
    };

    async function apiCall(path, method, body) {
        const token = (typeof Auth !== 'undefined' && Auth.getToken()) ? Auth.getToken() : '';
        const opts = { method: method || 'GET', headers: { 'X-Auth-Token': token } };
        if (body) { opts.headers['Content-Type'] = 'application/json'; opts.body = JSON.stringify(body); }
        const resp = await fetch('/api/' + path.replace(/^\//, ''), opts);
        const data = await resp.json();
        if (!resp.ok) throw new Error(data.error || ('HTTP ' + resp.status));
        return data;
    }

    function loadXLSX() {
        if (window.XLSX) return Promise.resolve();
        if (typeof OfficeParser !== 'undefined' && OfficeParser.loadXLSX) return OfficeParser.loadXLSX();
        return new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
            s.onload = () => window.XLSX ? resolve() : reject(new Error('SheetJS加载失败'));
            s.onerror = () => reject(new Error('SheetJS加载失败'));
            document.head.appendChild(s);
        });
    }

    function _rid(pfx) { return pfx + Math.random().toString(36).slice(2, 8); }
    function _pj(t, fb) { if (!t) return fb; try { const v = JSON.parse(t); return v == null ? fb : v; } catch (e) { return fb; } }

    /* ========================================================== */
    /* ============ 通用表单弹窗组件 FF ========================= */
    /* ========================================================== */
    const FF = (function () {
        let _mask = null, _cfg = null;

        function close() {
            if (_mask) { _mask.remove(); _mask = null; }
            _cfg = null;
            document.removeEventListener('keydown', _onKey);
        }
        function _onKey(e) { if (_mask && e.key === 'Escape') { e.preventDefault(); close(); } }

        function _fieldHtml(f, i) {
            const id = 'ff_' + i;
            const label = esc(f.label || '');
            const hint = f.hint ? '<div class="ff-hint">' + f.hint + '</div>' : '';

            if (f.type === 'html') return '<div class="fg">' + (f.html || '') + '</div>';

            if (f.type === 'check') {
                return '<div class="fg"><label class="pt" style="margin:0">'
                    + '<input type="checkbox" id="' + id + '" ' + (f.value ? 'checked' : '') + (f.onchange ? ' onchange="' + f.onchange + '"' : '') + '> '
                    + label + '</label>' + hint + '</div>';
            }
            if (f.type === 'checks') {
                const vals = Array.isArray(f.value) ? f.value : [];
                let box = '<div class="fg"><label>' + label + '</label><div class="ff-checks" id="' + id + '">';
                (f.options || []).forEach(o => {
                    const on = vals.indexOf(o.v) >= 0;
                    box += '<label class="ff-chk' + (on ? ' on' : '') + '">'
                        + '<input type="checkbox" data-v="' + esc(o.v) + '" ' + (on ? 'checked' : '')
                        + ' onchange="this.parentNode.classList.toggle(\'on\',this.checked)"> '
                        + esc(o.t || o.v) + '</label>';
                });
                if (!(f.options || []).length) box += '<span style="font-size:12px;color:var(--text2)">（暂无可选项）</span>';
                box += '</div>' + hint + '</div>';
                return box;
            }
            if (f.type === 'select') {
                let sel = '<div class="fg"><label>' + label + '</label><select id="' + id + '"' + (f.onchange ? ' onchange="' + f.onchange + '"' : '') + '>';
                (f.options || []).forEach(o => {
                    sel += '<option value="' + esc(o.v) + '"' + (String(f.value) === String(o.v) ? ' selected' : '') + '>' + esc(o.t || o.v) + '</option>';
                });
                sel += '</select>' + hint + '</div>';
                return sel;
            }
            if (f.type === 'textarea') {
                return '<div class="fg"><label>' + label + '</label>'
                    + '<textarea id="' + id + '" rows="' + (f.rows || 4) + '" placeholder="' + esc(f.placeholder || '') + '">'
                    + esc(f.value || '') + '</textarea>' + hint + '</div>';
            }
            const t = f.type || 'text';
            return '<div class="fg"><label>' + label + '</label>'
                + '<input type="' + t + '" id="' + id + '" value="' + esc(f.value != null ? f.value : '') + '"'
                + ' placeholder="' + esc(f.placeholder || '') + '"'
                + (t === 'number' ? ' step="' + (f.step || 'any') + '"' : '')
                + ' autocomplete="off">' + hint + '</div>';
        }

        function _collect() {
            const vals = {};
            (_cfg.fields || []).forEach((f, i) => {
                if (f.type === 'html') return;
                const el = document.getElementById('ff_' + i);
                if (!el) return;
                if (f.type === 'check') { vals[f.key] = el.checked; return; }
                if (f.type === 'checks') {
                    const arr = [];
                    el.querySelectorAll('input[type="checkbox"]').forEach(cb => { if (cb.checked) arr.push(cb.getAttribute('data-v')); });
                    vals[f.key] = arr; return;
                }
                if (f.type === 'number') { vals[f.key] = parseFloat(el.value) || 0; return; }
                vals[f.key] = el.value;
            });
            return vals;
        }
        function setErr(msg) { const e = document.getElementById('ffErr'); if (e) e.textContent = msg || ''; }

        async function _submit() {
            const btn = document.getElementById('ffOk');
            const vals = _collect();
            setErr('');
            if (btn) { btn.disabled = true; btn.textContent = '处理中...'; }
            try {
                const r = await _cfg.onSubmit(vals, setErr);
                if (r !== false) close();
            } catch (e) { setErr(e.message || '操作失败'); }
            if (btn) { btn.disabled = false; btn.textContent = _cfg.okText || '保存'; }
        }

        function open(cfg) {
            close();
            _cfg = cfg || {};
            _mask = document.createElement('div');
            _mask.className = 'ff-mask';
            let body = '';
            if (_cfg.tip) body += '<div class="ff-tip">' + _cfg.tip + '</div>';
            body += '<div class="ff-err" id="ffErr"></div>';
            (_cfg.fields || []).forEach((f, i) => { body += _fieldHtml(f, i); });

            _mask.innerHTML = '<div class="ff-box"' + (_cfg.wide ? ' style="max-width:660px"' : '') + '>'
                + '<div class="ff-hdr"><h3>' + esc(_cfg.title || '') + '</h3><button class="ff-x" id="ffX">×</button></div>'
                + '<div class="ff-body">' + body + '</div>'
                + '<div class="ff-ft"><button class="btn btn-s" id="ffCancel">取消</button>'
                + '<button class="btn btn-p btn-s" id="ffOk">' + esc(_cfg.okText || '保存') + '</button></div></div>';

            document.body.appendChild(_mask);
            document.getElementById('ffX').onclick = close;
            document.getElementById('ffCancel').onclick = close;
            document.getElementById('ffOk').onclick = _submit;
            _mask.onclick = (e) => { if (e.target === _mask) close(); };
            document.addEventListener('keydown', _onKey);

            setTimeout(() => {
                const first = _mask.querySelector('.ff-body input[type="text"],.ff-body input[type="password"],.ff-body textarea,.ff-body select');
                if (first) first.focus();
            }, 60);
            _mask.querySelectorAll('.ff-body input[type="text"],.ff-body input[type="password"],.ff-body input[type="number"]').forEach(el => {
                el.onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); _submit(); } };
            });
        }
        return { open, close, setErr };
    })();

    /* ========================================================== */
    /* ====== ★ 参数编辑区（引擎/模板/档位/步骤 共用） ========== */
    /* ==========================================================
       生成一段 HTML，字段 id 前缀由 pfx 决定
       返回的值用 readParamBox(pfx) 取 */
    function paramBoxHtml(p, pfx, opts) {
        opts = opts || {};
        p = p || { useTemp: false, temperature: 0.7, useMax: false, max_tokens: 4096, useTopP: false, top_p: 1, useFreq: false, frequency_penalty: 0 };
        const anyOn = p.useTemp || p.useMax || p.useTopP || p.useFreq;

        let html = '<div class="cfg-sec' + (anyOn || opts.forceOpen ? ' open' : '') + '" id="' + pfx + 'Sec">'
            + '<div class="cfg-sec-hdr" onclick="Admin.toggleParamSec(\'' + pfx + '\')">'
            + '<span class="cs-caret">' + (anyOn || opts.forceOpen ? '▼' : '▶') + '</span>'
            + '<span class="cs-title">⚙️ ' + (opts.title || '高级运行参数') + '</span>'
            + '<span class="cs-sub" id="' + pfx + 'Sub">' + (anyOn ? '已启用部分参数' : '全部未启用（最兼容）') + '</span>'
            + '</div><div class="cfg-sec-body">';

        /* 一键预设 */
        html += '<div style="font-size:11px;color:var(--text2);margin-bottom:6px">一键预设（点了自动填好，不用懂原理）：</div>'
            + '<div class="presets" style="margin-bottom:12px">';
        Object.keys(PARAM_PRESETS).forEach(k => {
            html += '<button type="button" onclick="Admin.applyParamPreset(\'' + pfx + '\',\'' + k + '\')">' + PARAM_PRESETS[k].label + '</button>';
        });
        html += '</div>';

        html += '<div class="pt"><input type="checkbox" id="' + pfx + '_useTemp" ' + (p.useTemp ? 'checked' : '') + ' onchange="Admin.onParamToggle(\'' + pfx + '\')"><label for="' + pfx + '_useTemp">🔥 Temperature<span class="cfg-help" title="随机性。0.2=严谨(翻译/提取/改写) 0.7=平衡 1.0=创意(写小说/起标题)">?</span></label></div>'
            + '<div class="ps" id="' + pfx + '_tempBox" style="' + (p.useTemp ? '' : 'display:none') + '"><input type="number" id="' + pfx + '_temperature" value="' + p.temperature + '" step="0.1" min="0" max="2"></div>'

            + '<div class="pt"><input type="checkbox" id="' + pfx + '_useMax" ' + (p.useMax ? 'checked' : '') + ' onchange="Admin.onParamToggle(\'' + pfx + '\')"><label for="' + pfx + '_useMax">📏 Max Tokens<span class="cfg-help" title="单次回复长度上限。太小长文会被截断；太大部分模型直接报错。不确定就别开">?</span></label></div>'
            + '<div class="ps" id="' + pfx + '_maxBox" style="' + (p.useMax ? '' : 'display:none') + '"><input type="number" id="' + pfx + '_max_tokens" value="' + p.max_tokens + '" min="1">'
            + '<div class="presets">'
            + ['4096:4K', '8192:8K', '16384:16K', '32768:32K', '65536:64K', '131072:128K'].map(x => {
                const a = x.split(':');
                return '<button type="button" onclick="document.getElementById(\'' + pfx + '_max_tokens\').value=' + a[0] + '">' + a[1] + '</button>';
            }).join('')
            + '</div></div>'

            + '<div class="pt"><input type="checkbox" id="' + pfx + '_useTopP" ' + (p.useTopP ? 'checked' : '') + ' onchange="Admin.onParamToggle(\'' + pfx + '\')"><label for="' + pfx + '_useTopP">🎲 Top P<span class="cfg-help" title="与 Temperature 作用类似，一般只开一个">?</span></label></div>'
            + '<div class="ps" id="' + pfx + '_topPBox" style="' + (p.useTopP ? '' : 'display:none') + '"><input type="number" id="' + pfx + '_top_p" value="' + p.top_p + '" step="0.05" min="0" max="1"></div>'

            + '<div class="pt"><input type="checkbox" id="' + pfx + '_useFreq" ' + (p.useFreq ? 'checked' : '') + ' onchange="Admin.onParamToggle(\'' + pfx + '\')"><label for="' + pfx + '_useFreq">🚫 Frequency Penalty<span class="cfg-help" title="抑制重复用词。Claude/Gemini 常不支持，报错请取消">?</span></label></div>'
            + '<div class="ps" id="' + pfx + '_freqBox" style="' + (p.useFreq ? '' : 'display:none') + '"><input type="number" id="' + pfx + '_frequency_penalty" value="' + p.frequency_penalty + '" step="0.1" min="-2" max="2"></div>';

        html += '</div></div>';
        return html;
    }

    function toggleParamSec(pfx) {
        const el = document.getElementById(pfx + 'Sec');
        if (!el) return;
        el.classList.toggle('open');
        const c = el.querySelector('.cs-caret');
        if (c) c.textContent = el.classList.contains('open') ? '▼' : '▶';
    }

    function onParamToggle(pfx) {
        const map = [['useTemp', 'tempBox'], ['useMax', 'maxBox'], ['useTopP', 'topPBox'], ['useFreq', 'freqBox']];
        let anyOn = false;
        map.forEach(([k, box]) => {
            const cb = document.getElementById(pfx + '_' + k);
            const bx = document.getElementById(pfx + '_' + box);
            if (cb && bx) { bx.style.display = cb.checked ? '' : 'none'; if (cb.checked) anyOn = true; }
        });
        const sub = document.getElementById(pfx + 'Sub');
        if (sub) sub.textContent = anyOn ? '已启用部分参数' : '全部未启用（最兼容）';
    }

    function applyParamPreset(pfx, key) {
        const p = PARAM_PRESETS[key];
        if (!p) return;
        const set = (n, v) => { const el = document.getElementById(pfx + '_' + n); if (el) { if (el.type === 'checkbox') el.checked = !!v; else el.value = v; } };
        set('useTemp', p.useTemp); set('temperature', p.temperature);
        set('useMax', p.useMax); set('max_tokens', p.max_tokens);
        set('useTopP', p.useTopP); set('top_p', p.top_p);
        set('useFreq', p.useFreq); set('frequency_penalty', p.frequency_penalty);
        onParamToggle(pfx);
        toast('已套用：' + p.label);
    }

    function readParamBox(pfx) {
        const g = (n) => document.getElementById(pfx + '_' + n);
        if (!g('useTemp')) return null;
        return {
            useTemp: g('useTemp').checked,
            temperature: parseFloat(g('temperature').value) || 0.7,
            useMax: g('useMax').checked,
            max_tokens: parseInt(g('max_tokens').value, 10) || 4096,
            useTopP: g('useTopP').checked,
            top_p: parseFloat(g('top_p').value) || 1,
            useFreq: g('useFreq').checked,
            frequency_penalty: parseFloat(g('frequency_penalty').value) || 0,
        };
    }
    function paramSummary(p) {
        if (!p) return '';
        const a = [];
        if (p.useTemp) a.push('T' + p.temperature);
        if (p.useMax) a.push('M' + (p.max_tokens >= 1024 ? (Math.round(p.max_tokens / 1024) + 'K') : p.max_tokens));
        if (p.useTopP) a.push('P' + p.top_p);
        if (p.useFreq) a.push('F' + p.frequency_penalty);
        return a.join('·');
    }

    /* ========================================================== */
    /* ============ ★ 档位编辑器 SlotEditor ==================== */
    /* ==========================================================
       独立弹层（不用 FF，因为要动态增删行 + 每行折叠参数） */
    const SlotEditor = (function () {
        let _mask = null;
        let _slots = [];
        let _onSave = null;
        let _title = '';
        let _advOpen = {};

        function close() { if (_mask) { _mask.remove(); _mask = null; } _advOpen = {}; }

        function open(cfg) {
            close();
            _slots = JSON.parse(JSON.stringify(cfg.slots || []));
            _onSave = cfg.onSave;
            _title = cfg.title || '⚡ 快捷模型档';
            _mask = document.createElement('div');
            _mask.className = 'ff-mask';
            _mask.innerHTML = '<div class="ff-box" style="max-width:720px">'
                + '<div class="ff-hdr"><h3>' + esc(_title) + '</h3><button class="ff-x" onclick="Admin.slotEditorClose()">×</button></div>'
                + '<div class="ff-body" id="slotEdBody"></div>'
                + '<div class="ff-ft">'
                + '<button class="btn btn-s" onclick="Admin.slotEditorClose()">取消</button>'
                + '<button class="btn btn-p btn-s" onclick="Admin.slotEditorSave()">💾 保存档位</button>'
                + '</div></div>';
            document.body.appendChild(_mask);
            _mask.onclick = (e) => { if (e.target === _mask) close(); };
            draw();
        }

        function _engOpts(cur) {
            let h = '<option value="">（用户当前引擎）</option>';
            /* 按引擎名去重（不同用户同名引擎只列一次，下发时按名匹配没意义，这里按id） */
            const seen = {};
            _engCacheAll.forEach(e => {
                const key = e.username + '|' + e.name;
                if (seen[key]) return;
                seen[key] = 1;
                h += '<option value="' + esc(e.id) + '"' + (cur === e.id ? ' selected' : '') + '>'
                    + esc(e.name) + '（' + esc(e.username) + '）' + (e.engineType === 'image' ? ' 🎨' : '')
                    + '</option>';
            });
            return h;
        }

        function draw() {
            const body = document.getElementById('slotEdBody');
            if (!body) return;

            let html = '<div class="ff-tip">'
                + '每个档位 = 用户输入框上方的一个小按钮，点一下即<b>临时</b>切换模型（不改动引擎配置）。<br>'
                + '· <b>标签</b>：按钮上的字，越短越好（如"快""强""图"）<br>'
                + '· <b>模型名</b>：留空 = 用引擎的默认模型（可做"恢复默认"按钮）<br>'
                + '· <b>引擎</b>：留空 = 在用户当前引擎上换模型；选了 = 点该档自动切到那个引擎<br>'
                + '· <b>参数</b>：勾了才生效，整组覆盖引擎参数'
                + '</div>';

            if (!_slots.length) {
                html += '<div style="padding:20px;text-align:center;color:var(--text2);border:1px dashed var(--border);border-radius:9px;margin-bottom:10px">还没有档位，点下方「➕ 添加档位」</div>';
            }

            _slots.forEach((s, i) => {
                const advOn = !!_advOpen[i];
                html += '<div class="slot-row' + (advOn ? ' adv-open' : '') + '">'
                    + '<div class="slot-row-top">'
                    + '<div class="fg" style="flex:0 0 78px;margin-bottom:0"><label>标签</label>'
                    + '<input value="' + esc(s.label || '') + '" oninput="Admin.slotEdSet(' + i + ',\'label\',this.value)" placeholder="快" maxlength="10"></div>'

                    + '<div class="fg" style="flex:1;min-width:130px;margin-bottom:0"><label>模型名（留空=引擎默认）</label>'
                    + '<input value="' + esc(s.model || '') + '" oninput="Admin.slotEdSet(' + i + ',\'model\',this.value)" placeholder="gpt-4o-mini"></div>'

                    + '<div class="fg" style="flex:1;min-width:150px;margin-bottom:0"><label>绑定引擎（选填）</label>'
                    + '<select onchange="Admin.slotEdSet(' + i + ',\'engineId\',this.value)">' + _engOpts(s.engineId || '') + '</select></div>'

                    + '<div class="fg" style="flex:0 0 46px;margin-bottom:0;text-align:center"><label>生图</label>'
                    + '<input type="checkbox" ' + (s.isImage ? 'checked' : '') + ' onchange="Admin.slotEdSet(' + i + ',\'isImage\',this.checked)" style="width:18px;height:18px;accent-color:var(--pri)"></div>'

                    + '<button class="btn btn-s" onclick="Admin.slotEdToggleAdv(' + i + ')" style="margin-bottom:2px" title="参数覆盖">'
                    + (s.useParams ? '⚙️' + (paramSummary(s.params) || '参数') : '⚙️参数') + '</button>'
                    + '<button class="btn btn-s btn-d" onclick="Admin.slotEdDel(' + i + ')" style="margin-bottom:2px">🗑️</button>'
                    + '</div>';

                /* 参数覆盖区 */
                html += '<div class="slot-row-adv">'
                    + '<div class="pt" style="margin-bottom:8px"><input type="checkbox" id="slotUP' + i + '" ' + (s.useParams ? 'checked' : '') + ' onchange="Admin.slotEdSet(' + i + ',\'useParams\',this.checked)">'
                    + '<label for="slotUP' + i + '">✅ 该档覆盖引擎参数（不勾则继承引擎参数）</label></div>'
                    + paramBoxHtml(s.params, 'slotP' + i, { title: '该档专用参数', forceOpen: true })
                    + '<button class="btn btn-s btn-p" onclick="Admin.slotEdSaveParams(' + i + ')" style="margin-top:8px">✔️ 记录该档参数</button>'
                    + '</div>';

                html += '</div>';
            });

            html += '<button class="btn btn-s" onclick="Admin.slotEdAdd()">➕ 添加档位</button>'
                + '<span style="font-size:11px;color:var(--text2);margin-left:10px">已配 ' + _slots.length + ' 档（最多12）</span>';

            body.innerHTML = html;
        }

        function add() {
            if (_slots.length >= 12) { toast('最多 12 个档位', 'er'); return; }
            _slots.push({ label: '', model: '', engineId: '', isImage: false, useParams: false, params: null });
            draw();
        }
        function del(i) { _slots.splice(i, 1); _advOpen = {}; draw(); }
        function set(i, k, v) { if (_slots[i]) _slots[i][k] = v; if (k === 'useParams') draw(); }
        function toggleAdv(i) { _advOpen[i] = !_advOpen[i]; draw(); }
        function saveParams(i) {
            const p = readParamBox('slotP' + i);
            if (_slots[i]) { _slots[i].params = p; _slots[i].useParams = true; }
            _advOpen[i] = false;
            draw();
            toast('✅ 已记录该档参数：' + (paramSummary(p) || '全部未启用'));
        }
        function save() {
            const clean = _slots.filter(s => (s.label || '').trim()).map(s => ({
                label: (s.label || '').trim().slice(0, 10),
                model: (s.model || '').trim(),
                engineId: (s.engineId || '').trim(),
                isImage: !!s.isImage,
                useParams: !!s.useParams,
                params: s.useParams ? (s.params || null) : null,
            }));
            if (_onSave) _onSave(clean);
        }
        function getSlots() { return _slots; }

        return { open, close, add, del, set, toggleAdv, saveParams, save, draw, getSlots };
    })();

    /* ========================================================== */
    /* ===================== 主框架 ============================= */
    /* ========================================================== */
    function open() {
        if (typeof Auth === 'undefined' || !Auth.isAdmin()) { toast('无管理员权限', 'er'); return; }
        const mo = document.getElementById('mo-admin');
        if (mo) mo.classList.add('show');
        switchTab(_curTab);
    }
    function close() {
        if (_curTab === 'presets' && _presetDirty) {
            if (!confirm('预设有未保存的修改，确定关闭？未保存内容会丢失。')) return;
        }
        FF.close(); SlotEditor.close();
        const mo = document.getElementById('mo-admin');
        if (mo) mo.classList.remove('show');
    }
    function switchTab(tab) {
        if (_curTab === 'presets' && tab !== 'presets' && _presetDirty) {
            if (!confirm('预设有未保存的修改，确定切换？未保存内容会丢失。')) return;
        }
        _curTab = tab;
        document.querySelectorAll('#adminTabs .admin-tab').forEach(b => b.classList.toggle('act', b.dataset.tab === tab));
        const body = document.getElementById('adminBody');
        if (!body) return;
        if (tab === 'users') renderUsers(body);
        else if (tab === 'engines') renderEngines(body);
        else if (tab === 'templates') renderTemplates(body);
        else if (tab === 'slottpl') renderSlotTpl(body);
        else if (tab === 'health') renderHealth(body);
        else if (tab === 'models') renderModels(body);
        else if (tab === 'presets') renderPresets(body);
        else if (tab === 'monitor') renderMonitor(body);
        else if (tab === 'config') renderConfig(body);
    }
    function fmtTime(ts) {
        if (!ts) return '从未';
        const d = Date.now() - ts;
        if (d < 60000) return '刚刚';
        if (d < 3600000) return Math.floor(d / 60000) + '分钟前';
        if (d < 86400000) return Math.floor(d / 3600000) + '小时前';
        return Math.floor(d / 86400000) + '天前';
    }

    /* 载入全部引擎（供档位选引擎用） */
    async function ensureEngines() {
        if (_engCacheAll.length) return;
        try { const d = await apiCall('admin/engines/list'); _engCacheAll = d.engines || []; } catch (e) { _engCacheAll = []; }
    }

    /* ========================================================== */
    /* ===================== 账号管理 =========================== */
    /* ========================================================== */
    async function renderUsers(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>';
        try {
            const data = await apiCall('admin/users/list');
            _usersCache = data.users || [];
            await ensureEngines();
            drawUsersTable(box, '');
        } catch (e) { box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>'; }
    }

    function drawUsersTable(box, kw) {
        kw = (kw || '').toLowerCase();
        let users = _usersCache;
        if (kw) users = users.filter(u => (u.username + ' ' + (u.name || '')).toLowerCase().includes(kw));
        let html = `<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px;align-items:center">
            <button class="btn btn-p btn-s" onclick="Admin.showCreateUser()">➕ 新增账号</button>
            <button class="btn btn-s" onclick="Admin.switchTab('slottpl')">⚡ 档位模板库（批量下发）</button>
            <label class="btn btn-s" style="cursor:pointer">📥 xlsx导入<input type="file" accept=".xlsx,.xls" onchange="Admin.importXLSX(this)" style="display:none"></label>
            <button class="btn btn-s" onclick="Admin.exportXLSX(false)">📤 导出(脱敏)</button>
            <button class="btn btn-s btn-d" onclick="Admin.exportXLSX(true)">📤 导出(含Key)</button>
            <button class="btn btn-s" onclick="Admin.switchTab('users')">🔄 刷新</button>
            <input type="text" placeholder="🔍 搜索姓名/账号" oninput="Admin.searchUsers(this.value)" style="margin-left:auto;padding:5px 10px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);font-size:12px;width:170px"></div>
            <div style="font-size:12px;color:var(--text2);margin-bottom:8px">共 ${_usersCache.length} 个账号${kw ? '，匹配 ' + users.length : ''}。⚡=快捷模型档，🔴=最近7天≥3个IP。</div>
            <div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>姓名</th><th>账号</th><th>角色</th><th>状态</th><th>引擎</th><th>⚡档位</th><th>可用分组</th><th>最后活跃</th><th>IP</th><th>操作</th></tr></thead><tbody>`;
        users.forEach(u => {
            let permTxt = '全部';
            try { const p = JSON.parse(u.permissions || '{}'); if (p.allowGroups && p.allowGroups.length) permTxt = p.allowGroups.join('/'); } catch (e) {}
            const sc = u.slotCount || 0;
            html += `<tr><td>${esc(u.name || '-')}</td><td>${esc(u.username)}</td><td>${u.role === 'admin' ? '👑' : '普通'}</td><td>${u.status === 'active' ? '<span style="color:#10b981">启用</span>' : '<span style="color:#ef4444">禁用</span>'}</td><td>${u.engineCount}</td><td>${sc ? '<span style="color:var(--pri);font-weight:600">' + sc + ' 档</span>' : '<span style="color:var(--text2)">用全局</span>'}</td><td style="font-size:11px;max-width:110px;overflow:hidden;text-overflow:ellipsis" title="${esc(permTxt)}">${esc(permTxt)}</td><td style="font-size:11px">${fmtTime(u.lastActive)}</td><td>${u.ipAbnormal ? '<span style="color:#ef4444;font-weight:bold">🔴' + u.ipCount + '</span>' : (u.ipCount || 0)}</td><td class="admin-ops"><button onclick='Admin.showSlots(${JSON.stringify(u.username)})' title="快捷模型档">⚡</button><button onclick='Admin.showEditUser(${JSON.stringify(u.username)})' title="编辑资料">✏️</button><button onclick='Admin.showPerm(${JSON.stringify(u.username)})' title="工作流权限">🎫</button><button onclick='Admin.showResetPwd(${JSON.stringify(u.username)})' title="改密">🔑</button><button onclick='Admin.toggleStatus(${JSON.stringify(u.username)},${JSON.stringify(u.status)})' title="启用/禁用">${u.status === 'active' ? '🚫' : '✅'}</button>${u.username !== 'admin' ? `<button onclick='Admin.delUser(${JSON.stringify(u.username)})' style="color:#ef4444" title="删除">🗑️</button>` : ''}</td></tr>`;
        });
        html += '</tbody></table></div>';
        box.innerHTML = html;
    }

    function searchUsers(kw) { drawUsersTable(document.getElementById('adminBody'), kw); }

    /* ★ 用户专属档位 */
    async function showSlots(username) {
        await ensureEngines();
        const u = _usersCache.find(x => x.username === username);
        const cur = _pj(u && u.quickModels, []);
        SlotEditor.open({
            title: '⚡ 快捷模型档 — ' + username + (cur.length ? '' : '（当前使用全局默认档）'),
            slots: cur,
            onSave: async (slots) => {
                try {
                    await apiCall('admin/users/slots', 'POST', { username, slots });
                    toast(slots.length ? ('✅ 已保存 ' + slots.length + ' 个档位（该用户刷新后生效）') : '✅ 已清空，将回落全局默认档');
                    SlotEditor.close();
                    switchTab('users');
                } catch (e) { toast('保存失败：' + e.message, 'er'); }
            }
        });
    }

    function showCreateUser() {
        FF.open({
            title: '➕ 新增账号',
            tip: '账号创建后不可改名；密码可随时重置。',
            okText: '创建',
            fields: [
                { key: 'name', label: '姓名', type: 'text', placeholder: '如：张三', hint: '用于存档命名与分享署名' },
                { key: 'username', label: '登录账号', type: 'text', placeholder: '如：zhangsan', hint: '英文/数字，创建后不可修改' },
                { key: 'password', label: '初始密码', type: 'text', placeholder: '如：pass123', hint: '明文填写，方便告知本人' },
                { key: 'role', label: '角色', type: 'select', value: 'user', options: [{ v: 'user', t: '普通用户' }, { v: 'admin', t: '👑 管理员（可进后台）' }] },
            ],
            onSubmit: async (v, setErr) => {
                if (!v.username.trim()) { setErr('登录账号不能为空'); return false; }
                if (!v.password.trim()) { setErr('初始密码不能为空'); return false; }
                await apiCall('admin/users/create', 'POST', { username: v.username.trim(), password: v.password.trim(), name: (v.name || '').trim(), role: v.role });
                toast('✅ 已创建账号 ' + v.username.trim());
                switchTab('users');
            }
        });
    }

    function showEditUser(username) {
        const u = _usersCache.find(x => x.username === username);
        if (!u) { toast('账号不存在', 'er'); return; }
        FF.open({
            title: '✏️ 编辑账号 — ' + username,
            okText: '保存',
            fields: [
                { key: 'name', label: '姓名', type: 'text', value: u.name || '' },
                { key: 'role', label: '角色', type: 'select', value: u.role, options: [{ v: 'user', t: '普通用户' }, { v: 'admin', t: '👑 管理员' }] },
                { key: 'status', label: '状态', type: 'select', value: u.status, options: [{ v: 'active', t: '启用' }, { v: 'disabled', t: '禁用（无法登录）' }] },
            ],
            onSubmit: async (v) => {
                await apiCall('admin/users/update', 'POST', { username, name: v.name, role: v.role, status: v.status });
                toast('✅ 已保存'); switchTab('users');
            }
        });
    }

    function showResetPwd(username) {
        FF.open({
            title: '🔑 重置密码 — ' + username,
            tip: '重置后该账号所有登录会话会失效，需重新登录。',
            okText: '重置',
            fields: [
                { key: 'p1', label: '新密码', type: 'text', placeholder: '明文填写', hint: '方便直接告知本人' },
                { key: 'p2', label: '再输一次', type: 'text', placeholder: '防止手误' },
            ],
            onSubmit: async (v, setErr) => {
                if (!v.p1.trim()) { setErr('密码不能为空'); return false; }
                if (v.p1 !== v.p2) { setErr('两次输入不一致'); return false; }
                await apiCall('admin/users/resetpwd', 'POST', { username, password: v.p1.trim() });
                toast('✅ 已重置密码');
            }
        });
    }

    function showPerm(username) {
        const u = _usersCache.find(x => x.username === username);
        let perm = {};
        try { perm = JSON.parse((u && u.permissions) || '{}'); } catch (e) {}
        const cur = perm.allowGroups || [];
        const groups = (typeof Workflow !== 'undefined' && Workflow.isLoaded()) ? Workflow.getGroups() : [];
        FF.open({
            title: '🎫 工作流权限 — ' + username,
            tip: '<b>全不勾 = 允许全部分组</b>（默认）。勾选后只能用勾中的分组。修改后用户端<b>实时生效</b>。',
            okText: '保存权限',
            fields: [{
                key: 'groups', label: '可用工作流分组', type: 'checks', value: cur,
                options: groups.map(g => ({ v: g, t: g })),
                hint: groups.length ? '' : '⚠️ 未读到分组，请先到「📋 预设 → 🏷️ 分组管理」创建'
            }],
            onSubmit: async (v) => {
                await apiCall('admin/users/perm', 'POST', { username, permissions: Object.assign({}, perm, { allowGroups: v.groups }) });
                toast(v.groups.length ? ('✅ 已限定 ' + v.groups.length + ' 个分组') : '✅ 已开放全部分组');
                switchTab('users');
            }
        });
    }

    function toggleStatus(username, cur) {
        const next = cur === 'active' ? 'disabled' : 'active';
        if (next === 'disabled' && !confirm('禁用【' + username + '】？该账号将无法登录。')) return;
        apiCall('admin/users/update', 'POST', { username, status: next })
            .then(() => { toast(next === 'active' ? '已启用' : '已禁用'); switchTab('users'); })
            .catch(e => toast('失败：' + e.message, 'er'));
    }
    function delUser(username) {
        if (!confirm('⚠️ 删除账号【' + username + '】？\n\n其公有引擎、登录会话会一并删除，不可恢复。')) return;
        apiCall('admin/users/delete', 'POST', { username })
            .then(() => { toast('✅ 已删除'); switchTab('users'); })
            .catch(e => toast('失败：' + e.message, 'er'));
    }

    function importXLSX(inputEl) { const file = inputEl.files && inputEl.files[0]; if (!file) return; loadXLSX().then(() => { const reader = new FileReader(); reader.onload = async (e) => { try { const wb = XLSX.read(e.target.result, { type: 'array' }); const sheet = wb.Sheets[wb.SheetNames[0]]; const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' }); if (!rows.length) { toast('表格无数据', 'er'); inputEl.value = ''; return; } if (!confirm('导入 ' + rows.length + ' 行（同账号覆盖+重配引擎）？')) { inputEl.value = ''; return; } toast('导入中...'); const res = await apiCall('admin/users/import', 'POST', { rows }); let msg = '✅ 账号 ' + res.userCount + '，引擎 ' + res.engCount; if (res.errors && res.errors.length) msg += '\n⚠️ ' + res.errors.join('；'); alert(msg); _engCacheAll = []; switchTab('users'); } catch (err) { toast('导入失败：' + err.message, 'er'); } inputEl.value = ''; }; reader.readAsArrayBuffer(file); }).catch(e => toast('加载解析库失败：' + e.message, 'er')); }
    async function exportXLSX(withKey) { if (withKey && !confirm('⚠️ 导出含明文Key，妥善保管！继续？')) return; try { await loadXLSX(); const res = await apiCall('admin/users/export?withkey=' + (withKey ? '1' : '0')); const ws = XLSX.utils.json_to_sheet(res.rows || [], { header: ['姓名', '账号', '密码', '角色', '引擎名称', '协议', 'BaseURL', 'APIKey', '模型', '输入单价', '输出单价', '缓存读单价', '缓存写单价'] }); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '账号'); XLSX.writeFile(wb, 'feifan-accounts-' + (withKey ? 'withkey-' : '') + new Date().toISOString().slice(0, 10) + '.xlsx'); toast('✅ 已导出'); } catch (e) { toast('导出失败：' + e.message, 'er'); } }

    /* ========================================================== */
    /* ===================== 引擎管理 =========================== */
    /* ========================================================== */
    async function renderEngines(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>';
        try {
            const usersData = await apiCall('admin/users/list');
            const users = usersData.users || [];
            _usersCache = users;
            const engData = await apiCall('admin/engines/list');
            const engs = engData.engines || [];
            _engCacheAll = engs;

            const byUser = {};
            engs.forEach(e => { if (!byUser[e.username]) byUser[e.username] = []; byUser[e.username].push(e); });

            let html = `<div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap;align-items:center">
                <button class="btn btn-p btn-s" onclick="Admin.switchTab('templates')">📦 引擎模板库</button>
                <button class="btn btn-s" onclick="Admin.switchTab('slottpl')">⚡ 档位模板库</button>
                <button class="btn btn-s" onclick="Admin.switchTab('health')">🔍 引擎体检</button>
                <button class="btn btn-s" onclick="Admin.switchTab('engines')">🔄 刷新</button>
                <span style="font-size:11px;color:var(--text2);margin-left:auto">🎨生图 · 💰缓存 · 🔗来自模板 · ⚙️含参数</span>
            </div>
            <div style="font-size:12px;color:var(--text2);margin-bottom:10px;line-height:1.7">
                💡 <b>推荐做法</b>：每人配 <b>1 个引擎</b>（通道），再用「⚡ 档位模板库」给他配多个模型档 —— 用户点按钮就能换模型，你不用配一堆引擎。<br>
                跨站/跨协议（如 Claude 原生）才需要配多个引擎。
            </div>`;

            users.forEach(u => {
                const ue = byUser[u.username] || [];
                html += `<div class="eng-user-block"><div class="eng-user-hdr"><b>${esc(u.name || u.username)}</b> <span style="color:var(--text2);font-size:11px">(${esc(u.username)})</span><span style="font-size:11px;color:var(--text2);margin-left:6px">${ue.length} 引擎 · ${u.slotCount || 0} 档位</span><button class="btn btn-s" onclick='Admin.showSlots(${JSON.stringify(u.username)})' style="margin-left:auto">⚡档位</button><button class="btn btn-p btn-s" onclick='Admin.showEngEdit(${JSON.stringify(u.username)},"")'>➕ 加引擎</button></div>`;
                if (!ue.length) html += '<div style="font-size:11px;color:var(--text2);padding:4px 0">（无引擎）</div>';
                ue.forEach(e => {
                    const imgTag = (e.engineType === 'image') ? '<span style="color:#f59e0b">🎨</span>' : '';
                    const cacheTag = e.useCache ? '<span style="color:#10b981">💰</span>' : '';
                    const tplTag = e.tplId ? '<span style="color:var(--pri)" title="来自模板，可一键同步">🔗</span>' : '';
                    const pSum = paramSummary(e.params);
                    const pTag = pSum ? '<span style="color:#8b5cf6;font-size:10px" title="引擎参数">⚙️' + esc(pSum) + '</span>' : '';
                    html += `<div class="eng-item"><span>📦 <b>${esc(e.name)}</b> ${imgTag}${cacheTag}${tplTag}${pTag} <span style="color:var(--text2);font-size:11px">${esc(e.protocol)} / ${esc(e.model || '用户自选')}</span></span>
                        <div style="margin-left:auto;display:flex;gap:4px"><button class="btn btn-s" onclick='Admin.showEngEdit(${JSON.stringify(e.username)},${JSON.stringify(e.id)})'>✏️改</button><button class="btn btn-s btn-d" onclick='Admin.delEng(${JSON.stringify(e.id)},${JSON.stringify(e.name)})'>🗑️</button></div></div>`;
                });
                html += `</div>`;
            });

            box.innerHTML = html;
            box._engs = engs;
        } catch (e) {
            box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>';
        }
    }

    /* ★ 引擎编辑（含参数） */
    function showEngEdit(username, engId) {
        const box = document.getElementById('adminBody');
        const engs = (box && box._engs) || _engCacheAll;
        const e = engId ? engs.find(x => x.id === engId) : null;

        FF.open({
            title: (e ? '✏️ 编辑引擎' : '➕ 新增引擎') + ' — ' + username,
            wide: true,
            tip: e && e.tplId
                ? '⚠️ 此引擎来自<b>模板</b>。手动改动后会<b>脱离模板管理</b>（模板同步不再覆盖它）。'
                : '💰 缓存：Claude 等模型开启后重复内容省钱。<br>🎨 生图：选后用户发消息即出图。<br>⚙️ 参数会随引擎下发给用户（用户可微调，也能一键恢复）。',
            okText: '保存',
            fields: [
                { key: 'name', label: '引擎名称', type: 'text', value: e ? e.name : '', placeholder: '如：主引擎', hint: '用户端看到的名字' },
                { key: 'engineType', label: '引擎类型', type: 'select', value: e ? (e.engineType || 'chat') : 'chat', options: [{ v: 'chat', t: '💬 对话' }, { v: 'image', t: '🎨 生图 / 改图' }] },
                { key: 'protocol', label: '协议', type: 'select', value: e ? e.protocol : 'openai', options: [{ v: 'openai', t: 'OpenAI / 通用' }, { v: 'anthropic', t: 'Claude 原生' }, { v: 'gemini', t: 'Gemini 原生' }] },
                { key: 'base', label: 'Base URL', type: 'text', value: e ? e.base : 'https://api.openai-proxy.org/v1' },
                { key: 'key', label: 'API Key', type: 'password', value: '', placeholder: e ? '••••（留空=不修改）' : 'sk-...', hint: e ? '留空表示不修改现有 Key' : '' },
                { key: 'model', label: '默认模型', type: 'text', value: e ? (e.model || '') : '', placeholder: '对话如 gpt-4o；生图如 gpt-image-1', hint: '留空则由用户在前台自己获取选择' },
                { key: 'useCache', label: '💰 开启 Prompt 缓存', type: 'check', value: e ? !!e.useCache : false, hint: '生图引擎无需开启' },
                { type: 'html', html: paramBoxHtml(e ? e.params : null, 'engP', { title: '引擎默认参数（会下发给用户）' }) },
                { key: 'priceIn', label: '输入单价（$/1M）', type: 'number', value: e ? (e.priceIn || 0) : 0, step: '0.01', hint: '留 0 则查「💵 模型库」' },
                { key: 'priceOut', label: '输出单价（$/1M）', type: 'number', value: e ? (e.priceOut || 0) : 0, step: '0.01' },
                { key: 'priceCR', label: '缓存读单价', type: 'number', value: e ? (e.priceCR || 0) : 0, step: '0.01' },
                { key: 'priceCW', label: '缓存写单价', type: 'number', value: e ? (e.priceCW || 0) : 0, step: '0.01' },
            ],
            onSubmit: async (v, setErr) => {
                if (!v.name.trim()) { setErr('引擎名称必填'); return false; }
                await apiCall('admin/engines/save', 'POST', {
                    username, id: engId || undefined,
                    name: v.name.trim(), engineType: v.engineType, protocol: v.protocol,
                    base: v.base.trim(), key: v.key.trim(), model: v.model.trim(),
                    useCache: v.useCache,
                    params: readParamBox('engP'),
                    priceIn: v.priceIn, priceOut: v.priceOut, priceCR: v.priceCR, priceCW: v.priceCW,
                    tplId: ''
                });
                toast('✅ 已保存');
                _engCacheAll = [];
                switchTab('engines');
            }
        });
    }

    function delEng(id, name) {
        if (!confirm('删除引擎【' + (name || '') + '】？\n该用户将无法再使用它。\n⚠️ 若有档位绑定了它，那些档位会失效。')) return;
        apiCall('admin/engines/delete', 'POST', { id })
            .then(() => { toast('✅ 已删除'); _engCacheAll = []; switchTab('engines'); })
            .catch(e => toast('失败：' + e.message, 'er'));
    }

    /* ========================================================== */
    /* ============ 引擎模板库（含参数） ======================== */
    /* ========================================================== */
    async function renderTemplates(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>';
        try {
            const tData = await apiCall('admin/templates/list');
            _tplCache = tData.templates || [];
            if (!_usersCache.length) { const uData = await apiCall('admin/users/list'); _usersCache = uData.users || []; }
            drawTemplates(box);
        } catch (e) {
            box.innerHTML = '<div style="color:#ef4444;padding:20px;line-height:1.8">加载失败：' + esc(e.message)
                + '<br><br>如提示"模板表不存在"，请先到 Cloudflare → D1 执行建表 SQL。</div>';
        }
    }

    function drawTemplates(box) {
        const selCount = Object.keys(_tplSel).filter(k => _tplSel[k]).length;
        let html = `<div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap;align-items:center">
            <button class="btn btn-p btn-s" onclick="Admin.showTplEdit('')">➕ 新建引擎模板</button>
            <button class="btn btn-s" onclick="Admin.switchTab('templates')">🔄 刷新</button>
            <button class="btn btn-s" onclick="Admin.switchTab('engines')">← 返回引擎列表</button>
        </div>
        <div style="font-size:12px;color:var(--text2);margin-bottom:12px;line-height:1.7">
            💡 <b>用法</b>：建好模板（含 Key + 参数）→ 勾模板 → 勾账号 → 一键下发。<br>
            🔄 换 Key / 换中转站 / 改参数后，点「同步」即可刷新所有已下发引擎。
        </div>`;

        if (!_tplCache.length) {
            html += '<div style="padding:30px;text-align:center;color:var(--text2);border:1px dashed var(--border);border-radius:10px">还没有模板，点上方「➕ 新建引擎模板」</div>';
            box.innerHTML = html; return;
        }

        html += '<div class="tpl-grid">';
        _tplCache.forEach(t => {
            const on = !!_tplSel[t.id];
            const pSum = paramSummary(t.params);
            html += `<div class="tpl-card${on ? ' sel' : ''}">
                <div class="tpl-card-top">
                    <input type="checkbox" ${on ? 'checked' : ''} onchange='Admin.toggleTplSel(${JSON.stringify(t.id)},this.checked)'>
                    <span class="tpl-name">${esc(t.name)}</span>
                </div>
                <div class="tpl-meta">${esc(t.protocol)} · ${esc(t.model || '用户自选')}<br>${esc((t.base || '').replace(/^https?:\/\//, '').slice(0, 34))}</div>
                <div class="tpl-badges">
                    ${t.engineType === 'image' ? '<span class="tpl-badge img">🎨生图</span>' : '<span class="tpl-badge">💬对话</span>'}
                    ${t.useCache ? '<span class="tpl-badge cache">💰缓存</span>' : ''}
                    ${pSum ? '<span class="tpl-badge param">⚙️' + esc(pSum) + '</span>' : ''}
                    ${t.hasKey ? '<span class="tpl-badge">🔑Key</span>' : '<span class="tpl-badge nokey">⚠️无Key</span>'}
                    ${t.deployed ? '<span class="tpl-badge deployed">已下发 ' + t.deployed + '</span>' : ''}
                </div>
                <div class="tpl-acts">
                    <button class="btn btn-s" onclick='Admin.showTplEdit(${JSON.stringify(t.id)})'>✏️改</button>
                    ${t.deployed ? `<button class="btn btn-s" onclick='Admin.syncTpl(${JSON.stringify(t.id)},${JSON.stringify(t.name)},${t.deployed})'>🔄同步</button>` : ''}
                    <button class="btn btn-s btn-d" onclick='Admin.delTpl(${JSON.stringify(t.id)},${JSON.stringify(t.name)},${t.deployed})'>🗑️</button>
                </div>
            </div>`;
        });
        html += '</div>';

        html += `<div class="deploy-panel">
            <div style="font-size:13px;font-weight:600;margin-bottom:4px">🚀 批量下发引擎到账号</div>
            <div style="font-size:11px;color:var(--text2)">已选 <b style="color:var(--pri)">${selCount}</b> 个模板。下面勾选账号：</div>
            <div class="deploy-users" id="deployUsers"></div>
            <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
                <button class="btn btn-s" onclick="Admin.deploySelectAllUsers(true)">全选账号</button>
                <button class="btn btn-s" onclick="Admin.deploySelectAllUsers(false)">全不选</button>
                <select id="deployMode" style="padding:5px 8px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);font-size:12px">
                    <option value="skip">同名引擎：跳过（安全，推荐）</option>
                    <option value="replace">同名引擎：覆盖（更新Key/参数用）</option>
                </select>
                <button class="btn btn-p btn-s" onclick="Admin.doDeploy()" style="margin-left:auto">🚀 开始下发</button>
            </div>
        </div>`;

        box.innerHTML = html;
        drawDeployUsers();
    }

    function drawDeployUsers() {
        const wrap = document.getElementById('deployUsers');
        if (!wrap) return;
        if (!_usersCache.length) { wrap.innerHTML = '<span style="font-size:12px;color:var(--text2)">（无账号）</span>'; return; }
        let html = '';
        _usersCache.forEach(u => {
            const on = !!_deployUserSel[u.username];
            html += `<label class="ff-chk${on ? ' on' : ''}"><input type="checkbox" ${on ? 'checked' : ''} onchange='Admin.toggleDeployUser(${JSON.stringify(u.username)},this.checked)'>${esc(u.name || u.username)} <span style="opacity:.6;font-size:10px">${esc(u.username)}</span></label>`;
        });
        wrap.innerHTML = html;
    }

    function toggleTplSel(id, on) { _tplSel[id] = on; drawTemplates(document.getElementById('adminBody')); }
    function toggleDeployUser(un, on) { _deployUserSel[un] = on; drawDeployUsers(); }
    function deploySelectAllUsers(val) { _usersCache.forEach(u => { _deployUserSel[u.username] = val; }); drawDeployUsers(); }

    function showTplEdit(tplId) {
        const t = tplId ? _tplCache.find(x => x.id === tplId) : null;
        FF.open({
            title: t ? ('✏️ 编辑模板 — ' + t.name) : '➕ 新建引擎模板',
            wide: true,
            tip: '模板是"引擎的配方"（含 Key + 参数），可批量下发给多个账号。<br>模板名会成为用户端看到的引擎名。',
            okText: '保存模板',
            fields: [
                { key: 'name', label: '模板 / 引擎名称', type: 'text', value: t ? t.name : '', placeholder: '如：主引擎', hint: '下发后用户看到的就是这个名字' },
                { key: 'engineType', label: '引擎类型', type: 'select', value: t ? (t.engineType || 'chat') : 'chat', options: [{ v: 'chat', t: '💬 对话' }, { v: 'image', t: '🎨 生图 / 改图' }] },
                { key: 'protocol', label: '协议', type: 'select', value: t ? t.protocol : 'openai', options: [{ v: 'openai', t: 'OpenAI / 通用' }, { v: 'anthropic', t: 'Claude 原生' }, { v: 'gemini', t: 'Gemini 原生' }] },
                { key: 'base', label: 'Base URL', type: 'text', value: t ? t.base : 'https://api.openai-proxy.org/v1' },
                { key: 'key', label: 'API Key', type: 'password', value: '', placeholder: t ? (t.hasKey ? '••••（留空=不修改）' : 'sk-...') : 'sk-...', hint: t && t.hasKey ? '留空表示不修改现有 Key' : '' },
                { key: 'model', label: '默认模型', type: 'text', value: t ? (t.model || '') : '', placeholder: '如 gpt-4o', hint: '留空则用户自己选' },
                { key: 'useCache', label: '💰 开启 Prompt 缓存', type: 'check', value: t ? !!t.useCache : false },
                { type: 'html', html: paramBoxHtml(t ? t.params : null, 'tplP', { title: '默认参数（随模板下发）' }) },
                { key: 'priceIn', label: '输入单价（$/1M）', type: 'number', value: t ? (t.priceIn || 0) : 0, step: '0.01', hint: '留 0 则查「💵 模型库」' },
                { key: 'priceOut', label: '输出单价（$/1M）', type: 'number', value: t ? (t.priceOut || 0) : 0, step: '0.01' },
                { key: 'priceCR', label: '缓存读单价', type: 'number', value: t ? (t.priceCR || 0) : 0, step: '0.01' },
                { key: 'priceCW', label: '缓存写单价', type: 'number', value: t ? (t.priceCW || 0) : 0, step: '0.01' },
            ],
            onSubmit: async (v, setErr) => {
                if (!v.name.trim()) { setErr('模板名必填'); return false; }
                await apiCall('admin/templates/save', 'POST', {
                    id: tplId || undefined,
                    name: v.name.trim(), engineType: v.engineType, protocol: v.protocol,
                    base: v.base.trim(), key: v.key.trim(), model: v.model.trim(),
                    useCache: v.useCache,
                    params: readParamBox('tplP'),
                    priceIn: v.priceIn, priceOut: v.priceOut, priceCR: v.priceCR, priceCW: v.priceCW
                });
                toast('✅ 模板已保存');
                switchTab('templates');
            }
        });
    }

    function delTpl(id, name, deployed) {
        let msg = '删除模板【' + name + '】？';
        if (deployed) msg += '\n\n已下发的 ' + deployed + ' 个引擎【不会被删除】，只解除关联。';
        if (!confirm(msg)) return;
        apiCall('admin/templates/delete', 'POST', { id })
            .then(() => { delete _tplSel[id]; toast('✅ 已删除模板'); switchTab('templates'); })
            .catch(e => toast('失败：' + e.message, 'er'));
    }

    function syncTpl(id, name, deployed) {
        FF.open({
            title: '🔄 同步模板 — ' + name,
            tip: '将把此模板的 <b>Key / BaseURL / 模型 / 协议 / 缓存 / 参数 / 单价</b> 推给所有由它下发的引擎（共 <b>' + deployed + '</b> 个）。',
            okText: '确认同步',
            fields: [{ key: 'syncName', label: '同时覆盖引擎名称', type: 'check', value: false, hint: '不勾=保留用户端现有引擎名（推荐）' }],
            onSubmit: async (v) => {
                const r = await apiCall('admin/templates/sync', 'POST', { templateId: id, syncName: v.syncName });
                toast('✅ 已同步 ' + (r.count || 0) + ' 个引擎');
                _engCacheAll = [];
                switchTab('templates');
            }
        });
    }

    async function doDeploy() {
        const tplIds = Object.keys(_tplSel).filter(k => _tplSel[k]);
        const usernames = Object.keys(_deployUserSel).filter(k => _deployUserSel[k]);
        if (!tplIds.length) { toast('请先勾选至少 1 个模板', 'er'); return; }
        if (!usernames.length) { toast('请先勾选至少 1 个账号', 'er'); return; }
        const modeEl = document.getElementById('deployMode');
        const mode = modeEl ? modeEl.value : 'skip';
        if (!confirm('确认下发？\n\n模板 ' + tplIds.length + ' × 账号 ' + usernames.length + '\n同名策略：' + (mode === 'replace' ? '覆盖' : '跳过'))) return;
        toast('下发中...');
        try {
            const r = await apiCall('admin/templates/deploy', 'POST', { templateIds: tplIds, usernames, mode });
            let msg = '✅ 下发完成\n\n新建：' + (r.created || 0) + '\n覆盖：' + (r.updated || 0) + '\n跳过：' + (r.skipped || 0);
            if (r.errors && r.errors.length) msg += '\n\n⚠️ 失败 ' + r.errors.length + ' 条：\n' + r.errors.slice(0, 6).join('\n');
            alert(msg);
            _deployUserSel = {}; _engCacheAll = [];
            switchTab('templates');
        } catch (e) { toast('下发失败：' + e.message, 'er'); }
    }

    /* ========================================================== */
    /* ============ ★ 档位模板库 =============================== */
    /* ========================================================== */
    async function renderSlotTpl(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>';
        try {
            const d = await apiCall('admin/slottpl/list');
            _slotTplCache = d.templates || [];
            if (!_usersCache.length) { const u = await apiCall('admin/users/list'); _usersCache = u.users || []; }
            await ensureEngines();
            drawSlotTpl(box);
        } catch (e) {
            box.innerHTML = '<div style="color:#ef4444;padding:20px;line-height:1.8">加载失败：' + esc(e.message)
                + '<br><br>如提示"档位模板表不存在"，请先执行建表 SQL（slot_templates）。</div>';
        }
    }

    function drawSlotTpl(box) {
        let html = `<div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap;align-items:center">
            <button class="btn btn-p btn-s" onclick="Admin.newSlotTpl()">➕ 新建档位模板</button>
            <button class="btn btn-s" onclick="Admin.switchTab('slottpl')">🔄 刷新</button>
            <button class="btn btn-s" onclick="Admin.switchTab('users')">← 返回账号列表</button>
        </div>
        <div style="font-size:12px;color:var(--text2);margin-bottom:12px;line-height:1.7">
            💡 <b>核心玩法</b>：给每人配 <b>1 个引擎</b> + <b>一套档位</b>，用户点按钮就能换模型/换参数，不用配一堆引擎。<br>
            档位可以：换模型 · 绑定其他引擎（跨站/跨协议）· 切生图 · 覆盖参数。
        </div>`;

        if (!_slotTplCache.length) {
            html += '<div style="padding:30px;text-align:center;color:var(--text2);border:1px dashed var(--border);border-radius:10px">还没有档位模板，点上方「➕ 新建档位模板」</div>';
            box.innerHTML = html; return;
        }

        html += '<div class="tpl-grid">';
        _slotTplCache.forEach(t => {
            const on = (_slotTplSel === t.id);
            const labels = (t.slots || []).map(s => s.label + (s.isImage ? '🎨' : '') + (s.useParams ? '⚙️' : '')).join(' / ');
            html += `<div class="tpl-card${on ? ' sel' : ''}">
                <div class="tpl-card-top">
                    <input type="radio" name="slotTplPick" ${on ? 'checked' : ''} onchange='Admin.pickSlotTpl(${JSON.stringify(t.id)})' style="width:15px;height:15px;accent-color:var(--pri);margin-top:2px">
                    <span class="tpl-name">${esc(t.name)}</span>
                </div>
                <div class="tpl-meta">${t.slotCount} 个档位<br>${esc(labels.slice(0, 60))}</div>
                <div class="tpl-acts">
                    <button class="btn btn-s" onclick='Admin.editSlotTpl(${JSON.stringify(t.id)})'>✏️ 编辑档位</button>
                    <button class="btn btn-s" onclick='Admin.renameSlotTpl(${JSON.stringify(t.id)})'>📝 改名</button>
                    <button class="btn btn-s btn-d" onclick='Admin.delSlotTpl(${JSON.stringify(t.id)},${JSON.stringify(t.name)})'>🗑️</button>
                </div>
            </div>`;
        });
        html += '</div>';

        const picked = _slotTplCache.find(x => x.id === _slotTplSel);
        html += `<div class="deploy-panel">
            <div style="font-size:13px;font-weight:600;margin-bottom:4px">🚀 下发档位到账号</div>
            <div style="font-size:11px;color:var(--text2)">当前选中模板：<b style="color:var(--pri)">${picked ? esc(picked.name) + '（' + picked.slotCount + '档）' : '（请在上方选一个）'}</b></div>
            <div class="deploy-users" id="slotDeployUsers"></div>
            <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
                <button class="btn btn-s" onclick="Admin.slotDeployAll(true)">全选账号</button>
                <button class="btn btn-s" onclick="Admin.slotDeployAll(false)">全不选</button>
                <select id="slotDeployMode" style="padding:5px 8px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);font-size:12px">
                    <option value="replace">整套替换（推荐，统一配置）</option>
                    <option value="append">追加（保留原有，同名跳过）</option>
                </select>
                <button class="btn btn-p btn-s" onclick="Admin.doSlotDeploy()" style="margin-left:auto">🚀 开始下发</button>
            </div>
        </div>`;

        box.innerHTML = html;
        drawSlotDeployUsers();
    }

    function drawSlotDeployUsers() {
        const wrap = document.getElementById('slotDeployUsers');
        if (!wrap) return;
        if (!_usersCache.length) { wrap.innerHTML = '<span style="font-size:12px;color:var(--text2)">（无账号）</span>'; return; }
        let html = '';
        _usersCache.forEach(u => {
            const on = !!_slotDeployUserSel[u.username];
            html += `<label class="ff-chk${on ? ' on' : ''}"><input type="checkbox" ${on ? 'checked' : ''} onchange='Admin.toggleSlotDeployUser(${JSON.stringify(u.username)},this.checked)'>${esc(u.name || u.username)} <span style="opacity:.6;font-size:10px">${(u.slotCount || 0) ? (u.slotCount + '档') : '无'}</span></label>`;
        });
        wrap.innerHTML = html;
    }
    function pickSlotTpl(id) { _slotTplSel = id; drawSlotTpl(document.getElementById('adminBody')); }
    function toggleSlotDeployUser(un, on) { _slotDeployUserSel[un] = on; drawSlotDeployUsers(); }
    function slotDeployAll(v) { _usersCache.forEach(u => { _slotDeployUserSel[u.username] = v; }); drawSlotDeployUsers(); }

    function newSlotTpl() {
        FF.open({
            title: '➕ 新建档位模板',
            tip: '先起个名字，保存后再点「✏️ 编辑档位」配具体档位。',
            okText: '创建',
            fields: [{ key: 'name', label: '模板名称', type: 'text', value: '', placeholder: '如：标准三档', hint: '仅后台可见，方便你识别' }],
            onSubmit: async (v, setErr) => {
                if (!v.name.trim()) { setErr('名称不能为空'); return false; }
                const r = await apiCall('admin/slottpl/save', 'POST', { name: v.name.trim(), slots: [] });
                toast('✅ 已创建，请点「✏️ 编辑档位」配置');
                _slotTplSel = r.id || '';
                switchTab('slottpl');
            }
        });
    }

    async function editSlotTpl(id) {
        await ensureEngines();
        const t = _slotTplCache.find(x => x.id === id);
        if (!t) { toast('模板不存在', 'er'); return; }
        SlotEditor.open({
            title: '⚡ 编辑档位模板 — ' + t.name,
            slots: t.slots || [],
            onSave: async (slots) => {
                try {
                    await apiCall('admin/slottpl/save', 'POST', { id: id, name: t.name, slots: slots });
                    toast('✅ 已保存 ' + slots.length + ' 个档位');
                    SlotEditor.close();
                    switchTab('slottpl');
                } catch (e) { toast('保存失败：' + e.message, 'er'); }
            }
        });
    }

    function renameSlotTpl(id) {
        const t = _slotTplCache.find(x => x.id === id);
        if (!t) return;
        FF.open({
            title: '📝 重命名档位模板',
            okText: '保存',
            fields: [{ key: 'name', label: '模板名称', type: 'text', value: t.name }],
            onSubmit: async (v, setErr) => {
                if (!v.name.trim()) { setErr('不能为空'); return false; }
                await apiCall('admin/slottpl/save', 'POST', { id: id, name: v.name.trim(), slots: t.slots || [] });
                toast('✅ 已改名'); switchTab('slottpl');
            }
        });
    }

    function delSlotTpl(id, name) {
        if (!confirm('删除档位模板【' + name + '】？\n\n已下发给用户的档位不受影响。')) return;
        apiCall('admin/slottpl/delete', 'POST', { id })
            .then(() => { if (_slotTplSel === id) _slotTplSel = ''; toast('✅ 已删除'); switchTab('slottpl'); })
            .catch(e => toast('失败：' + e.message, 'er'));
    }

    async function doSlotDeploy() {
        if (!_slotTplSel) { toast('请先选一个档位模板', 'er'); return; }
        const usernames = Object.keys(_slotDeployUserSel).filter(k => _slotDeployUserSel[k]);
        if (!usernames.length) { toast('请勾选至少 1 个账号', 'er'); return; }
        const modeEl = document.getElementById('slotDeployMode');
        const mode = modeEl ? modeEl.value : 'replace';
        const t = _slotTplCache.find(x => x.id === _slotTplSel);
        if (!confirm('把【' + (t ? t.name : '') + '】（' + (t ? t.slotCount : 0) + '档）下发给 ' + usernames.length + ' 个账号？\n\n模式：' + (mode === 'append' ? '追加' : '整套替换'))) return;
        try {
            const r = await apiCall('admin/slottpl/deploy', 'POST', { templateId: _slotTplSel, usernames, mode });
            let msg = '✅ 已下发给 ' + (r.count || 0) + ' 个账号（每个 ' + (r.slotCount || 0) + ' 档）\n\n用户刷新页面后生效。';
            if (r.errors && r.errors.length) msg += '\n\n⚠️ 失败：\n' + r.errors.slice(0, 6).join('\n');
            alert(msg);
            _slotDeployUserSel = {};
            switchTab('slottpl');
        } catch (e) { toast('下发失败：' + e.message, 'er'); }
    }

    /* ========================================================== */
    /* ============ 引擎体检（含档位模型校验） ================== */
    /* ========================================================== */
    async function renderHealth(box) {
        box.innerHTML = `<div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap;align-items:center">
                <button class="btn btn-p btn-s" onclick="Admin.runHealth()">🔍 开始体检</button>
                <label class="btn btn-s" style="cursor:pointer"><input type="checkbox" id="hcCheckSlots" checked style="accent-color:var(--pri);margin-right:5px">同时校验档位模型</label>
                <button class="btn btn-s" onclick="Admin.switchTab('engines')">← 返回引擎列表</button>
            </div>
            <div style="font-size:12px;color:var(--text2);margin-bottom:12px;line-height:1.7">
                验证每个引擎的 <b>BaseURL / Key / 权限 / 余额 / 网络</b>，并检查<b>默认模型</b>与<b>各账号档位里的模型</b>是否真的存在。<br>
                每个引擎最多等 12 秒；引擎多时分批进行。
            </div>
            <div id="healthArea"><div style="padding:26px;text-align:center;color:var(--text2);border:1px dashed var(--border);border-radius:10px">点上方按钮开始</div></div>`;
    }

    async function runHealth() {
        const area = document.getElementById('healthArea');
        if (!area) return;
        const chkSlots = (document.getElementById('hcCheckSlots') || {}).checked;
        area.innerHTML = '<div style="color:var(--text2);padding:14px;text-align:center">正在读取引擎与档位...</div>';

        try {
            const engData = await apiCall('admin/engines/list');
            const engs = engData.engines || [];
            _engCacheAll = engs;
            if (!engs.length) { area.innerHTML = '<div style="padding:20px;text-align:center;color:var(--text2)">还没有任何公有引擎</div>'; return; }

            /* 收集所有档位模型（去重） */
            let extraModels = [];
            if (chkSlots) {
                const uData = await apiCall('admin/users/list');
                _usersCache = uData.users || [];
                const set = {};
                _usersCache.forEach(u => {
                    _pj(u.quickModels, []).forEach(s => { if (s && s.model && String(s.model).trim()) set[String(s.model).trim()] = 1; });
                });
                extraModels = Object.keys(set);
            }

            const ids = engs.map(e => e.id);
            const BATCH = 10;
            const batches = [];
            for (let i = 0; i < ids.length; i += BATCH) batches.push(ids.slice(i, i + BATCH));

            _healthRows = [];
            let done = 0;

            const drawProgress = () => {
                const pct = Math.round(done / ids.length * 100);
                area.innerHTML = '<div style="font-size:13px;margin-bottom:6px">正在体检 <b>' + done + ' / ' + ids.length + '</b> ...'
                    + (extraModels.length ? ('（同时校验 ' + extraModels.length + ' 个档位模型）') : '') + '</div>'
                    + '<div class="health-bar"><div class="health-bar-in" style="width:' + pct + '%"></div></div>'
                    + drawHealthTable();
            };
            drawProgress();

            for (const b of batches) {
                try {
                    const r = await apiCall('admin/engines/health', 'POST', { ids: b, extraModels: extraModels });
                    _healthRows = _healthRows.concat(r.results || []);
                } catch (e) {
                    b.forEach(id => {
                        const eg = engs.find(x => x.id === id) || {};
                        _healthRows.push({ id, username: eg.username || '', name: eg.name || '', model: eg.model || '', ok: false, msg: '请求失败：' + e.message, ms: 0 });
                    });
                }
                done += b.length;
                drawProgress();
            }

            const okN = _healthRows.filter(r => r.ok).length;
            const badN = _healthRows.length - okN;
            const missN = _healthRows.filter(r => r.missModels && r.missModels.length).length;

            area.innerHTML = `<div class="health-sum">
                    <div class="health-card ok"><b>${okN}</b>正常</div>
                    <div class="health-card bad"><b>${badN}</b>异常</div>
                    ${missN ? `<div class="health-card bad"><b>${missN}</b>档位模型缺失</div>` : ''}
                    <div class="health-card"><b>${_healthRows.length}</b>总计</div>
                </div>` + drawHealthTable();

            if (badN || missN) toast('⚠️ 体检完成：' + badN + ' 异常' + (missN ? '，' + missN + ' 个引擎缺档位模型' : ''), 'er');
            else toast('✅ 体检完成：全部正常');
        } catch (e) {
            area.innerHTML = '<div style="color:#ef4444;padding:20px">体检失败：' + esc(e.message) + '</div>';
        }
    }

    function drawHealthTable() {
        if (!_healthRows.length) return '';
        const rows = _healthRows.slice().sort((a, b) => {
            const sa = (a.ok ? 1 : 0) - ((a.missModels && a.missModels.length) ? 0.5 : 0);
            const sb = (b.ok ? 1 : 0) - ((b.missModels && b.missModels.length) ? 0.5 : 0);
            return sa - sb;
        });
        let html = '<div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>状态</th><th>账号</th><th>引擎</th><th>默认模型</th><th>档位模型</th><th>延迟</th><th>说明</th></tr></thead><tbody>';
        rows.forEach(r => {
            let mdl = '-';
            if (r.ok) {
                if (r.modelOk === true) mdl = '<span style="color:#10b981">✅ 存在</span>';
                else if (r.modelOk === false) mdl = '<span style="color:#f59e0b">⚠️ 找不到<br><span style="font-size:10px">' + esc(r.model || '') + '</span></span>';
                else mdl = '<span style="color:var(--text2);font-size:11px">' + (r.model ? '未校验' : '未设默认') + '</span>';
            }
            let slotCol = '-';
            if (r.ok) {
                if (r.missModels && r.missModels.length) {
                    slotCol = '<span style="color:#ef4444;font-size:11px">❌ 缺 ' + r.missModels.length + ' 个<br>' + esc(r.missModels.slice(0, 3).join('、')) + (r.missModels.length > 3 ? '…' : '') + '</span>';
                } else {
                    slotCol = '<span style="color:#10b981;font-size:11px">✅ 全部可用</span>';
                }
            }
            const bad = (!r.ok) || (r.missModels && r.missModels.length);
            html += '<tr class="' + (bad ? 'health-row-bad' : 'health-row-ok') + '">'
                + '<td>' + (r.ok ? (bad ? '🟡 部分' : '🟢 正常') : '🔴 异常') + '</td>'
                + '<td>' + esc(r.username || '') + '</td>'
                + '<td>' + esc(r.name || '') + (r.engineType === 'image' ? ' 🎨' : '') + '</td>'
                + '<td>' + mdl + '</td>'
                + '<td>' + slotCol + '</td>'
                + '<td>' + (r.ms ? (r.ms + 'ms') : '-') + '</td>'
                + '<td class="health-msg">' + esc(r.msg || '') + '</td>'
                + '</tr>';
        });
        html += '</tbody></table></div>';
        return html;
    }

    /* ========================================================== */
    /* ====================== 模型库 ============================ */
    /* ========================================================== */
    async function renderModels(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>';
        try {
            const data = await apiCall('admin/models/list');
            const models = data.models || [];
            let html = `<div style="display:flex;gap:8px;margin-bottom:12px"><button class="btn btn-p btn-s" onclick="Admin.showModelEdit('')">➕ 新增模型</button><button class="btn btn-s" onclick="Admin.switchTab('models')">🔄 刷新</button></div><div style="font-size:12px;color:var(--text2);margin-bottom:8px">模型库：各模型单价（美元/1M token）。<b>档位换模型后，费用靠这里查</b>，请把用到的模型都录进来。</div><div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>模型名</th><th>输入</th><th>输出</th><th>缓存读</th><th>缓存写</th><th>操作</th></tr></thead><tbody>`;
            models.forEach(m => {
                html += `<tr><td>${esc(m.model_name)}</td><td>${m.price_in}</td><td>${m.price_out}</td><td>${m.price_cache_read}</td><td>${m.price_cache_write}</td><td class="admin-ops"><button onclick='Admin.showModelEdit(${JSON.stringify(m.model_name)})'>✏️</button><button onclick='Admin.delModel(${JSON.stringify(m.model_name)})' style="color:#ef4444">🗑️</button></td></tr>`;
            });
            html += '</tbody></table></div>';
            box.innerHTML = html;
            box._models = models;
        } catch (e) { box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>'; }
    }

    function showModelEdit(modelName) {
        const box = document.getElementById('adminBody');
        const models = (box && box._models) || [];
        const m = modelName ? models.find(x => x.model_name === modelName) : null;
        FF.open({
            title: m ? ('✏️ 编辑单价 — ' + m.model_name) : '➕ 新增模型单价',
            tip: '单位：<b>美元 / 100万 token</b>。引擎/档位没填单价时，系统按模型名来这里查。',
            okText: '保存',
            fields: [
                m ? { key: '_ro', label: '模型名（不可改）', type: 'text', value: m.model_name, hint: '如需改名请删除后重建' }
                  : { key: 'model_name', label: '模型名', type: 'text', value: '', placeholder: '如 gpt-4o', hint: '必须与引擎/档位里填的完全一致' },
                { key: 'priceIn', label: '输入单价', type: 'number', value: m ? m.price_in : 0, step: '0.01' },
                { key: 'priceOut', label: '输出单价', type: 'number', value: m ? m.price_out : 0, step: '0.01' },
                { key: 'priceCR', label: '缓存读单价', type: 'number', value: m ? m.price_cache_read : 0, step: '0.01' },
                { key: 'priceCW', label: '缓存写单价', type: 'number', value: m ? m.price_cache_write : 0, step: '0.01' },
            ],
            onSubmit: async (v, setErr) => {
                const name = m ? m.model_name : (v.model_name || '').trim();
                if (!name) { setErr('模型名必填'); return false; }
                await apiCall('admin/models/save', 'POST', { model_name: name, priceIn: v.priceIn, priceOut: v.priceOut, priceCR: v.priceCR, priceCW: v.priceCW });
                toast('✅ 已保存'); switchTab('models');
            }
        });
    }
    function delModel(name) {
        if (!confirm('删除模型【' + name + '】单价？')) return;
        apiCall('admin/models/delete', 'POST', { model_name: name })
            .then(() => { toast('✅ 已删除'); switchTab('models'); })
            .catch(e => toast('失败：' + e.message, 'er'));
    }

    /* ========================================================== */
    /* ============ 预设管理（含分支 + 步骤指定模型） =========== */
    /* ========================================================== */
    async function renderPresets(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中（解密明文）...</div>';
        try {
            let data = null;
            const res = await apiCall('admin/presets/get');
            if (res.presets) data = res.presets;
            else if (typeof Workflow !== 'undefined' && Workflow.getRawData) data = Workflow.getRawData();
            if (!data) data = { version: 3, groups: [], security: { sensitiveWords: [], alertWebhook: '', alertKeyword: '飞凡警报', simThreshold: 70, guard: true }, presets: [] };

            _presetData = JSON.parse(JSON.stringify(data));
            if (!Array.isArray(_presetData.groups)) _presetData.groups = [];

            if (typeof Workflow !== 'undefined' && Workflow.decrypt) {
                for (const p of (_presetData.presets || [])) {
                    for (const s of (p.steps || [])) {
                        const groups = [s.segments || []];
                        (s.variants || []).forEach(v => groups.push(v.segments || []));
                        for (const segs of groups) {
                            for (const seg of segs) {
                                if (seg.type === 'prompt') {
                                    try { seg._plain = await Workflow.decrypt(seg.hidden || ''); }
                                    catch (e) { seg._plain = '（解密失败）'; }
                                }
                            }
                        }
                    }
                }
            }
            _presetDirty = false;
            _curPresetIdx = (_presetData.presets && _presetData.presets.length) ? 0 : -1;
            _openSteps = {}; _stepSel = {}; _curVar = {};
            drawPresetLayout(box);
        } catch (e) { box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>'; }
    }

    function drawPresetLayout(box) {
        box.innerHTML = `<div class="preset-layout">
            <div class="preset-left">
                <div class="preset-left-tools">
                    <button class="btn btn-p btn-s btn-b" onclick="Admin.addPreset()">➕ 新建预设</button>
                    <button class="btn btn-s btn-b" onclick="Admin.manageGroups()" style="margin-top:6px">🏷️ 分组管理</button>
                    <input type="text" id="presetSearch" placeholder="🔍 搜索预设名" oninput="Admin.filterPresetList()">
                </div>
                <div class="preset-list-scroll" id="presetListScroll"></div>
                <div class="preset-left-ft">
                    <button class="btn btn-s btn-b" onclick="Admin.showSecurity()">🛡️ 安全设置</button>
                    <button class="btn btn-p btn-s btn-b" onclick="Admin.savePresets()" style="margin-top:6px">💾 保存到云端</button>
                    <div style="font-size:10px;color:var(--text2);margin-top:6px;display:flex;gap:8px"><a href="javascript:;" onclick="Admin.exportPresetsJSON()">导出备份</a><label style="cursor:pointer">导入JSON<input type="file" accept=".json" onchange="Admin.importPresetsJSON(this)" style="display:none"></label></div>
                    <div id="presetDirtyTip" style="font-size:11px;color:#ef4444;margin-top:4px"></div>
                </div>
            </div>
            <div class="preset-right" id="presetRight"></div>
        </div>`;
        drawPresetList(); drawPresetEditor();
    }

    function drawPresetList() {
        const scroll = document.getElementById('presetListScroll');
        if (!scroll) return;
        const kw = ((document.getElementById('presetSearch') || {}).value || '').toLowerCase();
        let html = '';
        (_presetData.presets || []).forEach((p, i) => {
            if (kw && !(p.name || '').toLowerCase().includes(kw)) return;
            const hasBranch = (p.steps || []).some(s => Array.isArray(s.variants) && s.variants.length);
            const hasModel = (p.steps || []).some(s => s.forceModel);
            html += `<div class="preset-litem${i === _curPresetIdx ? ' act' : ''}" onclick="Admin.selectPreset(${i})"><div class="pl-name">${esc(p.name || '未命名')}</div><div class="pl-meta">${esc(p.group || '无分组')} · ${(p.steps || []).length}步${hasBranch ? ' · 含分支' : ''}${hasModel ? ' · 指定模型' : ''}</div></div>`;
        });
        scroll.innerHTML = html || '<div style="font-size:12px;color:var(--text2);padding:10px;text-align:center">无预设</div>';
    }

    function filterPresetList() { drawPresetList(); }
    function selectPreset(i) { _curPresetIdx = i; _openSteps = {}; _stepSel = {}; _curVar = {}; drawPresetList(); drawPresetEditor(); }
    function markDirty() { _presetDirty = true; const t = document.getElementById('presetDirtyTip'); if (t) t.textContent = '● 有未保存修改'; }
    function _vidOf(si) { return _curVar[si] || DEFAULT_VID; }

    function _segsRef(si, vid) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        if (!vid || vid === DEFAULT_VID) { if (!Array.isArray(s.segments)) s.segments = []; return s.segments; }
        const v = (s.variants || []).find(x => x.id === vid);
        if (!v) { if (!Array.isArray(s.segments)) s.segments = []; return s.segments; }
        if (!Array.isArray(v.segments)) v.segments = [];
        return v.segments;
    }
    function _segsSummary(segs) {
        let p = 0, inp = 0, bl = 0;
        (segs || []).forEach(seg => { if (seg.type === 'prompt') p++; else if (seg.type === 'input') inp++; else if (seg.type === 'blank') bl++; });
        const a = [];
        if (p) a.push(p + '隐藏'); if (inp) a.push(inp + '输入'); if (bl) a.push(bl + '填空');
        return a.join('/');
    }

    function drawPresetEditor() {
        const right = document.getElementById('presetRight');
        if (!right) return;
        if (_curPresetIdx < 0 || !_presetData.presets[_curPresetIdx]) {
            right.innerHTML = '<div style="color:var(--text2);text-align:center;padding:60px">← 左侧选择或新建预设</div>'; return;
        }
        const p = _presetData.presets[_curPresetIdx];
        const pi = _curPresetIdx;
        const groupOpts = ['<option value="">（无分组）</option>']
            .concat((_presetData.groups || []).map(g => `<option value="${esc(g)}"${p.group === g ? ' selected' : ''}>${esc(g)}</option>`)).join('');

        let html = `<div class="preset-edit-hdr">
            <input value="${esc(p.name)}" onchange="Admin.updP(${pi},'name',this.value)" placeholder="预设名" style="font-weight:600;font-size:15px;flex:1;min-width:140px">
            <select onchange="Admin.updP(${pi},'group',this.value)" style="width:130px">${groupOpts}</select>
            <button class="btn btn-s" onclick="Admin.dupPreset(${pi})">📄 复制</button>
            <button class="btn btn-s btn-d" onclick="Admin.delPreset(${pi})">🗑️ 删预设</button>
        </div>
        <div style="font-size:11px;color:var(--text2);margin-bottom:8px">💡 拖 ⣿ 排序；步骤内可建 A/B 版本（最多3个）；🎯 可为该步指定模型/参数（用户无感）</div>
        <div class="preset-steps" id="presetStepsBox">`;

        (p.steps || []).forEach((s, si) => {
            const open = !!_openSteps[si];
            const vid = _vidOf(si);
            const curSegs = _segsRef(si, vid);
            const summary = _segsSummary(curSegs);
            const varCount = (s.variants || []).length;
            const mSum = s.forceModel ? ('🎯' + s.forceModel) : '';
            const pSum = s.forceParams ? ('⚙️' + (paramSummary(s.forceParams) || '')) : '';

            html += `<div class="pstep${open ? ' open' : ''}" draggable="true"
                    ondragstart="Admin.stepDragStart(event,${si})" ondragover="Admin.stepDragOver(event,${si})"
                    ondragleave="Admin.stepDragLeave(event,${si})" ondrop="Admin.stepDrop(event,${si})" ondragend="Admin.stepDragEnd(event)">
                <div class="pstep-bar" onclick="Admin.toggleStep(${si})">
                    <span class="pstep-drag" onclick="event.stopPropagation()" title="拖动排序">⣿</span>
                    <input type="checkbox" onclick="event.stopPropagation()" onchange="Admin.toggleStepSel(${si},this.checked)" ${_stepSel[si] ? 'checked' : ''} title="勾选后可拆分为新预设" style="flex-shrink:0">
                    <span class="pstep-caret">${open ? '▼' : '▶'}</span>
                    <span class="pstep-title">步骤${si + 1}：${esc(s.name || '未命名')}</span>
                    ${varCount ? '<span style="font-size:10px;color:#f59e0b;flex-shrink:0">🔀' + (varCount + 1) + '版</span>' : ''}
                    ${mSum ? '<span style="font-size:10px;color:#8b5cf6;flex-shrink:0" title="该步骤指定模型">' + esc(mSum) + '</span>' : ''}
                    ${pSum ? '<span style="font-size:10px;color:#8b5cf6;flex-shrink:0">' + esc(pSum) + '</span>' : ''}
                    ${summary ? '<span style="font-size:10px;color:var(--text2);flex-shrink:0">' + summary + '</span>' : ''}
                    ${s.engineName ? '<span style="font-size:10px;color:var(--pri)">🔌' + esc(s.engineName) + '</span>' : ''}
                    <span class="pstep-ops" onclick="event.stopPropagation()">
                        <button onclick="Admin.showStepModel(${si})" title="指定模型/参数">🎯</button>
                        <button onclick="Admin.dupStep(${si})" title="复制此步骤">📄</button>
                        <button onclick="Admin.moveStep(${si},-1)" title="上移">↑</button>
                        <button onclick="Admin.moveStep(${si},1)" title="下移">↓</button>
                        <button onclick="Admin.delStep(${si})" title="删除" style="color:#ef4444">🗑️</button>
                    </span>
                </div>
                <div class="pstep-body">
                    <div class="fr"><div class="fg"><label>步骤名</label><input value="${esc(s.name || '')}" onchange="Admin.updS(${si},'name',this.value)"></div><div class="fg"><label>绑定引擎名(选填)</label><input value="${esc(s.engineName || '')}" onchange="Admin.updS(${si},'engineName',this.value)" placeholder="该步自动用此公有引擎"></div></div>
                    ${_renderVariantTabs(si, s, vid)}
                    <div class="seg-list">`;

            curSegs.forEach((seg, gi) => {
                if (seg.type === 'prompt') {
                    const wc = (typeof cntW === 'function') ? cntW(seg._plain || '') : (seg._plain || '').length;
                    html += `<div class="seg-item seg-prompt"><div class="seg-label">🔒隐藏指令（明文，保存自动加密）<span style="color:var(--text2);float:right">${wc} 字</span></div><textarea class="seg-prompt-ta" onchange="Admin.updSegPrompt(${si},${gi},this.value)" placeholder="隐藏指令明文">${esc(seg._plain || '')}</textarea><button class="btn btn-s btn-d" onclick="Admin.delSeg(${si},${gi})">删</button></div>`;
                }
                else if (seg.type === 'input') html += `<div class="seg-item seg-input"><div class="seg-label">✍️输入框</div><input value="${esc(seg.placeholder || '')}" onchange="Admin.updSeg(${si},${gi},'placeholder',this.value)" placeholder="提示文字"><input value="${esc(seg.defaultValue || '')}" onchange="Admin.updSeg(${si},${gi},'defaultValue',this.value)" placeholder="默认值"><button class="btn btn-s btn-d" onclick="Admin.delSeg(${si},${gi})">删</button></div>`;
                else if (seg.type === 'blank') html += `<div class="seg-item seg-blank"><div class="seg-label">📝填空题（{}=空位）</div><input value="${esc(seg.template || '')}" onchange="Admin.updSeg(${si},${gi},'template',this.value)" placeholder="如：题材是{}，视角是{}"><button class="btn btn-s btn-d" onclick="Admin.delSeg(${si},${gi})">删</button></div>`;
            });

            html += `</div><div style="display:flex;gap:4px;margin-top:6px"><button class="btn btn-s" onclick="Admin.addSeg(${si},'prompt')">+隐藏指令</button><button class="btn btn-s" onclick="Admin.addSeg(${si},'input')">+输入框</button><button class="btn btn-s" onclick="Admin.addSeg(${si},'blank')">+填空题</button></div></div></div>`;
        });

        html += `</div><div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap">
                <button class="btn btn-p btn-s" onclick="Admin.addStep()">➕ 添加步骤</button>
                <button class="btn btn-s" onclick="Admin.stepsToNewPreset()">✂️ 勾选步骤另存为新预设</button>
            </div>`;
        right.innerHTML = html;
    }

    /* ★ 步骤指定模型 / 参数 */
    function showStepModel(si) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        if (!s) return;
        FF.open({
            title: '🎯 步骤指定模型 — ' + (s.name || ('步骤' + (si + 1))),
            wide: true,
            tip: '在这里写死这一步用什么模型/参数，<b>用户端完全无感</b>，不用他操作。<br>'
                + '通道仍用用户当前引擎（或上面填的"绑定引擎名"），只替换模型和参数。<br>'
                + '⚠️ 如果用户的引擎不支持这个模型，发送时会报错并提示他联系管理员。',
            okText: '保存',
            fields: [
                { key: 'forceModel', label: '指定模型（留空=不指定）', type: 'text', value: s.forceModel || '', placeholder: '如 claude-opus-4', hint: '必须是用户引擎所在中转站支持的模型名' },
                { key: 'forceIsImage', label: '🎨 该步骤走生图', type: 'check', value: !!s.forceIsImage, hint: '勾选后这一步会调生图接口' },
                { key: 'useForceParams', label: '⚙️ 该步骤覆盖参数', type: 'check', value: !!s.forceParams, hint: '勾选后下面参数生效（优先级最高，会盖掉引擎和档位的参数）' },
                { type: 'html', html: paramBoxHtml(s.forceParams, 'stepP', { title: '该步骤专用参数', forceOpen: true }) },
            ],
            onSubmit: (v) => {
                s.forceModel = (v.forceModel || '').trim();
                s.forceIsImage = !!v.forceIsImage;
                s.forceParams = v.useForceParams ? readParamBox('stepP') : null;
                markDirty(); drawPresetList(); drawPresetEditor();
                toast(s.forceModel ? ('✅ 该步骤将使用 ' + s.forceModel) : '✅ 已取消模型指定');
            }
        });
    }

    /* ---------- 版本标签页 ---------- */
    function _renderVariantTabs(si, step, curVid) {
        const vars = (step.variants || []);
        const total = 1 + vars.length;
        let html = '<div class="pv-tabs"><span class="pv-lbl">🔀 版本：</span>';
        const dLabel = step.defaultLabel || '默认';
        const dAct = (curVid === DEFAULT_VID) ? ' act' : '';
        html += '<span class="pv-tab' + dAct + '" style="user-select:none" title="点击切换｜双击改名"'
            + " onclick='Admin.selVariant(" + si + "," + JSON.stringify(DEFAULT_VID) + ")'"
            + " ondblclick='Admin.renameDefaultVar(" + si + ")'>" + esc(dLabel) + '</span>';
        vars.forEach(v => {
            const act = (curVid === v.id) ? ' act' : '';
            const vj = JSON.stringify(v.id);
            html += '<span class="pv-tab' + act + '" style="user-select:none" title="点击切换｜双击改名"'
                + " onclick='Admin.selVariant(" + si + "," + vj + ")'"
                + " ondblclick='Admin.renameVariant(" + si + "," + vj + ")'>" + esc(v.label || '未命名')
                + '<button class="pv-x"' + " onclick='event.stopPropagation();Admin.delVariant(" + si + "," + vj + ")'" + ' title="删除此版本">×</button></span>';
        });
        if (total < 3) html += '<button class="pv-add" onclick="Admin.addVariant(' + si + ')">+ 新建版本</button>';
        else html += '<span style="font-size:10px;color:var(--text2)">（已达3版上限）</span>';
        html += '</div>';
        return html;
    }

    function selVariant(si, vid) { _curVar[si] = vid; drawPresetEditor(); }
    function renameDefaultVar(si) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        FF.open({
            title: '默认版显示名', tip: '用户端会看到这个标签（如"标准版"）。', okText: '保存',
            fields: [{ key: 'label', label: '标签名', type: 'text', value: s.defaultLabel || '默认' }],
            onSubmit: (v) => { s.defaultLabel = (v.label || '').trim() || '默认'; markDirty(); drawPresetEditor(); }
        });
    }
    function addVariant(si) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        if (!Array.isArray(s.variants)) s.variants = [];
        if (s.variants.length >= 2) { toast('最多 3 个版本（含默认版）', 'er'); return; }
        FF.open({
            title: '➕ 新建步骤版本',
            tip: '版本用于同一步骤的 A/B 方案，用户端可自选。', okText: '创建',
            fields: [
                { key: 'label', label: '版本名称', type: 'text', value: 'A版', placeholder: '如：A版-悬疑向' },
                { key: 'copy', label: '从当前版本复制内容', type: 'check', value: true, hint: '推荐：复制后改几处即可' },
            ],
            onSubmit: (v, setErr) => {
                const label = (v.label || '').trim();
                if (!label) { setErr('版本名称不能为空'); return false; }
                const srcSegs = v.copy ? _segsRef(si, _vidOf(si)) : [];
                const nv = { id: _rid('v_'), label: label, segments: JSON.parse(JSON.stringify(srcSegs || [])) };
                s.variants.push(nv);
                _curVar[si] = nv.id;
                markDirty(); drawPresetList(); drawPresetEditor();
                toast('✅ 已新建版本「' + nv.label + '」');
            }
        });
    }
    function renameVariant(si, vid) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        const v = (s.variants || []).find(x => x.id === vid);
        if (!v) return;
        FF.open({
            title: '重命名版本', okText: '保存',
            fields: [{ key: 'label', label: '版本名称', type: 'text', value: v.label || '' }],
            onSubmit: (val, setErr) => { const nn = (val.label || '').trim(); if (!nn) { setErr('不能为空'); return false; } v.label = nn; markDirty(); drawPresetEditor(); }
        });
    }
    function delVariant(si, vid) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        const v = (s.variants || []).find(x => x.id === vid);
        if (!v) return;
        if (!confirm('删除版本「' + (v.label || '') + '」？该版本片段会一起删除。')) return;
        s.variants = (s.variants || []).filter(x => x.id !== vid);
        if (_curVar[si] === vid) _curVar[si] = DEFAULT_VID;
        markDirty(); drawPresetList(); drawPresetEditor();
        toast('已删除版本');
    }

    function updP(pi, f, v) { _presetData.presets[pi][f] = v; markDirty(); if (f === 'name' || f === 'group') drawPresetList(); }
    function updS(si, f, v) { _presetData.presets[_curPresetIdx].steps[si][f] = v; markDirty(); }
    function updSeg(si, gi, f, v) { const segs = _segsRef(si, _vidOf(si)); if (segs[gi]) segs[gi][f] = v; markDirty(); }
    function updSegPrompt(si, gi, v) { const segs = _segsRef(si, _vidOf(si)); if (!segs[gi]) return; segs[gi]._plain = v; segs[gi]._dirty = true; markDirty(); drawPresetEditor(); }
    function addSeg(si, type) {
        const seg = { type };
        if (type === 'prompt') { seg.hidden = ''; seg._plain = ''; seg._dirty = true; }
        else if (type === 'input') { seg.placeholder = '请输入...'; seg.defaultValue = ''; }
        else if (type === 'blank') { seg.template = ''; }
        _segsRef(si, _vidOf(si)).push(seg);
        markDirty(); drawPresetEditor();
    }
    function delSeg(si, gi) { _segsRef(si, _vidOf(si)).splice(gi, 1); markDirty(); drawPresetEditor(); }
    function toggleStep(si) { _openSteps[si] = !_openSteps[si]; drawPresetEditor(); }
    function toggleStepSel(si, c) { _stepSel[si] = c; }

    function addPreset() {
        FF.open({
            title: '➕ 新建预设', okText: '创建',
            fields: [
                { key: 'name', label: '预设名称', type: 'text', value: '', placeholder: '如：小说生成流程' },
                { key: 'group', label: '分组', type: 'select', value: '', options: [{ v: '', t: '（无分组）' }].concat((_presetData.groups || []).map(g => ({ v: g, t: g }))), hint: '分组用于给不同员工分配权限' },
            ],
            onSubmit: (v, setErr) => {
                const nm = (v.name || '').trim();
                if (!nm) { setErr('预设名称不能为空'); return false; }
                _presetData.presets.push({ id: _rid('p'), name: nm, group: v.group || '', steps: [] });
                _curPresetIdx = _presetData.presets.length - 1;
                _openSteps = {}; _stepSel = {}; _curVar = {};
                markDirty(); drawPresetList(); drawPresetEditor();
            }
        });
    }
    function dupPreset(pi) {
        const src = _presetData.presets[pi];
        if (!src) return;
        const copy = JSON.parse(JSON.stringify(src));
        copy.id = _rid('p');
        copy.name = (src.name || '未命名') + ' 副本';
        (copy.steps || []).forEach(s => { s.id = _rid('s'); (s.variants || []).forEach(v => { v.id = _rid('v_'); }); });
        _presetData.presets.splice(pi + 1, 0, copy);
        _curPresetIdx = pi + 1;
        _openSteps = {}; _stepSel = {}; _curVar = {};
        markDirty(); drawPresetList(); drawPresetEditor();
        toast('✅ 已复制预设');
    }
    function delPreset(pi) {
        const p = _presetData.presets[pi];
        if (!confirm('删除预设【' + ((p && p.name) || '') + '】？')) return;
        _presetData.presets.splice(pi, 1);
        _curPresetIdx = _presetData.presets.length ? 0 : -1;
        _openSteps = {}; _stepSel = {}; _curVar = {};
        markDirty(); drawPresetList(); drawPresetEditor();
    }
    function addStep() {
        const p = _presetData.presets[_curPresetIdx];
        if (!p.steps) p.steps = [];
        p.steps.push({ id: _rid('s'), name: '新步骤', order: p.steps.length + 1, segments: [] });
        _openSteps[p.steps.length - 1] = true;
        markDirty(); drawPresetList(); drawPresetEditor();
    }
    function dupStep(si) {
        const p = _presetData.presets[_curPresetIdx];
        const src = p.steps[si];
        if (!src) return;
        const copy = JSON.parse(JSON.stringify(src));
        copy.id = _rid('s');
        copy.name = (src.name || '未命名') + ' 副本';
        (copy.variants || []).forEach(v => { v.id = _rid('v_'); });
        p.steps.splice(si + 1, 0, copy);
        p.steps.forEach((s, i) => s.order = i + 1);
        _openSteps = {}; _stepSel = {}; _curVar = {};
        _openSteps[si + 1] = true;
        markDirty(); drawPresetEditor();
        toast('✅ 已复制步骤');
    }
    function delStep(si) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        if (!confirm('删除步骤【' + ((s && s.name) || '') + '】？')) return;
        _presetData.presets[_curPresetIdx].steps.splice(si, 1);
        _presetData.presets[_curPresetIdx].steps.forEach((x, i) => x.order = i + 1);
        _openSteps = {}; _stepSel = {}; _curVar = {};
        markDirty(); drawPresetList(); drawPresetEditor();
    }
    function moveStep(si, dir) {
        const steps = _presetData.presets[_curPresetIdx].steps;
        const j = si + dir;
        if (j < 0 || j >= steps.length) return;
        const t = steps[si]; steps[si] = steps[j]; steps[j] = t;
        steps.forEach((s, i) => s.order = i + 1);
        [_openSteps, _stepSel, _curVar].forEach(map => {
            const ta = map[si], tb = map[j];
            if (tb === undefined) delete map[si]; else map[si] = tb;
            if (ta === undefined) delete map[j]; else map[j] = ta;
        });
        markDirty(); drawPresetEditor();
    }

    function stepDragStart(ev, si) { _dragStepIdx = si; ev.currentTarget.classList.add('dragging'); try { ev.dataTransfer.effectAllowed = 'move'; ev.dataTransfer.setData('text/plain', String(si)); } catch (e) {} }
    function stepDragOver(ev, si) { if (_dragStepIdx === null || _dragStepIdx === si) return; ev.preventDefault(); try { ev.dataTransfer.dropEffect = 'move'; } catch (e) {} ev.currentTarget.classList.add('drag-over'); }
    function stepDragLeave(ev) { ev.currentTarget.classList.remove('drag-over'); }
    function stepDrop(ev, si) {
        ev.preventDefault(); ev.currentTarget.classList.remove('drag-over');
        let from = _dragStepIdx;
        if (from === null) { const raw = ev.dataTransfer ? ev.dataTransfer.getData('text/plain') : ''; from = raw === '' ? null : parseInt(raw, 10); }
        if (from === null || isNaN(from) || from === si) { _dragStepIdx = null; return; }
        const steps = _presetData.presets[_curPresetIdx].steps;
        const moved = steps.splice(from, 1)[0];
        steps.splice(si, 0, moved);
        steps.forEach((s, i) => s.order = i + 1);
        _openSteps = {}; _stepSel = {}; _curVar = {};
        _dragStepIdx = null;
        markDirty(); drawPresetEditor();
        toast('✅ 已重新排序');
    }
    function stepDragEnd() { _dragStepIdx = null; document.querySelectorAll('.pstep').forEach(el => { el.classList.remove('dragging'); el.classList.remove('drag-over'); }); }

    function stepsToNewPreset() {
        const p = _presetData.presets[_curPresetIdx];
        const picked = (p.steps || []).filter((s, i) => _stepSel[i]);
        if (!picked.length) { toast('请先勾选要提取的步骤', 'er'); return; }
        FF.open({
            title: '✂️ 拆分为新预设',
            tip: '已勾选 <b>' + picked.length + '</b> 个步骤，将复制成新预设（原预设不变）。',
            okText: '创建',
            fields: [
                { key: 'name', label: '新预设名称', type: 'text', value: (p.name || '') + ' - 拆分' },
                { key: 'group', label: '分组', type: 'select', value: p.group || '', options: [{ v: '', t: '（无分组）' }].concat((_presetData.groups || []).map(g => ({ v: g, t: g }))) },
            ],
            onSubmit: (v, setErr) => {
                const nm = (v.name || '').trim();
                if (!nm) { setErr('名称不能为空'); return false; }
                const np = { id: _rid('p'), name: nm, group: v.group || '', steps: JSON.parse(JSON.stringify(picked)) };
                np.steps.forEach((s, i) => { s.id = _rid('s'); s.order = i + 1; (s.variants || []).forEach(x => { x.id = _rid('v_'); }); });
                _presetData.presets.push(np);
                _stepSel = {};
                markDirty(); drawPresetList(); drawPresetEditor();
                toast('✅ 已拆分为新预设「' + np.name + '」');
            }
        });
    }

    function manageGroups() {
        if (!_presetData.groups) _presetData.groups = [];
        const right = document.getElementById('presetRight');
        if (!right) return;
        let html = '<div style="max-width:480px"><h3 style="margin-bottom:6px">🏷️ 分组管理</h3>';
        html += '<div style="font-size:11px;color:var(--text2);margin-bottom:12px;line-height:1.6">分组用于给员工分配权限（👥账号 → 🎫）。改名会自动同步所有预设。</div>';
        html += '<div style="display:flex;gap:6px;margin-bottom:12px"><input id="newGroupName" placeholder="输入新分组名" style="flex:1;padding:7px 10px;border:1px solid var(--border);border-radius:8px;background:var(--bg);color:var(--text)" onkeydown="if(event.key===\'Enter\'){event.preventDefault();Admin.addGroup();}"><button class="btn btn-p btn-s" onclick="Admin.addGroup()">➕ 添加</button></div>';
        html += '<div style="display:flex;flex-direction:column;gap:6px">';
        if (!_presetData.groups.length) html += '<div style="font-size:12px;color:var(--text2)">（暂无分组）</div>';
        _presetData.groups.forEach((g, i) => {
            const cnt = (_presetData.presets || []).filter(p => p.group === g).length;
            html += `<div style="display:flex;align-items:center;gap:8px;padding:8px 10px;background:var(--bg);border:1px solid var(--border);border-radius:8px"><span style="flex:1">${esc(g)} <span style="font-size:11px;color:var(--text2)">（${cnt}个预设）</span></span><button class="btn btn-s" onclick='Admin.renameGroup(${i})'>✏️改名</button><button class="btn btn-s btn-d" onclick='Admin.delGroup(${i})'>🗑️</button></div>`;
        });
        html += '</div><button class="btn btn-s" onclick="Admin.backToEditor()" style="margin-top:14px">← 返回编辑</button></div>';
        right.innerHTML = html;
    }
    function backToEditor() { drawPresetEditor(); }
    function addGroup() {
        const el = document.getElementById('newGroupName');
        const name = (el ? el.value : '').trim();
        if (!name) { toast('请输入分组名', 'er'); return; }
        if (_presetData.groups.includes(name)) { toast('分组已存在', 'er'); return; }
        _presetData.groups.push(name);
        markDirty(); manageGroups();
    }
    function renameGroup(i) {
        const old = _presetData.groups[i];
        FF.open({
            title: '重命名分组',
            tip: '所有属于此分组的预设会自动跟着改。<br>⚠️ 已配过此分组权限的账号，请到「👥账号 → 🎫」重新勾选。',
            okText: '保存',
            fields: [{ key: 'name', label: '分组名', type: 'text', value: old }],
            onSubmit: (v, setErr) => {
                const nn = (v.name || '').trim();
                if (!nn) { setErr('不能为空'); return false; }
                if (nn === old) return true;
                if (_presetData.groups.includes(nn)) { setErr('分组已存在'); return false; }
                _presetData.groups[i] = nn;
                (_presetData.presets || []).forEach(p => { if (p.group === old) p.group = nn; });
                markDirty(); manageGroups(); drawPresetList();
            }
        });
    }
    function delGroup(i) {
        const g = _presetData.groups[i];
        const cnt = (_presetData.presets || []).filter(p => p.group === g).length;
        if (!confirm('删除分组「' + g + '」？' + (cnt ? '\n该分组下 ' + cnt + ' 个预设将变为「无分组」。' : ''))) return;
        (_presetData.presets || []).forEach(p => { if (p.group === g) p.group = ''; });
        _presetData.groups.splice(i, 1);
        markDirty(); manageGroups(); drawPresetList();
    }

    function showSecurity() {
        const sec = _presetData.security || {};
        const right = document.getElementById('presetRight');
        right.innerHTML = `<div style="max-width:520px"><h3 style="margin-bottom:12px">🛡️ 安全设置</h3>
            <div class="fg"><label>敏感词（逗号隔开）<span class="cfg-help" title="用户输入含这些词时直接拦截并推钉钉">?</span></label><textarea id="ps_sensitive" rows="4">${esc((sec.sensitiveWords || []).join(','))}</textarea></div>
            <div class="fg"><label>钉钉 Webhook<span class="cfg-help" title="留空则不推送报警">?</span></label><input id="ps_webhook" value="${esc(sec.alertWebhook || '')}" placeholder="https://oapi.dingtalk.com/robot/send?access_token=..."></div>
            <div class="fr"><div class="fg"><label>报警关键词</label><input id="ps_keyword" value="${esc(sec.alertKeyword || '飞凡警报')}"></div><div class="fg"><label>相似度阈值%<span class="cfg-help" title="AI输出与隐藏指令相似度超过此值即判定泄露并屏蔽">?</span></label><input id="ps_sim" type="number" value="${sec.simThreshold || 70}"></div></div>
            <div class="pt"><input type="checkbox" id="ps_guard" ${sec.guard !== false ? 'checked' : ''}><label for="ps_guard">开启 GUARD 保密前缀（防套取指令）</label></div>
            <button class="btn btn-p" onclick="Admin.applySecurity()" style="margin-top:10px">✔️ 应用（需再点保存到云端）</button>
            <button class="btn btn-s" onclick="Admin.backToEditor()" style="margin-top:10px;margin-left:8px">← 返回编辑</button></div>`;
    }
    function applySecurity() {
        _presetData.security = {
            sensitiveWords: document.getElementById('ps_sensitive').value.split(',').map(s => s.trim()).filter(Boolean),
            alertWebhook: document.getElementById('ps_webhook').value.trim(),
            alertKeyword: document.getElementById('ps_keyword').value.trim() || '飞凡警报',
            simThreshold: parseInt(document.getElementById('ps_sim').value) || 70,
            guard: document.getElementById('ps_guard').checked
        };
        markDirty(); toast('已应用，记得点"💾 保存到云端"');
    }

    async function savePresets() {
        toast('加密并保存中...');
        try {
            for (const p of _presetData.presets) {
                for (const s of (p.steps || [])) {
                    const groups = [s.segments || []];
                    (s.variants || []).forEach(v => groups.push(v.segments || []));
                    for (const segs of groups) {
                        for (const seg of segs) {
                            if (seg.type === 'prompt') {
                                if (typeof Workflow !== 'undefined' && Workflow.encrypt) seg.hidden = await Workflow.encrypt(seg._plain || '');
                                else seg.hidden = '__PLAIN__' + (seg._plain || '');
                            }
                        }
                    }
                }
            }
            const clean = JSON.parse(JSON.stringify(_presetData));
            _stripPlain(clean);
            await apiCall('admin/presets/save', 'POST', { presets: clean });
            _presetDirty = false;
            const t = document.getElementById('presetDirtyTip');
            if (t) t.textContent = '';
            toast('✅ 已保存到云端，所有用户下次加载生效');
            if (typeof Workflow !== 'undefined' && Workflow.reload) await Workflow.reload(clean);
        } catch (e) { toast('保存失败：' + e.message, 'er'); }
    }
    function _stripPlain(data) {
        (data.presets || []).forEach(p => {
            (p.steps || []).forEach(s => {
                const groups = [s.segments || []];
                (s.variants || []).forEach(v => groups.push(v.segments || []));
                groups.forEach(segs => { segs.forEach(seg => { delete seg._plain; delete seg._dirty; }); });
            });
        });
        return data;
    }
    function exportPresetsJSON() {
        const clean = JSON.parse(JSON.stringify(_presetData));
        _stripPlain(clean);
        const blob = new Blob([JSON.stringify(clean, null, 2)], { type: 'application/json' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'presets-backup-' + new Date().toISOString().slice(0, 10) + '.json';
        a.click(); URL.revokeObjectURL(a.href);
        toast('✅ 已导出（不含明文）');
    }
    function importPresetsJSON(inputEl) {
        const file = inputEl.files && inputEl.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                _presetData = JSON.parse(e.target.result);
                if (!Array.isArray(_presetData.groups)) _presetData.groups = [];
                if (typeof Workflow !== 'undefined' && Workflow.decrypt) {
                    for (const p of (_presetData.presets || [])) {
                        for (const s of (p.steps || [])) {
                            const groups = [s.segments || []];
                            (s.variants || []).forEach(v => groups.push(v.segments || []));
                            for (const segs of groups) {
                                for (const seg of segs) {
                                    if (seg.type === 'prompt') {
                                        try { seg._plain = await Workflow.decrypt(seg.hidden || ''); } catch (er) { seg._plain = ''; }
                                    }
                                }
                            }
                        }
                    }
                }
                _curPresetIdx = _presetData.presets.length ? 0 : -1;
                _openSteps = {}; _stepSel = {}; _curVar = {};
                markDirty();
                drawPresetLayout(document.getElementById('adminBody'));
                toast('✅ 已导入（点保存生效）');
            } catch (err) { toast('JSON解析失败', 'er'); }
            inputEl.value = '';
        };
        reader.readAsText(file, 'utf-8');
    }

    /* ========================================================== */
    /* ======================= 监视 ============================= */
    /* ========================================================== */
    async function renderMonitor(box) { box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>'; try { const data = await apiCall('admin/monitor'); const logMap = {}; (data.logs || []).forEach(l => logMap[l.username] = l); const sessMap = {}; (data.sessions || []).forEach(s => sessMap[s.username] = s); const usernames = new Set([...Object.keys(logMap), ...Object.keys(sessMap)]); let html = `<div style="display:flex;gap:16px;margin-bottom:12px;flex-wrap:wrap"><div style="padding:10px 16px;background:var(--pri-l);border-radius:8px"><div style="font-size:11px;color:var(--text2)">当前在线（5分钟内）</div><div style="font-size:22px;font-weight:600;color:#10b981">${data.onlineCount || 0} 人</div></div><button class="btn btn-s" onclick="Admin.switchTab('monitor')" style="align-self:center">🔄刷新</button></div><div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>账号</th><th>对话次数</th><th>累计Token</th><th>不同IP</th><th>最后活跃</th></tr></thead><tbody>`; usernames.forEach(un => { const l = logMap[un] || {}; const s = sessMap[un] || {}; html += `<tr><td>${esc(un)}</td><td>${l.logCount || 0}</td><td>${(l.totalTokens || 0).toLocaleString()}</td><td>${(s.ipc || 0) >= 3 ? '<span style="color:#ef4444">🔴' + s.ipc + '</span>' : (s.ipc || 0)}</td><td style="font-size:11px">${fmtTime(s.last || 0)}</td></tr>`; }); html += '</tbody></table></div><h4 style="font-size:13px;margin:16px 0 8px">📋 最近100条</h4><div style="overflow-x:auto;max-height:280px;overflow-y:auto"><table class="admin-table"><thead><tr><th>时间</th><th>账号</th><th>对话</th><th>轮次</th><th>Token</th><th>模型</th></tr></thead><tbody>'; (data.recent || []).forEach(r => { html += `<tr><td style="font-size:11px">${new Date(r.created_at).toLocaleString()}</td><td>${esc(r.username)}</td><td style="max-width:150px;overflow:hidden;text-overflow:ellipsis">${esc(r.chat_name || '-')}</td><td>${r.rounds || 0}</td><td>${r.tokens || 0}</td><td>${esc(r.model || '-')}</td></tr>`; }); html += '</tbody></table></div>'; box.innerHTML = html; } catch (e) { box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>'; } }

    /* ========================================================== */
    /* ============ 全局设置（全局默认档 + 指令） =============== */
    /* ========================================================== */
    let _cfgOpen = { chunk: true, quick: true, cmds: false };
    let _globalSlots = [];

    async function renderConfig(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>';
        try {
            const data = await apiCall('admin/config/get');
            const cfg = data.config || {};
            await ensureEngines();
            _globalSlots = _pj(cfg.quickModels, []);
            if (!Array.isArray(_globalSlots)) _globalSlots = [];
            const cmdsRaw = cfg.quickCmds || '';
            const cmdCount = cmdsRaw.split('\n').filter(l => l.indexOf('|') > 0).length;

            box.innerHTML = `<div style="max-width:640px">
                <div style="font-size:12px;color:var(--text2);margin-bottom:12px">⚙️ 这里的设置对<b>所有用户</b>生效。改完点最下方「💾 保存全部」。</div>

                <div class="cfg-sec${_cfgOpen.chunk ? ' open' : ''}" id="cfgSecChunk">
                    <div class="cfg-sec-hdr" onclick="Admin.toggleCfgSec('chunk')">
                        <span class="cs-caret">${_cfgOpen.chunk ? '▼' : '▶'}</span>
                        <span class="cs-title">📐 物理打标</span>
                        <span class="cs-sub">当前每块 ${esc(cfg.chunkSize || '300')} 字</span>
                    </div>
                    <div class="cfg-sec-body">
                        <div class="fg"><label>每块字数<span class="cfg-help" title="上传长文时按此字数切块并标注位置，让AI能准确说出'第X字/占比Y%'">?</span></label>
                            <input id="cfg_chunkSize" type="number" value="${esc(cfg.chunkSize || '300')}" min="50" max="5000">
                            <div class="ff-hint">推荐 300。长篇小说可用 500；要精确定位用 200。</div>
                        </div>
                    </div>
                </div>

                <div class="cfg-sec${_cfgOpen.quick ? ' open' : ''}" id="cfgSecQuick">
                    <div class="cfg-sec-hdr" onclick="Admin.toggleCfgSec('quick')">
                        <span class="cs-caret">${_cfgOpen.quick ? '▼' : '▶'}</span>
                        <span class="cs-title">⚡ 全局默认档位</span>
                        <span class="cs-sub">已配 ${_globalSlots.length} 档（仅对"未单独配档"的用户生效）</span>
                    </div>
                    <div class="cfg-sec-body">
                        <div style="font-size:11px;color:var(--text2);margin-bottom:10px;line-height:1.7">
                            这是<b>兜底档位</b>：某用户没有专属档位时，用这套。<br>
                            要给特定用户定制，请到「👥 账号 → ⚡」或「⚡ 档位模板库」批量下发。
                        </div>
                        <div style="padding:10px;background:var(--bg);border:1px solid var(--border);border-radius:9px;margin-bottom:8px">
                            <div style="font-size:12px;margin-bottom:6px">当前全局档：<b>${_globalSlots.length ? esc(_globalSlots.map(s => s.label).join(' / ')) : '（无）'}</b></div>
                            <button class="btn btn-s btn-p" onclick="Admin.editGlobalSlots()">✏️ 编辑全局默认档</button>
                        </div>
                    </div>
                </div>

                <div class="cfg-sec${_cfgOpen.cmds ? ' open' : ''}" id="cfgSecCmds">
                    <div class="cfg-sec-hdr" onclick="Admin.toggleCfgSec('cmds')">
                        <span class="cs-caret">${_cfgOpen.cmds ? '▼' : '▶'}</span>
                        <span class="cs-title">⌨️ 快捷指令</span>
                        <span class="cs-sub">已配 ${cmdCount} 条</span>
                    </div>
                    <div class="cfg-sec-body">
                        <div class="fg"><label>指令列表<span class="cfg-help" title="用户在输入框打 / 即弹出面板，选中后内容自动填入">?</span></label>
                            <textarea id="cfg_quickCmds" rows="7" placeholder="每行一条，格式：名称|内容模板&#10;翻译成英文|请把以下内容翻译成地道的英文：">${esc(cmdsRaw)}</textarea>
                            <div class="ff-hint">格式：<code>名称|内容</code>，一行一条（内容不能换行）。</div>
                        </div>
                    </div>
                </div>

                <button class="btn btn-p" onclick="Admin.saveConfig()" style="margin-top:6px">💾 保存全部</button>
            </div>`;
        } catch (e) { box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>'; }
    }

    function toggleCfgSec(key) {
        _cfgOpen[key] = !_cfgOpen[key];
        const map = { chunk: 'cfgSecChunk', quick: 'cfgSecQuick', cmds: 'cfgSecCmds' };
        const el = document.getElementById(map[key]);
        if (el) {
            el.classList.toggle('open', _cfgOpen[key]);
            const c = el.querySelector('.cs-caret');
            if (c) c.textContent = _cfgOpen[key] ? '▼' : '▶';
        }
    }

    function editGlobalSlots() {
        SlotEditor.open({
            title: '⚡ 全局默认档位（兜底用）',
            slots: _globalSlots,
            onSave: (slots) => {
                _globalSlots = slots;
                SlotEditor.close();
                renderConfig(document.getElementById('adminBody'));
                toast('已记录 ' + slots.length + ' 档，别忘点「💾 保存全部」');
            }
        });
    }

    function saveConfig() {
        const chunkSize = document.getElementById('cfg_chunkSize').value;
        const quickCmds = document.getElementById('cfg_quickCmds').value;
        apiCall('admin/config/save', 'POST', {
            config: { chunkSize, quickModels: JSON.stringify(_globalSlots || []), quickCmds }
        }).then(() => {
            toast('✅ 已保存（用户刷新后生效）');
            if (typeof Chunker !== 'undefined') Chunker.setBlockSize(chunkSize);
            if (typeof window.applyGlobalConfig === 'function') {
                window.applyGlobalConfig({ quickModels: JSON.stringify(_globalSlots || []), quickCmds, chunkSize });
            }
        }).catch(e => toast('失败：' + e.message, 'er'));
    }

    /* ========================================================== */
    return {
        open, close, switchTab, apiCall,
        /* 账号 */
        showCreateUser, showEditUser, showResetPwd, showPerm, showSlots, toggleStatus, delUser, searchUsers,
        importXLSX, exportXLSX,
        /* 引擎 */
        showEngEdit, delEng,
        /* 引擎模板库 */
        showTplEdit, delTpl, syncTpl, doDeploy, toggleTplSel, toggleDeployUser, deploySelectAllUsers,
        /* 档位模板库 */
        newSlotTpl, editSlotTpl, renameSlotTpl, delSlotTpl, pickSlotTpl, toggleSlotDeployUser, slotDeployAll, doSlotDeploy,
        /* 档位编辑器 */
        slotEditorClose: SlotEditor.close, slotEditorSave: SlotEditor.save,
        slotEdAdd: SlotEditor.add, slotEdDel: SlotEditor.del, slotEdSet: SlotEditor.set,
        slotEdToggleAdv: SlotEditor.toggleAdv, slotEdSaveParams: SlotEditor.saveParams,
        /* 参数编辑区 */
        toggleParamSec, onParamToggle, applyParamPreset,
        /* 体检 */
        runHealth,
        /* 模型库 */
        showModelEdit, delModel,
        /* 预设 */
        addPreset, dupPreset, delPreset, selectPreset, filterPresetList,
        updP, updS, updSeg, updSegPrompt,
        toggleStep, toggleStepSel, stepsToNewPreset,
        addStep, dupStep, delStep, moveStep, addSeg, delSeg, backToEditor,
        showStepModel,
        selVariant, addVariant, renameVariant, delVariant, renameDefaultVar,
        stepDragStart, stepDragOver, stepDragLeave, stepDrop, stepDragEnd,
        manageGroups, addGroup, renameGroup, delGroup,
        showSecurity, applySecurity, savePresets, exportPresetsJSON, importPresetsJSON,
        /* 全局设置 */
        saveConfig, toggleCfgSec, editGlobalSlots,
        FF,
    };
})();

window.Admin = Admin;
