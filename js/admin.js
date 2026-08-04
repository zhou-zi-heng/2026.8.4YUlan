/* ===== 飞凡AI - 超管后台 (v3.3.0 批次3：表单化+模板库+体检) ===== */

const Admin = (function () {

    /* ---------- 模块级状态（全部提到顶部，避免 TDZ） ---------- */
    let _curTab = 'users';
    let _presetData = null;
    let _curPresetIdx = -1;
    let _openSteps = {};
    let _presetDirty = false;
    let _usersCache = [];
    let _stepSel = {};
    let _curVar = {};
    let _quickModelDraft = [];
    let _dragStepIdx = null;

    /* 批次3 新增状态 */
    let _tplCache = [];        // 模板列表
    let _tplSel = {};          // 勾选的模板 {tplId:true}
    let _deployUserSel = {};   // 勾选的账号 {username:true}
    let _healthRows = [];      // 体检结果

    const DEFAULT_VID = '__default__';

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

    /* ========================================================== */
    /* ============ ★ 通用表单弹窗组件 FF ======================= */
    /* ==========================================================
       用法：
       FF.open({
         title:'新增账号',
         tip:'说明文字（可选）',
         fields:[
           {key:'name',  label:'姓名',  type:'text', value:'', placeholder:'', hint:'小字提示'},
           {key:'role',  label:'角色',  type:'select', value:'user', options:[{v:'user',t:'普通'},{v:'admin',t:'管理员'}]},
           {key:'cache', label:'开缓存', type:'check', value:true},
           {key:'grp',   label:'分组',  type:'checks', value:['运营'], options:[{v:'运营',t:'运营'}]},
           {key:'note',  label:'备注',  type:'textarea', rows:4},
         ],
         okText:'保存',
         onSubmit: async (vals, setErr) => { ...; return true; }  // 返回 false 则不关闭
       });
       ========================================================== */
    const FF = (function () {
        let _mask = null;
        let _cfg = null;

        function close() {
            if (_mask) { _mask.remove(); _mask = null; }
            _cfg = null;
            document.removeEventListener('keydown', _onKey);
        }

        function _onKey(e) {
            if (!_mask) return;
            if (e.key === 'Escape') { e.preventDefault(); close(); }
        }

        function _fieldHtml(f, i) {
            const id = 'ff_' + i;
            const label = esc(f.label || '');
            const hint = f.hint ? '<div class="ff-hint">' + esc(f.hint) + '</div>' : '';

            if (f.type === 'check') {
                return '<div class="fg"><label class="pt" style="margin:0">'
                    + '<input type="checkbox" id="' + id + '" ' + (f.value ? 'checked' : '') + '> '
                    + label + '</label>' + hint + '</div>';
            }

            if (f.type === 'checks') {
                const vals = Array.isArray(f.value) ? f.value : [];
                let box = '<div class="fg"><label>' + label + '</label><div class="ff-checks" id="' + id + '">';
                (f.options || []).forEach((o, k) => {
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
                let sel = '<div class="fg"><label>' + label + '</label><select id="' + id + '">';
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

            /* text / password / number */
            const t = f.type || 'text';
            return '<div class="fg"><label>' + label + '</label>'
                + '<input type="' + t + '" id="' + id + '" value="' + esc(f.value != null ? f.value : '') + '"'
                + ' placeholder="' + esc(f.placeholder || '') + '"'
                + (t === 'number' ? ' step="' + (f.step || 'any') + '"' : '')
                + (f.autocomplete ? ' autocomplete="' + esc(f.autocomplete) + '"' : ' autocomplete="off"')
                + '>' + hint + '</div>';
        }

        function _collect() {
            const vals = {};
            (_cfg.fields || []).forEach((f, i) => {
                const el = document.getElementById('ff_' + i);
                if (!el) return;
                if (f.type === 'check') { vals[f.key] = el.checked; return; }
                if (f.type === 'checks') {
                    const arr = [];
                    el.querySelectorAll('input[type="checkbox"]').forEach(cb => { if (cb.checked) arr.push(cb.getAttribute('data-v')); });
                    vals[f.key] = arr;
                    return;
                }
                if (f.type === 'number') { vals[f.key] = parseFloat(el.value) || 0; return; }
                vals[f.key] = el.value;
            });
            return vals;
        }

        function setErr(msg) {
            const e = document.getElementById('ffErr');
            if (e) e.textContent = msg || '';
        }

        async function _submit() {
            const btn = document.getElementById('ffOk');
            const vals = _collect();
            setErr('');
            if (btn) { btn.disabled = true; btn.textContent = '处理中...'; }
            try {
                const r = await _cfg.onSubmit(vals, setErr);
                if (r !== false) close();
            } catch (e) {
                setErr(e.message || '操作失败');
            }
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

            _mask.innerHTML = '<div class="ff-box">'
                + '<div class="ff-hdr"><h3>' + esc(_cfg.title || '') + '</h3><button class="ff-x" id="ffX">×</button></div>'
                + '<div class="ff-body">' + body + '</div>'
                + '<div class="ff-ft"><button class="btn btn-s" id="ffCancel">取消</button>'
                + '<button class="btn btn-p btn-s" id="ffOk">' + esc(_cfg.okText || '保存') + '</button></div>'
                + '</div>';

            document.body.appendChild(_mask);

            document.getElementById('ffX').onclick = close;
            document.getElementById('ffCancel').onclick = close;
            document.getElementById('ffOk').onclick = _submit;
            _mask.onclick = (e) => { if (e.target === _mask) close(); };
            document.addEventListener('keydown', _onKey);

            /* 首个可输入项自动聚焦 */
            setTimeout(() => {
                const first = _mask.querySelector('.ff-body input[type="text"],.ff-body input[type="password"],.ff-body textarea,.ff-body select');
                if (first) first.focus();
            }, 60);

            /* 单行输入框按 Enter 直接提交 */
            _mask.querySelectorAll('.ff-body input[type="text"],.ff-body input[type="password"],.ff-body input[type="number"]').forEach(el => {
                el.onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); _submit(); } };
            });
        }

        return { open, close, setErr };
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
        FF.close();
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

    /* ========================================================== */
    /* ===================== 账号管理 =========================== */
    /* ========================================================== */
    async function renderUsers(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>';
        try { const data = await apiCall('admin/users/list'); _usersCache = data.users || []; drawUsersTable(box, ''); }
        catch (e) { box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>'; }
    }

    function drawUsersTable(box, kw) {
        kw = (kw || '').toLowerCase();
        let users = _usersCache;
        if (kw) users = users.filter(u => (u.username + ' ' + (u.name || '')).toLowerCase().includes(kw));
        let html = `<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px;align-items:center">
            <button class="btn btn-p btn-s" onclick="Admin.showCreateUser()">➕ 新增账号</button>
            <label class="btn btn-s" style="cursor:pointer">📥 xlsx导入<input type="file" accept=".xlsx,.xls" onchange="Admin.importXLSX(this)" style="display:none"></label>
            <button class="btn btn-s" onclick="Admin.exportXLSX(false)">📤 导出(脱敏)</button>
            <button class="btn btn-s btn-d" onclick="Admin.exportXLSX(true)">📤 导出(含Key)</button>
            <button class="btn btn-s" onclick="Admin.downloadTemplate()">📋 模板</button>
            <button class="btn btn-s" onclick="Admin.switchTab('users')">🔄 刷新</button>
            <input type="text" placeholder="🔍 搜索姓名/账号" oninput="Admin.searchUsers(this.value)" style="margin-left:auto;padding:5px 10px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);font-size:12px;width:180px"></div>
            <div style="font-size:12px;color:var(--text2);margin-bottom:8px">共 ${_usersCache.length} 个账号${kw ? '，匹配 ' + users.length : ''}。🔴=最近7天≥3个IP。</div>
            <div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>姓名</th><th>账号</th><th>角色</th><th>状态</th><th>引擎</th><th>可用分组</th><th>最后活跃</th><th>IP</th><th>操作</th></tr></thead><tbody>`;
        users.forEach(u => {
            let permTxt = '全部';
            try { const p = JSON.parse(u.permissions || '{}'); if (p.allowGroups && p.allowGroups.length) permTxt = p.allowGroups.join('/'); } catch (e) {}
            html += `<tr><td>${esc(u.name || '-')}</td><td>${esc(u.username)}</td><td>${u.role === 'admin' ? '👑' : '普通'}</td><td>${u.status === 'active' ? '<span style="color:#10b981">启用</span>' : '<span style="color:#ef4444">禁用</span>'}</td><td>${u.engineCount}</td><td style="font-size:11px;max-width:120px;overflow:hidden;text-overflow:ellipsis" title="${esc(permTxt)}">${esc(permTxt)}</td><td style="font-size:11px">${fmtTime(u.lastActive)}</td><td>${u.ipAbnormal ? '<span style="color:#ef4444;font-weight:bold">🔴' + u.ipCount + '</span>' : (u.ipCount || 0)}</td><td class="admin-ops"><button onclick='Admin.showEditUser(${JSON.stringify(u.username)})' title="编辑资料">✏️</button><button onclick='Admin.showPerm(${JSON.stringify(u.username)})' title="工作流权限">🎫</button><button onclick='Admin.showResetPwd(${JSON.stringify(u.username)})' title="改密">🔑</button><button onclick='Admin.toggleStatus(${JSON.stringify(u.username)},${JSON.stringify(u.status)})' title="启用/禁用">${u.status === 'active' ? '🚫' : '✅'}</button>${u.username !== 'admin' ? `<button onclick='Admin.delUser(${JSON.stringify(u.username)})' style="color:#ef4444" title="删除">🗑️</button>` : ''}</td></tr>`;
        });
        html += '</tbody></table></div>';
        box.innerHTML = html;
    }

    function searchUsers(kw) { drawUsersTable(document.getElementById('adminBody'), kw); }

    /* ★ 新增账号：一屏表单 */
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
                await apiCall('admin/users/create', 'POST', {
                    username: v.username.trim(), password: v.password.trim(),
                    name: (v.name || '').trim(), role: v.role
                });
                toast('✅ 已创建账号 ' + v.username.trim());
                switchTab('users');
            }
        });
    }

    /* ★ 编辑账号资料（姓名/角色/状态） */
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
                toast('✅ 已保存');
                switchTab('users');
            }
        });
    }

    /* ★ 改密：表单 */
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

    /* ★ 工作流权限：分组勾选（不再手打逗号） */
    function showPerm(username) {
        const u = _usersCache.find(x => x.username === username);
        let perm = {};
        try { perm = JSON.parse((u && u.permissions) || '{}'); } catch (e) {}
        const cur = perm.allowGroups || [];
        const groups = (typeof Workflow !== 'undefined' && Workflow.isLoaded()) ? Workflow.getGroups() : [];

        FF.open({
            title: '🎫 工作流权限 — ' + username,
            tip: '<b>全不勾 = 允许全部分组</b>（默认）。勾选后只能用勾中的分组。修改后用户端<b>实时生效</b>，无需重新登录。',
            okText: '保存权限',
            fields: [
                {
                    key: 'groups', label: '可用工作流分组', type: 'checks', value: cur,
                    options: groups.map(g => ({ v: g, t: g })),
                    hint: groups.length ? '' : '⚠️ 未读到分组，请先到「📋 预设 → 🏷️ 分组管理」创建分组'
                }
            ],
            onSubmit: async (v) => {
                await apiCall('admin/users/perm', 'POST', {
                    username, permissions: Object.assign({}, perm, { allowGroups: v.groups })
                });
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

    function importXLSX(inputEl) { const file = inputEl.files && inputEl.files[0]; if (!file) return; loadXLSX().then(() => { const reader = new FileReader(); reader.onload = async (e) => { try { const wb = XLSX.read(e.target.result, { type: 'array' }); const sheet = wb.Sheets[wb.SheetNames[0]]; const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' }); if (!rows.length) { toast('表格无数据', 'er'); inputEl.value = ''; return; } if (!confirm('导入 ' + rows.length + ' 行（同账号覆盖+重配引擎）？')) { inputEl.value = ''; return; } toast('导入中...'); const res = await apiCall('admin/users/import', 'POST', { rows }); let msg = '✅ 账号 ' + res.userCount + '，引擎 ' + res.engCount; if (res.errors && res.errors.length) msg += '\n⚠️ ' + res.errors.join('；'); alert(msg); switchTab('users'); } catch (err) { toast('导入失败：' + err.message, 'er'); } inputEl.value = ''; }; reader.readAsArrayBuffer(file); }).catch(e => toast('加载解析库失败：' + e.message, 'er')); }
    async function exportXLSX(withKey) { if (withKey && !confirm('⚠️ 导出含明文Key，妥善保管！继续？')) return; try { await loadXLSX(); const res = await apiCall('admin/users/export?withkey=' + (withKey ? '1' : '0')); const ws = XLSX.utils.json_to_sheet(res.rows || [], { header: ['姓名', '账号', '密码', '角色', '引擎名称', '协议', 'BaseURL', 'APIKey', '模型', '输入单价', '输出单价', '缓存读单价', '缓存写单价'] }); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '账号'); XLSX.writeFile(wb, 'feifan-accounts-' + (withKey ? 'withkey-' : '') + new Date().toISOString().slice(0, 10) + '.xlsx'); toast('✅ 已导出'); } catch (e) { toast('导出失败：' + e.message, 'er'); } }
    async function downloadTemplate() { try { await loadXLSX(); const rows = [{ 姓名: '张三', 账号: 'zhangsan', 密码: 'pass123', 角色: 'user', 引擎名称: '快速引擎', 协议: 'openai', BaseURL: 'https://api.openai-proxy.org/v1', APIKey: 'sk-xxx', 模型: '', 输入单价: '', 输出单价: '', 缓存读单价: '', 缓存写单价: '' }, { 姓名: '李四', 账号: 'lisi', 密码: 'pass456', 角色: 'user', 引擎名称: '便宜', 协议: 'openai', BaseURL: 'https://api.openai-proxy.org/v1', APIKey: 'sk-ds', 模型: '', 输入单价: '', 输出单价: '', 缓存读单价: '', 缓存写单价: '' }]; const ws = XLSX.utils.json_to_sheet(rows, { header: ['姓名', '账号', '密码', '角色', '引擎名称', '协议', 'BaseURL', 'APIKey', '模型', '输入单价', '输出单价', '缓存读单价', '缓存写单价'] }); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '账号'); XLSX.writeFile(wb, 'feifan-账号导入模板.xlsx'); toast('✅ 模板已下载'); } catch (e) { toast('生成模板失败：' + e.message, 'er'); } }

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

            const byUser = {};
            engs.forEach(e => { if (!byUser[e.username]) byUser[e.username] = []; byUser[e.username].push(e); });

            let html = `<div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap;align-items:center">
                <button class="btn btn-p btn-s" onclick="Admin.switchTab('templates')">📦 引擎模板库（批量下发）</button>
                <button class="btn btn-s" onclick="Admin.switchTab('health')">🔍 引擎体检</button>
                <button class="btn btn-s" onclick="Admin.switchTab('engines')">🔄 刷新</button>
                <span style="font-size:11px;color:var(--text2);margin-left:auto">🎨生图 · 💰缓存 · 🔗来自模板</span>
            </div>
            <div style="font-size:12px;color:var(--text2);margin-bottom:10px">给账号配公有引擎。人多时建议用<b>模板库批量下发</b>，不用一个个点。</div>`;

            users.forEach(u => {
                const ue = byUser[u.username] || [];
                html += `<div class="eng-user-block"><div class="eng-user-hdr"><b>${esc(u.name || u.username)}</b> <span style="color:var(--text2);font-size:11px">(${esc(u.username)})</span><span style="font-size:11px;color:var(--text2);margin-left:6px">${ue.length} 个引擎</span><button class="btn btn-p btn-s" style="margin-left:auto" onclick='Admin.showEngEdit(${JSON.stringify(u.username)},"")'>➕ 加引擎</button></div>`;
                if (!ue.length) html += '<div style="font-size:11px;color:var(--text2);padding:4px 0">（无引擎）</div>';
                ue.forEach(e => {
                    const imgTag = (e.engineType === 'image') ? '<span style="color:#f59e0b">🎨</span>' : '';
                    const cacheTag = e.useCache ? '<span style="color:#10b981">💰</span>' : '';
                    const tplTag = e.tplId ? '<span style="color:var(--pri)" title="来自模板，可在模板库一键同步">🔗</span>' : '';
                    html += `<div class="eng-item"><span>📦 <b>${esc(e.name)}</b> ${imgTag}${cacheTag}${tplTag} <span style="color:var(--text2);font-size:11px">${esc(e.protocol)} / ${esc(e.model || '用户自选')}</span></span>
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

    /* ★ 引擎编辑：表单弹窗 */
    function showEngEdit(username, engId) {
        const box = document.getElementById('adminBody');
        const engs = (box && box._engs) || [];
        const e = engId ? engs.find(x => x.id === engId) : null;

        FF.open({
            title: (e ? '✏️ 编辑引擎' : '➕ 新增引擎') + ' — ' + username,
            tip: e && e.tplId
                ? '⚠️ 此引擎来自<b>模板</b>。手动改动后会<b>脱离模板管理</b>（模板同步不再覆盖它）。'
                : '💰 缓存：Claude 等模型开启后重复内容省钱。<br>🎨 生图：选后用户发消息即出图。',
            okText: '保存',
            fields: [
                { key: 'name', label: '引擎名称', type: 'text', value: e ? e.name : '', placeholder: '如：快速引擎', hint: '用户端看到的名字' },
                { key: 'engineType', label: '引擎类型', type: 'select', value: e ? (e.engineType || 'chat') : 'chat', options: [{ v: 'chat', t: '💬 对话' }, { v: 'image', t: '🎨 生图 / 改图' }] },
                { key: 'protocol', label: '协议', type: 'select', value: e ? e.protocol : 'openai', options: [{ v: 'openai', t: 'OpenAI / 通用' }, { v: 'anthropic', t: 'Claude 原生' }, { v: 'gemini', t: 'Gemini 原生' }] },
                { key: 'base', label: 'Base URL', type: 'text', value: e ? e.base : 'https://api.openai-proxy.org/v1', placeholder: 'https://...' },
                { key: 'key', label: 'API Key', type: 'password', value: '', placeholder: e ? '••••（留空=不修改）' : 'sk-...', hint: e ? '留空表示不修改现有 Key' : '' },
                { key: 'model', label: '默认模型', type: 'text', value: e ? (e.model || '') : '', placeholder: '对话如 gpt-4o；生图如 gpt-image-1', hint: '留空则由用户在前台自己获取选择' },
                { key: 'useCache', label: '💰 开启 Prompt 缓存', type: 'check', value: e ? !!e.useCache : false, hint: '生图引擎无需开启' },
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
                    priceIn: v.priceIn, priceOut: v.priceOut, priceCR: v.priceCR, priceCW: v.priceCW,
                    tplId: ''   /* 手动编辑 = 脱离模板管理 */
                });
                toast('✅ 已保存');
                switchTab('engines');
            }
        });
    }

    function delEng(id, name) {
        if (!confirm('删除引擎【' + (name || '') + '】？\n该用户将无法再使用它。')) return;
        apiCall('admin/engines/delete', 'POST', { id })
            .then(() => { toast('✅ 已删除'); switchTab('engines'); })
            .catch(e => toast('失败：' + e.message, 'er'));
    }

    /* ========================================================== */
    /* ============ ★ 引擎模板库（批量下发） ==================== */
    /* ========================================================== */
    async function renderTemplates(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>';
        try {
            const tData = await apiCall('admin/templates/list');
            _tplCache = tData.templates || [];
            if (!_usersCache.length) {
                const uData = await apiCall('admin/users/list');
                _usersCache = uData.users || [];
            }
            drawTemplates(box);
        } catch (e) {
            box.innerHTML = '<div style="color:#ef4444;padding:20px;line-height:1.8">加载失败：' + esc(e.message)
                + '<br><br>如提示"模板表不存在"，请先到 Cloudflare → D1 执行建表 SQL。</div>';
        }
    }

    function drawTemplates(box) {
        const selCount = Object.keys(_tplSel).filter(k => _tplSel[k]).length;

        let html = `<div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap;align-items:center">
            <button class="btn btn-p btn-s" onclick="Admin.showTplEdit('')">➕ 新建模板</button>
            <button class="btn btn-s" onclick="Admin.switchTab('templates')">🔄 刷新</button>
            <button class="btn btn-s" onclick="Admin.switchTab('engines')">← 返回引擎列表</button>
        </div>
        <div style="font-size:12px;color:var(--text2);margin-bottom:12px;line-height:1.7">
            💡 <b>用法</b>：先建好模板（含 Key）→ 勾选模板 → 勾选账号 → 一键下发。<br>
            🔄 模板改了 Key/模型后，点模板上的「同步」即可刷新所有已下发的引擎，不用逐个改。
        </div>`;

        if (!_tplCache.length) {
            html += '<div style="padding:30px;text-align:center;color:var(--text2);border:1px dashed var(--border);border-radius:10px">还没有模板，点上方「➕ 新建模板」开始</div>';
            box.innerHTML = html;
            return;
        }

        html += '<div class="tpl-grid">';
        _tplCache.forEach(t => {
            const on = !!_tplSel[t.id];
            html += `<div class="tpl-card${on ? ' sel' : ''}">
                <div class="tpl-card-top">
                    <input type="checkbox" ${on ? 'checked' : ''} onchange='Admin.toggleTplSel(${JSON.stringify(t.id)},this.checked)'>
                    <span class="tpl-name">${esc(t.name)}</span>
                </div>
                <div class="tpl-meta">${esc(t.protocol)} · ${esc(t.model || '用户自选')}<br>${esc((t.base || '').replace(/^https?:\/\//, '').slice(0, 34))}</div>
                <div class="tpl-badges">
                    ${t.engineType === 'image' ? '<span class="tpl-badge img">🎨生图</span>' : '<span class="tpl-badge">💬对话</span>'}
                    ${t.useCache ? '<span class="tpl-badge cache">💰缓存</span>' : ''}
                    ${t.hasKey ? '<span class="tpl-badge">🔑已配Key</span>' : '<span class="tpl-badge nokey">⚠️无Key</span>'}
                    ${t.deployed ? '<span class="tpl-badge deployed">已下发 ' + t.deployed + '</span>' : ''}
                </div>
                <div class="tpl-acts">
                    <button class="btn btn-s" onclick='Admin.showTplEdit(${JSON.stringify(t.id)})'>✏️改</button>
                    ${t.deployed ? `<button class="btn btn-s" onclick='Admin.syncTpl(${JSON.stringify(t.id)},${JSON.stringify(t.name)},${t.deployed})' title="把模板最新内容推给已下发的引擎">🔄同步</button>` : ''}
                    <button class="btn btn-s btn-d" onclick='Admin.delTpl(${JSON.stringify(t.id)},${JSON.stringify(t.name)},${t.deployed})'>🗑️</button>
                </div>
            </div>`;
        });
        html += '</div>';

        /* 下发面板 */
        html += `<div class="deploy-panel">
            <div style="font-size:13px;font-weight:600;margin-bottom:4px">🚀 批量下发到账号</div>
            <div style="font-size:11px;color:var(--text2)">已选 <b style="color:var(--pri)">${selCount}</b> 个模板。下面勾选要下发的账号：</div>
            <div class="deploy-users" id="deployUsers"></div>
            <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
                <button class="btn btn-s" onclick="Admin.deploySelectAllUsers(true)">全选账号</button>
                <button class="btn btn-s" onclick="Admin.deploySelectAllUsers(false)">全不选</button>
                <select id="deployMode" style="padding:5px 8px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);font-size:12px">
                    <option value="skip">同名引擎：跳过（安全，推荐）</option>
                    <option value="replace">同名引擎：覆盖（更新Key/模型用）</option>
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
            html += `<label class="ff-chk${on ? ' on' : ''}">
                <input type="checkbox" ${on ? 'checked' : ''} onchange='Admin.toggleDeployUser(${JSON.stringify(u.username)},this.checked)'>
                ${esc(u.name || u.username)} <span style="opacity:.6;font-size:10px">${esc(u.username)}</span>
            </label>`;
        });
        wrap.innerHTML = html;
    }

    function toggleTplSel(id, on) {
        _tplSel[id] = on;
        drawTemplates(document.getElementById('adminBody'));
    }
    function toggleDeployUser(un, on) {
        _deployUserSel[un] = on;
        drawDeployUsers();
        /* 同步 label 高亮 */
        const wrap = document.getElementById('deployUsers');
        if (wrap) drawDeployUsers();
    }
    function deploySelectAllUsers(val) {
        _usersCache.forEach(u => { _deployUserSel[u.username] = val; });
        drawDeployUsers();
    }

    function showTplEdit(tplId) {
        const t = tplId ? _tplCache.find(x => x.id === tplId) : null;
        FF.open({
            title: t ? ('✏️ 编辑模板 — ' + t.name) : '➕ 新建引擎模板',
            tip: '模板是"引擎的配方"，存好后可批量下发给多个账号。<br>模板名会成为用户端看到的引擎名。',
            okText: '保存模板',
            fields: [
                { key: 'name', label: '模板 / 引擎名称', type: 'text', value: t ? t.name : '', placeholder: '如：快速引擎', hint: '下发后用户看到的就是这个名字' },
                { key: 'engineType', label: '引擎类型', type: 'select', value: t ? (t.engineType || 'chat') : 'chat', options: [{ v: 'chat', t: '💬 对话' }, { v: 'image', t: '🎨 生图 / 改图' }] },
                { key: 'protocol', label: '协议', type: 'select', value: t ? t.protocol : 'openai', options: [{ v: 'openai', t: 'OpenAI / 通用' }, { v: 'anthropic', t: 'Claude 原生' }, { v: 'gemini', t: 'Gemini 原生' }] },
                { key: 'base', label: 'Base URL', type: 'text', value: t ? t.base : 'https://api.openai-proxy.org/v1' },
                { key: 'key', label: 'API Key', type: 'password', value: '', placeholder: t ? (t.hasKey ? '••••（留空=不修改）' : 'sk-...') : 'sk-...', hint: t && t.hasKey ? '留空表示不修改现有 Key' : '' },
                { key: 'model', label: '默认模型', type: 'text', value: t ? (t.model || '') : '', placeholder: '如 gpt-4o / gpt-image-1', hint: '留空则用户自己选' },
                { key: 'useCache', label: '💰 开启 Prompt 缓存', type: 'check', value: t ? !!t.useCache : false },
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
                    priceIn: v.priceIn, priceOut: v.priceOut, priceCR: v.priceCR, priceCW: v.priceCW
                });
                toast('✅ 模板已保存');
                switchTab('templates');
            }
        });
    }

    function delTpl(id, name, deployed) {
        let msg = '删除模板【' + name + '】？';
        if (deployed) msg += '\n\n已下发的 ' + deployed + ' 个引擎【不会被删除】，只会解除与模板的关联（之后不能一键同步）。';
        if (!confirm(msg)) return;
        apiCall('admin/templates/delete', 'POST', { id })
            .then(() => { delete _tplSel[id]; toast('✅ 已删除模板'); switchTab('templates'); })
            .catch(e => toast('失败：' + e.message, 'er'));
    }

    function syncTpl(id, name, deployed) {
        FF.open({
            title: '🔄 同步模板 — ' + name,
            tip: '将把此模板的 <b>Key / BaseURL / 模型 / 协议 / 缓存开关 / 单价</b> 推给所有由它下发的引擎（共 <b>' + deployed + '</b> 个）。<br>适合换 Key、换中转站、统一改模型时用。',
            okText: '确认同步',
            fields: [
                { key: 'syncName', label: '同时覆盖引擎名称', type: 'check', value: false, hint: '不勾选=保留用户端现有引擎名（推荐，避免名字突变让用户困惑）' }
            ],
            onSubmit: async (v) => {
                const r = await apiCall('admin/templates/sync', 'POST', { templateId: id, syncName: v.syncName });
                toast('✅ 已同步 ' + (r.count || 0) + " 个引擎");
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
        const total = tplIds.length * usernames.length;

        if (!confirm('确认下发？\n\n模板 ' + tplIds.length + ' 个 × 账号 ' + usernames.length + ' 个 = 最多 ' + total + ' 条引擎记录\n\n同名策略：' + (mode === 'replace' ? '覆盖' : '跳过'))) return;

        toast('下发中...');
        try {
            const r = await apiCall('admin/templates/deploy', 'POST', { templateIds: tplIds, usernames, mode });
            let msg = '✅ 下发完成\n\n新建：' + (r.created || 0) + '\n覆盖：' + (r.updated || 0) + '\n跳过（同名已存在）：' + (r.skipped || 0);
            if (r.errors && r.errors.length) msg += '\n\n⚠️ 失败 ' + r.errors.length + ' 条：\n' + r.errors.slice(0, 6).join('\n');
            alert(msg);
            _deployUserSel = {};
            switchTab('templates');
        } catch (e) {
            toast('下发失败：' + e.message, 'er');
        }
    }

    /* ========================================================== */
    /* ============ ★ 引擎体检 ================================== */
    /* ========================================================== */
    async function renderHealth(box) {
        box.innerHTML = `<div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap;align-items:center">
                <button class="btn btn-p btn-s" onclick="Admin.runHealth()">🔍 开始体检（测所有引擎）</button>
                <button class="btn btn-s" onclick="Admin.switchTab('engines')">← 返回引擎列表</button>
            </div>
            <div style="font-size:12px;color:var(--text2);margin-bottom:12px;line-height:1.7">
                逐个请求各引擎的模型列表接口，验证 <b>BaseURL / Key / 权限 / 余额 / 网络</b> 是否正常，并检查<b>默认模型名是否真的存在</b>。<br>
                每个引擎最多等 12 秒；引擎多时分批进行，耐心等一下。
            </div>
            <div id="healthArea"><div style="padding:26px;text-align:center;color:var(--text2);border:1px dashed var(--border);border-radius:10px">点上方按钮开始</div></div>`;
    }

    async function runHealth() {
        const area = document.getElementById('healthArea');
        if (!area) return;
        area.innerHTML = '<div style="color:var(--text2);padding:14px;text-align:center">正在读取引擎列表...</div>';

        try {
            const engData = await apiCall('admin/engines/list');
            const engs = engData.engines || [];
            if (!engs.length) {
                area.innerHTML = '<div style="padding:20px;text-align:center;color:var(--text2)">还没有任何公有引擎</div>';
                return;
            }

            const ids = engs.map(e => e.id);
            const BATCH = 10;
            const batches = [];
            for (let i = 0; i < ids.length; i += BATCH) batches.push(ids.slice(i, i + BATCH));

            _healthRows = [];
            let done = 0;

            const drawProgress = () => {
                const pct = Math.round(done / ids.length * 100);
                area.innerHTML = '<div style="font-size:13px;margin-bottom:6px">正在体检 <b>' + done + ' / ' + ids.length + '</b> ...</div>'
                    + '<div class="health-bar"><div class="health-bar-in" style="width:' + pct + '%"></div></div>'
                    + drawHealthTable();
            };
            drawProgress();

            for (const b of batches) {
                try {
                    const r = await apiCall('admin/engines/health', 'POST', { ids: b });
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

            area.innerHTML = `<div class="health-sum">
                    <div class="health-card ok"><b>${okN}</b>正常</div>
                    <div class="health-card bad"><b>${badN}</b>异常</div>
                    <div class="health-card"><b>${_healthRows.length}</b>总计</div>
                </div>` + drawHealthTable();

            if (badN) toast('⚠️ 体检完成：' + badN + ' 个引擎异常', 'er');
            else toast('✅ 体检完成：全部正常');

        } catch (e) {
            area.innerHTML = '<div style="color:#ef4444;padding:20px">体检失败：' + esc(e.message) + '</div>';
        }
    }

    function drawHealthTable() {
        if (!_healthRows.length) return '';
        /* 异常排前面 */
        const rows = _healthRows.slice().sort((a, b) => (a.ok === b.ok) ? 0 : (a.ok ? 1 : -1));
        let html = '<div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>状态</th><th>账号</th><th>引擎</th><th>模型校验</th><th>延迟</th><th>说明</th></tr></thead><tbody>';
        rows.forEach(r => {
            let mdl = '-';
            if (r.ok) {
                if (r.modelOk === true) mdl = '<span style="color:#10b981">✅ 存在</span>';
                else if (r.modelOk === false) mdl = '<span style="color:#f59e0b">⚠️ 列表中找不到<br><span style="font-size:10px">' + esc(r.model || '') + '</span></span>';
                else mdl = '<span style="color:var(--text2);font-size:11px">' + (r.model ? '未校验' : '未设默认模型') + '</span>';
            }
            html += '<tr class="' + (r.ok ? 'health-row-ok' : 'health-row-bad') + '">'
                + '<td>' + (r.ok ? '🟢 正常' : '🔴 异常') + '</td>'
                + '<td>' + esc(r.username || '') + '</td>'
                + '<td>' + esc(r.name || '') + (r.engineType === 'image' ? ' 🎨' : '') + '</td>'
                + '<td>' + mdl + '</td>'
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
            let html = `<div style="display:flex;gap:8px;margin-bottom:12px"><button class="btn btn-p btn-s" onclick="Admin.showModelEdit('')">➕ 新增模型</button><button class="btn btn-s" onclick="Admin.switchTab('models')">🔄 刷新</button></div><div style="font-size:12px;color:var(--text2);margin-bottom:8px">模型库：配各模型单价（美元/1M token）。引擎未指定单价时自动查这里。</div><div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>模型名</th><th>输入</th><th>输出</th><th>缓存读</th><th>缓存写</th><th>操作</th></tr></thead><tbody>`;
            models.forEach(m => {
                html += `<tr><td>${esc(m.model_name)}</td><td>${m.price_in}</td><td>${m.price_out}</td><td>${m.price_cache_read}</td><td>${m.price_cache_write}</td><td class="admin-ops"><button onclick='Admin.showModelEdit(${JSON.stringify(m.model_name)})'>✏️</button><button onclick='Admin.delModel(${JSON.stringify(m.model_name)})' style="color:#ef4444">🗑️</button></td></tr>`;
            });
            html += '</tbody></table></div>';
            box.innerHTML = html;
            box._models = models;
        } catch (e) {
            box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>';
        }
    }

    function showModelEdit(modelName) {
        const box = document.getElementById('adminBody');
        const models = (box && box._models) || [];
        const m = modelName ? models.find(x => x.model_name === modelName) : null;
        FF.open({
            title: m ? ('✏️ 编辑单价 — ' + m.model_name) : '➕ 新增模型单价',
            tip: '单位：<b>美元 / 100万 token</b>。引擎里没填单价时，系统按模型名来这里查。',
            okText: '保存',
            fields: [
                m ? { key: '_ro', label: '模型名（不可改）', type: 'text', value: m.model_name, hint: '如需改名请删除后重建' }
                  : { key: 'model_name', label: '模型名', type: 'text', value: '', placeholder: '如 gpt-4o', hint: '必须与引擎里填的模型名完全一致' },
                { key: 'priceIn', label: '输入单价', type: 'number', value: m ? m.price_in : 0, step: '0.01' },
                { key: 'priceOut', label: '输出单价', type: 'number', value: m ? m.price_out : 0, step: '0.01' },
                { key: 'priceCR', label: '缓存读单价', type: 'number', value: m ? m.price_cache_read : 0, step: '0.01' },
                { key: 'priceCW', label: '缓存写单价', type: 'number', value: m ? m.price_cache_write : 0, step: '0.01' },
            ],
            onSubmit: async (v, setErr) => {
                const name = m ? m.model_name : (v.model_name || '').trim();
                if (!name) { setErr('模型名必填'); return false; }
                await apiCall('admin/models/save', 'POST', {
                    model_name: name, priceIn: v.priceIn, priceOut: v.priceOut, priceCR: v.priceCR, priceCW: v.priceCW
                });
                toast('✅ 已保存');
                switchTab('models');
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
    /* ==================== 预设管理（含分支） =================== */
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

            /* 递归解密：默认版 + 全部分支 */
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
        } catch (e) {
            box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>';
        }
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
        drawPresetList();
        drawPresetEditor();
    }

    function drawPresetList() {
        const scroll = document.getElementById('presetListScroll');
        if (!scroll) return;
        const kw = ((document.getElementById('presetSearch') || {}).value || '').toLowerCase();
        let html = '';
        (_presetData.presets || []).forEach((p, i) => {
            if (kw && !(p.name || '').toLowerCase().includes(kw)) return;
            const hasBranch = (p.steps || []).some(s => Array.isArray(s.variants) && s.variants.length);
            html += `<div class="preset-litem${i === _curPresetIdx ? ' act' : ''}" onclick="Admin.selectPreset(${i})"><div class="pl-name">${esc(p.name || '未命名')}</div><div class="pl-meta">${esc(p.group || '无分组')} · ${(p.steps || []).length}步${hasBranch ? ' · 含分支' : ''}</div></div>`;
        });
        scroll.innerHTML = html || '<div style="font-size:12px;color:var(--text2);padding:10px;text-align:center">无预设</div>';
    }

    function filterPresetList() { drawPresetList(); }
    function selectPreset(i) { _curPresetIdx = i; _openSteps = {}; _stepSel = {}; _curVar = {}; drawPresetList(); drawPresetEditor(); }
    function markDirty() { _presetDirty = true; const t = document.getElementById('presetDirtyTip'); if (t) t.textContent = '● 有未保存修改'; }

    function _vidOf(si) { return _curVar[si] || DEFAULT_VID; }

    function _segsRef(si, vid) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        if (!vid || vid === DEFAULT_VID) {
            if (!Array.isArray(s.segments)) s.segments = [];
            return s.segments;
        }
        const v = (s.variants || []).find(x => x.id === vid);
        if (!v) {
            if (!Array.isArray(s.segments)) s.segments = [];
            return s.segments;
        }
        if (!Array.isArray(v.segments)) v.segments = [];
        return v.segments;
    }

    function _segsSummary(segs) {
        let p = 0, inp = 0, bl = 0;
        (segs || []).forEach(seg => {
            if (seg.type === 'prompt') p++;
            else if (seg.type === 'input') inp++;
            else if (seg.type === 'blank') bl++;
        });
        const arr = [];
        if (p) arr.push(p + '隐藏');
        if (inp) arr.push(inp + '输入');
        if (bl) arr.push(bl + '填空');
        return arr.join('/');
    }

    function drawPresetEditor() {
        const right = document.getElementById('presetRight');
        if (!right) return;
        if (_curPresetIdx < 0 || !_presetData.presets[_curPresetIdx]) {
            right.innerHTML = '<div style="color:var(--text2);text-align:center;padding:60px">← 左侧选择或新建预设</div>';
            return;
        }
        const p = _presetData.presets[_curPresetIdx];
        const pi = _curPresetIdx;

        const groupOpts = ['<option value="">（无分组）</option>']
            .concat((_presetData.groups || []).map(g =>
                `<option value="${esc(g)}"${p.group === g ? ' selected' : ''}>${esc(g)}</option>`
            )).join('');

        let html = `<div class="preset-edit-hdr">
            <input value="${esc(p.name)}" onchange="Admin.updP(${pi},'name',this.value)" placeholder="预设名" style="font-weight:600;font-size:15px;flex:1">
            <select onchange="Admin.updP(${pi},'group',this.value)" style="width:140px">${groupOpts}</select>
            <button class="btn btn-s" onclick="Admin.dupPreset(${pi})" title="复制整个预设">📄 复制</button>
            <button class="btn btn-s btn-d" onclick="Admin.delPreset(${pi})">🗑️ 删预设</button>
        </div>
        <div style="font-size:11px;color:var(--text2);margin-bottom:8px">💡 拖动 ⣿ 排序步骤；步骤内可建 A/B 版本（最多3个），用户端自选</div>
        <div class="preset-steps" id="presetStepsBox">`;

        (p.steps || []).forEach((s, si) => {
            const open = !!_openSteps[si];
            const vid = _vidOf(si);
            const curSegs = _segsRef(si, vid);
            const summary = _segsSummary(curSegs);
            const varCount = (s.variants || []).length;

            html += `<div class="pstep${open ? ' open' : ''}" draggable="true"
                    ondragstart="Admin.stepDragStart(event,${si})"
                    ondragover="Admin.stepDragOver(event,${si})"
                    ondragleave="Admin.stepDragLeave(event,${si})"
                    ondrop="Admin.stepDrop(event,${si})"
                    ondragend="Admin.stepDragEnd(event)">
                <div class="pstep-bar" onclick="Admin.toggleStep(${si})">
                    <span class="pstep-drag" onclick="event.stopPropagation()" title="拖动排序">⣿</span>
                    <input type="checkbox" onclick="event.stopPropagation()" onchange="Admin.toggleStepSel(${si},this.checked)" ${_stepSel[si] ? 'checked' : ''} title="勾选后可拆分为新预设" style="flex-shrink:0">
                    <span class="pstep-caret">${open ? '▼' : '▶'}</span>
                    <span class="pstep-title">步骤${si + 1}：${esc(s.name || '未命名')}</span>
                    ${varCount ? '<span style="font-size:10px;color:#f59e0b;flex-shrink:0">🔀' + (varCount + 1) + '版</span>' : ''}
                    ${summary ? '<span style="font-size:10px;color:var(--text2);flex-shrink:0">' + summary + '</span>' : ''}
                    ${s.engineName ? '<span style="font-size:10px;color:var(--pri)">🔌' + esc(s.engineName) + '</span>' : ''}
                    <span class="pstep-ops" onclick="event.stopPropagation()">
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

        html += `</div>
            <div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap">
                <button class="btn btn-p btn-s" onclick="Admin.addStep()">➕ 添加步骤</button>
                <button class="btn btn-s" onclick="Admin.stepsToNewPreset()">✂️ 勾选步骤另存为新预设</button>
            </div>`;
        right.innerHTML = html;
    }

    /* ---------- 版本标签页（引号已修正：属性统一单引号） ---------- */
    function _renderVariantTabs(si, step, curVid) {
        const vars = (step.variants || []);
        const total = 1 + vars.length;
        let html = '<div class="pv-tabs"><span class="pv-lbl">🔀 版本：</span>';

        const dLabel = step.defaultLabel || '默认';
        const dAct = (curVid === DEFAULT_VID) ? ' act' : '';
        html += '<span class="pv-tab' + dAct + '"'
            + ' style="user-select:none"'
            + ' title="点击切换｜双击改名"'
            + " onclick='Admin.selVariant(" + si + "," + JSON.stringify(DEFAULT_VID) + ")'"
            + " ondblclick='Admin.renameDefaultVar(" + si + ")'"
            + '>' + esc(dLabel) + '</span>';

        vars.forEach(v => {
            const act = (curVid === v.id) ? ' act' : '';
            const vidJson = JSON.stringify(v.id);
            html += '<span class="pv-tab' + act + '"'
                + ' style="user-select:none"'
                + ' title="点击切换｜双击改名"'
                + " onclick='Admin.selVariant(" + si + "," + vidJson + ")'"
                + " ondblclick='Admin.renameVariant(" + si + "," + vidJson + ")'"
                + '>' + esc(v.label || '未命名')
                + '<button class="pv-x"'
                + " onclick='event.stopPropagation();Admin.delVariant(" + si + "," + vidJson + ")'"
                + ' title="删除此版本">×</button>'
                + '</span>';
        });

        if (total < 3) {
            html += '<button class="pv-add" onclick="Admin.addVariant(' + si + ')">+ 新建版本</button>';
        } else {
            html += '<span style="font-size:10px;color:var(--text2)">（已达3版上限）</span>';
        }

        html += '</div>';
        return html;
    }

    function selVariant(si, vid) { _curVar[si] = vid; drawPresetEditor(); }

    function renameDefaultVar(si) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        FF.open({
            title: '默认版显示名',
            tip: '用户端会看到这个标签（如改成"标准版""通用版"）。',
            okText: '保存',
            fields: [{ key: 'label', label: '标签名', type: 'text', value: s.defaultLabel || '默认' }],
            onSubmit: (v) => {
                s.defaultLabel = (v.label || '').trim() || '默认';
                markDirty(); drawPresetEditor();
            }
        });
    }

    function addVariant(si) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        if (!Array.isArray(s.variants)) s.variants = [];
        if (s.variants.length >= 2) { toast('最多 3 个版本（含默认版）', 'er'); return; }

        FF.open({
            title: '➕ 新建步骤版本',
            tip: '版本用于同一步骤的 A/B 方案，用户端可自选。',
            okText: '创建',
            fields: [
                { key: 'label', label: '版本名称', type: 'text', value: 'A版', placeholder: '如：A版-悬疑向', hint: '用户端会看到这个名字' },
                { key: 'copy', label: '从当前版本复制内容', type: 'check', value: true, hint: '推荐：复制后改几处即可，比从零写快得多' },
            ],
            onSubmit: (v, setErr) => {
                const label = (v.label || '').trim();
                if (!label) { setErr('版本名称不能为空'); return false; }
                const srcSegs = v.copy ? _segsRef(si, _vidOf(si)) : [];
                const nv = {
                    id: _rid('v_'),
                    label: label,
                    segments: JSON.parse(JSON.stringify(srcSegs || [])),
                };
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
            title: '重命名版本',
            okText: '保存',
            fields: [{ key: 'label', label: '版本名称', type: 'text', value: v.label || '' }],
            onSubmit: (val, setErr) => {
                const nn = (val.label || '').trim();
                if (!nn) { setErr('不能为空'); return false; }
                v.label = nn;
                markDirty(); drawPresetEditor();
            }
        });
    }

    function delVariant(si, vid) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        const v = (s.variants || []).find(x => x.id === vid);
        if (!v) return;
        if (!confirm('删除版本「' + (v.label || '') + '」？该版本的所有片段会一起删除。')) return;
        s.variants = (s.variants || []).filter(x => x.id !== vid);
        if (_curVar[si] === vid) _curVar[si] = DEFAULT_VID;
        markDirty(); drawPresetList(); drawPresetEditor();
        toast('已删除版本');
    }

    function updP(pi, f, v) { _presetData.presets[pi][f] = v; markDirty(); if (f === 'name' || f === 'group') drawPresetList(); }
    function updS(si, f, v) { _presetData.presets[_curPresetIdx].steps[si][f] = v; markDirty(); }
    function updSeg(si, gi, f, v) { const segs = _segsRef(si, _vidOf(si)); if (segs[gi]) segs[gi][f] = v; markDirty(); }
    function updSegPrompt(si, gi, v) {
        const segs = _segsRef(si, _vidOf(si));
        if (!segs[gi]) return;
        segs[gi]._plain = v;
        segs[gi]._dirty = true;
        markDirty();
        drawPresetEditor();
    }
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
    function toggleStepSel(si, checked) { _stepSel[si] = checked; }

    function addPreset() {
        FF.open({
            title: '➕ 新建预设',
            okText: '创建',
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
        (copy.steps || []).forEach(s => {
            s.id = _rid('s');
            (s.variants || []).forEach(v => { v.id = _rid('v_'); });
        });
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

    /* ---------- 步骤拖拽排序 ---------- */
    function stepDragStart(ev, si) {
        _dragStepIdx = si;
        ev.currentTarget.classList.add('dragging');
        try {
            ev.dataTransfer.effectAllowed = 'move';
            ev.dataTransfer.setData('text/plain', String(si));
        } catch (e) {}
    }
    function stepDragOver(ev, si) {
        if (_dragStepIdx === null || _dragStepIdx === si) return;
        ev.preventDefault();
        try { ev.dataTransfer.dropEffect = 'move'; } catch (e) {}
        ev.currentTarget.classList.add('drag-over');
    }
    function stepDragLeave(ev) { ev.currentTarget.classList.remove('drag-over'); }
    function stepDrop(ev, si) {
        ev.preventDefault();
        ev.currentTarget.classList.remove('drag-over');
        let from = _dragStepIdx;
        if (from === null) {
            const raw = ev.dataTransfer ? ev.dataTransfer.getData('text/plain') : '';
            from = raw === '' ? null : parseInt(raw, 10);
        }
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
    function stepDragEnd() {
        _dragStepIdx = null;
        document.querySelectorAll('.pstep').forEach(el => {
            el.classList.remove('dragging'); el.classList.remove('drag-over');
        });
    }

    function stepsToNewPreset() {
        const p = _presetData.presets[_curPresetIdx];
        const picked = (p.steps || []).filter((s, i) => _stepSel[i]);
        if (!picked.length) { toast('请先勾选要提取的步骤', 'er'); return; }

        FF.open({
            title: '✂️ 拆分为新预设',
            tip: '已勾选 <b>' + picked.length + '</b> 个步骤，将复制成一个新预设（原预设不变）。',
            okText: '创建',
            fields: [
                { key: 'name', label: '新预设名称', type: 'text', value: (p.name || '') + ' - 拆分' },
                { key: 'group', label: '分组', type: 'select', value: p.group || '', options: [{ v: '', t: '（无分组）' }].concat((_presetData.groups || []).map(g => ({ v: g, t: g }))) },
            ],
            onSubmit: (v, setErr) => {
                const nm = (v.name || '').trim();
                if (!nm) { setErr('名称不能为空'); return false; }
                const np = { id: _rid('p'), name: nm, group: v.group || '', steps: JSON.parse(JSON.stringify(picked)) };
                np.steps.forEach((s, i) => {
                    s.id = _rid('s'); s.order = i + 1;
                    (s.variants || []).forEach(x => { x.id = _rid('v_'); });
                });
                _presetData.presets.push(np);
                _stepSel = {};
                markDirty(); drawPresetList(); drawPresetEditor();
                toast('✅ 已拆分为新预设「' + np.name + '」');
            }
        });
    }

    /* ---------- 分组管理 ---------- */
    function manageGroups() {
        if (!_presetData.groups) _presetData.groups = [];
        const right = document.getElementById('presetRight');
        if (!right) return;
        let html = '<div style="max-width:480px"><h3 style="margin-bottom:6px">🏷️ 分组管理</h3>';
        html += '<div style="font-size:11px;color:var(--text2);margin-bottom:12px;line-height:1.6">分组用于给员工分配权限（👥账号 → 🎫 工作流权限）。改名会自动同步所有预设。</div>';
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
            tip: '所有属于此分组的预设会自动跟着改。<br>⚠️ 已给员工配过此分组权限的，请记得到「👥账号 → 🎫」重新勾选。',
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

    /* ---------- 安全设置 ---------- */
    function showSecurity() {
        const sec = _presetData.security || {};
        const right = document.getElementById('presetRight');
        right.innerHTML = `<div style="max-width:520px"><h3 style="margin-bottom:12px">🛡️ 安全设置</h3>
            <div class="fg"><label>敏感词（逗号隔开）<span class="cfg-help" title="用户输入含这些词时直接拦截并推钉钉">?</span></label><textarea id="ps_sensitive" rows="4">${esc((sec.sensitiveWords || []).join(','))}</textarea></div>
            <div class="fg"><label>钉钉 Webhook<span class="cfg-help" title="留空则不推送报警">?</span></label><input id="ps_webhook" value="${esc(sec.alertWebhook || '')}" placeholder="https://oapi.dingtalk.com/robot/send?access_token=..."></div>
            <div class="fr"><div class="fg"><label>报警关键词</label><input id="ps_keyword" value="${esc(sec.alertKeyword || '飞凡警报')}"></div><div class="fg"><label>相似度阈值%<span class="cfg-help" title="AI输出与隐藏指令相似度超过此值即判定泄露并屏蔽">?</span></label><input id="ps_sim" type="number" value="${sec.simThreshold || 70}"></div></div>
            <div class="pt"><input type="checkbox" id="ps_guard" ${sec.guard !== false ? 'checked' : ''}><label for="ps_guard">开启 GUARD 保密前缀（防套取指令）</label></div>
            <div style="font-size:11px;color:var(--text2);margin:8px 0">💡 分组请到「🏷️ 分组管理」中维护。</div>
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
        markDirty();
        toast('已应用，记得点"💾 保存到云端"');
    }

    /* ---------- 保存 / 导出 / 导入（递归处理分支） ---------- */
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
        } catch (e) {
            toast('保存失败：' + e.message, 'er');
        }
    }

    /* 🔴 递归剔除 _plain / _dirty（默认版 + 全部分支） */
    function _stripPlain(data) {
        (data.presets || []).forEach(p => {
            (p.steps || []).forEach(s => {
                const groups = [s.segments || []];
                (s.variants || []).forEach(v => groups.push(v.segments || []));
                groups.forEach(segs => {
                    segs.forEach(seg => { delete seg._plain; delete seg._dirty; });
                });
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
        a.click();
        URL.revokeObjectURL(a.href);
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
                                        try { seg._plain = await Workflow.decrypt(seg.hidden || ''); }
                                        catch (er) { seg._plain = ''; }
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
            } catch (err) {
                toast('JSON解析失败', 'er');
            }
            inputEl.value = '';
        };
        reader.readAsText(file, 'utf-8');
    }

    /* ========================================================== */
    /* ======================= 监视 ============================= */
    /* ========================================================== */
    async function renderMonitor(box) { box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>'; try { const data = await apiCall('admin/monitor'); const logMap = {}; (data.logs || []).forEach(l => logMap[l.username] = l); const sessMap = {}; (data.sessions || []).forEach(s => sessMap[s.username] = s); const usernames = new Set([...Object.keys(logMap), ...Object.keys(sessMap)]); let html = `<div style="display:flex;gap:16px;margin-bottom:12px;flex-wrap:wrap"><div style="padding:10px 16px;background:var(--pri-l);border-radius:8px"><div style="font-size:11px;color:var(--text2)">当前在线（5分钟内）</div><div style="font-size:22px;font-weight:600;color:#10b981">${data.onlineCount || 0} 人</div></div><button class="btn btn-s" onclick="Admin.switchTab('monitor')" style="align-self:center">🔄刷新</button></div><div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>账号</th><th>对话次数</th><th>累计Token</th><th>不同IP</th><th>最后活跃</th></tr></thead><tbody>`; usernames.forEach(un => { const l = logMap[un] || {}; const s = sessMap[un] || {}; html += `<tr><td>${esc(un)}</td><td>${l.logCount || 0}</td><td>${(l.totalTokens || 0).toLocaleString()}</td><td>${(s.ipc || 0) >= 3 ? '<span style="color:#ef4444">🔴' + s.ipc + '</span>' : (s.ipc || 0)}</td><td style="font-size:11px">${fmtTime(s.last || 0)}</td></tr>`; }); html += '</tbody></table></div><h4 style="font-size:13px;margin:16px 0 8px">📋 最近100条</h4><div style="overflow-x:auto;max-height:280px;overflow-y:auto"><table class="admin-table"><thead><tr><th>时间</th><th>账号</th><th>对话</th><th>轮次</th><th>Token</th><th>模型</th></tr></thead><tbody>'; (data.recent || []).forEach(r => { html += `<tr><td style="font-size:11px">${new Date(r.created_at).toLocaleString()}</td><td>${esc(r.username)}</td><td style="max-width:150px;overflow:hidden;text-overflow:ellipsis">${esc(r.chat_name || '-')}</td><td>${r.rounds || 0}</td><td>${r.tokens || 0}</td><td>${esc(r.model || '-')}</td></tr>`; }); html += '</tbody></table></div>'; box.innerHTML = html; } catch (e) { box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>'; } }

    /* ========================================================== */
    /* ============ ★ 全局设置（分组折叠 + ?提示） ============== */
    /* ========================================================== */
    let _cfgOpen = { chunk: true, quick: true, cmds: false };

    async function renderConfig(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>';
        try {
            const data = await apiCall('admin/config/get');
            const cfg = data.config || {};

            let quick = [];
            try { quick = JSON.parse(cfg.quickModels || '[]'); } catch (e) { quick = []; }
            if (!Array.isArray(quick)) quick = [];
            _quickModelDraft = quick;

            const cmdsRaw = cfg.quickCmds || '';
            const cmdCount = cmdsRaw.split('\n').filter(l => l.indexOf('|') > 0).length;

            box.innerHTML = `<div style="max-width:620px">
                <div style="font-size:12px;color:var(--text2);margin-bottom:12px">⚙️ 这里的设置对<b>所有用户</b>生效。改完点最下方「💾 保存全部」。</div>

                <!-- 打标 -->
                <div class="cfg-sec${_cfgOpen.chunk ? ' open' : ''}" id="cfgSecChunk">
                    <div class="cfg-sec-hdr" onclick="Admin.toggleCfgSec('chunk')">
                        <span class="cs-caret">${_cfgOpen.chunk ? '▼' : '▶'}</span>
                        <span class="cs-title">📐 物理打标</span>
                        <span class="cs-sub">当前每块 ${esc(cfg.chunkSize || '300')} 字</span>
                    </div>
                    <div class="cfg-sec-body">
                        <div class="fg">
                            <label>每块字数<span class="cfg-help" title="上传长文时按此字数切块并标注位置，让AI能准确说出'第X字/占比Y%'。块越小定位越准但提示词越长">?</span></label>
                            <input id="cfg_chunkSize" type="number" value="${esc(cfg.chunkSize || '300')}" min="50" max="5000">
                            <div class="ff-hint">推荐 300。小说类长文可用 500；需要精确定位可用 200。</div>
                        </div>
                    </div>
                </div>

                <!-- 快捷模型档 -->
                <div class="cfg-sec${_cfgOpen.quick ? ' open' : ''}" id="cfgSecQuick">
                    <div class="cfg-sec-hdr" onclick="Admin.toggleCfgSec('quick')">
                        <span class="cs-caret">${_cfgOpen.quick ? '▼' : '▶'}</span>
                        <span class="cs-title">⚡ 快捷模型档</span>
                        <span class="cs-sub">已配 ${quick.length} 档</span>
                    </div>
                    <div class="cfg-sec-body">
                        <div style="font-size:11px;color:var(--text2);margin-bottom:10px;line-height:1.7">
                            用户输入框上方会出现这些小按钮，点一下即<b>临时</b>换模型（<b>不改动引擎配置</b>，切引擎/换对话自动复位）。<br>
                            · <b>标签</b>：按钮上的字，越短越好（如"快""强""图"）<br>
                            · <b>模型名</b>：留空 = 用当前引擎的默认模型（可当"恢复默认"按钮用）<br>
                            · <b>生图</b>：勾选后点该档即进入生图模式
                        </div>
                        <div id="quickModelList"></div>
                        <button class="btn btn-s" onclick="Admin.addQuickModel()" style="margin-top:6px">➕ 添加一档</button>
                    </div>
                </div>

                <!-- 快捷指令 -->
                <div class="cfg-sec${_cfgOpen.cmds ? ' open' : ''}" id="cfgSecCmds">
                    <div class="cfg-sec-hdr" onclick="Admin.toggleCfgSec('cmds')">
                        <span class="cs-caret">${_cfgOpen.cmds ? '▼' : '▶'}</span>
                        <span class="cs-title">⌨️ 快捷指令</span>
                        <span class="cs-sub">已配 ${cmdCount} 条</span>
                    </div>
                    <div class="cfg-sec-body">
                        <div class="fg">
                            <label>指令列表<span class="cfg-help" title="用户在输入框打 / 即弹出面板，选中后内容自动填入输入框">?</span></label>
                            <textarea id="cfg_quickCmds" rows="7" placeholder="每行一条，格式：名称|内容模板&#10;翻译成英文|请把以下内容翻译成地道的英文：&#10;总结要点|请用要点列出以下内容的核心：">${esc(cmdsRaw)}</textarea>
                            <div class="ff-hint">格式：<code>名称|内容</code>，一行一条。内容必须写在同一行（不能换行）。</div>
                        </div>
                    </div>
                </div>

                <button class="btn btn-p" onclick="Admin.saveConfig()" style="margin-top:6px">💾 保存全部</button>
            </div>`;

            drawQuickModelList();
        } catch (e) {
            box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>';
        }
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

    function drawQuickModelList() {
        const box = document.getElementById('quickModelList');
        if (!box) return;
        let html = '';
        _quickModelDraft.forEach((q, i) => {
            html += `<div class="fr" style="margin-bottom:6px;align-items:flex-end">
                <div class="fg" style="flex:0 0 88px;margin-bottom:0"><label>标签</label><input value="${esc(q.label || '')}" onchange="Admin.updQuickModel(${i},'label',this.value)" placeholder="如：快"></div>
                <div class="fg" style="margin-bottom:0"><label>模型名（留空=引擎默认）</label><input value="${esc(q.model || '')}" onchange="Admin.updQuickModel(${i},'model',this.value)" placeholder="如 gpt-4o-mini"></div>
                <div class="fg" style="flex:0 0 50px;margin-bottom:0;text-align:center"><label>生图</label><input type="checkbox" ${q.isImage ? 'checked' : ''} onchange="Admin.updQuickModel(${i},'isImage',this.checked)" style="width:18px;height:18px;accent-color:var(--pri)"></div>
                <button class="btn btn-s btn-d" onclick="Admin.delQuickModel(${i})" style="margin-bottom:2px">🗑️</button>
            </div>`;
        });
        if (!_quickModelDraft.length) html = '<div style="font-size:11px;color:var(--text2);padding:6px 0">（暂无档位，点下方添加）</div>';
        box.innerHTML = html;
    }

    function addQuickModel() { _quickModelDraft.push({ label: '', model: '', isImage: false }); drawQuickModelList(); }
    function delQuickModel(i) { _quickModelDraft.splice(i, 1); drawQuickModelList(); }
    function updQuickModel(i, field, val) { if (_quickModelDraft[i]) _quickModelDraft[i][field] = val; }

    function saveConfig() {
        const chunkSize = document.getElementById('cfg_chunkSize').value;
        const quickCmds = document.getElementById('cfg_quickCmds').value;

        const quick = (_quickModelDraft || [])
            .filter(q => (q.label || '').trim())
            .map(q => ({ label: (q.label || '').trim(), model: (q.model || '').trim(), isImage: !!q.isImage }));

        apiCall('admin/config/save', 'POST', {
            config: { chunkSize, quickModels: JSON.stringify(quick), quickCmds }
        }).then(() => {
            toast('✅ 已保存（用户刷新页面后生效）');
            if (typeof Chunker !== 'undefined') Chunker.setBlockSize(chunkSize);
            if (typeof window.applyGlobalConfig === 'function') {
                window.applyGlobalConfig({ quickModels: JSON.stringify(quick), quickCmds, chunkSize });
            }
        }).catch(e => toast('失败：' + e.message, 'er'));
    }

    /* ========================================================== */
    return {
        open, close, switchTab, apiCall,
        /* 账号 */
        showCreateUser, showEditUser, showResetPwd, showPerm, toggleStatus, delUser, searchUsers,
        importXLSX, exportXLSX, downloadTemplate,
        /* 引擎 */
        showEngEdit, delEng,
        /* 模板库 */
        showTplEdit, delTpl, syncTpl, doDeploy, toggleTplSel, toggleDeployUser, deploySelectAllUsers,
        /* 体检 */
        runHealth,
        /* 模型库 */
        showModelEdit, delModel,
        /* 预设 */
        addPreset, dupPreset, delPreset, selectPreset, filterPresetList,
        updP, updS, updSeg, updSegPrompt,
        toggleStep, toggleStepSel, stepsToNewPreset,
        addStep, dupStep, delStep, moveStep, addSeg, delSeg, backToEditor,
        /* 分支 */
        selVariant, addVariant, renameVariant, delVariant, renameDefaultVar,
        /* 拖拽 */
        stepDragStart, stepDragOver, stepDragLeave, stepDrop, stepDragEnd,
        /* 分组 / 安全 / 保存 */
        manageGroups, addGroup, renameGroup, delGroup,
        showSecurity, applySecurity, savePresets, exportPresetsJSON, importPresetsJSON,
        /* 全局设置 */
        saveConfig, addQuickModel, delQuickModel, updQuickModel, toggleCfgSec,
        /* 表单组件（供外部复用） */
        FF,
    };
})();

window.Admin = Admin;
