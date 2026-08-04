/* ===== 飞凡AI - 超管后台 (v3.2.0 批次2：分支/拖拽/快捷档) ===== */

const Admin = (function () {

    /* ---------- 模块级状态（★全部提到顶部，避免 TDZ） ---------- */
    let _curTab = 'users';
    let _presetData = null;
    let _curPresetIdx = -1;        // 当前编辑的预设索引
    let _openSteps = {};           // 步骤折叠状态
    let _presetDirty = false;      // 未保存标记
    let _usersCache = [];
    let _stepSel = {};             // 步骤勾选（用于拆分）
    let _curVar = {};              // ★ 每个步骤当前选中的版本id：{ stepIdx: variantId }
    let _quickModelDraft = [];     // ★ 快捷模型档草稿
    let _dragStepIdx = null;       // ★ 拖拽中的步骤索引

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

    function _rid(pfx) { return pfx + Math.random().toString(36).slice(2, 8); }

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
            <div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>姓名</th><th>账号</th><th>角色</th><th>状态</th><th>引擎</th><th>权限</th><th>最后活跃</th><th>IP</th><th>操作</th></tr></thead><tbody>`;
        users.forEach(u => {
            let permTxt = '全部';
            try { const p = JSON.parse(u.permissions || '{}'); if (p.allowGroups && p.allowGroups.length) permTxt = p.allowGroups.join('/'); } catch (e) {}
            html += `<tr><td>${esc(u.name || '-')}</td><td>${esc(u.username)}</td><td>${u.role === 'admin' ? '👑' : '普通'}</td><td>${u.status === 'active' ? '<span style="color:#10b981">启用</span>' : '<span style="color:#ef4444">禁用</span>'}</td><td>${u.engineCount}</td><td style="font-size:11px;max-width:110px;overflow:hidden;text-overflow:ellipsis">${esc(permTxt)}</td><td style="font-size:11px">${fmtTime(u.lastActive)}</td><td>${u.ipAbnormal ? '<span style="color:#ef4444;font-weight:bold">🔴' + u.ipCount + '</span>' : (u.ipCount || 0)}</td><td class="admin-ops"><button onclick='Admin.showPerm(${JSON.stringify(u.username)},${JSON.stringify(u.permissions || "{}")})' title="权限">🎫</button><button onclick='Admin.showResetPwd(${JSON.stringify(u.username)})' title="改密">🔑</button><button onclick='Admin.toggleStatus(${JSON.stringify(u.username)},${JSON.stringify(u.status)})' title="启用/禁用">${u.status === 'active' ? '🚫' : '✅'}</button>${u.username !== 'admin' ? `<button onclick='Admin.delUser(${JSON.stringify(u.username)})' style="color:#ef4444">🗑️</button>` : ''}</td></tr>`;
        });
        html += '</tbody></table></div>';
        box.innerHTML = html;
    }

    function searchUsers(kw) { drawUsersTable(document.getElementById('adminBody'), kw); }
    function showCreateUser() { const name = prompt('姓名：', ''); if (name === null) return; const username = prompt('账号：', ''); if (!username || !username.trim()) { toast('账号不能为空', 'er'); return; } const password = prompt('密码：', ''); if (!password || !password.trim()) { toast('密码不能为空', 'er'); return; } const isAdmin = confirm('设为管理员？\n✅=管理员 ❌=普通'); apiCall('admin/users/create', 'POST', { username: username.trim(), password: password.trim(), name: name.trim(), role: isAdmin ? 'admin' : 'user' }).then(() => { toast('✅ 已创建'); switchTab('users'); }).catch(e => toast('失败：' + e.message, 'er')); }
    function showResetPwd(username) { const p = prompt('为【' + username + '】设新密码：', ''); if (!p || !p.trim()) return; apiCall('admin/users/resetpwd', 'POST', { username, password: p.trim() }).then(() => toast('✅ 已重置')).catch(e => toast('失败：' + e.message, 'er')); }
    function toggleStatus(username, cur) { const next = cur === 'active' ? 'disabled' : 'active'; apiCall('admin/users/update', 'POST', { username, status: next }).then(() => { toast(next === 'active' ? '已启用' : '已禁用'); switchTab('users'); }).catch(e => toast('失败：' + e.message, 'er')); }
    function delUser(username) { if (!confirm('删除账号【' + username + '】？其引擎、会话也删除。')) return; apiCall('admin/users/delete', 'POST', { username }).then(() => { toast('✅ 已删除'); switchTab('users'); }).catch(e => toast('失败：' + e.message, 'er')); }
    function showPerm(username, permJson) { let perm = {}; try { perm = JSON.parse(permJson); } catch (e) {} const cur = (perm.allowGroups || []).join(','); const groups = (typeof Workflow !== 'undefined' && Workflow.isLoaded()) ? Workflow.getGroups().join('、') : '（预设未加载）'; const v = prompt('设置【' + username + '】可用工作流分组\n\n可选：' + groups + '\n\n多个用英文逗号；留空=全部：', cur); if (v === null) return; const arr = v.split(',').map(s => s.trim()).filter(Boolean); apiCall('admin/users/perm', 'POST', { username, permissions: Object.assign({}, perm, { allowGroups: arr }) }).then(() => { toast('✅ 权限已更新（实时生效）'); switchTab('users'); }).catch(e => toast('失败：' + e.message, 'er')); }
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
            const engData = await apiCall('admin/engines/list');
            const engs = engData.engines || [];

            const byUser = {};
            engs.forEach(e => { if (!byUser[e.username]) byUser[e.username] = []; byUser[e.username].push(e); });

            let html = `<div style="font-size:12px;color:var(--text2);margin-bottom:10px">给账号配公有引擎。🎨=生图引擎，💰=已开缓存（省钱）。模型可留空让用户自选。</div>`;

            users.forEach(u => {
                const ue = byUser[u.username] || [];
                html += `<div class="eng-user-block"><div class="eng-user-hdr"><b>${esc(u.name || u.username)}</b> <span style="color:var(--text2);font-size:11px">(${esc(u.username)})</span><button class="btn btn-p btn-s" style="margin-left:auto" onclick='Admin.showEngEdit(${JSON.stringify(u.username)},"")'>➕ 加引擎</button></div>`;
                if (!ue.length) html += '<div style="font-size:11px;color:var(--text2);padding:4px 0">（无引擎）</div>';
                ue.forEach(e => {
                    const imgTag = (e.engineType === 'image') ? '<span style="color:#f59e0b">🎨生图</span>' : '';
                    const cacheTag = e.useCache ? '<span style="color:#10b981">💰缓存</span>' : '';
                    html += `<div class="eng-item"><span>📦 ${esc(e.name)} ${imgTag} ${cacheTag} <span style="color:var(--text2);font-size:11px">${esc(e.protocol)} / ${esc(e.model || '用户自选')}</span></span>
                        <div style="margin-left:auto;display:flex;gap:4px"><button class="btn btn-s" onclick='Admin.showEngEdit(${JSON.stringify(e.username)},${JSON.stringify(e.id)})'>✏️改</button><button class="btn btn-s btn-d" onclick='Admin.delEng(${JSON.stringify(e.id)})'>🗑️</button></div></div>`;
                });
                html += `</div>`;
            });

            box.innerHTML = html;
            box._engs = engs;
        } catch (e) {
            box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>';
        }
    }

    function showEngEdit(username, engId) {
        const box = document.getElementById('adminBody');
        const engs = (box && box._engs) || [];
        const e = engId ? engs.find(x => x.id === engId) : null;
        const etype = e ? (e.engineType || 'chat') : 'chat';

        box.innerHTML = `<div style="max-width:520px"><h3 style="margin-bottom:12px">${e ? '✏️ 编辑' : '➕ 新增'}公有引擎 — ${esc(username)}</h3>

            <div class="fg"><label>引擎名称</label><input id="ee_name" value="${e ? esc(e.name) : ''}"></div>

            <div class="fg" style="padding:10px;background:var(--pri-l);border-radius:8px">
                <label>🔧 引擎类型</label>
                <select id="ee_etype">
                    <option value="chat"${etype !== 'image' ? ' selected' : ''}>💬 对话</option>
                    <option value="image"${etype === 'image' ? ' selected' : ''}>🎨 生图 / 改图</option>
                </select>
                <div style="font-size:11px;color:var(--text2);margin-top:4px">选「生图」后，用户发消息即出图（支持带参考图改图）</div>
            </div>

            <div class="fg"><label>协议</label><select id="ee_proto">
                <option value="openai"${!e || e.protocol === 'openai' ? ' selected' : ''}>OpenAI/通用</option>
                <option value="anthropic"${e && e.protocol === 'anthropic' ? ' selected' : ''}>Claude原生</option>
                <option value="gemini"${e && e.protocol === 'gemini' ? ' selected' : ''}>Gemini原生</option>
            </select></div>

            <div class="fg"><label>Base URL</label><input id="ee_base" value="${e ? esc(e.base) : 'https://api.openai-proxy.org/v1'}"></div>

            <div class="fg"><label>API Key ${e ? '<span style="color:var(--text2);font-size:11px">（留空=不改）</span>' : ''}</label><input id="ee_key" type="password" placeholder="${e ? '••••（留空不变）' : 'sk-...'}"></div>

            <div class="fg"><label>默认模型 <span style="color:var(--text2);font-size:11px">（可留空，用户自选）</span></label><input id="ee_model" value="${e ? esc(e.model || '') : ''}" placeholder="对话如 gpt-4o；生图如 gpt-image-1" oninput="Admin.autoDetectEngType()"></div>

            <div class="fg" style="padding:10px;background:var(--pri-l);border-radius:8px"><label class="pt" style="margin:0"><input type="checkbox" id="ee_cache" ${e && e.useCache ? 'checked' : ''}> 💰 开启 Prompt 缓存（对话用，重复内容省钱）</label><div style="font-size:11px;color:var(--text2);margin-top:4px">开启后，用户对话底部"💰命中"数字&gt;0即生效省钱。生图引擎无需开启</div></div>

            <div class="fg"><label style="font-size:11px;color:var(--text2)">单价可留空 → 查"模型库"。也可指定：</label></div>
            <div class="fr"><div class="fg"><label>输入</label><input id="ee_pi" type="number" step="0.01" value="${e ? (e.priceIn || 0) : 0}"></div><div class="fg"><label>输出</label><input id="ee_po" type="number" step="0.01" value="${e ? (e.priceOut || 0) : 0}"></div></div>
            <div class="fr"><div class="fg"><label>缓存读</label><input id="ee_pcr" type="number" step="0.01" value="${e ? (e.priceCR || 0) : 0}"></div><div class="fg"><label>缓存写</label><input id="ee_pcw" type="number" step="0.01" value="${e ? (e.priceCW || 0) : 0}"></div></div>

            <div style="display:flex;gap:8px;margin-top:12px"><button class="btn btn-p" onclick='Admin.saveEng(${JSON.stringify(username)},${JSON.stringify(e ? e.id : "")})'>💾 保存</button><button class="btn" onclick="Admin.switchTab('engines')">取消</button></div></div>`;
    }

    function autoDetectEngType() {
        const modelEl = document.getElementById('ee_model');
        const etEl = document.getElementById('ee_etype');
        if (!modelEl || !etEl) return;
        const m = (modelEl.value || '').toLowerCase();
        if (/image|dall-?e|flux|stable-?diffusion|midjourney|imagen|seedream|kolors|画图|绘图/.test(m)) {
            etEl.value = 'image';
        }
    }

    function saveEng(username, engId) {
        const g = (id) => { const el = document.getElementById(id); return el ? el.value : ''; };
        const body = {
            username,
            id: engId || undefined,
            name: g('ee_name').trim(),
            engineType: g('ee_etype'),
            protocol: g('ee_proto'),
            base: g('ee_base').trim(),
            key: g('ee_key').trim(),
            model: g('ee_model').trim(),
            useCache: document.getElementById('ee_cache').checked,
            priceIn: parseFloat(g('ee_pi')) || 0,
            priceOut: parseFloat(g('ee_po')) || 0,
            priceCR: parseFloat(g('ee_pcr')) || 0,
            priceCW: parseFloat(g('ee_pcw')) || 0
        };
        if (!body.name) { toast('引擎名必填', 'er'); return; }
        apiCall('admin/engines/save', 'POST', body)
            .then(() => { toast('✅ 已保存'); switchTab('engines'); })
            .catch(e => toast('失败：' + e.message, 'er'));
    }

    function delEng(id) { if (!confirm('删除这个引擎？')) return; apiCall('admin/engines/delete', 'POST', { id }).then(() => { toast('✅ 已删除'); switchTab('engines'); }).catch(e => toast('失败：' + e.message, 'er')); }

    /* ========================================================== */
    /* ====================== 模型库 ============================ */
    /* ========================================================== */
    async function renderModels(box) { box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>'; try { const data = await apiCall('admin/models/list'); const models = data.models || []; let html = `<div style="display:flex;gap:8px;margin-bottom:12px"><button class="btn btn-p btn-s" onclick="Admin.showModelEdit('')">➕ 新增模型</button><button class="btn btn-s" onclick="Admin.switchTab('models')">🔄 刷新</button></div><div style="font-size:12px;color:var(--text2);margin-bottom:8px">模型库：配各模型单价（美元/1M token）。引擎未指定单价时自动查这里。</div><div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>模型名</th><th>输入</th><th>输出</th><th>缓存读</th><th>缓存写</th><th>操作</th></tr></thead><tbody>`; models.forEach(m => { html += `<tr><td>${esc(m.model_name)}</td><td>${m.price_in}</td><td>${m.price_out}</td><td>${m.price_cache_read}</td><td>${m.price_cache_write}</td><td class="admin-ops"><button onclick='Admin.showModelEdit(${JSON.stringify(m.model_name)})'>✏️</button><button onclick='Admin.delModel(${JSON.stringify(m.model_name)})' style="color:#ef4444">🗑️</button></td></tr>`; }); html += '</tbody></table></div>'; box.innerHTML = html; box._models = models; } catch (e) { box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>'; } }
    function showModelEdit(modelName) { const box = document.getElementById('adminBody'); const models = (box && box._models) || []; const m = modelName ? models.find(x => x.model_name === modelName) : null; box.innerHTML = `<div style="max-width:460px"><h3 style="margin-bottom:12px">${m ? '✏️ 编辑' : '➕ 新增'}模型单价</h3><div class="fg"><label>模型名 ${m ? '（不可改）' : ''}</label><input id="mm_name" value="${m ? esc(m.model_name) : ''}" ${m ? 'disabled' : ''} placeholder="如 gpt-4o"></div><div class="fr"><div class="fg"><label>输入</label><input id="mm_pi" type="number" step="0.01" value="${m ? m.price_in : 0}"></div><div class="fg"><label>输出</label><input id="mm_po" type="number" step="0.01" value="${m ? m.price_out : 0}"></div></div><div class="fr"><div class="fg"><label>缓存读</label><input id="mm_pcr" type="number" step="0.01" value="${m ? m.price_cache_read : 0}"></div><div class="fg"><label>缓存写</label><input id="mm_pcw" type="number" step="0.01" value="${m ? m.price_cache_write : 0}"></div></div><div style="display:flex;gap:8px;margin-top:12px"><button class="btn btn-p" onclick='Admin.saveModel(${JSON.stringify(m ? m.model_name : "")})'>💾 保存</button><button class="btn" onclick="Admin.switchTab('models')">取消</button></div></div>`; }
    function saveModel(existName) { const g = (id) => { const el = document.getElementById(id); return el ? el.value : ''; }; const name = existName || g('mm_name').trim(); if (!name) { toast('模型名必填', 'er'); return; } apiCall('admin/models/save', 'POST', { model_name: name, priceIn: parseFloat(g('mm_pi')) || 0, priceOut: parseFloat(g('mm_po')) || 0, priceCR: parseFloat(g('mm_pcr')) || 0, priceCW: parseFloat(g('mm_pcw')) || 0 }).then(() => { toast('✅ 已保存'); switchTab('models'); }).catch(e => toast('失败：' + e.message, 'er')); }
    function delModel(name) { if (!confirm('删除模型【' + name + '】单价？')) return; apiCall('admin/models/delete', 'POST', { model_name: name }).then(() => { toast('✅ 已删除'); switchTab('models'); }).catch(e => toast('失败：' + e.message, 'er')); }

    /* ========================================================== */
    /* ==================== 预设管理（含分支） =================== */
    /* ========================================================== */

    /* 遍历一个步骤的所有片段集合（默认版 + 全部分支），callback(segArr) */
    function _eachSegGroup(step, cb) {
        if (!step) return;
        cb(step.segments || []);
        (step.variants || []).forEach(v => cb(v.segments || []));
    }

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

            /* ★ 递归解密：默认版 + 全部分支 */
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
            _openSteps = {};
            _stepSel = {};
            _curVar = {};
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

    /* 当前步骤选中的版本id（默认=默认版） */
    function _vidOf(si) { return _curVar[si] || DEFAULT_VID; }

    /* 取某步骤某版本的片段数组引用（可直接修改） */
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

    /* 片段概要（用于步骤条） */
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
        <div style="font-size:11px;color:var(--text2);margin-bottom:8px">💡 拖动 ⣿ 可排序步骤；步骤内可建 A/B 版本（最多3个），用户端自选</div>
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

    /* ---------- 版本标签页 ---------- */
    function _renderVariantTabs(si, step, curVid) {
        const vars = (step.variants || []);
        const total = 1 + vars.length;
        let html = '<div class="pv-tabs"><span class="pv-lbl">🔀 版本：</span>';

        /* 默认版（属性统一用单引号包裹，内部 JSON.stringify 出双引号，不冲突） */
        const dLabel = step.defaultLabel || '默认';
        const dAct = (curVid === DEFAULT_VID) ? ' act' : '';
        html += '<span class="pv-tab' + dAct + '"'
            + ' style="user-select:none"'
            + ' title="点击切换｜双击改名"'
            + " onclick='Admin.selVariant(" + si + "," + JSON.stringify(DEFAULT_VID) + ")'"
            + " ondblclick='Admin.renameDefaultVar(" + si + ")'"
            + '>' + esc(dLabel) + '</span>';

        /* 分支版 */
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

        /* 新建（含默认版最多 3 个） */
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
        const nv = prompt('默认版显示名（用户端看到的标签）：', s.defaultLabel || '默认');
        if (nv === null) return;
        s.defaultLabel = nv.trim() || '默认';
        markDirty(); drawPresetEditor();
    }

    function addVariant(si) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        if (!Array.isArray(s.variants)) s.variants = [];
        if (s.variants.length >= 2) { toast('最多 3 个版本（含默认版）', 'er'); return; }

        const label = prompt('新版本名称（用户端会看到）：', 'A版');
        if (label === null) return;

        const copyFrom = confirm('从当前版本复制内容？\n✅=复制（推荐，改几处即可）\n❌=空白新建');
        const srcSegs = copyFrom ? _segsRef(si, _vidOf(si)) : [];

        const v = {
            id: _rid('v_'),
            label: label.trim() || ('版本' + (s.variants.length + 1)),
            segments: JSON.parse(JSON.stringify(srcSegs || [])),
        };
        s.variants.push(v);
        _curVar[si] = v.id;
        markDirty(); drawPresetList(); drawPresetEditor();
        toast('✅ 已新建版本「' + v.label + '」');
    }

    function renameVariant(si, vid) {
        const s = _presetData.presets[_curPresetIdx].steps[si];
        const v = (s.variants || []).find(x => x.id === vid);
        if (!v) return;
        const nv = prompt('重命名版本：', v.label || '');
        if (nv === null) return;
        v.label = nv.trim() || v.label;
        markDirty(); drawPresetEditor();
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

    /* ---------- 片段编辑（自动作用于当前选中版本） ---------- */
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

    /* ---------- 预设 增删复制 ---------- */
    function addPreset() {
        const p = { id: _rid('p'), name: '新预设', group: '', steps: [] };
        _presetData.presets.push(p);
        _curPresetIdx = _presetData.presets.length - 1;
        _openSteps = {}; _stepSel = {}; _curVar = {};
        markDirty(); drawPresetList(); drawPresetEditor();
    }

    function dupPreset(pi) {
        const src = _presetData.presets[pi];
        if (!src) return;
        const copy = JSON.parse(JSON.stringify(src));
        copy.id = _rid('p');
        copy.name = (src.name || '未命名') + ' 副本';
        /* ★ 重生成步骤id + 全部 variant id */
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
        if (!confirm('删除此预设？')) return;
        _presetData.presets.splice(pi, 1);
        _curPresetIdx = _presetData.presets.length ? 0 : -1;
        _openSteps = {}; _stepSel = {}; _curVar = {};
        markDirty(); drawPresetList(); drawPresetEditor();
    }

    /* ---------- 步骤 增删复制移动 ---------- */
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
        /* ★ 重生成 variant id，避免共用 */
        (copy.variants || []).forEach(v => { v.id = _rid('v_'); });
        p.steps.splice(si + 1, 0, copy);
        p.steps.forEach((s, i) => s.order = i + 1);
        _remapStepIdxState(si, si + 1, 'insert');
        _openSteps[si + 1] = true;
        markDirty(); drawPresetEditor();
        toast('✅ 已复制步骤');
    }

    function delStep(si) {
        if (!confirm('删除此步骤？')) return;
        const p = _presetData.presets[_curPresetIdx];
        p.steps.splice(si, 1);
        p.steps.forEach((s, i) => s.order = i + 1);
        _openSteps = {}; _stepSel = {}; _curVar = {};
        markDirty(); drawPresetList(); drawPresetEditor();
    }

    function moveStep(si, dir) {
        const steps = _presetData.presets[_curPresetIdx].steps;
        const j = si + dir;
        if (j < 0 || j >= steps.length) return;
        const t = steps[si]; steps[si] = steps[j]; steps[j] = t;
        steps.forEach((s, i) => s.order = i + 1);
        _swapStepIdxState(si, j);
        markDirty(); drawPresetEditor();
    }

    /* 索引态交换（折叠/勾选/版本选择） */
    function _swapStepIdxState(a, b) {
        [_openSteps, _stepSel, _curVar].forEach(map => {
            const ta = map[a], tb = map[b];
            if (tb === undefined) delete map[a]; else map[a] = tb;
            if (ta === undefined) delete map[b]; else map[b] = ta;
        });
    }
    /* 插入/删除时整体后移（简化：插入后清空，避免错位） */
    function _remapStepIdxState(from, to, mode) {
        _openSteps = {}; _stepSel = {}; _curVar = {};
    }

    /* ---------- 步骤拖拽排序 ---------- */
    function stepDragStart(ev, si) {
        _dragStepIdx = si;
        const el = ev.currentTarget;
        el.classList.add('dragging');
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
    function stepDragLeave(ev, si) {
        ev.currentTarget.classList.remove('drag-over');
    }
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

        /* 拖拽后索引全乱，直接重置折叠/勾选/版本态 */
        _openSteps = {}; _stepSel = {}; _curVar = {};
        _dragStepIdx = null;
        markDirty(); drawPresetEditor();
        toast('✅ 已重新排序');
    }
    function stepDragEnd(ev) {
        _dragStepIdx = null;
        document.querySelectorAll('.pstep').forEach(el => {
            el.classList.remove('dragging');
            el.classList.remove('drag-over');
        });
    }

    /* ---------- 勾选步骤 → 另存为新预设 ---------- */
    function stepsToNewPreset() {
        const p = _presetData.presets[_curPresetIdx];
        const picked = (p.steps || []).filter((s, i) => _stepSel[i]);
        if (!picked.length) { toast('请先勾选要提取的步骤', 'er'); return; }
        const name = prompt('新预设名称：', (p.name || '') + ' - 拆分');
        if (name === null) return;
        const np = {
            id: _rid('p'),
            name: name.trim() || '拆分预设',
            group: p.group || '',
            steps: JSON.parse(JSON.stringify(picked)),
        };
        /* ★ 重生成步骤id + variant id */
        np.steps.forEach((s, i) => {
            s.id = _rid('s');
            s.order = i + 1;
            (s.variants || []).forEach(v => { v.id = _rid('v_'); });
        });
        _presetData.presets.push(np);
        _stepSel = {};
        markDirty(); drawPresetList(); drawPresetEditor();
        toast('✅ 已拆分为新预设「' + np.name + '」（共 ' + picked.length + ' 步）');
    }

    /* ---------- 分组管理 ---------- */
    function manageGroups() {
        if (!_presetData.groups) _presetData.groups = [];
        const right = document.getElementById('presetRight');
        if (!right) return;
        let html = '<div style="max-width:480px"><h3 style="margin-bottom:12px">🏷️ 分组管理</h3>';
        html += '<div style="display:flex;gap:6px;margin-bottom:12px"><input id="newGroupName" placeholder="输入新分组名" style="flex:1;padding:7px 10px;border:1px solid var(--border);border-radius:8px;background:var(--bg);color:var(--text)"><button class="btn btn-p btn-s" onclick="Admin.addGroup()">➕ 添加</button></div>';
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
        const nv = prompt('重命名分组：', old);
        if (!nv || !nv.trim() || nv.trim() === old) return;
        const nn = nv.trim();
        if (_presetData.groups.includes(nn)) { toast('分组已存在', 'er'); return; }
        _presetData.groups[i] = nn;
        (_presetData.presets || []).forEach(p => { if (p.group === old) p.group = nn; });
        markDirty(); manageGroups(); drawPresetList();
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
            <div class="fg"><label>敏感词（逗号隔开）</label><textarea id="ps_sensitive" rows="4">${esc((sec.sensitiveWords || []).join(','))}</textarea></div>
            <div class="fg"><label>钉钉Webhook</label><input id="ps_webhook" value="${esc(sec.alertWebhook || '')}"></div>
            <div class="fr"><div class="fg"><label>报警关键词</label><input id="ps_keyword" value="${esc(sec.alertKeyword || '飞凡警报')}"></div><div class="fg"><label>相似度阈值%</label><input id="ps_sim" type="number" value="${sec.simThreshold || 70}"></div></div>
            <div class="pt"><input type="checkbox" id="ps_guard" ${sec.guard !== false ? 'checked' : ''}><label for="ps_guard">开启GUARD保密前缀</label></div>
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

    /* ---------- 保存 / 导出 / 导入（★递归处理分支） ---------- */
    async function savePresets() {
        toast('加密并保存中...');
        try {
            /* ① 加密所有隐藏指令（默认版 + 全部分支） */
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

            /* ② 深拷贝后递归清理明文（🔴 安全红线） */
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

    /* 🔴 递归剔除 _plain / _dirty（默认版 + 全部分支），绝不让明文出站 */
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
                /* 递归解密 */
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
    /* ==================== 全局设置（含快捷档） ================= */
    /* ========================================================== */
    async function renderConfig(box) {
        box.innerHTML = '<div style="color:var(--text2);padding:20px;text-align:center">加载中...</div>';
        try {
            const data = await apiCall('admin/config/get');
            const cfg = data.config || {};

            let quick = [];
            try { quick = JSON.parse(cfg.quickModels || '[]'); } catch (e) { quick = []; }
            if (!Array.isArray(quick)) quick = [];
            _quickModelDraft = quick;

            box.innerHTML = `<div style="max-width:560px">
                <h4 style="font-size:13px;margin-bottom:12px">⚙️ 全局参数（所有用户生效）</h4>

                <div class="fg">
                    <label>📐 物理打标：每块字数</label>
                    <input id="cfg_chunkSize" type="number" value="${esc(cfg.chunkSize || '300')}" min="50" max="5000">
                    <div style="font-size:11px;color:var(--text2);margin-top:4px">默认300。给AI定位长文用。</div>
                </div>

                <div style="margin:16px 0;padding:12px;background:var(--pri-l);border-radius:8px">
                    <h4 style="font-size:13px;margin-bottom:6px">⚡ 快捷模型档</h4>
                    <div style="font-size:11px;color:var(--text2);margin-bottom:10px;line-height:1.6">
                        用户输入框上方会出现这些按钮，点一下即临时切换模型（<b>不改动引擎配置</b>，切引擎/换对话自动复位）。<br>
                        · 标签：按钮上的字，越短越好（如"快""强""图"）<br>
                        · 模型名：留空=用当前引擎的默认模型<br>
                        · 生图：勾选后点该档即进入生图模式
                    </div>
                    <div id="quickModelList"></div>
                    <button class="btn btn-s" onclick="Admin.addQuickModel()" style="margin-top:6px">➕ 添加一档</button>
                </div>

                <div class="fg">
                    <label>⌨️ 快捷指令（用户在输入框打 "/" 唤起）</label>
                    <textarea id="cfg_quickCmds" rows="6" placeholder="每行一条，格式：名称|内容模板&#10;翻译成英文|请把以下内容翻译成地道的英文：&#10;总结要点|请用要点列出以下内容的核心：">${esc(cfg.quickCmds || '')}</textarea>
                    <div style="font-size:11px;color:var(--text2);margin-top:4px">格式：<code>名称|内容</code>，一行一条。内容里可含换行请用一行写完。</div>
                </div>

                <button class="btn btn-p" onclick="Admin.saveConfig()">💾 保存</button>
            </div>`;

            drawQuickModelList();
        } catch (e) {
            box.innerHTML = '<div style="color:#ef4444;padding:20px">加载失败：' + e.message + '</div>';
        }
    }

    function drawQuickModelList() {
        const box = document.getElementById('quickModelList');
        if (!box) return;
        let html = '';
        _quickModelDraft.forEach((q, i) => {
            html += `<div class="fr" style="margin-bottom:6px;align-items:flex-end">
                <div class="fg" style="flex:0 0 90px;margin-bottom:0"><label>标签</label><input value="${esc(q.label || '')}" onchange="Admin.updQuickModel(${i},'label',this.value)" placeholder="如：快"></div>
                <div class="fg" style="margin-bottom:0"><label>模型名（留空=引擎默认）</label><input value="${esc(q.model || '')}" onchange="Admin.updQuickModel(${i},'model',this.value)" placeholder="如 gpt-4o-mini"></div>
                <div class="fg" style="flex:0 0 52px;margin-bottom:0;text-align:center"><label>生图</label><input type="checkbox" ${q.isImage ? 'checked' : ''} onchange="Admin.updQuickModel(${i},'isImage',this.checked)" style="width:18px;height:18px;accent-color:var(--pri)"></div>
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
            .map(q => ({
                label: (q.label || '').trim(),
                model: (q.model || '').trim(),
                isImage: !!q.isImage
            }));

        apiCall('admin/config/save', 'POST', {
            config: {
                chunkSize: chunkSize,
                quickModels: JSON.stringify(quick),
                quickCmds: quickCmds
            }
        }).then(() => {
            toast('✅ 已保存（用户刷新后生效）');
            if (typeof Chunker !== 'undefined') Chunker.setBlockSize(chunkSize);
            /* 当前管理员页面立即生效 */
            if (typeof window.applyGlobalConfig === 'function') {
                window.applyGlobalConfig({ quickModels: JSON.stringify(quick), quickCmds: quickCmds, chunkSize: chunkSize });
            }
        }).catch(e => toast('失败：' + e.message, 'er'));
    }

    return {
        open, close, switchTab, apiCall,
        /* 账号 */
        showCreateUser, showResetPwd, toggleStatus, delUser, showPerm, searchUsers,
        importXLSX, exportXLSX, downloadTemplate,
        /* 引擎 */
        showEngEdit, saveEng, delEng, autoDetectEngType,
        /* 模型库 */
        showModelEdit, saveModel, delModel,
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
        saveConfig, addQuickModel, delQuickModel, updQuickModel,
    };
})();

window.Admin = Admin;
