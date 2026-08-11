/* ===== 飞凡AI - 快照系统 (v2.3.5) ===== */
/* 自动覆盖式快照 + 导入导出 + 全版本兼容 + 智能 key 保护 + 加密分享对话 */

const Snapshot = (function () {

    let _autoTimer = null;
    let _lastSnapHash = '';

    /* ==========================================================
       ===== 加密相关常量 =======================================
       ========================================================== */
    // 内置应用密钥（无感层）。普通人扒不到，达成"外人打开是乱码"的目标。
    // ⚠️ 如需更高安全，分享时叠加用户口令（见 shareChat 的 password 参数）。
    const APP_SECRET = 'FeiFanAI-2026-Sx9#kQ2$mZ7&pL4!vR8@nW3^bT6*cY1%hG5';
    const PBKDF2_ITERATIONS = 100000;
    const SUPPORTS_CRYPTO = !!(window.crypto && window.crypto.subtle);

    function _fallbackHash(str) {
        let h = 2166136261;
        for (let i = 0; i < str.length; i++) {
            h ^= str.charCodeAt(i);
            h = Math.imul(h, 16777619);
        }
        return 'f:' + (h >>> 0).toString(16) + ':' + str.length;
    }

    async function contentHash(obj) {
        const str = JSON.stringify(obj);
        if (SUPPORTS_CRYPTO) {
            const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
            return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
        }
        return _fallbackHash(str);
    }

    function startAuto(intervalMin, getStateFn) {
        stopAuto();
        const min = parseInt(intervalMin, 10) || 0;
        if (min <= 0) {
            console.log('[Snapshot] 自动快照已关闭');
            return;
        }
        const ms = min * 60 * 1000;
        console.log('[Snapshot] 自动快照启动，间隔 ' + min + ' 分钟');
        _autoTimer = setInterval(async () => {
            try {
                const state = getStateFn();
                const h = await contentHash(state);
                if (h === _lastSnapHash) {
                    console.log('[Snapshot] 数据无变化，跳过');
                    return;
                }
                await DB.saveAutoSnapshot(state);
                _lastSnapHash = h;
                console.log('[Snapshot] 自动快照完成 @ ' + new Date().toLocaleTimeString());
            } catch (e) {
                console.error('[Snapshot] 自动快照失败', e);
            }
        }, ms);
    }

    function stopAuto() {
        if (_autoTimer) {
            clearInterval(_autoTimer);
            _autoTimer = null;
        }
    }

    async function snapNow(state) {
        await DB.saveAutoSnapshot(state);
        _lastSnapHash = await contentHash(state);
    }

    /* ---------- 导出快照（支持不带 key） ---------- */
    function exportToFile(state, options) {
        const opts = options || {};
        const includeKey = opts.includeKey !== false;

        const data = JSON.parse(JSON.stringify(state));

        if (!includeKey) {
            if (data.profiles) {
                for (const id in data.profiles) {
                    if (data.profiles[id]) data.profiles[id].key = '';
                }
            }
        }

        const wrap = {
            __feifan_snapshot__: true,
            version: APP_VERSION,
            exportedAt: new Date().toISOString(),
            includeKey: includeKey,
            data: data,
        };
        const json = JSON.stringify(wrap, null, 2);
        const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        const suffix = includeKey ? '' : '-nokey';
        dl(json, 'feifan-backup-' + ts + suffix + '.json', 'application/json');
        toast(includeKey ? '✅ 已导出快照（含 API Key）' : '✅ 已导出快照（不含 API Key）');
    }

    /* ---------- 解析快照（兼容所有历史版本） ---------- */
    function detectAndNormalize(raw) {
        if (!raw || typeof raw !== 'object') {
            throw new Error('快照内容无效');
        }

        if (raw.__feifan_snapshot__ && raw.data) {
            return { state: normalizeCurrentState(raw.data), source: 'feifan-v' + (raw.version || '?') };
        }

        if (raw.chats && typeof raw.chats === 'object') {
            return { state: normalizeOldState(raw), source: 'legacy-direct' };
        }

        if (raw.conversations && typeof raw.conversations === 'object') {
            const fixed = Object.assign({}, raw, { chats: raw.conversations });
            delete fixed.conversations;
            return { state: normalizeOldState(fixed), source: 'legacy-conversations' };
        }

        if (raw.profiles && typeof raw.profiles === 'object') {
            const fixed = Object.assign({ chats: {} }, raw);
            return { state: normalizeOldState(fixed), source: 'legacy-profiles-only' };
        }

        if (Array.isArray(raw.messages)) {
            const id = (raw.id) || ('imp_' + Date.now());
            const chat = {
                id: id,
                title: raw.title || '导入的对话',
                messages: raw.messages,
                systemPrompt: raw.systemPrompt || '',
                knowledgeBase: raw.knowledgeBase || [],
                isPinned: !!raw.isPinned,
                isArchived: !!raw.isArchived,
                createdAt: raw.createdAt || Date.now(),
                updatedAt: raw.updatedAt || Date.now(),
            };
            return {
                state: normalizeOldState({ chats: { [id]: chat }, currentChatId: id }),
                source: 'legacy-single-chat',
            };
        }

        throw new Error('无法识别此快照格式');
    }

    function normalizeOldState(old) {
        const n = {
            profiles: {},
            chats: {},
            chatOrder: [],
            currentChatId: null,
            currentEngId: 'zenmux',
            theme: 'light',
            snapInterval: 5,
            archiveInterval: 10,
            userName: '',
            uiMode: 'chat',
            defaultMode: 'free',
            folders: [],
            folderCollapsed: {},
            _engineMigrated: false,
        };

        if (old.profiles && typeof old.profiles === 'object') {
            for (const k in old.profiles) {
                const p = old.profiles[k] || {};
                n.profiles[k] = {
                    id: p.id || k,
                    name: p.name || k,
                    base: p.base || p.baseUrl || p.endpoint || 'https://api.openai.com/v1',
                    key: p.key || p.apiKey || '',
                    model: p.model || p.modelName || 'gpt-4o-mini',
                    useTemp: p.useTemp !== false,
                    temperature: p.temperature !== undefined ? p.temperature : 0.7,
                    useMax: p.useMax !== false,
                    max_tokens: p.max_tokens || p.maxTokens || 4096,
                    useTopP: !!p.useTopP,
                    top_p: p.top_p !== undefined ? p.top_p : 1,
                    useFreq: !!p.useFreq,
                    frequency_penalty: p.frequency_penalty !== undefined ? p.frequency_penalty : 0,
                };
            }
        }

        if (old.chats && typeof old.chats === 'object') {
            for (const k in old.chats) {
                const c = old.chats[k] || {};
                const messages = Array.isArray(c.messages) ? c.messages.map(normalizeOldMessage).filter(Boolean) : [];
                n.chats[k] = {
                    id: c.id || k,
                    title: c.title || '导入的对话',
                    messages: messages,
                    systemPrompt: c.systemPrompt || c.system || '',
                    knowledgeBase: Array.isArray(c.knowledgeBase) ? c.knowledgeBase :
                                   (Array.isArray(c.kb) ? c.kb : []),
                    isPinned: !!(c.isPinned || c.pinned),
                    isArchived: !!(c.isArchived || c.archived),
                    refPool: Array.isArray(c.refPool) ? c.refPool : [],
                    folderId: c.folderId || null,
                    mode: c.mode || 'free',
                    modeLocked: !!c.modeLocked,
                    createdAt: c.createdAt || c.created || Date.now(),
                    updatedAt: c.updatedAt || c.updated || Date.now(),
                };
            }
        }

        if (Array.isArray(old.chatOrder)) {
            n.chatOrder = old.chatOrder.filter(id => n.chats[id]);
        }
        const missingIds = Object.keys(n.chats).filter(id => !n.chatOrder.includes(id));
        missingIds.sort((a, b) => (n.chats[b].updatedAt || 0) - (n.chats[a].updatedAt || 0));
        n.chatOrder = n.chatOrder.concat(missingIds);

        n.currentChatId = old.currentChatId && n.chats[old.currentChatId]
            ? old.currentChatId
            : (n.chatOrder[0] || null);
        n.currentEngId = old.currentEngId || (Object.keys(n.profiles)[0] || 'zenmux');
        n.theme = old.theme || 'light';
        n.snapInterval = (old.snapInterval !== undefined) ? old.snapInterval : 5;
        n.archiveInterval = (old.archiveInterval !== undefined) ? old.archiveInterval : 10;
        n.userName = old.userName || '';
        n.uiMode = old.uiMode || 'chat';
        n.defaultMode = old.defaultMode || 'free';
        n.folders = Array.isArray(old.folders) ? old.folders : [];
        n.folderCollapsed = old.folderCollapsed && typeof old.folderCollapsed === 'object' ? old.folderCollapsed : {};
        n._engineMigrated = !!old._engineMigrated;

        return n;
    }

    function normalizeCurrentState(raw) {
        let src;
        try { src = JSON.parse(JSON.stringify(raw || {})); }
        catch (e) { throw new Error('快照状态无法序列化'); }
        const n = Object.assign({
            profiles: {}, chats: {}, chatOrder: [], currentChatId: null, currentEngId: '',
            theme: 'light', snapInterval: 5, archiveInterval: 10, userName: '',
            uiMode: 'chat', defaultMode: 'free', folders: [], folderCollapsed: {},
            _engineMigrated: false,
        }, src);
        if (!n.profiles || typeof n.profiles !== 'object' || Array.isArray(n.profiles)) n.profiles = {};
        if (!n.chats || typeof n.chats !== 'object' || Array.isArray(n.chats)) n.chats = {};
        if (!Array.isArray(n.folders)) n.folders = [];
        if (!n.folderCollapsed || typeof n.folderCollapsed !== 'object') n.folderCollapsed = {};
        for (const id of Object.keys(n.chats)) {
            const c = n.chats[id];
            if (!c || typeof c !== 'object') { delete n.chats[id]; continue; }
            c.id = c.id || id;
            c.title = c.title || '导入的对话';
            c.messages = Array.isArray(c.messages) ? c.messages.map(m => {
                if (!m || typeof m !== 'object') return null;
                return Object.assign({}, m, {
                    id: m.id || gId(), role: m.role || 'user',
                    content: m.content !== undefined ? m.content : '',
                    _streaming: false, _interrupted: !!m._interrupted,
                });
            }).filter(Boolean) : [];
            c.knowledgeBase = Array.isArray(c.knowledgeBase) ? c.knowledgeBase : [];
            c.refPool = Array.isArray(c.refPool) ? c.refPool : [];
            c.folderId = c.folderId || null;
        }
        const seen = new Set();
        n.chatOrder = (Array.isArray(n.chatOrder) ? n.chatOrder : []).filter(id => {
            if (!n.chats[id] || seen.has(id)) return false;
            seen.add(id); return true;
        });
        Object.keys(n.chats).forEach(id => { if (!seen.has(id)) n.chatOrder.push(id); });
        if (!n.currentChatId || !n.chats[n.currentChatId]) n.currentChatId = n.chatOrder[0] || null;
        if (!n.currentEngId) n.currentEngId = Object.keys(n.profiles)[0] || '';
        return n;
    }

    function normalizeOldMessage(m) {
        if (!m || typeof m !== 'object') return null;
        return Object.assign({}, m, {
            id: m.id || gId(),
            role: m.role || 'user',
            content: m.content !== undefined ? m.content : (m.text || ''),
            attachments: Array.isArray(m.attachments) ? m.attachments : [],
            _time: m._time || m.time || '',
            _streaming: false,
            _interrupted: !!m._interrupted,
        });
    }

    /* ---------- 智能 key 保护 ---------- */
    function protectLocalKeys(incoming, current) {
        if (!incoming || !incoming.profiles || !current || !current.profiles) {
            return { state: incoming, protectedCount: 0 };
        }
        const result = JSON.parse(JSON.stringify(incoming));
        let protectedCount = 0;
        for (const id in result.profiles) {
            const incP = result.profiles[id];
            const curP = current.profiles[id];
            if (curP && curP.key && (!incP.key || incP.key.trim() === '')) {
                incP.key = curP.key;
                protectedCount++;
            }
        }
        if (protectedCount > 0) {
            console.log('[Snapshot] 已保护 ' + protectedCount + ' 个本地 API Key');
        }
        return { state: result, protectedCount: protectedCount };
    }

    async function importFromFile(file) {
        return new Promise((resolve, reject) => {
            const r = new FileReader();
            r.onload = async (e) => {
                try {
                    const raw = JSON.parse(e.target.result);
                    const { state, source } = detectAndNormalize(raw);
                    console.log('[Snapshot] 导入来源:', source);
                    resolve({ state, source });
                } catch (err) {
                    reject(new Error('解析快照失败：' + err.message));
                }
            };
            r.onerror = () => reject(new Error('文件读取失败'));
            r.readAsText(file, 'utf-8');
        });
    }

    function mergeStates(current, incoming) {
        const merged = JSON.parse(JSON.stringify(current));
        merged.profiles = merged.profiles || {};
        merged.chats = merged.chats || {};
        merged.folders = Array.isArray(merged.folders) ? merged.folders : [];
        const suffix = Date.now().toString(36);
        let seq = 0;
        const uniqueId = (base, exists) => {
            let id;
            do { id = String(base || 'item') + '_import_' + suffix + '_' + (++seq); } while (exists(id));
            return id;
        };

        const profileMap = {};
        for (const id in (incoming.profiles || {})) {
            const p = JSON.parse(JSON.stringify(incoming.profiles[id] || {}));
            let targetId = id;
            if (merged.profiles[targetId]) {
                if (JSON.stringify(merged.profiles[targetId]) === JSON.stringify(p)) { profileMap[id] = targetId; continue; }
                targetId = uniqueId(id, x => !!merged.profiles[x]);
            }
            p.id = targetId;
            merged.profiles[targetId] = p;
            profileMap[id] = targetId;
        }

        const folderMap = {};
        for (const folder of (incoming.folders || [])) {
            if (!folder || !folder.id) continue;
            const copy = JSON.parse(JSON.stringify(folder));
            let targetId = copy.id;
            const existing = merged.folders.find(f => f && f.id === targetId);
            if (existing) {
                if ((existing.name || '') === (copy.name || '')) { folderMap[copy.id] = targetId; continue; }
                targetId = uniqueId(copy.id, x => merged.folders.some(f => f && f.id === x));
            }
            folderMap[copy.id] = targetId;
            copy.id = targetId;
            merged.folders.push(copy);
        }

        const chatMap = {};
        for (const id in (incoming.chats || {})) {
            const chat = JSON.parse(JSON.stringify(incoming.chats[id] || {}));
            let targetId = id;
            if (merged.chats[targetId]) {
                if (JSON.stringify(merged.chats[targetId]) === JSON.stringify(chat)) { chatMap[id] = targetId; continue; }
                targetId = uniqueId(id, x => !!merged.chats[x]);
            }
            chat.id = targetId;
            if (chat.folderId && folderMap[chat.folderId]) chat.folderId = folderMap[chat.folderId];
            (chat.messages || []).forEach(m => {
                if (m && m._engId && profileMap[m._engId]) m._engId = profileMap[m._engId];
            });
            merged.chats[targetId] = chat;
            chatMap[id] = targetId;
        }
        const seen = new Set();
        const newOrder = [];
        (incoming.chatOrder || []).forEach(id => {
            const mapped = chatMap[id] || id;
            if (merged.chats[mapped] && !seen.has(mapped)) { newOrder.push(mapped); seen.add(mapped); }
        });
        (current.chatOrder || []).forEach(id => {
            if (merged.chats[id] && !seen.has(id)) { newOrder.push(id); seen.add(id); }
        });
        Object.keys(merged.chats).forEach(id => {
            if (!seen.has(id)) { newOrder.push(id); seen.add(id); }
        });
        merged.chatOrder = newOrder;
        return merged;
    }

    /* ==========================================================
       ===== WebCrypto 加密工具 (v2.3.5) ========================
       ========================================================== */

    // ArrayBuffer <-> Base64
    function _ab2b64(buf) {
        const bytes = new Uint8Array(buf);
        let bin = '';
        for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
        return btoa(bin);
    }
    function _b642ab(b64) {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        return bytes.buffer;
    }

    // 从口令字符串派生 AES-GCM 密钥
    async function _deriveKey(passphrase, salt) {
        const enc = new TextEncoder();
        const baseKey = await crypto.subtle.importKey(
            'raw', enc.encode(passphrase),
            { name: 'PBKDF2' }, false, ['deriveKey']
        );
        return crypto.subtle.deriveKey(
            { name: 'PBKDF2', salt: salt, iterations: PBKDF2_ITERATIONS, hash: 'SHA-256' },
            baseKey,
            { name: 'AES-GCM', length: 256 },
            false, ['encrypt', 'decrypt']
        );
    }

    /* ---------- 加密分享数据 ---------- */
    // password: 可选用户口令（叠加在应用密钥之上）
    async function encryptShareData(plainObj, password) {
        if (!SUPPORTS_CRYPTO) throw new Error('当前浏览器不支持加密（WebCrypto）');

        const json = JSON.stringify(plainObj);
        const enc = new TextEncoder();

        // 组合口令：应用密钥 +（用户口令）
        const passphrase = APP_SECRET + (password ? ('::' + password) : '');

        const salt = crypto.getRandomValues(new Uint8Array(16));
        const iv = crypto.getRandomValues(new Uint8Array(12));
        const key = await _deriveKey(passphrase, salt);

        const cipherBuf = await crypto.subtle.encrypt(
            { name: 'AES-GCM', iv: iv },
            key,
            enc.encode(json)
        );

        return {
            __feifan_enc__: true,
            version: APP_VERSION,
            alg: 'AES-GCM',
            hasPassword: !!password,
            salt: _ab2b64(salt),
            iv: _ab2b64(iv),
            cipher: _ab2b64(cipherBuf),
            encryptedAt: new Date().toISOString(),
        };
    }

    /* ---------- 解密分享数据 ---------- */
    async function decryptShareData(encObj, password) {
        if (!SUPPORTS_CRYPTO) throw new Error('当前浏览器不支持解密（WebCrypto）');
        if (!encObj || !encObj.__feifan_enc__) throw new Error('不是加密分享文件');

        const passphrase = APP_SECRET + (password ? ('::' + password) : '');
        const salt = new Uint8Array(_b642ab(encObj.salt));
        const iv = new Uint8Array(_b642ab(encObj.iv));
        const cipherBuf = _b642ab(encObj.cipher);

        const key = await _deriveKey(passphrase, salt);

        let plainBuf;
        try {
            plainBuf = await crypto.subtle.decrypt(
                { name: 'AES-GCM', iv: iv },
                key,
                cipherBuf
            );
        } catch (e) {
            // 解密失败 = 口令错误或文件损坏
            throw new Error(encObj.hasPassword ? '口令错误或文件损坏' : '文件损坏或非本应用生成');
        }

        const json = new TextDecoder().decode(plainBuf);
        return JSON.parse(json);
    }

    /* ==========================================================
       ===== 分享对话功能 (v2.3.5：支持加密) ====================
       ========================================================== */

    /* ---------- 构建分享 payload（明文对象） ---------- */
    function _buildSharePayload(chat, opts) {
        const chatCopy = JSON.parse(JSON.stringify(chat));
        if (chatCopy.messages) {
            chatCopy.messages.forEach(function (m) { m._streaming = false; });
        }
        if (!opts.includeKB) chatCopy.knowledgeBase = [];

        return {
            __feifan_share__: true,
            version: APP_VERSION,
            sharedAt: new Date().toISOString(),
            sharedBy: opts.sharedBy || '',
            chat: chatCopy,
        };
    }

    /* ---------- 分享导出（加密 / 明文） ---------- */
    // options: { includeKB, sharedBy, encrypt(默认true), password(可选) }
    async function shareChat(chat, options) {
        const opts = options || {};
        if (!chat) { toast('请先选择一个对话', 'er'); return; }

        const payload = _buildSharePayload(chat, opts);
        const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        const safeName = (chat.title || 'chat').replace(/[^\w\u4e00-\u9fff-]/g, '_').slice(0, 30);

        const doEncrypt = opts.encrypt !== false; // 默认加密

        if (doEncrypt && SUPPORTS_CRYPTO) {
            try {
                const encObj = await encryptShareData(payload, opts.password || '');
                const json = JSON.stringify(encObj, null, 2);
                dl(json, safeName + '-' + ts + '.feifan-enc.json', 'application/json');
                toast(opts.password
                    ? '✅ 已导出加密分享（含口令），对方需输入口令打开'
                    : '✅ 已导出加密分享，仅飞凡AI用户可打开');
                return;
            } catch (e) {
                console.error('[shareChat] 加密失败，降级明文', e);
                toast('加密失败，已降级为明文分享：' + e.message, 'er');
            }
        }

        // 明文降级
        const json = JSON.stringify(payload, null, 2);
        dl(json, safeName + '-' + ts + '.feifan-share.json', 'application/json');
        toast('✅ 对话已导出为分享文件（明文）');
    }

    /* ---------- 从已解析对象导入分享（明文/加密统一入口） ---------- */
    // raw: 已 JSON.parse 的对象；password: 解密口令（如需要）
    async function normalizeSharedObject(raw, password) {
        let payload = raw;

        // 加密文件 → 先解密
        if (raw && raw.__feifan_enc__) {
            payload = await decryptShareData(raw, password || '');
        }

        if (!payload || !payload.__feifan_share__ || !payload.chat) {
            throw new Error('不是有效的分享对话文件');
        }

        const chat = payload.chat;
        chat.id = gId();
        chat.title = (chat.title || '分享的对话') + ' (分享)';
        chat.updatedAt = Date.now();
        chat.isPinned = false;
        chat.isArchived = false;

        if (chat.messages) {
            chat.messages.forEach(function (m) {
                if (!m.id) m.id = gId();
                m._streaming = false;
            });
        } else {
            chat.messages = [];
        }
        if (!chat.knowledgeBase) chat.knowledgeBase = [];
        if (!chat.systemPrompt) chat.systemPrompt = '';

        return {
            chat: chat,
            source: 'feifan-share-v' + (payload.version || '?'),
            sharedAt: payload.sharedAt,
            sharedBy: payload.sharedBy,
        };
    }

    /* ---------- 导入分享的对话（从文件，自动识别明文/加密） ---------- */
    // password: 可选；若加密文件带口令而未提供，会抛出 NEED_PASSWORD 错误
    function importSharedChat(file, password) {
        return new Promise(function (resolve, reject) {
            var r = new FileReader();
            r.onload = async function (e) {
                try {
                    var raw = JSON.parse(e.target.result);

                    // 加密文件且需要口令但没给 → 通知上层弹窗
                    if (raw && raw.__feifan_enc__ && raw.hasPassword && !password) {
                        var err = new Error('NEED_PASSWORD');
                        err.code = 'NEED_PASSWORD';
                        reject(err);
                        return;
                    }

                    var result = await normalizeSharedObject(raw, password);
                    resolve(result);
                } catch (err) {
                    reject(new Error('解析分享文件失败：' + err.message));
                }
            };
            r.onerror = function () { reject(new Error('文件读取失败')); };
            r.readAsText(file, 'utf-8');
        });
    }

    /* ---------- 检测文件类型（share / enc / snapshot / unknown） ---------- */
    function detectFileType(file) {
        return new Promise(function (resolve) {
            var r = new FileReader();
            r.onload = function (e) {
                try {
                    var raw = JSON.parse(e.target.result);
                    if (raw.__feifan_enc__) {
                        resolve(raw.hasPassword ? 'enc-pwd' : 'enc');
                    } else if (raw.__feifan_share__) {
                        resolve('share');
                    } else if (raw.__feifan_snapshot__ || raw.chats || raw.profiles || raw.conversations) {
                        resolve('snapshot');
                    } else {
                        resolve('unknown');
                    }
                } catch (err) {
                    resolve('unknown');
                }
            };
            r.onerror = function () { resolve('unknown'); };
            r.readAsText(file, 'utf-8');
        });
    }

    /* ---------- 返回公开方法 ---------- */
    return {
        startAuto: startAuto,
        stopAuto: stopAuto,
        snapNow: snapNow,
        exportToFile: exportToFile,
        importFromFile: importFromFile,
        mergeStates: mergeStates,
        protectLocalKeys: protectLocalKeys,
        detectAndNormalize: detectAndNormalize,
        // 分享 + 加密 (v2.3.5)
        shareChat: shareChat,
        importSharedChat: importSharedChat,
        normalizeSharedObject: normalizeSharedObject,
        detectFileType: detectFileType,
        encryptShareData: encryptShareData,
        decryptShareData: decryptShareData,
        SUPPORTS_CRYPTO: SUPPORTS_CRYPTO,
        // 给 archive.js 复用：构建明文 payload
        _buildSharePayload: _buildSharePayload,
    };
})();

window.Snapshot = Snapshot;
