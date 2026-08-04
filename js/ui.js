/* ===== 飞凡AI - UI 渲染层 (v3.4.0 批次4) ===== */
/* 改动：移除思考气泡 / 代码块·Mermaid·表格改外置工具条（不遮挡）/ 新增节点导航 + 大纲 */

const UI = (function () {

    /* ---------- Markdown 渲染器 ---------- */
    let _md = null;
    function getMD() {
        if (_md) return _md;
        if (typeof markdownit === 'undefined') return null;
        _md = markdownit({
            html: false,
            linkify: true,
            breaks: true,
            highlight: function (str, lang) {
                if (typeof hljs !== 'undefined' && lang && hljs.getLanguage(lang)) {
                    try { return hljs.highlight(str, { language: lang, ignoreIllegals: true }).value; }
                    catch (e) {}
                }
                return '';
            },
        });
        try {
            if (window.markdownitSub) _md.use(window.markdownitSub);
            if (window.markdownitSup) _md.use(window.markdownitSup);
            if (window.markdownitMark) _md.use(window.markdownitMark);
            if (window.markdownitFootnote) _md.use(window.markdownitFootnote);
            if (window.markdownitTaskLists) _md.use(window.markdownitTaskLists, { enabled: false });
            if (window.markdownitEmoji && window.markdownitEmoji.full) _md.use(window.markdownitEmoji.full);
        } catch (e) { console.warn('[MD plugin]', e); }
        return _md;
    }

    function renderMarkdown(text) {
        const md = getMD();
        let html;
        if (md) html = md.render(text || '');
        else html = '<p>' + esc(text) + '</p>';
        if (window.DOMPurify) {
            html = window.DOMPurify.sanitize(html, {
                ADD_TAGS: ['mark', 'sub', 'sup'],
                ADD_ATTR: ['target'],
            });
        }
        return html;
    }

    /* ---------- KaTeX 数学公式 ---------- */
    function renderMath(container) {
        if (!window.katex) return;
        const blockRegex = /\$\$([\s\S]+?)\$\$/g;
        const inlineRegex = /\$([^\$\n]+?)\$/g;
        function processNode(node) {
            if (node.nodeType === Node.TEXT_NODE) {
                const text = node.textContent;
                if (!text.includes('$')) return;
                const parent = node.parentNode;
                if (!parent || parent.tagName === 'CODE' || parent.tagName === 'PRE') return;
                let html = esc(text);
                html = html.replace(blockRegex, (_, expr) => {
                    try { return katex.renderToString(expr, { displayMode: true, throwOnError: false }); }
                    catch (e) { return _; }
                });
                html = html.replace(inlineRegex, (_, expr) => {
                    try { return katex.renderToString(expr, { displayMode: false, throwOnError: false }); }
                    catch (e) { return _; }
                });
                if (html !== esc(text)) {
                    const span = document.createElement('span');
                    span.innerHTML = html;
                    parent.replaceChild(span, node);
                }
            } else if (node.nodeType === Node.ELEMENT_NODE) {
                if (node.tagName === 'CODE' || node.tagName === 'PRE') return;
                Array.from(node.childNodes).forEach(processNode);
            }
        }
        processNode(container);
    }

    /* ========================================================== */
    /* ====== 代码块：外置工具条（不遮挡内容） ================== */
    /* ==========================================================
       结构：
       <div class="blk blk-code">
         <div class="blk-bar">
            <span class="blk-lang">PYTHON</span>
            <span class="blk-spacer"></span>
            <button>📋 复制</button>
            <button>展开全部</button>
         </div>
         <pre><code>...</code></pre>
       </div>
       ========================================================== */
    function wrapCodeBlocks(container) {
        const pres = container.querySelectorAll('pre');
        pres.forEach(pre => {
            if (pre.parentNode && pre.parentNode.classList && pre.parentNode.classList.contains('blk-code')) return;
            const code = pre.querySelector('code');
            if (!code) return;

            const langClass = (code.className || '').match(/language-([\w#+-]+)/);
            const lang = langClass ? langClass[1] : '';
            const text = code.textContent;
            const lineCount = text.split('\n').length;
            const long = lineCount > 16;

            const blk = document.createElement('div');
            blk.className = 'blk blk-code' + (long ? ' blk-collapsed' : '');

            /* --- 工具条 --- */
            const bar = document.createElement('div');
            bar.className = 'blk-bar';

            const langTag = document.createElement('span');
            langTag.className = 'blk-lang';
            langTag.textContent = lang ? lang : 'CODE';
            bar.appendChild(langTag);

            const info = document.createElement('span');
            info.className = 'blk-info';
            info.textContent = lineCount + ' 行';
            bar.appendChild(info);

            const spacer = document.createElement('span');
            spacer.className = 'blk-spacer';
            bar.appendChild(spacer);

            const copyBtn = document.createElement('button');
            copyBtn.className = 'blk-btn';
            copyBtn.textContent = '📋 复制';
            copyBtn.onclick = (e) => {
                e.stopPropagation();
                navigator.clipboard.writeText(text).then(() => {
                    copyBtn.textContent = '✓ 已复制';
                    copyBtn.classList.add('done');
                    setTimeout(() => { copyBtn.textContent = '📋 复制'; copyBtn.classList.remove('done'); }, 1500);
                }).catch(() => toast('复制失败', 'er'));
            };
            bar.appendChild(copyBtn);

            if (long) {
                const tgl = document.createElement('button');
                tgl.className = 'blk-btn';
                tgl.textContent = '⤢ 展开';
                tgl.onclick = (e) => {
                    e.stopPropagation();
                    blk.classList.toggle('blk-collapsed');
                    tgl.textContent = blk.classList.contains('blk-collapsed') ? '⤢ 展开' : '⤡ 收起';
                };
                bar.appendChild(tgl);
            }

            blk.appendChild(bar);

            /* --- 内容 --- */
            pre.parentNode.replaceChild(blk, pre);
            blk.appendChild(pre);
        });
    }

    /* ========================================================== */
    /* ====== 表格：外置工具条 ================================== */
    /* ========================================================== */
    function wrapTables(container) {
        const tables = container.querySelectorAll('table');
        tables.forEach(table => {
            if (table.parentNode && table.parentNode.classList && table.parentNode.classList.contains('blk-table')) return;

            const rowN = table.querySelectorAll('tr').length;

            const blk = document.createElement('div');
            blk.className = 'blk blk-table';

            const bar = document.createElement('div');
            bar.className = 'blk-bar';

            const tag = document.createElement('span');
            tag.className = 'blk-lang';
            tag.textContent = '表格';
            bar.appendChild(tag);

            const info = document.createElement('span');
            info.className = 'blk-info';
            info.textContent = rowN + ' 行';
            bar.appendChild(info);

            const spacer = document.createElement('span');
            spacer.className = 'blk-spacer';
            bar.appendChild(spacer);

            const xlsBtn = document.createElement('button');
            xlsBtn.className = 'blk-btn';
            xlsBtn.textContent = '📊 导出Excel';
            xlsBtn.onclick = (e) => { e.stopPropagation(); exportTableToXlsx(table, xlsBtn); };
            bar.appendChild(xlsBtn);

            const cpBtn = document.createElement('button');
            cpBtn.className = 'blk-btn';
            cpBtn.textContent = '📋 复制';
            cpBtn.onclick = (e) => {
                e.stopPropagation();
                const lines = [];
                table.querySelectorAll('tr').forEach(tr => {
                    const cells = [];
                    tr.querySelectorAll('th,td').forEach(td => cells.push((td.textContent || '').trim()));
                    if (cells.length) lines.push(cells.join('\t'));
                });
                navigator.clipboard.writeText(lines.join('\n')).then(() => {
                    cpBtn.textContent = '✓ 已复制';
                    cpBtn.classList.add('done');
                    setTimeout(() => { cpBtn.textContent = '📋 复制'; cpBtn.classList.remove('done'); }, 1500);
                }).catch(() => toast('复制失败', 'er'));
            };
            bar.appendChild(cpBtn);

            blk.appendChild(bar);

            /* 表格外再包一层横向滚动容器 */
            const scroll = document.createElement('div');
            scroll.className = 'blk-scroll';
            table.parentNode.replaceChild(blk, table);
            scroll.appendChild(table);
            blk.appendChild(scroll);
        });
    }

    function exportTableToXlsx(tableEl, btnEl) {
        const doExport = () => {
            try {
                const rows = [];
                tableEl.querySelectorAll('tr').forEach(tr => {
                    const cells = [];
                    tr.querySelectorAll('th,td').forEach(td => cells.push((td.textContent || '').trim()));
                    if (cells.length) rows.push(cells);
                });
                if (!rows.length) { if (typeof toast === 'function') toast('表格为空', 'er'); return; }
                const ws = XLSX.utils.aoa_to_sheet(rows);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
                XLSX.writeFile(wb, '表格-' + new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19) + '.xlsx');
                if (btnEl) {
                    btnEl.textContent = '✓ 已导出';
                    btnEl.classList.add('done');
                    setTimeout(() => { btnEl.textContent = '📊 导出Excel'; btnEl.classList.remove('done'); }, 1600);
                }
            } catch (e) {
                if (typeof toast === 'function') toast('导出失败：' + e.message, 'er');
            }
        };

        if (window.XLSX) doExport();
        else if (typeof OfficeParser !== 'undefined' && OfficeParser.loadXLSX) {
            if (btnEl) btnEl.textContent = '⏳ 加载中...';
            OfficeParser.loadXLSX().then(doExport).catch(e => {
                if (typeof toast === 'function') toast('加载Excel库失败：' + e.message, 'er');
                if (btnEl) btnEl.textContent = '📊 导出Excel';
            });
        } else {
            if (typeof toast === 'function') toast('Excel库不可用', 'er');
        }
    }

    /* ========================================================== */
    /* ====== Mermaid：外置工具条 =============================== */
    /* ========================================================== */
    function renderMermaid(container) {
        if (!window.mermaid) return;
        const blocks = container.querySelectorAll('pre code.language-mermaid, pre code.language-mmd');
        blocks.forEach((codeEl, i) => {
            const code = codeEl.textContent;

            const blk = document.createElement('div');
            blk.className = 'blk blk-mermaid';

            const bar = document.createElement('div');
            bar.className = 'blk-bar';
            const tag = document.createElement('span');
            tag.className = 'blk-lang';
            tag.textContent = '图表';
            bar.appendChild(tag);
            const spacer = document.createElement('span');
            spacer.className = 'blk-spacer';
            bar.appendChild(spacer);
            const dlBtn = document.createElement('button');
            dlBtn.className = 'blk-btn';
            dlBtn.textContent = '🖼️ 下载PNG';
            dlBtn.disabled = true;
            bar.appendChild(dlBtn);
            const srcBtn = document.createElement('button');
            srcBtn.className = 'blk-btn';
            srcBtn.textContent = '📋 源码';
            srcBtn.onclick = (e) => {
                e.stopPropagation();
                navigator.clipboard.writeText(code).then(() => {
                    srcBtn.textContent = '✓ 已复制';
                    srcBtn.classList.add('done');
                    setTimeout(() => { srcBtn.textContent = '📋 源码'; srcBtn.classList.remove('done'); }, 1500);
                }).catch(() => toast('复制失败', 'er'));
            };
            bar.appendChild(srcBtn);
            blk.appendChild(bar);

            const holder = document.createElement('div');
            holder.className = 'mermaid-holder';
            holder.id = 'mer-' + Date.now() + '-' + i;
            holder.textContent = code;
            blk.appendChild(holder);

            /* 替换原 pre，包裹进外置工具条结构 */
            const pre = codeEl.closest('pre');
            if (pre && pre.parentNode) pre.parentNode.replaceChild(blk, pre);

            const bindDl = () => {
                const svg = holder.querySelector('svg');
                if (!svg) {
                    dlBtn.textContent = '⚠️ 渲染失败';
                    return;
                }
                dlBtn.disabled = false;
                dlBtn.onclick = (e) => { e.stopPropagation(); downloadSvgAsPng(svg, dlBtn); };
            };

            try {
                const r = mermaid.run({ nodes: [holder] });
                if (r && typeof r.then === 'function') {
                    r.then(bindDl).catch(err => {
                        console.warn('[Mermaid]', err);
                        dlBtn.textContent = '⚠️ 语法有误';
                    });
                } else {
                    setTimeout(bindDl, 350);
                }
            } catch (e) {
                console.warn('[Mermaid]', e);
                dlBtn.textContent = '⚠️ 渲染失败';
            }
        });
    }

    function downloadSvgAsPng(svgEl, btnEl) {
        try {
            const clone = svgEl.cloneNode(true);
            clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

            let w = (svgEl.viewBox && svgEl.viewBox.baseVal && svgEl.viewBox.baseVal.width)
                ? svgEl.viewBox.baseVal.width : svgEl.getBoundingClientRect().width;
            let h = (svgEl.viewBox && svgEl.viewBox.baseVal && svgEl.viewBox.baseVal.height)
                ? svgEl.viewBox.baseVal.height : svgEl.getBoundingClientRect().height;
            w = Math.max(w, 100); h = Math.max(h, 100);

            const svgStr = new XMLSerializer().serializeToString(clone);
            const svgBlob = new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' });
            const url = URL.createObjectURL(svgBlob);

            const img = new Image();
            img.onload = () => {
                const scale = 2;
                const canvas = document.createElement('canvas');
                canvas.width = w * scale; canvas.height = h * scale;
                const ctx = canvas.getContext('2d');
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.scale(scale, scale);
                ctx.drawImage(img, 0, 0, w, h);
                URL.revokeObjectURL(url);

                canvas.toBlob((blob) => {
                    if (!blob) { if (typeof toast === 'function') toast('导出失败', 'er'); return; }
                    const a = document.createElement('a');
                    a.href = URL.createObjectURL(blob);
                    a.download = '图表-' + new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19) + '.png';
                    document.body.appendChild(a); a.click(); document.body.removeChild(a);
                    setTimeout(() => URL.revokeObjectURL(a.href), 800);
                    if (btnEl) {
                        btnEl.textContent = '✓ 已下载';
                        btnEl.classList.add('done');
                        setTimeout(() => { btnEl.textContent = '🖼️ 下载PNG'; btnEl.classList.remove('done'); }, 1600);
                    }
                }, 'image/png');
            };
            img.onerror = () => {
                URL.revokeObjectURL(url);
                if (typeof toast === 'function') toast('图表转换失败', 'er');
            };
            img.src = url;
        } catch (e) {
            if (typeof toast === 'function') toast('下载失败：' + e.message, 'er');
        }
    }

    /* ---------- 完整渲染流程 ---------- */
    function fullRender(bubElement, text) {
        bubElement.classList.remove('streaming');
        bubElement.innerHTML = renderMarkdown(text);
        try {
            renderMath(bubElement);
            wrapCodeBlocks(bubElement);
            wrapTables(bubElement);
            renderMermaid(bubElement);
            bubElement.querySelectorAll('img').forEach(img => {
                img.onclick = () => {
                    const lb = document.getElementById('lightbox');
                    const lbImg = document.getElementById('lbImg');
                    if (lb && lbImg) { lbImg.src = img.src; lb.classList.add('show'); }
                };
            });
        } catch (e) { console.warn('[fullRender]', e); }
    }

    function streamRender(bubElement, text) {
        if (!bubElement.classList.contains('streaming')) bubElement.classList.add('streaming');
        bubElement.textContent = text;
    }

    /* ---------- 创建消息 DOM ---------- */
    function createMessageNode(msg, options) {
        const opts = options || {};
        const wrap = document.createElement('div');
        wrap.className = 'msg ' + msg.role;
        wrap.dataset.msgId = msg.id || '';

        if (opts.selectMode) {
            const selBox = document.createElement('div');
            selBox.className = 'msg-sel';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = opts.selectedMsgs && opts.selectedMsgs.includes(msg.id);
            cb.onchange = () => { if (opts.onSelectToggle) opts.onSelectToggle(msg.id, cb.checked); };
            selBox.appendChild(cb);
            wrap.appendChild(selBox);
        }

        const av = document.createElement('div');
        av.className = 'av';
        av.textContent = msg.role === 'user' ? '👤' : '🤖';
        wrap.appendChild(av);

        const m = document.createElement('div');
        m.className = 'm';

        if (msg.attachments && msg.attachments.length) {
            const mf = document.createElement('div');
            mf.className = 'mf';
            msg.attachments.forEach(a => {
                const fb = document.createElement('span');
                fb.className = 'fb';
                fb.textContent = '📎 ' + (a.name || 'file');
                mf.appendChild(fb);
            });
            m.appendChild(mf);
        }

        /* ★ 思考气泡已移除（_reasoning 数据仍保留在消息对象里，不渲染） */

        const bub = document.createElement('div');
        bub.className = 'bub';
        if (msg.role === 'assistant') {
            if (msg._streaming) streamRender(bub, msg.content || '');
            else fullRender(bub, msg.content || '');
        } else {
            bub.textContent = typeof msg.content === 'string' ? msg.content : JSON.stringify(msg.content);
        }
        m.appendChild(bub);

        const mm = document.createElement('div');
        mm.className = 'mm';

        if (msg._time) {
            const t = document.createElement('span');
            t.className = 'msg-time';
            t.textContent = msg._time;
            mm.appendChild(t);
        }

        if (msg.role === 'assistant' && msg.content) {
            const wc = cntW(msg.content);
            if (wc > 0) {
                const w = document.createElement('span');
                w.textContent = wc + ' 字';
                mm.appendChild(w);
            }
        }

        const cpBtn = document.createElement('button');
        cpBtn.textContent = '📋 复制';
        cpBtn.onclick = () => {
            navigator.clipboard.writeText(typeof msg.content === 'string' ? msg.content : JSON.stringify(msg.content))
                .then(() => toast('已复制')).catch(() => toast('复制失败', 'er'));
        };
        mm.appendChild(cpBtn);

        if (opts.onDelete) {
            const dBtn = document.createElement('button');
            dBtn.textContent = '🗑️ 删除';
            dBtn.onclick = () => { if (confirm('删除这条消息？')) opts.onDelete(msg); };
            mm.appendChild(dBtn);
        }

        if (msg.role === 'user' && opts.onEdit) {
            const eBtn = document.createElement('button');
            eBtn.textContent = '✏️ 编辑';
            eBtn.title = '改完后会删除这条及之后所有消息并重新生成';
            eBtn.onclick = () => opts.onEdit(msg);
            mm.appendChild(eBtn);
        }

        if (msg.role === 'assistant' && opts.onRegen) {
            const rBtn = document.createElement('button');
            rBtn.textContent = '🔄 重答';
            rBtn.onclick = () => opts.onRegen(msg);
            mm.appendChild(rBtn);
        }

        if (msg.role === 'assistant' && msg._interrupted && msg.content && opts.onContinue) {
            const cBtn = document.createElement('button');
            cBtn.textContent = '▶ 继续';
            cBtn.title = '接着未写完的内容继续写';
            cBtn.style.color = 'var(--pri)';
            cBtn.onclick = () => opts.onContinue(msg);
            mm.appendChild(cBtn);
        }

        m.appendChild(mm);
        wrap.appendChild(m);
        return { wrap: wrap, bub: bub };
    }

    /* ---------- 渲染整个消息列表 ---------- */
    function renderMessages(container, messages, options) {
        container.innerHTML = '';
        if (!messages || !messages.length) {
            container.innerHTML = '<div class="empty"><div class="ico">🚀</div><p>开始新的对话吧</p>'
                + '<p style="font-size:12px;opacity:.6;margin-top:4px">支持拖拽 / 粘贴上传文档与图片</p></div>';
            return null;
        }
        let lastBub = null;
        messages.forEach(m => {
            const node = createMessageNode(m, options);
            container.appendChild(node.wrap);
            lastBub = node.bub;
        });
        container.scrollTop = container.scrollHeight;
        return lastBub;
    }

    /* ---------- 滚动到底部（节流） ---------- */
    const scrollToBottom = rafThrottle(function (container) {
        container.scrollTop = container.scrollHeight;
    });

    /* ---------- 流式更新（节流 + 智能跟随） ---------- */
    function makeStreamUpdater(bub, container) {
        let lastUpdate = 0;
        let pending = '';
        let timer = null;
        const INTERVAL = 50;
        const NEAR_BOTTOM = 80;
        function isNearBottom() {
            if (!container) return true;
            return (container.scrollHeight - container.scrollTop - container.clientHeight) < NEAR_BOTTOM;
        }
        function flush() {
            if (!pending) return;
            const stick = isNearBottom();
            streamRender(bub, pending);
            if (stick) scrollToBottom(container);
            lastUpdate = Date.now();
            timer = null;
        }
        return function update(fullText) {
            pending = fullText;
            const now = Date.now();
            const elapsed = now - lastUpdate;
            if (elapsed >= INTERVAL) flush();
            else if (!timer) timer = setTimeout(flush, INTERVAL - elapsed);
        };
    }

    /* ========================================================== */
    /* ============ ★ 节点导航（右侧竖条） ====================== */
    /* ========================================================== */
    let _navObserver = null;
    let _navMsgIds = [];

    /* 计算节点高度权重：按内容字数成比例，带上下限 */
    function _weightOf(msg) {
        const txt = typeof msg.content === 'string' ? msg.content : '';
        const n = (typeof cntW === 'function') ? cntW(txt) : txt.length;
        /* 用平方根压缩：避免一条5000字的把其他全挤成1px */
        return Math.max(1, Math.sqrt(Math.max(n, 1)));
    }

    function _previewOf(msg) {
        let t = typeof msg.content === 'string' ? msg.content : '[多媒体内容]';
        t = t.replace(/!\[[^\]]*\]\([^)]*\)/g, '[图]')
             .replace(/```[\s\S]*?```/g, '[代码]')
             .replace(/\s+/g, ' ')
             .trim();
        return t.slice(0, 40) + (t.length > 40 ? '…' : '');
    }

    /* 渲染右侧导航条
       opts: { minCount:3 (少于N条不显示), onJump(msgId) } */
    function renderMsgNav(navEl, msgsEl, messages, opts) {
        opts = opts || {};
        const minCount = opts.minCount || 3;

        if (_navObserver) { try { _navObserver.disconnect(); } catch (e) {} _navObserver = null; }

        if (!navEl || !msgsEl) return;
        if (!messages || messages.length < minCount) {
            navEl.classList.remove('show');
            navEl.innerHTML = '';
            _navMsgIds = [];
            return;
        }

        navEl.classList.add('show');
        navEl.innerHTML = '';
        _navMsgIds = messages.map(m => m.id);

        const weights = messages.map(_weightOf);
        const total = weights.reduce((a, b) => a + b, 0) || 1;

        messages.forEach((m, i) => {
            const node = document.createElement('div');
            node.className = 'nav-node nav-' + (m.role === 'user' ? 'u' : 'a');
            node.dataset.mid = m.id;
            /* flex-grow 按权重分配，minHeight 保底 */
            node.style.flexGrow = String(weights[i]);

            /* 悬停预览 */
            const tipTxt = (m.role === 'user' ? '👤 ' : '🤖 ') + (m._time ? ('[' + m._time + '] ') : '') + _previewOf(m);
            node.title = tipTxt;

            const tip = document.createElement('span');
            tip.className = 'nav-tip';
            tip.textContent = tipTxt;
            node.appendChild(tip);

            node.onclick = () => {
                const target = msgsEl.querySelector('.msg[data-msg-id="' + m.id + '"]');
                if (!target) return;
                /* 滚到消息顶部，留 12px 余量 */
                const top = target.offsetTop - 12;
                try { msgsEl.scrollTo({ top: top, behavior: 'smooth' }); }
                catch (e) { msgsEl.scrollTop = top; }
                /* 闪一下提示位置 */
                target.classList.add('msg-flash');
                setTimeout(() => target.classList.remove('msg-flash'), 900);
                if (opts.onJump) opts.onJump(m.id);
            };

            navEl.appendChild(node);
        });

        /* 视口高亮：IntersectionObserver 原生实现，零性能负担 */
        try {
            _navObserver = new IntersectionObserver((entries) => {
                entries.forEach(en => {
                    const mid = en.target.dataset.msgId;
                    if (!mid) return;
                    const node = navEl.querySelector('.nav-node[data-mid="' + mid + '"]');
                    if (node) node.classList.toggle('in-view', en.isIntersecting);
                });
            }, { root: msgsEl, threshold: 0.01, rootMargin: '-10% 0px -60% 0px' });

            msgsEl.querySelectorAll('.msg').forEach(el => _navObserver.observe(el));
        } catch (e) { /* 老浏览器不支持则跳过高亮 */ }
    }

    /* ---------- 大纲列表（弹窗用） ---------- */
    function renderOutline(listEl, msgsEl, messages, onJumpDone) {
        if (!listEl) return;
        if (!messages || !messages.length) {
            listEl.innerHTML = '<div style="font-size:12px;color:var(--text2);padding:14px;text-align:center">（当前对话无消息）</div>';
            return;
        }
        listEl.innerHTML = '';
        let uN = 0, aN = 0;
        messages.forEach(m => {
            const isU = m.role === 'user';
            if (isU) uN++; else aN++;
            const row = document.createElement('div');
            row.className = 'ol-item ol-' + (isU ? 'u' : 'a');
            const wc = (typeof cntW === 'function' && typeof m.content === 'string') ? cntW(m.content) : 0;
            row.innerHTML = '<span class="ol-badge">' + (isU ? '问' + uN : '答' + aN) + '</span>'
                + '<span class="ol-txt">' + esc(_previewOf(m)) + '</span>'
                + '<span class="ol-meta">' + (wc ? (wc + '字') : '') + (m._time ? (' · ' + esc(m._time)) : '') + '</span>';
            row.onclick = () => {
                const target = msgsEl.querySelector('.msg[data-msg-id="' + m.id + '"]');
                if (target) {
                    const top = target.offsetTop - 12;
                    try { msgsEl.scrollTo({ top: top, behavior: 'smooth' }); }
                    catch (e) { msgsEl.scrollTop = top; }
                    target.classList.add('msg-flash');
                    setTimeout(() => target.classList.remove('msg-flash'), 900);
                }
                if (onJumpDone) onJumpDone();
            };
            listEl.appendChild(row);
        });
    }

    return {
        renderMarkdown: renderMarkdown,
        renderMath: renderMath,
        renderMermaid: renderMermaid,
        wrapCodeBlocks: wrapCodeBlocks,
        wrapTables: wrapTables,
        fullRender: fullRender,
        streamRender: streamRender,
        createMessageNode: createMessageNode,
        renderMessages: renderMessages,
        scrollToBottom: scrollToBottom,
        makeStreamUpdater: makeStreamUpdater,
        /* 批次4 新增 */
        renderMsgNav: renderMsgNav,
        renderOutline: renderOutline,
    };
})();
