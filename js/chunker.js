/* ===== 飞凡AI - 物理分块打标引擎 (v3.0 · 字数口径统一 · 一把尺到底) ===== */
/* 全站唯一计数标准 = cntW（utils.js，与消息气泡同源）。
   切块：逐字累加"字数"到 DEFAULT_SIZE 即切一块，每块字数精确（末块除外）。
   显示：只有【块号 + 本块字数 + 全文百分比】，绝不出现"字符/字符位置"。
   全文总字数 = 各块字数之和，同一算法，无任何歧义。 */

const Chunker = (function () {

    let DEFAULT_SIZE = 300;   // ★ 每块目标"字数"（cntW口径），非字符数

    function setBlockSize(n) {
        n = parseInt(n, 10);
        if (n && n >= 50 && n <= 5000) DEFAULT_SIZE = n;
    }
    function getBlockSize() { return DEFAULT_SIZE; }

    /* ---------- 字符工具 ---------- */
    function _chars(s) { return [...String(s || '')]; }
    function _isLatinWordChar(ch) { return /[A-Za-z0-9]/.test(ch); }

    /* ---------- 唯一计数标准：复用全局 cntW，保证与气泡永远同源 ---------- */
    function _wc(text) {
        if (typeof cntW === 'function') return cntW(text);
        // 兜底（与 utils.js 的 cntW 同款算法）
        if (!text) return 0;
        const s = String(text);
        let han;
        try { han = (s.match(/[\u4e00-\u9fff\u3400-\u4dbf\u{20000}-\u{2a6df}\u3000-\u303f\uff00-\uffef]/gu) || []).length; }
        catch (e) { han = (s.match(/[\u4e00-\u9fff\u3400-\u4dbf\u3000-\u303f\uff00-\uffef]/g) || []).length; }
        const words = (s.match(/[a-zA-Z0-9]+(?:['’\-][a-zA-Z0-9]+)*/g) || []).length;
        return han + words;
    }

    /* ========== 净化：去无意义空格 + 清特殊字符 ========== */
    function clean(text) {
        let s = String(text || '');
        s = s.replace(/[\uFEFF\u200B\u200C\u200D\u2060\u180E]/g, '');
        s = s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '');
        s = s.replace(/\u3000/g, ' ');

        const arr = _chars(s);
        const out = [];
        for (let i = 0; i < arr.length; i++) {
            const ch = arr[i];
            if (ch === ' ' || ch === '\t') {
                let prev = '';
                for (let j = out.length - 1; j >= 0; j--) {
                    if (out[j] !== ' ' && out[j] !== '\t') { prev = out[j]; break; }
                    if (out[j] === '\n') { prev = '\n'; break; }
                }
                let next = '';
                for (let k = i + 1; k < arr.length; k++) {
                    if (arr[k] !== ' ' && arr[k] !== '\t') { next = arr[k]; break; }
                    if (arr[k] === '\n') { next = '\n'; break; }
                }
                const keep = _isLatinWordChar(prev) && _isLatinWordChar(next);
                if (keep) out.push(' ');
            } else {
                out.push(ch);
            }
        }
        s = out.join('');
        s = s.replace(/\n[ \t]*\n[ \t\n]*/g, '\n\n');
        s = s.replace(/[ \t]+\n/g, '\n').replace(/\n[ \t]+/g, '\n');
        s = s.trim();
        return s;
    }

    /* ========== 核心：按"字数"逐字累加切块 ==========
       逐字符扫描，实时用 cntW 累计"字数"，够 size 就切一块。
       因为 cntW 里英文/数字连续算1词，切块时按"字符边界"切，
       但用 cntW 复算每块字数 —— 保证"每块字数=各块之和=全文"绝对成立。 */
    function chunk(text, opts) {
        opts = opts || {};
        const size = opts.size || DEFAULT_SIZE;
        const doClean = opts.clean !== false;

        const cleaned = doClean ? clean(text) : String(text || '');
        const chars = _chars(cleaned);
        const n = chars.length;

        const totalWords = _wc(cleaned);   // 全文总字数（唯一标准）
        if (!n) return { total: 0, size: size, blocks: [], marked: '', cleaned: cleaned };

        const blocks = [];
        let idx = 1;
        let segStart = 0;              // 当前块起始字符下标
        let cumWordsBefore = 0;        // 已完成块的累计字数

        let i = 0;
        while (i < n) {
            // 从 segStart 向后推进，直到本段 cntW 达到/超过 size，或到文末
            // 用"试探"方式：每推进一个字符，复算本段字数
            let j = i;
            // 每步至少前进1字符，避免死循环
            j++;
            const curSeg = chars.slice(segStart, j).join('');
            const curWords = _wc(curSeg);

            if (curWords >= size || j >= n) {
                // 收一块
                const body = chars.slice(segStart, j).join('');
                const blockWords = _wc(body);
                const wStart = cumWordsBefore + 1;
                const wEnd = cumWordsBefore + blockWords;
                const pctStart = totalWords ? +((cumWordsBefore / totalWords) * 100).toFixed(1) : 0;
                const pctEnd = totalWords ? +((wEnd / totalWords) * 100).toFixed(1) : 0;

                blocks.push({
                    no: idx++,
                    words: blockWords,
                    wStart: wStart,
                    wEnd: wEnd,
                    pctStart: pctStart,
                    pctEnd: pctEnd,
                    text: body,
                });
                cumWordsBefore = wEnd;
                segStart = j;
                i = j;
            } else {
                i = j;
            }
        }

        return {
            total: totalWords,
            size: size,
            blocks: blocks,
            marked: _render(blocks, totalWords),
            cleaned: cleaned,
        };
    }

    /* ---------- 渲染标记文本（全字数口径，只有一种计数） ---------- */
    function _render(blocks, totalWords) {
        const avg = blocks.length ? Math.round(totalWords / blocks.length) : 0;

        let out = '=== 文档分块索引（权威字数数据，禁止自行估算）===\n' +
            '【重要规则】本文档由系统精确切分，以下所有"字数""百分比"均为系统实测准确值，' +
            '与常用文字软件（WPS/Word）的"字数"口径一致。\n' +
            '当你需要说明某段情节的位置、字数或占比时，必须直接引用下方标记中的现成数字，' +
            '严禁自己数字、估算或换算——你的估算一定不准，标记里的数字才是唯一标准。\n' +
            '· 全文精确总字数：' + totalWords + ' 字。\n' +
            '· 共 ' + blocks.length + ' 块，每块约 ' + avg + ' 字。\n' +
            '· 引用规则：情节起于「块X」、止于「块Y」，其字数 = 块Y末字 − 块X首字 + 1；占比直接取两端标记百分比。\n' +
            '· 若情节在某块中间开始/结束，就近取该块边界，并注明"约"。\n\n';

        out += '=== 全文进度速查表（直接查，不要算）===\n';
        const milestones = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100];
        milestones.forEach(pct => {
            const wordNo = Math.round(totalWords * pct / 100);
            let inBlock = blocks.length;
            for (const b of blocks) { if (wordNo <= b.wEnd) { inBlock = b.no; break; } }
            out += '· 全文 ' + pct + '% ≈ 第 ' + wordNo + ' 字（位于块' + inBlock + '）\n';
        });
        out += '\n';

        blocks.forEach(b => {
            out += '▌块' + b.no + '｜首字' + b.wStart + '·末字' + b.wEnd +
                '｜本块' + b.words + '字｜全文进度' + b.pctStart + '%→' + b.pctEnd + '%\n' +
                b.text + '\n\n';
        });
        return out;
    }

    /* ---------- 打标一组附件 ---------- */
    function chunkAttachments(atts, opts) {
        return atts.map(a => {
            if (!a.text || a.type === 'image') return a;
            const r = chunk(a.text, opts);
            return Object.assign({}, a, {
                text: r.marked,
                _chunked: true,
                _chunkInfo: { total: r.total, blocks: r.blocks.length }
            });
        });
    }

    /* ---------- 预览单个附件对象 ---------- */
    function previewOne(att) {
        if (!att || !att.text || att.type === 'image') {
            return '[该附件为图片或无文本，不参与打标]';
        }
        const r = chunk(att.text, {});
        const info = '【文件：' + (att.fileName || att.name || '未命名') +
            '｜全文总字数：' + r.total + ' 字｜分 ' + r.blocks.length + ' 块】\n\n';
        return info + r.marked;
    }

    return {
        chunk: chunk,
        clean: clean,
        chunkAttachments: chunkAttachments,
        previewOne: previewOne,
        setBlockSize: setBlockSize,
        getBlockSize: getBlockSize,
        DEFAULT_SIZE: DEFAULT_SIZE,
    };
})();

window.Chunker = Chunker;
