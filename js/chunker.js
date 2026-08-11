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

    function _isCountedHan(ch) {
        try { return /[\u4e00-\u9fff\u3400-\u4dbf\u{20000}-\u{2a6df}\u3000-\u303f\uff00-\uffef]/u.test(ch); }
        catch (e) { return /[\u4e00-\u9fff\u3400-\u4dbf\u3000-\u303f\uff00-\uffef]/.test(ch); }
    }

    /* 与 cntW 相同口径扫描词元，并记录完整词元边界，避免把英文单词从中间切开。 */
    function _scanTokens(chars) {
        const tokens = [];
        let i = 0;
        while (i < chars.length) {
            const ch = chars[i];
            if (_isCountedHan(ch)) {
                tokens.push({ start: i, end: i + 1 });
                i++;
                continue;
            }
            if (_isLatinWordChar(ch)) {
                const start = i;
                i++;
                while (i < chars.length) {
                    if (_isLatinWordChar(chars[i])) { i++; continue; }
                    if ((chars[i] === "'" || chars[i] === '’' || chars[i] === '-')
                        && i + 1 < chars.length && _isLatinWordChar(chars[i + 1])) {
                        i += 2;
                        continue;
                    }
                    break;
                }
                tokens.push({ start: start, end: i });
                continue;
            }
            i++;
        }
        return tokens;
    }

    /* ========== 核心：按完整词元单次扫描切块（O(n)） ========== */
    function chunk(text, opts) {
        opts = opts || {};
        const size = Math.max(1, parseInt(opts.size, 10) || DEFAULT_SIZE);
        const doClean = opts.clean !== false;

        const cleaned = doClean ? clean(text) : String(text || '');
        const chars = _chars(cleaned);
        const n = chars.length;
        if (!n) return { total: 0, size: size, blocks: [], marked: '', cleaned: cleaned };

        const tokens = _scanTokens(chars);
        const totalWords = tokens.length;
        if (!totalWords) {
            const only = { no: 1, words: 0, wStart: 0, wEnd: 0, pctStart: 0, pctEnd: 0, text: cleaned };
            return { total: 0, size: size, blocks: [only], marked: _render([only], 0), cleaned: cleaned };
        }

        const blocks = [];
        let charStart = 0;
        for (let tokenStart = 0; tokenStart < totalWords; tokenStart += size) {
            const tokenEnd = Math.min(tokenStart + size, totalWords);
            const charEnd = tokenEnd < totalWords ? tokens[tokenEnd].start : n;
            const blockWords = tokenEnd - tokenStart;
            blocks.push({
                no: blocks.length + 1,
                words: blockWords,
                wStart: tokenStart + 1,
                wEnd: tokenEnd,
                pctStart: +((tokenStart / totalWords) * 100).toFixed(1),
                pctEnd: +((tokenEnd / totalWords) * 100).toFixed(1),
                text: chars.slice(charStart, charEnd).join(''),
            });
            charStart = charEnd;
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
