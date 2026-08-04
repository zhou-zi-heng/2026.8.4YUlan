/* ===== 飞凡AI - 导出引擎 (v3.3.0 批次3) ===== */
/* 真 docx（docx.js，带标题/列表/表格/代码/图片格式）+ PDF（html2pdf）
   两个库都是按需 CDN 加载，不用时零开销 */

const Exporter = (function () {

    /* ---------- 按需加载 docx.js（双CDN兜底） ---------- */
    let _docxLoading = null;
    function loadDocx() {
        if (window.docx) return Promise.resolve();
        if (_docxLoading) return _docxLoading;
        _docxLoading = new Promise((resolve, reject) => {
            const urls = [
                'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.min.js',
                'https://unpkg.com/docx@8.5.0/build/index.umd.min.js',
                'https://cdnjs.cloudflare.com/ajax/libs/docx/8.5.0/index.umd.min.js',
            ];
            let idx = 0;
            const tryLoad = () => {
                if (idx >= urls.length) { reject(new Error('docx.js 所有 CDN 均加载失败，请检查网络')); return; }
                const s = document.createElement('script');
                s.src = urls[idx++];
                s.onload = () => window.docx ? resolve() : tryLoad();
                s.onerror = () => tryLoad();
                document.head.appendChild(s);
            };
            tryLoad();
        });
        return _docxLoading;
    }

    /* ---------- 按需加载 html2pdf（双CDN兜底） ---------- */
    let _pdfLoading = null;
    function loadPdf() {
        if (window.html2pdf) return Promise.resolve();
        if (_pdfLoading) return _pdfLoading;
        _pdfLoading = new Promise((resolve, reject) => {
            const urls = [
                'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js',
                'https://cdn.jsdelivr.net/npm/html2pdf.js@0.10.1/dist/html2pdf.bundle.min.js',
            ];
            let idx = 0;
            const tryLoad = () => {
                if (idx >= urls.length) { reject(new Error('html2pdf 所有 CDN 均加载失败，请检查网络')); return; }
                const s = document.createElement('script');
                s.src = urls[idx++];
                s.onload = () => window.html2pdf ? resolve() : tryLoad();
                s.onerror = () => tryLoad();
                document.head.appendChild(s);
            };
            tryLoad();
        });
        return _pdfLoading;
    }

    /* ========== 简易 Markdown 解析（逐行，供 docx 用） ========== */
    function parseMarkdownToBlocks(md) {
        const lines = String(md || '').split('\n');
        const blocks = [];
        let i = 0;

        while (i < lines.length) {
            const line = lines[i];

            /* 代码块 ``` */
            if (/^\s*```/.test(line)) {
                const codeLines = [];
                i++;
                while (i < lines.length && !/^\s*```/.test(lines[i])) { codeLines.push(lines[i]); i++; }
                i++;
                blocks.push({ type: 'code', text: codeLines.join('\n') });
                continue;
            }

            /* 表格（连续以 | 开头的行） */
            if (/^\s*\|.*\|\s*$/.test(line)) {
                const tableLines = [];
                while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) { tableLines.push(lines[i]); i++; }
                blocks.push({ type: 'table', rows: parseTableLines(tableLines) });
                continue;
            }

            /* 图片 ![alt](dataUrl 或 http) */
            const imgM = line.match(/^\s*!\[[^\]]*\]\((data:image\/[^)]+|https?:\/\/[^)]+)\)/);
            if (imgM) {
                blocks.push({ type: 'image', src: imgM[1] });
                i++;
                continue;
            }

            /* 标题 */
            const hM = line.match(/^(#{1,6})\s+(.*)$/);
            if (hM) { blocks.push({ type: 'heading', level: hM[1].length, text: hM[2] }); i++; continue; }

            /* 分割线 */
            if (/^\s*([-*_])\s*\1\s*\1[\s\-*_]*$/.test(line)) { blocks.push({ type: 'hr' }); i++; continue; }

            /* 引用 */
            if (/^\s*>\s?/.test(line)) { blocks.push({ type: 'quote', text: line.replace(/^\s*>\s?/, '') }); i++; continue; }

            /* 无序列表 */
            const ulM = line.match(/^\s*[-*+]\s+(.*)$/);
            if (ulM) { blocks.push({ type: 'listItem', ordered: false, text: ulM[1] }); i++; continue; }

            /* 有序列表 */
            const olM = line.match(/^\s*(\d+)\.\s+(.*)$/);
            if (olM) { blocks.push({ type: 'listItem', ordered: true, num: olM[1], text: olM[2] }); i++; continue; }

            /* 空行 */
            if (line.trim() === '') { blocks.push({ type: 'empty' }); i++; continue; }

            /* 普通段落 */
            blocks.push({ type: 'paragraph', text: line });
            i++;
        }
        return blocks;
    }

    function parseTableLines(lines) {
        const rows = [];
        lines.forEach(ln => {
            if (/^\s*\|[\s:\-|]+\|\s*$/.test(ln)) return;  // 跳过分隔行
            const cells = ln.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => c.trim());
            rows.push(cells);
        });
        return rows;
    }

    /* ---------- 行内格式（加粗/斜体/行内代码/删除线）→ docx TextRun[] ---------- */
    function parseInlineRuns(text, d, baseOpts) {
        const base = baseOpts || {};
        const font = base.font || 'Microsoft YaHei';
        const size = base.size || 24;       // 半磅：24 = 12pt
        const color = base.color;
        const italics = base.italics;

        const runs = [];
        let remaining = String(text || '');
        const re = /(\*\*([^*]+)\*\*)|(\*([^*]+)\*)|(`([^`]+)`)|(~~([^~]+)~~)/;
        let m;

        const mk = (t, extra) => new d.TextRun(Object.assign({ text: t, font: font, size: size, color: color, italics: italics }, extra || {}));

        while ((m = re.exec(remaining)) !== null) {
            const before = remaining.slice(0, m.index);
            if (before) runs.push(mk(before));

            if (m[2] !== undefined) runs.push(mk(m[2], { bold: true }));
            else if (m[4] !== undefined) runs.push(mk(m[4], { italics: true }));
            else if (m[6] !== undefined) runs.push(mk(m[6], { font: 'Consolas', size: 22, shading: { fill: 'F0F0F0' } }));
            else if (m[8] !== undefined) runs.push(mk(m[8], { strike: true }));

            remaining = remaining.slice(m.index + m[0].length);
        }
        if (remaining) runs.push(mk(remaining));
        if (!runs.length) runs.push(mk(''));
        return runs;
    }

    /* dataURL / http → Uint8Array（http 需 fetch，可能被 CORS 拦） */
    function dataUrlToUint8(dataUrl) {
        const m = String(dataUrl).match(/^data:[^;]+;base64,(.+)$/);
        if (!m) return null;
        try {
            const bin = atob(m[1]);
            const bytes = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
            return bytes;
        } catch (e) { return null; }
    }

    async function fetchImageBytes(url) {
        try {
            const resp = await fetch(url);
            if (!resp.ok) return null;
            const buf = await resp.arrayBuffer();
            return new Uint8Array(buf);
        } catch (e) { return null; }
    }

    /* ========== 生成真 docx ========== */
    /* sections: [{roleLabel:'👤 我 (10:23)', content:'markdown文本'}] */
    async function exportDocx(sections, title) {
        await loadDocx();
        const d = window.docx;
        const children = [];

        /* 文档大标题 */
        children.push(new d.Paragraph({
            children: [new d.TextRun({ text: title || '对话记录', bold: true, size: 36, font: 'Microsoft YaHei' })],
            spacing: { after: 120 },
        }));
        children.push(new d.Paragraph({
            children: [new d.TextRun({ text: '导出时间：' + new Date().toLocaleString(), size: 18, color: '888888', font: 'Microsoft YaHei' })],
            spacing: { after: 240 },
        }));

        for (const sec of sections) {
            /* 角色标签 */
            if (sec.roleLabel) {
                children.push(new d.Paragraph({
                    children: [new d.TextRun({ text: sec.roleLabel, bold: true, size: 26, font: 'Microsoft YaHei', color: '667EEA' })],
                    spacing: { before: 240, after: 80 },
                    border: { bottom: { color: 'E5E7EB', space: 2, style: d.BorderStyle.SINGLE, size: 6 } },
                }));
            }

            const blocks = parseMarkdownToBlocks(sec.content);

            for (const b of blocks) {
                if (b.type === 'heading') {
                    const lv = b.level <= 1 ? d.HeadingLevel.HEADING_1
                        : b.level === 2 ? d.HeadingLevel.HEADING_2
                        : d.HeadingLevel.HEADING_3;
                    children.push(new d.Paragraph({
                        children: parseInlineRuns(b.text, d, { size: b.level <= 2 ? 28 : 26 }),
                        heading: lv,
                        spacing: { before: 160, after: 70 },
                    }));

                } else if (b.type === 'paragraph') {
                    children.push(new d.Paragraph({
                        children: parseInlineRuns(b.text, d),
                        spacing: { after: 90, line: 320 },
                    }));

                } else if (b.type === 'listItem') {
                    if (b.ordered) {
                        children.push(new d.Paragraph({
                            children: [new d.TextRun({ text: b.num + '. ', font: 'Microsoft YaHei', size: 24 })].concat(parseInlineRuns(b.text, d)),
                            indent: { left: 420 },
                            spacing: { after: 50 },
                        }));
                    } else {
                        children.push(new d.Paragraph({
                            children: parseInlineRuns(b.text, d),
                            bullet: { level: 0 },
                            spacing: { after: 50 },
                        }));
                    }

                } else if (b.type === 'quote') {
                    children.push(new d.Paragraph({
                        children: parseInlineRuns(b.text, d, { color: '666666', italics: true }),
                        indent: { left: 420 },
                        spacing: { after: 90 },
                        border: { left: { color: '667EEA', space: 8, style: d.BorderStyle.SINGLE, size: 12 } },
                    }));

                } else if (b.type === 'code') {
                    b.text.split('\n').forEach(cl => {
                        children.push(new d.Paragraph({
                            children: [new d.TextRun({ text: cl || ' ', font: 'Consolas', size: 20 })],
                            shading: { fill: 'F6F8FA' },
                            spacing: { after: 0, line: 260 },
                        }));
                    });
                    children.push(new d.Paragraph({ text: '', spacing: { after: 90 } }));

                } else if (b.type === 'hr') {
                    children.push(new d.Paragraph({
                        text: '',
                        border: { bottom: { color: 'CCCCCC', space: 1, style: d.BorderStyle.SINGLE, size: 6 } },
                        spacing: { before: 90, after: 120 },
                    }));

                } else if (b.type === 'table') {
                    if (b.rows.length) {
                        const colN = Math.max.apply(null, b.rows.map(r => r.length));
                        const tableRows = b.rows.map((row, ri) => {
                            const cells = row.slice();
                            while (cells.length < colN) cells.push('');
                            return new d.TableRow({
                                children: cells.map(cell => new d.TableCell({
                                    children: [new d.Paragraph({
                                        children: parseInlineRuns(cell, d, { size: 22 }).map(r => { if (ri === 0) r.bold = true; return r; }),
                                    })],
                                    shading: ri === 0 ? { fill: 'EEF0FF' } : undefined,
                                    margins: { top: 60, bottom: 60, left: 90, right: 90 },
                                })),
                            });
                        });
                        children.push(new d.Table({
                            rows: tableRows,
                            width: { size: 100, type: d.WidthType.PERCENTAGE },
                        }));
                        children.push(new d.Paragraph({ text: '', spacing: { after: 120 } }));
                    }

                } else if (b.type === 'image') {
                    let bytes = dataUrlToUint8(b.src);
                    if (!bytes && /^https?:\/\//.test(b.src)) bytes = await fetchImageBytes(b.src);
                    if (bytes) {
                        try {
                            children.push(new d.Paragraph({
                                children: [new d.ImageRun({ data: bytes, transformation: { width: 420, height: 315 } })],
                                spacing: { after: 120 },
                            }));
                        } catch (e) {
                            children.push(new d.Paragraph({ children: [new d.TextRun({ text: '[图片插入失败]', color: '999999', size: 20, font: 'Microsoft YaHei' })] }));
                        }
                    } else {
                        children.push(new d.Paragraph({ children: [new d.TextRun({ text: '[图片无法读取，可能是跨域链接]', color: '999999', size: 20, font: 'Microsoft YaHei' })] }));
                    }
                }
                /* empty 不处理，段落间已有 spacing */
            }
        }

        const doc = new d.Document({
            sections: [{
                properties: { page: { margin: { top: 1200, bottom: 1200, left: 1200, right: 1200 } } },
                children: children,
            }],
        });

        const blob = await d.Packer.toBlob(doc);
        _downloadBlob(blob, _safeName(title) + '-' + _ts() + '.docx');
    }

    /* ========== 生成 PDF（html2pdf 拍照 HTML） ========== */
    async function exportPdf(html, title) {
        await loadPdf();
        const wrap = document.createElement('div');
        wrap.style.cssText = 'position:fixed;left:-99999px;top:0;width:760px;padding:24px;'
            + 'font-family:"Microsoft YaHei",-apple-system,sans-serif;line-height:1.75;color:#222;background:#fff;';
        wrap.innerHTML = '<h1 style="font-size:21px;margin-bottom:4px">' + _esc(title || '对话记录') + '</h1>'
            + '<div style="font-size:11px;color:#999;margin-bottom:16px">导出时间：' + _esc(new Date().toLocaleString()) + '</div>'
            + '<style>'
            + 'pre{background:#f6f8fa;border-radius:6px;padding:10px;overflow:visible;white-space:pre-wrap;word-break:break-all;font-size:12px}'
            + 'code{font-family:Consolas,monospace;font-size:12px}'
            + 'table{border-collapse:collapse;width:100%;font-size:12px;margin:8px 0}'
            + 'th,td{border:1px solid #ddd;padding:5px 8px;text-align:left}'
            + 'th{background:#eef0ff}'
            + 'img{max-width:100%;height:auto}'
            + 'blockquote{border-left:3px solid #667eea;padding-left:10px;color:#666;margin:8px 0}'
            + 'h1,h2,h3{page-break-after:avoid}'
            + '</style>'
            + html;
        document.body.appendChild(wrap);
        try {
            await window.html2pdf().set({
                margin: [10, 10, 12, 10],
                filename: _safeName(title) + '-' + _ts() + '.pdf',
                image: { type: 'jpeg', quality: 0.95 },
                html2canvas: { scale: 2, useCORS: true, logging: false, windowWidth: 760 },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
                pagebreak: { mode: ['avoid-all', 'css'] },
            }).from(wrap).save();
        } finally {
            document.body.removeChild(wrap);
        }
    }

    /* ---------- 工具 ---------- */
    function _ts() { return new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19); }
    function _safeName(t) {
        return String(t || 'chat').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim().slice(0, 50) || 'chat';
    }
    function _esc(t) { const d = document.createElement('div'); d.textContent = String(t == null ? '' : t); return d.innerHTML; }
    function _downloadBlob(blob, filename) {
        const u = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = u; a.download = filename;
        document.body.appendChild(a); a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(u), 1000);
    }

    return {
        exportDocx: exportDocx,
        exportPdf: exportPdf,
        parseMarkdownToBlocks: parseMarkdownToBlocks,
        loadDocx: loadDocx,
        loadPdf: loadPdf,
    };
})();

window.Exporter = Exporter;
