/* ===== 飞凡AI - Office 解析（docx / xlsx / xls） v3.1 ===== */
/* docx: JSZip 抠 document.xml 正文（更全，接近WPS） + mammoth 兜底 */

const OfficeParser = (function () {

    /* ---------- 按需加载 SheetJS ---------- */
    let _xlsxLoading = null;
    function loadXLSX() {
        if (window.XLSX) return Promise.resolve();
        if (_xlsxLoading) return _xlsxLoading;
        _xlsxLoading = new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
            s.onload = () => { if (window.XLSX) resolve(); else reject(new Error('SheetJS 加载后未注入')); };
            s.onerror = () => {
                const s2 = document.createElement('script');
                s2.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
                s2.onload = () => window.XLSX ? resolve() : reject(new Error('SheetJS 加载失败'));
                s2.onerror = () => reject(new Error('SheetJS CDN 全部失败'));
                document.head.appendChild(s2);
            };
            document.head.appendChild(s);
        });
        return _xlsxLoading;
    }

    /* ---------- 按需加载 JSZip ---------- */
    let _jszipLoading = null;
    function loadJSZip() {
        if (window.JSZip) return Promise.resolve();
        if (_jszipLoading) return _jszipLoading;
        _jszipLoading = new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
            s.onload = () => { if (window.JSZip) resolve(); else reject(new Error('JSZip 加载后未注入')); };
            s.onerror = () => {
                const s2 = document.createElement('script');
                s2.src = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
                s2.onload = () => window.JSZip ? resolve() : reject(new Error('JSZip 加载失败'));
                s2.onerror = () => reject(new Error('JSZip CDN 全部失败'));
                document.head.appendChild(s2);
            };
            document.head.appendChild(s);
        });
        return _jszipLoading;
    }

    /* ---------- 从 document.xml 抠"正文"文字 ----------
       只取正文段落 <w:p>，其中的文字节点 <w:t>；段落之间加换行。
       换行符 <w:br/> 和制表 <w:tab/> 做基本还原。
       不含文本框(<w:txbxContent>可选跳过)、页眉页脚、脚注尾注(不同xml文件)。 */
    function extractDocxText(xmlStr) {
        // 用浏览器 DOMParser 解析 XML
        let doc;
        try {
            doc = new DOMParser().parseFromString(xmlStr, 'application/xml');
        } catch (e) { return null; }
        if (!doc) return null;
        // 解析出错检测
        if (doc.getElementsByTagName('parsererror').length) return null;

        // 命名空间下，标签名可能带前缀 w:，用 localName 匹配更稳
        const body = doc.getElementsByTagName('*');
        // 收集所有段落 <w:p>
        const paras = [];
        const all = doc.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.localName === 'p') paras.push(el);
        }

        const lines = [];
        paras.forEach(p => {
            // 跳过文本框内的段落（txbxContent 内的 p）
            let inTextbox = false;
            let anc = p.parentNode;
            while (anc && anc.localName) {
                if (anc.localName === 'txbxContent') { inTextbox = true; break; }
                anc = anc.parentNode;
            }
            if (inTextbox) return;

            // 遍历段落内后代，拼文字
            let line = '';
            const nodes = p.getElementsByTagName('*');
            for (let k = 0; k < nodes.length; k++) {
                const node = nodes[k];
                const ln = node.localName;
                if (ln === 't') {
                    line += node.textContent || '';
                } else if (ln === 'tab') {
                    line += '\t';
                } else if (ln === 'br' || ln === 'cr') {
                    line += '\n';
                }
            }
            lines.push(line);
        });

        return lines.join('\n');
    }

    /* ---------- DOCX 解析：JSZip 优先，失败回退 mammoth ---------- */
    async function parseDocx(file) {
        const buf = await file.arrayBuffer();

        // ① 尝试 JSZip 抠正文
        try {
            await loadJSZip();
            const zip = await window.JSZip.loadAsync(buf);
            const docXmlFile = zip.file('word/document.xml');
            if (docXmlFile) {
                const xmlStr = await docXmlFile.async('string');
                const text = extractDocxText(xmlStr);
                if (text && text.trim()) {
                    return {
                        type: 'document',
                        fileName: file.name,
                        text: text,
                        meta: { ext: 'docx', engine: 'jszip' },
                    };
                }
            }
        } catch (e) {
            console.warn('[OfficeParser] JSZip 解析失败，回退 mammoth:', e.message);
        }

        // ② 回退 mammoth
        if (!window.mammoth) throw new Error('JSZip 解析失败，且 mammoth 未加载');
        const result = await window.mammoth.extractRawText({ arrayBuffer: buf });
        return {
            type: 'document',
            fileName: file.name,
            text: result.value || '',
            meta: { ext: 'docx', engine: 'mammoth', warnings: (result.messages || []).length },
        };
    }

    /* ---------- XLSX / XLS 解析 → Markdown 表格 ---------- */
    async function parseExcel(file, ext) {
        await loadXLSX();
        const buf = await file.arrayBuffer();
        const wb = window.XLSX.read(buf, { type: 'array' });
        let combined = '';
        const sheetInfos = [];
        wb.SheetNames.forEach(name => {
            const sheet = wb.Sheets[name];
            const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            const clean = rows.filter(r => r.some(c => c !== '' && c !== null && c !== undefined));
            sheetInfos.push({ name: name, rows: clean.length });
            if (clean.length === 0) return;
            const md = CSVParser.rowsToMarkdown(clean.map(r => r.map(c => String(c))));
            combined += '### 📋 工作表：' + name + '\n\n' + md + '\n\n';
        });
        return {
            type: 'table',
            fileName: file.name,
            text: combined.trim() || '（空文件）',
            meta: { ext: ext, sheets: sheetInfos },
        };
    }

    return {
        parseDocx: parseDocx,
        parseExcel: parseExcel,
        loadXLSX: loadXLSX,
        loadJSZip: loadJSZip,
    };
})();

window.OfficeParser = OfficeParser;
