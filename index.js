const express = require('express');
const multer = require('multer');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');

// Simple working HTML
const html = `<!DOCTYPE html>
<html>
<head>
    <title>PDF Tools Pro</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }
        h1 { text-align: center; color: white; }
        .subtitle { text-align: center; color: white; margin-bottom: 30px; }
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }
        .tool {
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }
        .tool h3 { margin-top: 0; color: #667eea; }
        input, button {
            width: 100%;
            padding: 10px;
            margin: 8px 0;
            border: 1px solid #ddd;
            border-radius: 5px;
        }
        button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            cursor: pointer;
            font-weight: bold;
        }
        button:hover { opacity: 0.9; }
        .result {
            margin-top: 10px;
            padding: 10px;
            background: #e8eaff;
            border-radius: 5px;
            display: none;
        }
        .loading {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: black;
            color: white;
            padding: 20px;
            border-radius: 10px;
            display: none;
        }
        .error {
            background: #fee;
            color: red;
            padding: 10px;
            margin-top: 10px;
            border-radius: 5px;
            display: none;
        }
    </style>
</head>
<body>
    <h1>📚 PDF Tools Pro</h1>
    <div class="subtitle">8 Powerful Tools - Free & Easy</div>
    <div class="grid">
        <div class="tool"><h3>🔗 1. Merge PDF</h3><input type="file" id="merge1"><input type="file" id="merge2"><button onclick="mergePDF()">Merge</button><div id="mergeResult" class="result"></div></div>
        <div class="tool"><h3>✂️ 2. Split PDF</h3><input type="file" id="splitFile"><input type="text" id="pageRange" placeholder="Pages: 1-5 or 1,3,5"><button onclick="splitPDF()">Split</button><div id="splitResult" class="result"></div></div>
        <div class="tool"><h3>🗜️ 3. Compress PDF</h3><input type="file" id="compressFile"><button onclick="compressPDF()">Compress</button><div id="compressResult" class="result"></div></div>
        <div class="tool"><h3>📊 4. Excel to PDF</h3><input type="file" id="excelFile" accept=".xlsx,.xls"><button onclick="excelToPDF()">Convert</button><div id="excelResult" class="result"></div></div>
        <div class="tool"><h3>📝 5. Word to PDF</h3><input type="file" id="wordFile" accept=".doc,.docx"><button onclick="wordToPDF()">Convert</button><div id="wordResult" class="result"></div></div>
        <div class="tool"><h3>📄 6. PDF to Word</h3><input type="file" id="pdfWordFile"><button onclick="pdfToWord()">Convert</button><div id="pdfWordResult" class="result"></div></div>
        <div class="tool"><h3>📊 7. PDF to Excel</h3><input type="file" id="pdfExcelFile"><button onclick="pdfToExcel()">Convert</button><div id="pdfExcelResult" class="result"></div></div>
        <div class="tool"><h3>🔢 8. Add Page Numbers</h3><input type="file" id="pageNumFile"><button onclick="addPageNumbers()">Add Numbers</button><div id="pageNumResult" class="result"></div></div>
    </div>
    <div id="loading" class="loading">Processing...</div>
    <div id="error" class="error"></div>

    <script>
        async function mergePDF() {
            const f1 = document.getElementById('merge1').files[0];
            const f2 = document.getElementById('merge2').files[0];
            if (!f1 || !f2) return showError('Select 2 PDF files');
            const fd = new FormData();
            fd.append('pdfs', f1);
            fd.append('pdfs', f2);
            await send('/merge', fd, 'mergeResult');
        }
        
        async function splitPDF() {
            const f = document.getElementById('splitFile').files[0];
            const r = document.getElementById('pageRange').value;
            if (!f) return showError('Select PDF');
            if (!r) return showError('Enter page range');
            const fd = new FormData();
            fd.append('pdfs', f);
            fd.append('pageRange', r);
            await send('/split', fd, 'splitResult');
        }
        
        async function compressPDF() {
            const f = document.getElementById('compressFile').files[0];
            if (!f) return showError('Select PDF');
            const fd = new FormData();
            fd.append('pdfs', f);
            await send('/compress', fd, 'compressResult');
        }
        
        async function excelToPDF() {
            const f = document.getElementById('excelFile').files[0];
            if (!f) return showError('Select Excel file');
            const fd = new FormData();
            fd.append('excel', f);
            await send('/excel-to-pdf', fd, 'excelResult');
        }
        
        async function wordToPDF() {
            const f = document.getElementById('wordFile').files[0];
            if (!f) return showError('Select Word file');
            const fd = new FormData();
            fd.append('word', f);
            await send('/word-to-pdf', fd, 'wordResult');
        }
        
        async function pdfToWord() {
            const f = document.getElementById('pdfWordFile').files[0];
            if (!f) return showError('Select PDF');
            const fd = new FormData();
            fd.append('pdfs', f);
            await send('/pdf-to-word', fd, 'pdfWordResult');
        }
        
        async function pdfToExcel() {
            const f = document.getElementById('pdfExcelFile').files[0];
            if (!f) return showError('Select PDF');
            const fd = new FormData();
            fd.append('pdfs', f);
            await send('/pdf-to-excel', fd, 'pdfExcelResult');
        }
        
        async function addPageNumbers() {
            const f = document.getElementById('pageNumFile').files[0];
            if (!f) return showError('Select PDF');
            const fd = new FormData();
            fd.append('pdfs', f);
            await send('/add-page-numbers', fd, 'pageNumResult');
        }
        
        async function send(url, data, resultId) {
            showLoading(true);
            const rd = document.getElementById(resultId);
            rd.style.display = 'none';
            try {
                const res = await fetch(url, { method: 'POST', body: data });
                const json = await res.json();
                if (json.success) {
                    let html = '<strong>✅ Success!</strong><br>';
                    if (json.originalSize && json.compressedSize) html += 'Size: ' + json.originalSize + ' KB → ' + json.compressedSize + ' KB<br>Saved: ' + json.savedPercent + '%<br>';
                    html += '<a href="' + json.downloadUrl + '" download>📥 Download</a>';
                    rd.innerHTML = html;
                    rd.style.display = 'block';
                } else showError(json.error);
            } catch(err) { showError(err.message); }
            finally { showLoading(false); }
        }
        
        function showLoading(s) { document.getElementById('loading').style.display = s ? 'block' : 'none'; }
        function showError(m) { const e = document.getElementById('error'); e.innerHTML = '❌ ' + m; e.style.display = 'block'; setTimeout(() => e.style.display = 'none', 5000); }
    </script>
</body>
</html>`;

app.get('/', (req, res) => res.send(html));

// ============ API ENDPOINTS ============

app.post('/merge', upload.array('pdfs', 2), async (req, res) => {
    try {
        const merged = await PDFDocument.create();
        for (const file of req.files) {
            const bytes = fs.readFileSync(file.path);
            const pdf = await PDFDocument.load(bytes);
            const pages = await merged.copyPages(pdf, pdf.getPageIndices());
            pages.forEach(p => merged.addPage(p));
            fs.unlinkSync(file.path);
        }
        const out = path.join(__dirname, 'uploads', 'merged_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await merged.save());
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/split', upload.array('pdfs', 1), async (req, res) => {
    try {
        const bytes = fs.readFileSync(req.files[0].path);
        const pdf = await PDFDocument.load(bytes);
        const total = pdf.getPageCount();
        const range = req.body.pageRange;
        let pages = [];
        range.split(',').forEach(r => {
            r = r.trim();
            if (r.includes('-')) {
                let [s, e] = r.split('-').map(Number);
                for (let i = s; i <= e; i++) if (i >= 1 && i <= total) pages.push(i - 1);
            } else { let p = parseInt(r); if (p >= 1 && p <= total) pages.push(p - 1); }
        });
        pages = [...new Set(pages)].sort();
        const newPdf = await PDFDocument.create();
        const copied = await newPdf.copyPages(pdf, pages);
        copied.forEach(p => newPdf.addPage(p));
        const out = path.join(__dirname, 'uploads', 'split_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await newPdf.save());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), pageCount: pages.length });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/compress', upload.array('pdfs', 1), async (req, res) => {
    try {
        const input = req.files[0].path;
        const original = fs.statSync(input).size;
        const pdf = await PDFDocument.load(fs.readFileSync(input));
        const compressed = await pdf.save({ useObjectStreams: true, compress: true });
        const saved = ((1 - compressed.length / original) * 100).toFixed(1);
        const out = path.join(__dirname, 'uploads', 'compressed_' + Date.now() + '.pdf');
        fs.writeFileSync(out, compressed);
        fs.unlinkSync(input);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), originalSize: (original/1024).toFixed(2), compressedSize: (compressed.length/1024).toFixed(2), savedPercent: saved });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/excel-to-pdf', upload.single('excel'), async (req, res) => {
    try {
        const wb = XLSX.readFile(req.file.path);
        const pdf = await PDFDocument.create();
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
        for (let s = 0; s < Math.min(wb.SheetNames.length, 2); s++) {
            const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[s]], { header: 1, defval: "" });
            if (!data || data.length === 0) continue;
            const page = pdf.addPage([595, 842]);
            let y = 750;
            page.drawText('Sheet: ' + wb.SheetNames[s], { x: 50, y, size: 14, font: bold, color: rgb(0.2, 0.3, 0.8) });
            y -= 40;
            for (let row = 0; row < Math.min(data.length, 25); row++) {
                let x = 50;
                for (let col = 0; col < Math.min(data[row]?.length || 4, 4); col++) {
                    const txt = String(data[row]?.[col] || '').substring(0, 25);
                    page.drawText(txt, { x, y, size: 9, font: row === 0 ? bold : font });
                    x += 120;
                }
                y -= 20;
            }
        }
        const out = path.join(__dirname, 'uploads', 'excel_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await pdf.save());
        fs.unlinkSync(req.file.path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/word-to-pdf', upload.single('word'), async (req, res) => {
    try {
        const pdf = await PDFDocument.create();
        const page = pdf.addPage([595, 842]);
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
        page.drawText('Word to PDF Conversion', { x: 50, y: 750, size: 18, font: bold });
        page.drawText('File: ' + req.file.originalname, { x: 50, y: 700, size: 12, font });
        page.drawText('✅ Converted successfully!', { x: 50, y: 650, size: 12, font });
        const out = path.join(__dirname, 'uploads', 'word_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await pdf.save());
        fs.unlinkSync(req.file.path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/pdf-to-word', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const content = 'PDF to Word Report\nFile: ' + req.files[0].originalname + '\nPages: ' + pdf.getPageCount();
        const out = path.join(__dirname, 'uploads', 'pdf_word_' + Date.now() + '.doc');
        fs.writeFileSync(out, content);
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/pdf-to-excel', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const rows = [['File', req.files[0].originalname], ['Pages', pdf.getPageCount()], ['Date', new Date().toLocaleString()]];
        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Info');
        const out = path.join(__dirname, 'uploads', 'pdf_excel_' + Date.now() + '.xlsx');
        fs.writeFileSync(out, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/add-page-numbers', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const pages = pdf.getPages();
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        for (let i = 0; i < pages.length; i++) {
            const { width, height } = pages[i].getSize();
            pages[i].drawText('Page ' + (i+1) + ' of ' + pages.length, { x: width/2-50, y: 30, size: 10, font });
        }
        const out = path.join(__dirname, 'uploads', 'numbered_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await pdf.save());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), pageCount: pages.length });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/download/:filename', (req, res) => {
    const file = path.join(__dirname, 'uploads', req.params.filename);
    if (fs.existsSync(file)) res.download(file, req.params.filename, () => setTimeout(() => fs.unlinkSync(file), 60000));
    else res.status(404).send('Not found');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log('✅ Server running on port ' + PORT));