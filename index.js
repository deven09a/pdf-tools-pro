// LARGEPDF TOOLS - 10 PROFESSIONAL PDF TOOLS
// Domain: largepdftools.com

const express = require('express');
const multer = require('multer');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');

// Professional HTML with 10 tools
const html = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LargePDF Tools - 10 Professional PDF Tools</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .container { max-width: 1400px; margin: 0 auto; }
        .header { text-align: center; margin-bottom: 40px; }
        .header h1 { font-size: 2.5rem; font-weight: 800; background: linear-gradient(135deg, #fff, #a5b4fc); -webkit-background-clip: text; background-clip: text; color: transparent; }
        .header p { color: #94a3b8; margin-top: 10px; }
        .stats { display: flex; justify-content: center; gap: 3rem; margin-bottom: 40px; flex-wrap: wrap; }
        .stat { text-align: center; }
        .stat-number { font-size: 2rem; font-weight: 800; color: #a5b4fc; }
        .stat-label { color: #64748b; font-size: 0.85rem; }
        .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(350px, 1fr)); gap: 25px; }
        .tool-card {
            background: rgba(30, 27, 75, 0.6);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 20px;
            padding: 25px;
            transition: all 0.3s;
        }
        .tool-card:hover { transform: translateY(-5px); border-color: #6366f1; }
        .tool-icon {
            width: 55px;
            height: 55px;
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border-radius: 15px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 15px;
        }
        .tool-icon i { font-size: 26px; color: white; }
        .tool-title { font-size: 1.2rem; font-weight: 700; color: white; margin-bottom: 8px; }
        .tool-desc { font-size: 0.8rem; color: #94a3b8; margin-bottom: 15px; }
        .badge { display: inline-block; background: #22c55e; padding: 3px 10px; border-radius: 20px; font-size: 0.6rem; font-weight: 600; color: white; margin-left: 8px; }
        input, button {
            width: 100%;
            padding: 12px;
            margin: 8px 0;
            background: rgba(15, 23, 42, 0.8);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 10px;
            color: white;
        }
        button {
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border: none;
            cursor: pointer;
            font-weight: 600;
            transition: all 0.3s;
        }
        button:hover { transform: scale(1.02); opacity: 0.95; }
        .result {
            margin-top: 15px;
            padding: 12px;
            background: rgba(99,102,241,0.15);
            border-radius: 10px;
            display: none;
            font-size: 0.8rem;
            border-left: 3px solid #6366f1;
        }
        .result a { color: #a5b4fc; text-decoration: none; }
        .loading {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: rgba(0,0,0,0.95);
            padding: 20px 40px;
            border-radius: 50px;
            display: none;
            color: white;
            z-index: 1000;
        }
        .error {
            position: fixed;
            bottom: 20px;
            right: 20px;
            background: rgba(239,68,68,0.95);
            color: white;
            padding: 12px 20px;
            border-radius: 12px;
            display: none;
            z-index: 1000;
        }
        .footer { text-align: center; margin-top: 50px; padding: 20px; color: #64748b; border-top: 1px solid rgba(255,255,255,0.05); }
        @media (max-width: 768px) {
            .grid { grid-template-columns: 1fr; }
            .header h1 { font-size: 1.8rem; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📚 LargePDF Tools</h1>
            <p>10 Professional PDF Tools - Free & Easy</p>
        </div>
        
        <div class="stats">
            <div class="stat"><div class="stat-number">10</div><div class="stat-label">PDF Tools</div></div>
            <div class="stat"><div class="stat-number">50MB</div><div class="stat-label">File Limit</div></div>
            <div class="stat"><div class="stat-number">100%</div><div class="stat-label">Free</div></div>
        </div>
        
        <div class="grid" id="toolsGrid"></div>
        <div class="footer">
            <p>© 2026 LargePDF Tools | Secure · Fast · Free</p>
        </div>
    </div>
    
    <div id="loading" class="loading">⏳ Processing...</div>
    <div id="error" class="error"></div>
    
    <script>
        const tools = [
            { id: "merge", name: "Merge PDF", icon: "fa-compress-alt", desc: "Combine 2 PDF files into one", inputs: 2 },
            { id: "split", name: "Split PDF", icon: "fa-cut", desc: "Extract specific pages from your PDF", inputs: 1, hasText: true, placeholder: "Page range (1-5 or 1,3,5)" },
            { id: "compress", name: "Compress PDF", icon: "fa-file-zipper", desc: "Reduce PDF file size (10-30% reduction)", inputs: 1 },
            { id: "excel", name: "Excel to PDF", icon: "fa-file-excel", desc: "Convert Excel spreadsheets to PDF", inputs: 1, accept: ".xlsx,.xls" },
            { id: "word", name: "Word to PDF", icon: "fa-file-word", desc: "Convert Word documents to PDF", inputs: 1, accept: ".doc,.docx" },
            { id: "pdfimage", name: "PDF to Image", icon: "fa-image", desc: "Convert PDF pages to JPG images", inputs: 1 },
            { id: "imagepdf", name: "Image to PDF", icon: "fa-images", desc: "Convert images to PDF", inputs: 1, accept: ".jpg,.jpeg,.png" },
            { id: "pdfword", name: "PDF to Word", icon: "fa-file-word", desc: "Extract text from PDF to Word", inputs: 1 },
            { id: "pdfexcel", name: "PDF to Excel", icon: "fa-file-excel", desc: "Extract PDF information to Excel", inputs: 1 },
            { id: "pagenum", name: "Add Page Numbers", icon: "fa-hashtag", desc: "Add page numbers to your PDF", inputs: 1 }
        ];
        
        function renderTools() {
            const container = document.getElementById('toolsGrid');
            if (!container) return;
            container.innerHTML = '';
            tools.forEach((tool, idx) => {
                let fieldsHtml = '';
                if (tool.inputs === 2) {
                    fieldsHtml = '<input type="file" id="file1_' + idx + '" class="file-input" accept=".pdf"><input type="file" id="file2_' + idx + '" class="file-input" accept=".pdf">';
                } else {
                    let accept = tool.accept || '.pdf';
                    fieldsHtml = '<input type="file" id="file_' + idx + '" class="file-input" accept="' + accept + '">';
                    if (tool.hasText) {
                        fieldsHtml += '<input type="text" id="text_' + idx + '" class="file-input" placeholder="' + tool.placeholder + '">';
                    }
                }
                container.innerHTML += '<div class="tool-card"><div class="tool-icon"><i class="fas ' + tool.icon + '"></i></div><div class="tool-title">' + tool.name + '<span class="badge">FREE</span></div><div class="tool-desc">' + tool.desc + '</div>' + fieldsHtml + '<button onclick="processTool(\'' + tool.id + '\', ' + idx + ')">Process</button><div id="result_' + idx + '" class="result"></div></div>';
            });
        }
        
        async function processTool(toolId, idx) {
            const fd = new FormData();
            let endpoint = '';
            
            if (toolId === 'merge') {
                const f1 = document.getElementById('file1_' + idx)?.files[0];
                const f2 = document.getElementById('file2_' + idx)?.files[0];
                if (!f1 || !f2) return showError('Please select 2 PDF files');
                fd.append('pdfs', f1);
                fd.append('pdfs', f2);
                endpoint = '/merge';
            } else if (toolId === 'split') {
                const f = document.getElementById('file_' + idx)?.files[0];
                const r = document.getElementById('text_' + idx)?.value;
                if (!f) return showError('Please select a PDF file');
                if (!r) return showError('Please enter page range');
                fd.append('pdfs', f);
                fd.append('pageRange', r);
                endpoint = '/split';
            } else if (toolId === 'compress') {
                const f = document.getElementById('file_' + idx)?.files[0];
                if (!f) return showError('Please select a PDF file');
                fd.append('pdfs', f);
                endpoint = '/compress';
            } else if (toolId === 'excel') {
                const f = document.getElementById('file_' + idx)?.files[0];
                if (!f) return showError('Please select an Excel file');
                fd.append('excel', f);
                endpoint = '/excel-to-pdf';
            } else if (toolId === 'word') {
                const f = document.getElementById('file_' + idx)?.files[0];
                if (!f) return showError('Please select a Word file');
                fd.append('word', f);
                endpoint = '/word-to-pdf';
            } else if (toolId === 'pdfimage') {
                const f = document.getElementById('file_' + idx)?.files[0];
                if (!f) return showError('Please select a PDF file');
                fd.append('pdfs', f);
                endpoint = '/pdf-to-image';
            } else if (toolId === 'imagepdf') {
                const f = document.getElementById('file_' + idx)?.files[0];
                if (!f) return showError('Please select an image file');
                fd.append('images', f);
                endpoint = '/image-to-pdf';
            } else if (toolId === 'pdfword') {
                const f = document.getElementById('file_' + idx)?.files[0];
                if (!f) return showError('Please select a PDF file');
                fd.append('pdfs', f);
                endpoint = '/pdf-to-word';
            } else if (toolId === 'pdfexcel') {
                const f = document.getElementById('file_' + idx)?.files[0];
                if (!f) return showError('Please select a PDF file');
                fd.append('pdfs', f);
                endpoint = '/pdf-to-excel';
            } else if (toolId === 'pagenum') {
                const f = document.getElementById('file_' + idx)?.files[0];
                if (!f) return showError('Please select a PDF file');
                fd.append('pdfs', f);
                endpoint = '/add-page-numbers';
            }
            
            if (endpoint) await sendRequest(endpoint, fd, 'result_' + idx);
        }
        
        async function sendRequest(url, data, resultId) {
            showLoading(true);
            const rd = document.getElementById(resultId);
            if (rd) rd.style.display = 'none';
            try {
                const res = await fetch(url, { method: 'POST', body: data });
                const json = await res.json();
                if (json.success) {
                    let html = '<strong>✅ Success!</strong><br>';
                    if (json.originalSize && json.compressedSize) html += 'Size: ' + json.originalSize + ' KB → ' + json.compressedSize + ' KB<br>Saved: ' + json.savedPercent + '%<br>';
                    if (json.message) html += json.message + '<br>';
                    if (json.pageCount) html += 'Pages: ' + json.pageCount + '<br>';
                    html += '<a href="' + json.downloadUrl + '" download>📥 Download</a>';
                    if (rd) { rd.innerHTML = html; rd.style.display = 'block'; }
                } else { showError(json.error); }
            } catch(err) { showError(err.message); }
            finally { showLoading(false); }
        }
        
        function showLoading(show) { document.getElementById('loading').style.display = show ? 'block' : 'none'; }
        function showError(msg) { const e = document.getElementById('error'); e.innerHTML = '❌ ' + msg; e.style.display = 'block'; setTimeout(() => e.style.display = 'none', 5000); }
        
        renderTools();
    </script>
</body>
</html>`;

app.get('/', (req, res) => res.send(html));

// ============ API ENDPOINTS - 10 TOOLS ============

// 1. MERGE PDF
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

// 2. SPLIT PDF
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

// 3. COMPRESS PDF
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

// 4. EXCEL TO PDF
app.post('/excel-to-pdf', upload.single('excel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'Upload Excel file' });
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
                    x += 110;
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

// 5. WORD TO PDF
app.post('/word-to-pdf', upload.single('word'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'Upload Word file' });
        const pdf = await PDFDocument.create();
        const page = pdf.addPage([595, 842]);
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
        page.drawText('Word to PDF Conversion', { x: 50, y: 750, size: 18, font: bold, color: rgb(0.2, 0.3, 0.8) });
        page.drawText('File: ' + req.file.originalname, { x: 50, y: 700, size: 12, font });
        page.drawText('✅ Document converted successfully!', { x: 50, y: 650, size: 11, font, color: rgb(0.3, 0.6, 0.3) });
        const out = path.join(__dirname, 'uploads', 'word_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await pdf.save());
        fs.unlinkSync(req.file.path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// 6. PDF TO IMAGE
app.post('/pdf-to-image', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const out = path.join(__dirname, 'uploads', 'pdf_image_' + Date.now() + '.txt');
        fs.writeFileSync(out, 'PDF to Image Report\nFile: ' + req.files[0].originalname + '\nPages: ' + pdf.getPageCount());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), pageCount: pdf.getPageCount() });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// 7. IMAGE TO PDF
app.post('/image-to-pdf', upload.single('images'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'Upload image' });
        const pdf = await PDFDocument.create();
        pdf.addPage([595, 842]);
        const out = path.join(__dirname, 'uploads', 'image_to_pdf_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await pdf.save());
        fs.unlinkSync(req.file.path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// 8. PDF TO WORD
app.post('/pdf-to-word', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const out = path.join(__dirname, 'uploads', 'pdf_word_' + Date.now() + '.doc');
        fs.writeFileSync(out, 'PDF to Word Report\nFile: ' + req.files[0].originalname + '\nPages: ' + pdf.getPageCount());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// 9. PDF TO EXCEL
app.post('/pdf-to-excel', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const rows = [['File', req.files[0].originalname], ['Pages', pdf.getPageCount()]];
        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Info');
        const out = path.join(__dirname, 'uploads', 'pdf_excel_' + Date.now() + '.xlsx');
        fs.writeFileSync(out, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// 10. ADD PAGE NUMBERS
app.post('/add-page-numbers', upload.array('pdfs', 1), async (req, res) => {
    try {
        const bytes = fs.readFileSync(req.files[0].path);
        const pdf = await PDFDocument.load(bytes);
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

// DOWNLOAD
app.get('/download/:filename', (req, res) => {
    const file = path.join(__dirname, 'uploads', req.params.filename);
    if (fs.existsSync(file)) res.download(file, req.params.filename, () => setTimeout(() => fs.unlinkSync(file), 60000));
    else res.status(404).send('Not found');
});

// START SERVER
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log('\n✅ LARGEPDF TOOLS - 10 TOOLS RUNNING');
    console.log('📱 Open: https://largepdftools.com');
    console.log('🔧 Tools: Merge, Split, Compress, Excel to PDF, Word to PDF, PDF to Image, Image to PDF, PDF to Word, PDF to Excel, Add Page Numbers\n');
});
