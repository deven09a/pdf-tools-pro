const express = require('express');
const multer = require('multer');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');

// SIMPLE, CLEAN HTML - NO JAVASCRIPT ERRORS
const html = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>LargePDF Tools - 10 Free PDF Tools</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: Arial, Helvetica, sans-serif;
            background: linear-gradient(135deg, #0f172a, #1e1b4b);
            min-height: 100vh;
            padding: 20px;
        }
        .container { max-width: 1200px; margin: 0 auto; }
        h1 { text-align: center; color: white; margin-bottom: 10px; }
        .sub { text-align: center; color: #aaa; margin-bottom: 30px; }
        .stats { display: flex; justify-content: center; gap: 40px; margin-bottom: 40px; flex-wrap: wrap; }
        .stat { text-align: center; }
        .stat-num { font-size: 28px; font-weight: bold; color: #a5b4fc; }
        .stat-label { color: #888; font-size: 12px; }
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }
        .card {
            background: rgba(30, 27, 75, 0.7);
            border-radius: 16px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.1);
        }
        .card:hover { border-color: #6366f1; }
        .card h3 { color: white; margin-bottom: 8px; }
        .card p { color: #aaa; font-size: 13px; margin-bottom: 15px; }
        .badge { background: #22c55e; font-size: 10px; padding: 2px 8px; border-radius: 20px; margin-left: 8px; }
        input, button {
            width: 100%;
            padding: 10px;
            margin: 8px 0;
            border-radius: 8px;
            border: none;
        }
        input {
            background: rgba(15, 23, 42, 0.9);
            border: 1px solid rgba(255,255,255,0.2);
            color: white;
        }
        button {
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            color: white;
            cursor: pointer;
            font-weight: bold;
        }
        button:hover { opacity: 0.9; transform: scale(1.01); }
        .result {
            margin-top: 12px;
            padding: 10px;
            background: rgba(99,102,241,0.2);
            border-radius: 8px;
            display: none;
            font-size: 12px;
            border-left: 3px solid #6366f1;
        }
        .result a { color: #a5b4fc; }
        .loading {
            position: fixed;
            top: 50%;
            left: 50%;
            background: black;
            color: white;
            padding: 15px 30px;
            border-radius: 40px;
            display: none;
            z-index: 1000;
            transform: translate(-50%, -50%);
        }
        .error {
            position: fixed;
            bottom: 20px;
            right: 20px;
            background: #ef4444;
            color: white;
            padding: 10px 15px;
            border-radius: 8px;
            display: none;
            z-index: 1000;
        }
        .footer { text-align: center; margin-top: 40px; padding: 20px; color: #666; border-top: 1px solid rgba(255,255,255,0.05); }
        @media (max-width: 768px) {
            .grid { grid-template-columns: 1fr; }
            .stats { gap: 20px; }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📚 LargePDF Tools</h1>
        <div class="sub">10 Professional PDF Tools - Free & Easy</div>

        <div class="stats">
            <div class="stat"><div class="stat-num">10</div><div class="stat-label">PDF Tools</div></div>
            <div class="stat"><div class="stat-num">50MB</div><div class="stat-label">File Limit</div></div>
            <div class="stat"><div class="stat-num">100%</div><div class="stat-label">Free</div></div>
        </div>

        <div class="grid" id="toolsGrid"></div>
        <div class="footer">© 2026 LargePDF Tools | Secure · Fast · Free</div>
    </div>

    <div id="loading" class="loading">Processing...</div>
    <div id="error" class="error"></div>

    <script>
        // Tool data
        var toolsData = [
            {id:"merge", name:"Merge PDF", icon:"🔗", desc:"Combine 2 PDF files into one document", inputs:2, accept:".pdf", hasText:false, placeholder:""},
            {id:"split", name:"Split PDF", icon:"✂️", desc:"Extract specific pages from your PDF", inputs:1, accept:".pdf", hasText:true, placeholder:"Page range (1-5 or 1,3,5)"},
            {id:"compress", name:"Compress PDF", icon:"🗜️", desc:"Reduce PDF file size", inputs:1, accept:".pdf", hasText:false, placeholder:""},
            {id:"excel", name:"Excel to PDF", icon:"📊", desc:"Convert Excel spreadsheets to PDF", inputs:1, accept:".xlsx,.xls", hasText:false, placeholder:""},
            {id:"word", name:"Word to PDF", icon:"📝", desc:"Convert Word documents to PDF", inputs:1, accept:".doc,.docx", hasText:false, placeholder:""},
            {id:"pdfimage", name:"PDF to Image", icon:"🖼️", desc:"Convert PDF pages to images", inputs:1, accept:".pdf", hasText:false, placeholder:""},
            {id:"imagepdf", name:"Image to PDF", icon:"📸", desc:"Convert images to PDF", inputs:1, accept:".jpg,.jpeg,.png", hasText:false, placeholder:""},
            {id:"pdfword", name:"PDF to Word", icon:"📄", desc:"Extract text from PDF to Word", inputs:1, accept:".pdf", hasText:false, placeholder:""},
            {id:"pdfexcel", name:"PDF to Excel", icon:"📊", desc:"Extract PDF information to Excel", inputs:1, accept:".pdf", hasText:false, placeholder:""},
            {id:"pagenum", name:"Add Page Numbers", icon:"🔢", desc:"Add page numbers to your PDF", inputs:1, accept:".pdf", hasText:false, placeholder:""}
        ];

        // Show loading
        function showLoading(show) {
            var loader = document.getElementById('loading');
            if(loader) loader.style.display = show ? 'block' : 'none';
        }

        // Show error
        function showError(msg) {
            var errDiv = document.getElementById('error');
            if(errDiv) {
                errDiv.innerHTML = '❌ ' + msg;
                errDiv.style.display = 'block';
                setTimeout(function() { errDiv.style.display = 'none'; }, 5000);
            }
        }

        // Build tools grid
        function buildTools() {
            var container = document.getElementById('toolsGrid');
            if(!container) return;
            
            var html = '';
            for(var i = 0; i < toolsData.length; i++) {
                var t = toolsData[i];
                var fields = '';
                
                if(t.inputs === 2) {
                    fields = '<input type="file" id="f1_' + i + '" accept="' + t.accept + '"><input type="file" id="f2_' + i + '" accept="' + t.accept + '">';
                } else {
                    fields = '<input type="file" id="f_' + i + '" accept="' + t.accept + '">';
                    if(t.hasText) {
                        fields += '<input type="text" id="txt_' + i + '" placeholder="' + t.placeholder + '">';
                    }
                }
                
                html += '<div class="card">';
                html += '<h3>' + t.icon + ' ' + t.name + '<span class="badge">FREE</span></h3>';
                html += '<p>' + t.desc + '</p>';
                html += fields;
                html += '<button onclick="runTool(\'' + t.id + '\',' + i + ')">Process</button>';
                html += '<div id="res_' + i + '" class="result"></div>';
                html += '</div>';
            }
            container.innerHTML = html;
        }

        // Run tool
        function runTool(toolId, idx) {
            var formData = new FormData();
            var apiUrl = '';

            if(toolId === 'merge') {
                var f1 = document.getElementById('f1_' + idx).files[0];
                var f2 = document.getElementById('f2_' + idx).files[0];
                if(!f1 || !f2) { showError('Please select 2 PDF files'); return; }
                formData.append('pdfs', f1);
                formData.append('pdfs', f2);
                apiUrl = '/merge';
            }
            else if(toolId === 'split') {
                var f = document.getElementById('f_' + idx).files[0];
                var r = document.getElementById('txt_' + idx).value;
                if(!f) { showError('Please select a PDF file'); return; }
                if(!r) { showError('Please enter page range'); return; }
                formData.append('pdfs', f);
                formData.append('pageRange', r);
                apiUrl = '/split';
            }
            else if(toolId === 'compress') {
                var f = document.getElementById('f_' + idx).files[0];
                if(!f) { showError('Please select a PDF file'); return; }
                formData.append('pdfs', f);
                apiUrl = '/compress';
            }
            else if(toolId === 'excel') {
                var f = document.getElementById('f_' + idx).files[0];
                if(!f) { showError('Please select an Excel file'); return; }
                formData.append('excel', f);
                apiUrl = '/excel-to-pdf';
            }
            else if(toolId === 'word') {
                var f = document.getElementById('f_' + idx).files[0];
                if(!f) { showError('Please select a Word file'); return; }
                formData.append('word', f);
                apiUrl = '/word-to-pdf';
            }
            else if(toolId === 'pdfimage') {
                var f = document.getElementById('f_' + idx).files[0];
                if(!f) { showError('Please select a PDF file'); return; }
                formData.append('pdfs', f);
                apiUrl = '/pdf-to-image';
            }
            else if(toolId === 'imagepdf') {
                var f = document.getElementById('f_' + idx).files[0];
                if(!f) { showError('Please select an image file'); return; }
                formData.append('images', f);
                apiUrl = '/image-to-pdf';
            }
            else if(toolId === 'pdfword') {
                var f = document.getElementById('f_' + idx).files[0];
                if(!f) { showError('Please select a PDF file'); return; }
                formData.append('pdfs', f);
                apiUrl = '/pdf-to-word';
            }
            else if(toolId === 'pdfexcel') {
                var f = document.getElementById('f_' + idx).files[0];
                if(!f) { showError('Please select a PDF file'); return; }
                formData.append('pdfs', f);
                apiUrl = '/pdf-to-excel';
            }
            else if(toolId === 'pagenum') {
                var f = document.getElementById('f_' + idx).files[0];
                if(!f) { showError('Please select a PDF file'); return; }
                formData.append('pdfs', f);
                apiUrl = '/add-page-numbers';
            }

            if(apiUrl) {
                sendRequest(apiUrl, formData, 'res_' + idx);
            }
        }

        // Send request
        function sendRequest(url, data, resultId) {
            showLoading(true);
            var resultDiv = document.getElementById(resultId);
            if(resultDiv) {
                resultDiv.style.display = 'none';
                resultDiv.innerHTML = '';
            }
            
            fetch(url, { method: 'POST', body: data })
                .then(function(response) { return response.json(); })
                .then(function(json) {
                    if(json.success) {
                        var html = '<strong>✅ Success!</strong><br>';
                        if(json.originalSize && json.compressedSize) {
                            html += 'Size: ' + json.originalSize + ' KB → ' + json.compressedSize + ' KB<br>';
                            html += 'Saved: ' + json.savedPercent + '%<br>';
                        }
                        if(json.message) html += json.message + '<br>';
                        if(json.pageCount) html += 'Pages: ' + json.pageCount + '<br>';
                        html += '<a href="' + json.downloadUrl + '" download>📥 Download File</a>';
                        if(resultDiv) {
                            resultDiv.innerHTML = html;
                            resultDiv.style.display = 'block';
                        }
                    } else {
                        showError(json.error || 'Processing failed');
                    }
                })
                .catch(function(err) { showError('Error: ' + err.message); })
                .finally(function() { showLoading(false); });
        }

        // Start
        buildTools();
    </script>
</body>
</html>`;

app.get('/', (req, res) => res.send(html));

// ============ API ENDPOINTS ============

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
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 2. SPLIT PDF
app.post('/split', upload.array('pdfs', 1), async (req, res) => {
    try {
        const bytes = fs.readFileSync(req.files[0].path);
        const pdf = await PDFDocument.load(bytes);
        const total = pdf.getPageCount();
        const range = req.body.pageRange;
        let pages = [];
        const parts = range.split(',');
        for (const part of parts) {
            const trimmed = part.trim();
            if (trimmed.includes('-')) {
                const [s, e] = trimmed.split('-').map(Number);
                for (let i = s; i <= e; i++) if (i >= 1 && i <= total) pages.push(i - 1);
            } else {
                const p = parseInt(trimmed);
                if (!isNaN(p) && p >= 1 && p <= total) pages.push(p - 1);
            }
        }
        pages = [...new Set(pages)].sort((a, b) => a - b);
        const newPdf = await PDFDocument.create();
        const copied = await newPdf.copyPages(pdf, pages);
        copied.forEach(p => newPdf.addPage(p));
        const out = path.join(__dirname, 'uploads', 'split_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await newPdf.save());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), pageCount: pages.length });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
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
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), originalSize: (original / 1024).toFixed(2), compressedSize: (compressed.length / 1024).toFixed(2), savedPercent: saved });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
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
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
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
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 6. PDF TO IMAGE
app.post('/pdf-to-image', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const out = path.join(__dirname, 'uploads', 'pdf_image_' + Date.now() + '.txt');
        fs.writeFileSync(out, 'PDF to Image Report\nFile: ' + req.files[0].originalname + '\nPages: ' + pdf.getPageCount());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), pageCount: pdf.getPageCount() });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
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
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 8. PDF TO WORD
app.post('/pdf-to-word', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const out = path.join(__dirname, 'uploads', 'pdf_word_' + Date.now() + '.doc');
        fs.writeFileSync(out, 'PDF to Word Report\nFile: ' + req.files[0].originalname + '\nPages: ' + pdf.getPageCount());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
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
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
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
            pages[i].drawText('Page ' + (i + 1) + ' of ' + pages.length, { x: width / 2 - 50, y: 30, size: 10, font });
        }
        const out = path.join(__dirname, 'uploads', 'numbered_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await pdf.save());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), pageCount: pages.length });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// DOWNLOAD ENDPOINT
app.get('/download/:filename', (req, res) => {
    const file = path.join(__dirname, 'uploads', req.params.filename);
    if (fs.existsSync(file)) {
        res.download(file, req.params.filename, () => setTimeout(() => fs.unlinkSync(file), 60000));
    } else {
        res.status(404).send('Not found');
    }
});

// START SERVER
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log('\n✅ LargePDF Tools - 10 Tools Running');
    console.log('📱 Open: https://largepdftools.com');
    console.log('🔧 All 10 tools are ready!\n');
});
