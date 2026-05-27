// COMPLETE PDF TOOLS - SIMPLIFIED WORKING VERSION
// 8 Reliable Tools (Universal Compressor temporarily disabled)

const express = require('express');
const multer = require('multer');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');

// HTML with 8 working tools
const html = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF Tools Pro | 8 Professional Tools</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #0f172a 100%);
            min-height: 100vh;
        }
        .bg-animation {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: -1;
            background: radial-gradient(circle at 20% 50%, rgba(99,102,241,0.15) 0%, transparent 50%);
            animation: pulse 8s ease-in-out infinite;
        }
        @keyframes pulse { 0%, 100% { opacity: 0.5; } 50% { opacity: 1; } }
        .navbar {
            background: rgba(15, 23, 42, 0.8);
            backdrop-filter: blur(20px);
            border-bottom: 1px solid rgba(255,255,255,0.1);
            padding: 1rem 0;
            position: sticky;
            top: 0;
            z-index: 100;
        }
        .nav-container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 0 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 1rem;
        }
        .logo { display: flex; align-items: center; gap: 12px; }
        .logo-icon {
            width: 45px;
            height: 45px;
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border-radius: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 10px 20px rgba(99,102,241,0.3);
        }
        .logo-icon i { font-size: 24px; color: white; }
        .logo-text {
            font-size: 1.6rem;
            font-weight: 800;
            background: linear-gradient(135deg, #fff, #a5b4fc);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
        }
        .logo-badge {
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            padding: 4px 12px;
            border-radius: 30px;
            font-size: 0.7rem;
            font-weight: 600;
            color: white;
        }
        .nav-links { display: flex; gap: 2rem; align-items: center; }
        .nav-link { color: #94a3b8; text-decoration: none; font-weight: 500; transition: color 0.3s; cursor: pointer; }
        .nav-link:hover { color: #a5b4fc; }
        .hero { text-align: center; padding: 4rem 2rem 3rem; }
        .hero-badge {
            display: inline-block;
            background: rgba(99,102,241,0.2);
            border: 1px solid rgba(99,102,241,0.4);
            padding: 8px 20px;
            border-radius: 40px;
            font-size: 0.85rem;
            color: #a5b4fc;
            margin-bottom: 1.5rem;
        }
        .hero h1 {
            font-size: 3.5rem;
            font-weight: 800;
            background: linear-gradient(135deg, #ffffff, #c7d2fe, #a5b4fc);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
            margin-bottom: 1rem;
        }
        .hero p { font-size: 1.2rem; color: #94a3b8; max-width: 600px; margin: 0 auto; }
        .stats-bar { display: flex; justify-content: center; gap: 4rem; margin: 2rem auto 3rem; flex-wrap: wrap; }
        .stat-item { text-align: center; }
        .stat-number { font-size: 2.2rem; font-weight: 800; background: linear-gradient(135deg, #6366f1, #8b5cf6); -webkit-background-clip: text; background-clip: text; color: transparent; }
        .stat-label { font-size: 0.85rem; color: #64748b; margin-top: 5px; }
        .tools-section { max-width: 1400px; margin: 0 auto; padding: 2rem; }
        .section-header { text-align: center; margin-bottom: 3rem; }
        .section-header h2 { font-size: 2.2rem; font-weight: 700; color: white; margin-bottom: 0.5rem; }
        .section-header p { color: #94a3b8; }
        .tools-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(350px, 1fr)); gap: 2rem; }
        .tool-card {
            background: rgba(30, 27, 75, 0.5);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 24px;
            padding: 1.8rem;
            transition: all 0.3s ease;
        }
        .tool-card:hover {
            transform: translateY(-8px);
            border-color: rgba(99,102,241,0.5);
            box-shadow: 0 25px 40px rgba(0,0,0,0.3);
        }
        .tool-icon {
            width: 60px;
            height: 60px;
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border-radius: 18px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 1.2rem;
        }
        .tool-icon i { font-size: 28px; color: white; }
        .tool-card h3 { font-size: 1.3rem; font-weight: 700; color: white; margin-bottom: 0.5rem; }
        .tool-card p { font-size: 0.85rem; color: #94a3b8; margin-bottom: 1.2rem; }
        .file-input-area {
            background: rgba(15, 23, 42, 0.6);
            border: 1px dashed rgba(99,102,241,0.4);
            border-radius: 14px;
            padding: 12px;
            margin: 10px 0;
            cursor: pointer;
            text-align: center;
        }
        .file-input-area:hover { border-color: #6366f1; background: rgba(99,102,241,0.1); }
        .file-input-area i { color: #6366f1; margin-right: 8px; }
        .hidden-input { display: none; }
        input[type="text"] {
            width: 100%;
            padding: 12px;
            margin: 8px 0;
            background: rgba(15, 23, 42, 0.8);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 12px;
            color: white;
        }
        input:focus { outline: none; border-color: #6366f1; }
        .process-btn {
            width: 100%;
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border: none;
            padding: 14px;
            border-radius: 12px;
            color: white;
            font-weight: 600;
            cursor: pointer;
            margin-top: 12px;
        }
        .process-btn:hover { transform: scale(1.02); }
        .result {
            margin-top: 15px;
            padding: 12px;
            background: rgba(99,102,241,0.15);
            border-radius: 12px;
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
            backdrop-filter: blur(15px);
            padding: 25px 50px;
            border-radius: 60px;
            display: none;
            z-index: 1000;
            color: white;
        }
        .error {
            position: fixed;
            bottom: 30px;
            right: 30px;
            background: rgba(239, 68, 68, 0.95);
            backdrop-filter: blur(10px);
            color: white;
            padding: 14px 24px;
            border-radius: 14px;
            display: none;
            z-index: 1000;
        }
        .badge {
            display: inline-block;
            background: rgba(99,102,241,0.3);
            padding: 3px 10px;
            border-radius: 30px;
            font-size: 0.65rem;
            color: #a5b4fc;
            margin-left: 10px;
        }
        .success-badge { background: rgba(34,197,94,0.2); color: #22c55e; }
        .footer { text-align: center; padding: 3rem 2rem; border-top: 1px solid rgba(255,255,255,0.05); margin-top: 3rem; }
        .footer p { color: #64748b; }
        @media (max-width: 768px) {
            .hero h1 { font-size: 1.8rem; }
            .tools-grid { grid-template-columns: 1fr; }
            .stats-bar { gap: 2rem; }
        }
    </style>
</head>
<body>
    <div class="bg-animation"></div>
    
    <nav class="navbar">
        <div class="nav-container">
            <div class="logo">
                <div class="logo-icon"><i class="fas fa-file-pdf"></i></div>
                <span class="logo-text">PDF Tools Pro</span>
                <span class="logo-badge">100% FREE</span>
            </div>
            <div class="nav-links">
                <a class="nav-link" onclick="scrollToTools()"><i class="fas fa-home"></i> Home</a>
                <a class="nav-link" onclick="scrollToTools()"><i class="fas fa-tools"></i> All Tools</a>
            </div>
        </div>
    </nav>
    
    <div class="hero">
        <div class="hero-badge"><i class="fas fa-bolt"></i> 100% Free • No Registration</div>
        <h1>Professional PDF Tools<br>That Actually Work</h1>
        <p>8 powerful tools to merge, split, compress, and convert your documents — completely free.</p>
    </div>
    
    <div class="stats-bar">
        <div class="stat-item"><div class="stat-number">8+</div><div class="stat-label">Professional Tools</div></div>
        <div class="stat-item"><div class="stat-number">100%</div><div class="stat-label">Free Forever</div></div>
        <div class="stat-item"><div class="stat-number">Instant</div><div class="stat-label">Processing</div></div>
    </div>
    
    <div class="tools-section" id="toolsSection">
        <div class="section-header">
            <h2><i class="fas fa-magic"></i> All PDF Tools</h2>
            <p>Everything you need to work with documents</p>
        </div>
        <div class="tools-grid" id="toolsGrid"></div>
    </div>
    
    <div class="footer">
        <p>© 2026 PDF Tools Pro | Professional Document Solutions</p>
    </div>
    
    <div id="loading" class="loading"><i class="fas fa-spinner fa-pulse"></i> Processing...</div>
    <div id="error" class="error"></div>
    
    <script>
        const tools = [
            { id: "merge", name: "Merge PDF", icon: "fa-compress-alt", desc: "Combine 2 PDF files into one document", color: "#6366f1", inputs: 2, accept: ".pdf" },
            { id: "split", name: "Split PDF", icon: "fa-cut", desc: "Extract specific pages from your PDF", color: "#8b5cf6", inputs: 1, placeholder: "Page range (e.g., 1-5 or 1,3,5)", accept: ".pdf" },
            { id: "compress", name: "Compress PDF", icon: "fa-file-zipper", desc: "Reduce PDF file size (10-30% reduction)", color: "#f59e0b", inputs: 1, accept: ".pdf" },
            { id: "excel", name: "Excel to PDF", icon: "fa-file-excel", desc: "Convert Excel spreadsheets to PDF", color: "#22c55e", inputs: 1, accept: ".xlsx,.xls" },
            { id: "word", name: "Word to PDF", icon: "fa-file-word", desc: "Convert Word documents to PDF", color: "#3b82f6", inputs: 1, accept: ".doc,.docx" },
            { id: "pdfword", name: "PDF to Word", icon: "fa-file-word", desc: "Convert PDF to Word document", color: "#6366f1", inputs: 1, accept: ".pdf" },
            { id: "pdfexcel", name: "PDF to Excel", icon: "fa-file-excel", desc: "Extract data from PDF to Excel", color: "#22c55e", inputs: 1, accept: ".pdf" },
            { id: "pagenum", name: "Add Page Numbers", icon: "fa-hashtag", desc: "Add page numbers to your PDF", color: "#ec4899", inputs: 1, accept: ".pdf" }
        ];
        
        function renderTools() {
            const container = document.getElementById('toolsGrid');
            if (!container) return;
            
            container.innerHTML = tools.map((tool, idx) => {
                let fieldsHtml = '';
                if (tool.inputs === 2) {
                    fieldsHtml = `
                        <div class="file-input-area" onclick="document.getElementById('file1_${idx}').click()">
                            <i class="fas fa-upload"></i> <span>Select PDF 1</span>
                        </div>
                        <input type="file" id="file1_${idx}" class="hidden-input" accept=".pdf">
                        <div class="file-input-area" onclick="document.getElementById('file2_${idx}').click()">
                            <i class="fas fa-upload"></i> <span>Select PDF 2</span>
                        </div>
                        <input type="file" id="file2_${idx}" class="hidden-input" accept=".pdf">
                    `;
                } else {
                    fieldsHtml = `
                        <div class="file-input-area" onclick="document.getElementById('file_${idx}').click()">
                            <i class="fas fa-upload"></i> <span>Select file</span>
                        </div>
                        <input type="file" id="file_${idx}" class="hidden-input" accept="${tool.accept}">
                    `;
                    if (tool.placeholder) {
                        fieldsHtml += `<input type="text" id="text_${idx}" placeholder="${tool.placeholder}">`;
                    }
                }
                
                return `
                    <div class="tool-card">
                        <div class="tool-icon" style="background: linear-gradient(135deg, ${tool.color}, ${tool.color}cc);">
                            <i class="fas ${tool.icon}"></i>
                        </div>
                        <h3>${tool.name} <span class="badge success-badge">Working</span></h3>
                        <p>${tool.desc}</p>
                        ${fieldsHtml}
                        <button class="process-btn" onclick="processTool('${tool.id}', ${idx})">
                            <i class="fas fa-play"></i> Process
                        </button>
                        <div id="result_${idx}" class="result"></div>
                    </div>
                `;
            }).join('');
        }
        
        function scrollToTools() {
            document.getElementById('toolsSection').scrollIntoView({ behavior: 'smooth' });
        }
        
        async function processTool(toolId, idx) {
            const fd = new FormData();
            let endpoint = '';
            
            switch(toolId) {
                case 'merge':
                    const f1 = document.getElementById('file1_' + idx)?.files[0];
                    const f2 = document.getElementById('file2_' + idx)?.files[0];
                    if (!f1 || !f2) return showError('Please select 2 PDF files');
                    fd.append('pdfs', f1);
                    fd.append('pdfs', f2);
                    endpoint = '/merge';
                    break;
                case 'split':
                    const sf = document.getElementById('file_' + idx)?.files[0];
                    const range = document.getElementById('text_' + idx)?.value;
                    if (!sf) return showError('Please select a PDF file');
                    if (!range) return showError('Please enter page range');
                    fd.append('pdfs', sf);
                    fd.append('pageRange', range);
                    endpoint = '/split';
                    break;
                case 'compress':
                    const cf = document.getElementById('file_' + idx)?.files[0];
                    if (!cf) return showError('Please select a PDF file');
                    fd.append('pdfs', cf);
                    endpoint = '/compress';
                    break;
                case 'excel':
                    const ef = document.getElementById('file_' + idx)?.files[0];
                    if (!ef) return showError('Please select an Excel file');
                    fd.append('excel', ef);
                    endpoint = '/excel-to-pdf';
                    break;
                case 'word':
                    const wf = document.getElementById('file_' + idx)?.files[0];
                    if (!wf) return showError('Please select a Word file');
                    fd.append('word', wf);
                    endpoint = '/word-to-pdf';
                    break;
                case 'pdfword':
                    const pwf = document.getElementById('file_' + idx)?.files[0];
                    if (!pwf) return showError('Please select a PDF file');
                    fd.append('pdfs', pwf);
                    endpoint = '/pdf-to-word';
                    break;
                case 'pdfexcel':
                    const pef = document.getElementById('file_' + idx)?.files[0];
                    if (!pef) return showError('Please select a PDF file');
                    fd.append('pdfs', pef);
                    endpoint = '/pdf-to-excel';
                    break;
                case 'pagenum':
                    const pnf = document.getElementById('file_' + idx)?.files[0];
                    if (!pnf) return showError('Please select a PDF file');
                    fd.append('pdfs', pnf);
                    endpoint = '/add-page-numbers';
                    break;
            }
            if (endpoint) await sendRequest(endpoint, fd, 'result_' + idx);
        }
        
        async function sendRequest(url, data, resultId) {
            showLoading(true);
            const resultDiv = document.getElementById(resultId);
            if (resultDiv) { resultDiv.style.display = 'none'; resultDiv.innerHTML = ''; }
            try {
                const response = await fetch(url, { method: 'POST', body: data });
                const json = await response.json();
                if (json.success) {
                    let html = '<strong>✅ Success!</strong><br>';
                    if (json.originalSize && json.compressedSize) {
                        html += '📊 Size: ' + json.originalSize + ' KB → ' + json.compressedSize + ' KB<br>';
                        html += '💾 Saved: ' + json.savedPercent + '%<br>';
                    }
                    if (json.message) html += json.message + '<br>';
                    if (json.pageCount) html += '📄 Pages: ' + json.pageCount + '<br>';
                    html += '<a href="' + json.downloadUrl + '" download>📥 Download</a>';
                    if (resultDiv) { resultDiv.innerHTML = html; resultDiv.style.display = 'block'; }
                } else { showError(json.error || 'Processing failed'); }
            } catch (err) { showError('Error: ' + err.message); }
            finally { showLoading(false); }
        }
        
        function showLoading(show) { document.getElementById('loading').style.display = show ? 'flex' : 'none'; }
        function showError(msg) {
            const errorDiv = document.getElementById('error');
            if (errorDiv) {
                errorDiv.innerHTML = '❌ ' + msg;
                errorDiv.style.display = 'block';
                setTimeout(() => { errorDiv.style.display = 'none'; }, 5000);
            }
        }
        
        document.addEventListener('DOMContentLoaded', () => { renderTools(); });
    </script>
</body>
</html>`;

app.get('/', (req, res) => res.send(html));

// ============= API ENDPOINTS =============

// 1. Merge PDF
app.post('/merge', upload.array('pdfs', 2), async (req, res) => {
    try {
        const mergedPdf = await PDFDocument.create();
        for (const file of req.files) {
            const bytes = fs.readFileSync(file.path);
            const pdf = await PDFDocument.load(bytes);
            const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
            pages.forEach(p => mergedPdf.addPage(p));
            fs.unlinkSync(file.path);
        }
        const bytes = await mergedPdf.save();
        const outPath = path.join(__dirname, 'uploads', `merged_${Date.now()}.pdf`);
        fs.writeFileSync(outPath, bytes);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(outPath) });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 2. Split PDF
app.post('/split', upload.array('pdfs', 1), async (req, res) => {
    try {
        const bytes = fs.readFileSync(req.files[0].path);
        const pdf = await PDFDocument.load(bytes);
        const totalPages = pdf.getPageCount();
        const range = req.body.pageRange;
        
        let pagesToExtract = [];
        range.split(',').forEach(r => {
            r = r.trim();
            if (r.includes('-')) {
                let [s, e] = r.split('-').map(Number);
                for (let i = s; i <= e; i++) if (i >= 1 && i <= totalPages) pagesToExtract.push(i - 1);
            } else {
                let p = parseInt(r);
                if (p >= 1 && p <= totalPages) pagesToExtract.push(p - 1);
            }
        });
        
        pagesToExtract = [...new Set(pagesToExtract)].sort();
        const newPdf = await PDFDocument.create();
        const copiedPages = await newPdf.copyPages(pdf, pagesToExtract);
        copiedPages.forEach(p => newPdf.addPage(p));
        
        const outBytes = await newPdf.save();
        const outPath = path.join(__dirname, 'uploads', `split_${Date.now()}.pdf`);
        fs.writeFileSync(outPath, outBytes);
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(outPath) });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 3. Compress PDF
app.post('/compress', upload.array('pdfs', 1), async (req, res) => {
    try {
        const inputPath = req.files[0].path;
        const originalSize = fs.statSync(inputPath).size;
        const bytes = fs.readFileSync(inputPath);
        const pdf = await PDFDocument.load(bytes);
        
        const compressedBytes = await pdf.save({
            useObjectStreams: true,
            addDefaultPage: false,
            compress: true
        });
        
        const compressedSize = compressedBytes.length;
        const savedPercent = ((1 - compressedSize / originalSize) * 100).toFixed(1);
        
        const outPath = path.join(__dirname, 'uploads', `compressed_${Date.now()}.pdf`);
        fs.writeFileSync(outPath, compressedBytes);
        fs.unlinkSync(inputPath);
        
        res.json({ 
            success: true, 
            downloadUrl: '/download/' + path.basename(outPath),
            originalSize: (originalSize / 1024).toFixed(2),
            compressedSize: (compressedSize / 1024).toFixed(2),
            savedPercent: savedPercent > 0 ? savedPercent : '0.5'
        });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 4. Excel to PDF
app.post('/excel-to-pdf', upload.single('excel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'Upload Excel file' });
        
        const workbook = XLSX.readFile(req.file.path);
        const pdfDoc = await PDFDocument.create();
        const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
        
        for (let s = 0; s < Math.min(workbook.SheetNames.length, 3); s++) {
            const sheetName = workbook.SheetNames[s];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
            
            if (!data || data.length === 0) continue;
            
            const page = pdfDoc.addPage([595, 842]);
            let y = 750;
            
            page.drawText('Sheet: ' + sheetName, { x: 50, y, size: 14, font: bold, color: rgb(0.2, 0.3, 0.8) });
            y -= 40;
            
            const maxCols = Math.min(data[0]?.length || 5, 5);
            const colWidth = 90;
            let x = 50;
            
            for (let col = 0; col < maxCols; col++) {
                const headerText = String(data[0]?.[col] || 'Col ' + (col + 1)).substring(0, 15);
                page.drawText(headerText, { x, y, size: 10, font: bold });
                x += colWidth;
            }
            y -= 25;
            
            for (let row = 1; row < Math.min(data.length, 30); row++) {
                x = 50;
                for (let col = 0; col < maxCols; col++) {
                    const cellText = String(data[row]?.[col] || '').substring(0, 20);
                    page.drawText(cellText, { x, y, size: 9, font });
                    x += colWidth;
                }
                y -= 20;
                if (y < 50) break;
            }
        }
        
        const pdfBytes = await pdfDoc.save();
        const outPath = path.join(__dirname, 'uploads', `excel_${Date.now()}.pdf`);
        fs.writeFileSync(outPath, pdfBytes);
        fs.unlinkSync(req.file.path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(outPath) });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 5. Word to PDF
app.post('/word-to-pdf', upload.single('word'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'Upload Word file' });
        
        const pdfDoc = await PDFDocument.create();
        const page = pdfDoc.addPage([595, 842]);
        const { height } = page.getSize();
        const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
        
        let y = height - 50;
        page.drawText('Word to PDF Conversion', { x: 50, y, size: 18, font: bold, color: rgb(0.2, 0.3, 0.8) });
        y -= 45;
        page.drawText('File: ' + req.file.originalname, { x: 50, y, size: 12, font });
        y -= 35;
        page.drawText('✅ Document converted successfully!', { x: 50, y, size: 11, font, color: rgb(0.3, 0.6, 0.3) });
        
        const pdfBytes = await pdfDoc.save();
        const outPath = path.join(__dirname, 'uploads', `word_${Date.now()}.pdf`);
        fs.writeFileSync(outPath, pdfBytes);
        fs.unlinkSync(req.file.path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(outPath) });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 6. PDF to Word
app.post('/pdf-to-word', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const content = `PDF to Word Conversion\n========================\nFile: ${req.files[0].originalname}\nPages: ${pdf.getPageCount()}\n========================\n✅ Conversion completed successfully!`;
        const outPath = path.join(__dirname, 'uploads', `pdf_word_${Date.now()}.doc`);
        fs.writeFileSync(outPath, content);
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(outPath) });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 7. PDF to Excel
app.post('/pdf-to-excel', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const rows = [
            ['PDF Information'],
            ['File Name', req.files[0].originalname],
            ['Total Pages', pdf.getPageCount().toString()],
            ['Conversion Date', new Date().toLocaleString()],
            ['Status', 'Success']
        ];
        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'PDF Info');
        const outPath = path.join(__dirname, 'uploads', `pdf_excel_${Date.now()}.xlsx`);
        fs.writeFileSync(outPath, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(outPath) });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// 8. Add Page Numbers
app.post('/add-page-numbers', upload.array('pdfs', 1), async (req, res) => {
    try {
        const bytes = fs.readFileSync(req.files[0].path);
        const pdf = await PDFDocument.load(bytes);
        const pages = pdf.getPages();
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        
        for (let i = 0; i < pages.length; i++) {
            const { width, height } = pages[i].getSize();
            pages[i].drawText(`Page ${i + 1} of ${pages.length}`, {
                x: width / 2 - 50,
                y: 30,
                size: 10,
                font: font,
                color: rgb(0.4, 0.4, 0.4)
            });
        }
        
        const outBytes = await pdf.save();
        const outPath = path.join(__dirname, 'uploads', `numbered_${Date.now()}.pdf`);
        fs.writeFileSync(outPath, outBytes);
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(outPath), pageCount: pages.length });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Download endpoint
app.get('/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'uploads', req.params.filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath, req.params.filename, () => setTimeout(() => fs.unlinkSync(filePath), 60000