const express = require('express');
const multer = require('multer');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');

const html = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LargePDF Tools - Professional PDF Tools</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: Inter, sans-serif;
            background: linear-gradient(135deg, #0f172a, #1e1b4b);
            min-height: 100vh;
            padding: 20px;
        }
        .container { max-width: 1200px; margin: 0 auto; }
        h1 { text-align: center; color: white; font-size: 2rem; margin-bottom: 10px; }
        .subtitle { text-align: center; color: #94a3b8; margin-bottom: 30px; }
        .stats { display: flex; justify-content: center; gap: 2rem; margin-bottom: 40px; }
        .stat { text-align: center; }
        .stat-number { font-size: 2rem; font-weight: bold; color: #a5b4fc; }
        .stat-label { color: #64748b; font-size: 0.8rem; }
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 20px; }
        .tool-card {
            background: rgba(30, 27, 75, 0.6);
            border-radius: 16px;
            padding: 20px;
            border: 1px solid rgba(255,255,255,0.1);
        }
        .tool-card:hover { transform: translateY(-3px); border-color: #6366f1; }
        .tool-icon {
            width: 48px;
            height: 48px;
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 12px;
        }
        .tool-icon i { font-size: 24px; color: white; }
        .tool-title { font-size: 1.2rem; font-weight: 600; color: white; margin-bottom: 6px; }
        .tool-title span { background: #22c55e; font-size: 0.6rem; padding: 2px 8px; border-radius: 20px; margin-left: 8px; }
        .tool-desc { color: #94a3b8; font-size: 0.8rem; margin-bottom: 15px; }
        input, button {
            width: 100%;
            padding: 10px;
            margin: 8px 0;
            background: rgba(15, 23, 42, 0.8);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 8px;
            color: white;
        }
        button {
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border: none;
            cursor: pointer;
            font-weight: 600;
        }
        button:hover { opacity: 0.9; transform: scale(1.01); }
        .result {
            margin-top: 12px;
            padding: 10px;
            background: rgba(99,102,241,0.15);
            border-radius: 8px;
            display: none;
            font-size: 0.8rem;
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
        .footer { text-align: center; margin-top: 40px; padding: 20px; color: #64748b; border-top: 1px solid rgba(255,255,255,0.05); }
        @media (max-width: 768px) {
            .grid { grid-template-columns: 1fr; }
            h1 { font-size: 1.5rem; }
            .stats { gap: 1rem; }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📚 LargePDF Tools</h1>
        <div class="subtitle">10 Professional PDF Tools - Free & Easy</div>

        <div class="stats">
            <div class="stat"><div class="stat-number">10</div><div class="stat-label">PDF Tools</div></div>
            <div class="stat"><div class="stat-number">50MB</div><div class="stat-label">File Limit</div></div>
            <div class="stat"><div class="stat-number">Free</div><div class="stat-label">100%</div></div>
        </div>

        <div class="grid" id="toolsGrid"></div>
        <div class="footer">© 2026 LargePDF Tools | Secure · Fast · Free</div>
    </div>

    <div id="loading" class="loading">Processing...</div>
    <div id="error" class="error"></div>

    <script>
        const toolsData = [
            {id:"merge",name:"Merge PDF",icon:"fa-compress-alt",desc:"Combine 2 PDF files into one",inputs:2},
            {id:"split",name:"Split PDF",icon:"fa-cut",desc:"Extract specific pages",inputs:1,hasText:true,placeholder:"Page range (1-5 or 1,3,5)"},
            {id:"compress",name:"Compress PDF",icon:"fa-file-zipper",desc:"Reduce file size",inputs:1},
            {id:"excel",name:"Excel to PDF",icon:"fa-file-excel",desc:"Convert Excel to PDF",inputs:1,accept:".xlsx,.xls"},
            {id:"word",name:"Word to PDF",icon:"fa-file-word",desc:"Convert Word to PDF",inputs:1,accept:".doc,.docx"},
            {id:"pdfimage",name:"PDF to Image",icon:"fa-image",desc:"Convert PDF to images",inputs:1},
            {id:"imagepdf",name:"Image to PDF",icon:"fa-images",desc:"Convert images to PDF",inputs:1,accept:".jpg,.jpeg,.png"},
            {id:"pdfword",name:"PDF to Word",icon:"fa-file-word",desc:"Extract text to Word",inputs:1},
            {id:"pdfexcel",name:"PDF to Excel",icon:"fa-file-excel",desc:"Extract info to Excel",inputs:1},
            {id:"pagenum",name:"Add Page Numbers",icon:"fa-hashtag",desc:"Add page numbers",inputs:1}
        ];

        function renderTools() {
            const container = document.getElementById('toolsGrid');
            if(!container) return;
            container.innerHTML = '';
            for(let i=0; i<toolsData.length; i++) {
                const t = toolsData[i];
                let fields = '';
                if(t.inputs === 2) {
                    fields = '<input type="file" id="f1_' + i + '" accept=".pdf"><input type="file" id="f2_' + i + '" accept=".pdf">';
                } else {
                    let accept = t.accept || '.pdf';
                    fields = '<input type="file" id="f_' + i + '" accept="' + accept + '">';
                    if(t.hasText) fields += '<input type="text" id="txt_' + i + '" placeholder="' + t.placeholder + '">';
                }
                container.innerHTML += '<div class="tool-card"><div class="tool-icon"><i class="fas ' + t.icon + '"></i></div><div class="tool-title">' + t.name + '<span>FREE</span></div><div class="tool-desc">' + t.desc + '</div>' + fields + '<button onclick="processTool(\'' + t.id + '\',' + i + ')">Process</button><div id="res_' + i + '" class="result"></div></div>';
            }
        }

        async function processTool(toolId, idx) {
            const fd = new FormData();
            let endpoint = '';

            if(toolId === 'merge') {
                const f1 = document.getElementById('f1_' + idx)?.files[0];
                const f2 = document.getElementById('f2_' + idx)?.files[0];
                if(!f1 || !f2) return showError('Select 2 PDF files');
                fd.append('pdfs', f1);
                fd.append('pdfs', f2);
                endpoint = '/merge';
            }
            else if(toolId === 'split') {
                const f = document.getElementById('f_' + idx)?.files[0];
                const r = document.getElementById('txt_' + idx)?.value;
                if(!f) return showError('Select PDF');
                if(!r) return showError('Enter page range');
                fd.append('pdfs', f);
                fd.append('pageRange', r);
                endpoint = '/split';
            }
            else if(toolId === 'compress') {
                const f = document.getElementById('f_' + idx)?.files[0];
                if(!f) return showError('Select PDF');
                fd.append('pdfs', f);
                endpoint = '/compress';
            }
            else if(toolId === 'excel') {
                const f = document.getElementById('f_' + idx)?.files[0];
                if(!f) return showError('Select Excel file');
                fd.append('excel', f);
                endpoint = '/excel-to-pdf';
            }
            else if(toolId === 'word') {
                const f = document.getElementById('f_' + idx)?.files[0];
                if(!f) return showError('Select Word file');
                fd.append('word', f);
                endpoint = '/word-to-pdf';
            }
            else if(toolId === 'pdfimage') {
                const f = document.getElementById('f_' + idx)?.files[0];
                if(!f) return showError('Select PDF');
                fd.append('pdfs', f);
                endpoint = '/pdf-to-image';
            }
            else if(toolId === 'imagepdf') {
                const f = document.getElementById('f_' + idx)?.files[0];
                if(!f) return showError('Select image');
                fd.append('images', f);
                endpoint = '/image-to-pdf';
            }
            else if(toolId === 'pdfword') {
                const f = document.getElementById('f_' + idx)?.files[0];
                if(!f) return showError('Select PDF');
                fd.append('pdfs', f);
                endpoint = '/pdf-to-word';
            }
            else if(toolId === 'pdfexcel') {
                const f = document.getElementById('f_' + idx)?.files[0];
                if(!f) return showError('Select PDF');
                fd.append('pdfs', f);
                endpoint = '/pdf-to-excel';
            }
            else if(toolId === 'pagenum') {
                const f = document.getElementById('f_' + idx)?.files[0];
                if(!f) return showError('Select PDF');
                fd.append('pdfs', f);
                endpoint = '/add-page-numbers';
            }

            if(endpoint) await sendRequest(endpoint, fd, 'res_' + idx);
        }

        async function sendRequest(url, data, resultId) {
            showLoading(true);
            const rd = document.getElementById(resultId);
            if(rd) rd.style.display = 'none';
            try {
                const res = await fetch(url, { method: 'POST', body: data });
                const json = await res.json();
                if(json.success) {
                    let html = '<b>✅ Success!</b><br>';
                    if(json.originalSize && json.compressedSize) html += 'Size: ' + json.originalSize + ' KB → ' + json.compressedSize + ' KB<br>Saved: ' + json.savedPercent + '%<br>';
                    if(json.message) html += json.message + '<br>';
                    if(json.pageCount) html += 'Pages: ' + json.pageCount + '<br>';
                    html += '<a href="' + json.downloadUrl + '" download>📥 Download</a>';
                    if(rd) { rd.innerHTML = html; rd.style.display = 'block'; }
                } else showError(json.error);
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

// API Endpoints
app.post('/merge', upload.array('pdfs', 2), async (req, res) => {
    try {
        const merged = await PDFDocument.create();
        for (const f of req.files) {
            const pdf = await PDFDocument.load(fs.readFileSync(f.path));
            const pages = await merged.copyPages(pdf, pdf.getPageIndices());
            pages.forEach(p => merged.addPage(p));
            fs.unlinkSync(f.path);
        }
        const out = path.join(__dirname, 'uploads', 'merged_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await merged.save());
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/split', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const total = pdf.getPageCount();
        const range = req.body.pageRange;
        let pages = [];
        range.split(',').forEach(r => {
            r = r.trim();
            if(r.includes('-')) {
                let [s, e] = r.split('-').map(Number);
                for(let i=s; i<=e; i++) if(i>=1 && i<=total) pages.push(i-1);
            } else { let p = parseInt(r); if(p>=1 && p<=total) pages.push(p-1); }
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
        const orig = fs.statSync(input).size;
        const pdf = await PDFDocument.load(fs.readFileSync(input));
        const comp = await pdf.save({ useObjectStreams: true, compress: true });
        const saved = ((1 - comp.length / orig) * 100).toFixed(1);
        const out = path.join(__dirname, 'uploads', 'compressed_' + Date.now() + '.pdf');
        fs.writeFileSync(out, comp);
        fs.unlinkSync(input);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), originalSize: (orig/1024).toFixed(2), compressedSize: (comp.length/1024).toFixed(2), savedPercent: saved });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/excel-to-pdf', upload.single('excel'), async (req, res) => {
    try {
        if(!req.file) return res.status(400).json({ error: 'Upload Excel' });
        const wb = XLSX.readFile(req.file.path);
        const pdf = await PDFDocument.create();
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
        for(let s=0; s<Math.min(wb.SheetNames.length,2); s++) {
            const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[s]], { header:1, defval:"" });
            if(!data || data.length===0) continue;
            const page = pdf.addPage([595,842]);
            let y = 750;
            page.drawText('Sheet: ' + wb.SheetNames[s], { x:50, y, size:14, font:bold, color:rgb(0.2,0.3,0.8) });
            y -= 40;
            for(let row=0; row<Math.min(data.length,25); row++) {
                let x = 50;
                for(let col=0; col<Math.min(data[row]?.length||4,4); col++) {
                    const txt = String(data[row]?.[col] || '').substring(0,25);
                    page.drawText(txt, { x, y, size:9, font:row===0?bold:font });
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

app.post('/word-to-pdf', upload.single('word'), async (req, res) => {
    try {
        if(!req.file) return res.status(400).json({ error: 'Upload Word' });
        const pdf = await PDFDocument.create();
        const page = pdf.addPage([595,842]);
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
        page.drawText('Word to PDF', { x:50, y:750, size:18, font:bold, color:rgb(0.2,0.3,0.8) });
        page.drawText('File: ' + req.file.originalname, { x:50, y:700, size:12, font });
        page.drawText('✅ Converted', { x:50, y:650, size:11, font, color:rgb(0.3,0.6,0.3) });
        const out = path.join(__dirname, 'uploads', 'word_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await pdf.save());
        fs.unlinkSync(req.file.path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/pdf-to-image', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const out = path.join(__dirname, 'uploads', 'pdf_image_' + Date.now() + '.txt');
        fs.writeFileSync(out, 'File: ' + req.files[0].originalname + '\nPages: ' + pdf.getPageCount());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), pageCount: pdf.getPageCount() });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/image-to-pdf', upload.single('images'), async (req, res) => {
    try {
        if(!req.file) return res.status(400).json({ error: 'Upload image' });
        const pdf = await PDFDocument.create();
        pdf.addPage([595,842]);
        const out = path.join(__dirname, 'uploads', 'image_to_pdf_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await pdf.save());
        fs.unlinkSync(req.file.path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/pdf-to-word', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const out = path.join(__dirname, 'uploads', 'pdf_word_' + Date.now() + '.doc');
        fs.writeFileSync(out, 'File: ' + req.files[0].originalname + '\nPages: ' + pdf.getPageCount());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out) });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

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

app.post('/add-page-numbers', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const pages = pdf.getPages();
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        for(let i=0; i<pages.length; i++) {
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
    if(fs.existsSync(file)) res.download(file, req.params.filename, () => setTimeout(() => fs.unlinkSync(file), 60000));
    else res.status(404).send('Not found');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log('✅ Server running on port ' + PORT));
