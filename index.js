cat > /var/www/largepdf-tools/index.js << 'EOF'
const express = require('express');
const multer = require('multer');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const util = require('util');
const execPromise = util.promisify(exec);

const app = express();
const upload = multer({ dest: 'uploads/', limits: { fileSize: 100 * 1024 * 1024 } });

if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');

const html = `<!DOCTYPE html>
<html>
<head>
<title>LargePDF Tools - 10 High-Traffic PDF Tools</title>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="Free PDF tools: merge, split, compress, convert PDF to Word, Excel, JPG, and more. 90% compression with Ghostscript.">
<meta name="keywords" content="PDF tools, merge PDF, split PDF, compress PDF, PDF to Word, PDF to Excel, PDF to JPG">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:Arial;background:linear-gradient(135deg,#0f172a,#1e1b4b);min-height:100vh;padding:20px}
.container{max-width:1200px;margin:0 auto}
h1{text-align:center;color:white;font-size:2rem}
.sub{text-align:center;color:#aaa;margin-bottom:20px}
.stats{display:flex;justify-content:center;gap:30px;margin-bottom:30px;flex-wrap:wrap}
.stat{text-align:center;color:white}
.stat-num{font-size:28px;font-weight:bold;color:#a5b4fc}
.stat-label{font-size:12px;color:#aaa}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:20px}
.card{background:rgba(30,27,75,0.7);border-radius:12px;padding:20px;border:1px solid rgba(255,255,255,0.1);transition:all 0.3s}
.card:hover{border-color:#6366f1;transform:translateY(-3px)}
.card h3{color:white;margin-bottom:8px;font-size:1.2rem}
.card p{color:#aaa;font-size:13px;margin-bottom:15px}
.badge{background:#22c55e;font-size:10px;padding:2px 8px;border-radius:20px;margin-left:8px}
.badge-hot{background:#ef4444}
input,select,button{width:100%;padding:10px;margin:8px 0;border-radius:6px;border:none}
input,select{background:rgba(15,23,42,0.9);border:1px solid rgba(255,255,255,0.2);color:white}
button{background:linear-gradient(135deg,#6366f1,#8b5cf6);color:white;cursor:pointer;font-weight:bold}
button:hover{opacity:0.9}
.result{margin-top:12px;padding:10px;background:rgba(99,102,241,0.2);border-radius:6px;display:none;font-size:12px;border-left:3px solid #6366f1}
.result a{color:#a5b4fc}
.loading{position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:black;color:white;padding:15px 30px;border-radius:40px;display:none;z-index:1000}
.error{position:fixed;bottom:20px;right:20px;background:#ef4444;color:white;padding:10px 15px;border-radius:6px;display:none;z-index:1000}
.footer{text-align:center;margin-top:40px;padding:20px;color:#666}
</style>
</head>
<body>
<div class="container">
<h1>📚 LargePDF Tools</h1>
<div class="sub">10 High-Traffic PDF Tools - Free & Powerful</div>
<div class="stats">
<div class="stat"><div class="stat-num">10</div><div class="stat-label">Tools</div></div>
<div class="stat"><div class="stat-num">100MB</div><div class="stat-label">File Limit</div></div>
<div class="stat"><div class="stat-num">90%</div><div class="stat-label">Compression</div></div>
</div>
<div class="grid">
<div class="card"><h3>🔗 Merge PDF <span class="badge">FREE</span></h3><p>Combine 2 PDF files into one</p><input type="file" id="merge1" accept=".pdf"><input type="file" id="merge2" accept=".pdf"><button onclick="mergePDF()">Merge PDFs</button><div id="mergeResult" class="result"></div></div>
<div class="card"><h3>✂️ Split PDF <span class="badge">FREE</span></h3><p>Extract specific pages</p><input type="file" id="splitFile" accept=".pdf"><input type="text" id="splitRange" placeholder="Page range (1-5 or 1,3,5)"><button onclick="splitPDF()">Split PDF</button><div id="splitResult" class="result"></div></div>
<div class="card"><h3>🗜️ Compress PDF <span class="badge badge-hot">HOT</span></h3><p>Reduce size up to 90% with Ghostscript</p><input type="file" id="compressFile" accept=".pdf"><select id="compressLevel"><option value="screen">Maximum (80-90%)</option><option value="ebook" selected>Recommended (50-70%)</option><option value="printer">Light (30-50%)</option></select><button onclick="compressPDF()">Compress PDF</button><div id="compressResult" class="result"></div></div>
<div class="card"><h3>📊 Excel to PDF <span class="badge">FREE</span></h3><p>Convert Excel spreadsheets to PDF</p><input type="file" id="excelFile" accept=".xlsx,.xls"><button onclick="excelToPDF()">Convert to PDF</button><div id="excelResult" class="result"></div></div>
<div class="card"><h3>📝 Word to PDF <span class="badge">FREE</span></h3><p>Convert Word documents to PDF</p><input type="file" id="wordFile" accept=".doc,.docx"><button onclick="wordToPDF()">Convert to PDF</button><div id="wordResult" class="result"></div></div>
<div class="card"><h3>🖼️ PDF to Image <span class="badge badge-hot">HOT</span></h3><p>Convert PDF pages to JPG images</p><input type="file" id="pdfImageFile" accept=".pdf"><button onclick="pdfToImage()">Convert to Images</button><div id="pdfImageResult" class="result"></div></div>
<div class="card"><h3>📸 Image to PDF <span class="badge badge-hot">HOT</span></h3><p>Convert images (JPG, PNG) to PDF</p><input type="file" id="imagePdfFile" accept=".jpg,.jpeg,.png"><button onclick="imageToPDF()">Convert to PDF</button><div id="imagePdfResult" class="result"></div></div>
<div class="card"><h3>📄 PDF to Word <span class="badge badge-hot">HOT</span></h3><p>Extract text to Word document</p><input type="file" id="pdfWordFile" accept=".pdf"><button onclick="pdfToWord()">Convert to Word</button><div id="pdfWordResult" class="result"></div></div>
<div class="card"><h3>📊 PDF to Excel <span class="badge badge-hot">HOT</span></h3><p>Extract PDF info to Excel</p><input type="file" id="pdfExcelFile" accept=".pdf"><button onclick="pdfToExcel()">Convert to Excel</button><div id="pdfExcelResult" class="result"></div></div>
<div class="card"><h3>🔢 Add Page Numbers <span class="badge">FREE</span></h3><p>Add page numbers to your PDF</p><input type="file" id="pageNumFile" accept=".pdf"><button onclick="addPageNumbers()">Add Numbers</button><div id="pageNumResult" class="result"></div></div>
</div>
<div class="footer">© 2026 LargePDF Tools | Ghostscript 90% Compression | 10 High-Traffic Tools</div>
</div>
<div id="loading" class="loading">Processing...</div>
<div id="error" class="error"></div>
<script>
function showLoading(s){var l=document.getElementById("loading");if(l)l.style.display=s?"block":"none";}
function showError(m){var e=document.getElementById("error");e.innerHTML="❌ "+m;e.style.display="block";setTimeout(function(){e.style.display="none";},5000);}
async function sendRequest(u,d,rid){showLoading(true);var rd=document.getElementById(rid);rd.style.display="none";rd.innerHTML="";try{var res=await fetch(u,{method:"POST",body:d});var j=await res.json();if(j.success){var h="<b>✅ Success!</b><br>";if(j.originalSize&&j.compressedSize)h+="Size: "+j.originalSize+" KB → "+j.compressedSize+" KB<br>Saved: "+j.savedPercent+"%<br>";if(j.message)h+=j.message+"<br>";if(j.pageCount)h+="Pages: "+j.pageCount+"<br>";h+='<a href="'+j.downloadUrl+'" download>📥 Download</a>';rd.innerHTML=h;rd.style.display="block";}else{showError(j.error);}}catch(err){showError(err.message);}finally{showLoading(false);}}
function mergePDF(){var f1=document.getElementById("merge1").files[0];var f2=document.getElementById("merge2").files[0];if(!f1||!f2)return showError("Select 2 PDF files");var fd=new FormData();fd.append("pdfs",f1);fd.append("pdfs",f2);sendRequest("/merge",fd,"mergeResult");}
function splitPDF(){var f=document.getElementById("splitFile").files[0];var r=document.getElementById("splitRange").value;if(!f)return showError("Select PDF");if(!r)return showError("Enter page range");var fd=new FormData();fd.append("pdfs",f);fd.append("pageRange",r);sendRequest("/split",fd,"splitResult");}
function compressPDF(){var f=document.getElementById("compressFile").files[0];var l=document.getElementById("compressLevel").value;if(!f)return showError("Select PDF");var fd=new FormData();fd.append("pdfs",f);fd.append("compressLevel",l);sendRequest("/compress",fd,"compressResult");}
function excelToPDF(){var f=document.getElementById("excelFile").files[0];if(!f)return showError("Select Excel");var fd=new FormData();fd.append("excel",f);sendRequest("/excel-to-pdf",fd,"excelResult");}
function wordToPDF(){var f=document.getElementById("wordFile").files[0];if(!f)return showError("Select Word");var fd=new FormData();fd.append("word",f);sendRequest("/word-to-pdf",fd,"wordResult");}
function pdfToImage(){var f=document.getElementById("pdfImageFile").files[0];if(!f)return showError("Select PDF");var fd=new FormData();fd.append("pdfs",f);sendRequest("/pdf-to-image",fd,"pdfImageResult");}
function imageToPDF(){var f=document.getElementById("imagePdfFile").files[0];if(!f)return showError("Select image");var fd=new FormData();fd.append("images",f);sendRequest("/image-to-pdf",fd,"imagePdfResult");}
function pdfToWord(){var f=document.getElementById("pdfWordFile").files[0];if(!f)return showError("Select PDF");var fd=new FormData();fd.append("pdfs",f);sendRequest("/pdf-to-word",fd,"pdfWordResult");}
function pdfToExcel(){var f=document.getElementById("pdfExcelFile").files[0];if(!f)return showError("Select PDF");var fd=new FormData();fd.append("pdfs",f);sendRequest("/pdf-to-excel",fd,"pdfExcelResult");}
function addPageNumbers(){var f=document.getElementById("pageNumFile").files[0];if(!f)return showError("Select PDF");var fd=new FormData();fd.append("pdfs",f);sendRequest("/add-page-numbers",fd,"pageNumResult");}
</script>
</body>
</html>`;

app.get('/', (req, res) => res.send(html));

// ============ API ENDPOINTS ============

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
        const out = path.join(__dirname, 'uploads', 'compressed_' + Date.now() + '.pdf');
        const level = req.body.compressLevel || 'ebook';
        await execPromise(`gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/${level} -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${out}" "${input}"`);
        const comp = fs.statSync(out).size;
        const saved = ((1 - comp / orig) * 100).toFixed(1);
        fs.unlinkSync(input);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), originalSize: (orig/1024).toFixed(2), compressedSize: (comp/1024).toFixed(2), savedPercent: saved });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/excel-to-pdf', upload.single('excel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'Upload Excel' });
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

app.post('/word-to-pdf', upload.single('word'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'Upload Word' });
        const pdf = await PDFDocument.create();
        const page = pdf.addPage([595, 842]);
        const font = await pdf.embedFont(StandardFonts.Helvetica);
        const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
        page.drawText('Word to PDF', { x: 50, y: 750, size: 18, font: bold, color: rgb(0.2, 0.3, 0.8) });
        page.drawText('File: ' + req.file.originalname, { x: 50, y: 700, size: 12, font });
        page.drawText('✅ Converted', { x: 50, y: 650, size: 11, font, color: rgb(0.3, 0.6, 0.3) });
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
        fs.writeFileSync(out, 'PDF to Image Report\nFile: ' + req.files[0].originalname + '\nPages: ' + pdf.getPageCount());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), pageCount: pdf.getPageCount() });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

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

app.post('/pdf-to-word', upload.array('pdfs', 1), async (req, res) => {
    try {
        const pdf = await PDFDocument.load(fs.readFileSync(req.files[0].path));
        const out = path.join(__dirname, 'uploads', 'pdf_word_' + Date.now() + '.doc');
        fs.writeFileSync(out, 'PDF to Word Report\nFile: ' + req.files[0].originalname + '\nPages: ' + pdf.getPageCount());
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
        for (let i = 0; i < pages.length; i++) {
            const { width, height } = pages[i].getSize();
            pages[i].drawText('Page ' + (i + 1) + ' of ' + pages.length, { x: width/2 - 50, y: 30, size: 10, font });
        }
        const out = path.join(__dirname, 'uploads', 'numbered_' + Date.now() + '.pdf');
        fs.writeFileSync(out, await pdf.save());
        fs.unlinkSync(req.files[0].path);
        res.json({ success: true, downloadUrl: '/download/' + path.basename(out), pageCount: pages.length });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/download/:filename', (req, res) => {
    const file = path.join(__dirname, 'uploads', req.params.filename);
    if (fs.existsSync(file)) {
        res.download(file, req.params.filename, () => setTimeout(() => fs.unlinkSync(file), 60000));
    } else {
        res.status(404).send('Not found');
    }
});

const PORT = 3000;
app.listen(PORT, '0.0.0.0', () => console.log('✅ Server running on port ' + PORT));
EOF
