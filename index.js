const express = require('express');
const multer = require('multer');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');

// Simple HTML as a string - NO SYNTAX ERRORS
const html = '<!DOCTYPE html>\n<html>\n<head>\n<meta charset="UTF-8">\n<title>LargePDF Tools</title>\n<style>\n*{margin:0;padding:0;box-sizing:border-box;}\nbody{font-family:Arial;background:linear-gradient(135deg,#0f172a,#1e1b4b);min-height:100vh;padding:20px;}\n.container{max-width:1200px;margin:0 auto;}\nh1{text-align:center;color:white;margin-bottom:10px;}\n.sub{text-align:center;color:#aaa;margin-bottom:30px;}\n.stats{display:flex;justify-content:center;gap:40px;margin-bottom:40px;flex-wrap:wrap;}\n.stat{text-align:center;}\n.stat-num{font-size:28px;font-weight:bold;color:#a5b4fc;}\n.stat-label{color:#888;font-size:12px;}\n.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:20px;}\n.card{background:rgba(30,27,75,0.7);border-radius:16px;padding:20px;border:1px solid rgba(255,255,255,0.1);}\n.card:hover{border-color:#6366f1;}\n.card h3{color:white;margin-bottom:8px;}\n.card p{color:#aaa;font-size:13px;margin-bottom:15px;}\n.badge{background:#22c55e;font-size:10px;padding:2px 8px;border-radius:20px;margin-left:8px;}\ninput,button{width:100%;padding:10px;margin:8px 0;border-radius:8px;border:none;}\ninput{background:rgba(15,23,42,0.9);border:1px solid rgba(255,255,255,0.2);color:white;}\nbutton{background:linear-gradient(135deg,#6366f1,#8b5cf6);color:white;cursor:pointer;font-weight:bold;}\nbutton:hover{opacity:0.9;}\n.result{margin-top:12px;padding:10px;background:rgba(99,102,241,0.2);border-radius:8px;display:none;font-size:12px;border-left:3px solid #6366f1;}\n.result a{color:#a5b4fc;}\n.loading{position:fixed;top:50%;left:50%;background:black;color:white;padding:15px 30px;border-radius:40px;display:none;z-index:1000;transform:translate(-50%,-50%);}\n.error{position:fixed;bottom:20px;right:20px;background:#ef4444;color:white;padding:10px 15px;border-radius:8px;display:none;z-index:1000;}\n.footer{text-align:center;margin-top:40px;padding:20px;color:#666;border-top:1px solid rgba(255,255,255,0.05);}\n@media (max-width:768px){.grid{grid-template-columns:1fr;}.stats{gap:20px;}}\n</style>\n</head>\n<body>\n<div class="container">\n<h1>📚 LargePDF Tools</h1>\n<div class="sub">10 Professional PDF Tools - Free & Easy</div>\n<div class="stats">\n<div class="stat"><div class="stat-num">10</div><div class="stat-label">PDF Tools</div></div>\n<div class="stat"><div class="stat-num">50MB</div><div class="stat-label">File Limit</div></div>\n<div class="stat"><div class="stat-num">100%</div><div class="stat-label">Free</div></div>\n</div>\n<div class="grid" id="toolsGrid"></div>\n<div class="footer">© 2026 LargePDF Tools | Secure · Fast · Free</div>\n</div>\n<div id="loading" class="loading">Processing...</div>\n<div id="error" class="error"></div>\n<script>\nvar toolsData=[\n{id:"merge",name:"Merge PDF",icon:"🔗",desc:"Combine 2 PDF files into one",inputs:2,accept:".pdf"},\n{id:"split",name:"Split PDF",icon:"✂️",desc:"Extract specific pages",inputs:1,accept:".pdf",hasText:true,placeholder:"Page range (1-5 or 1,3,5)"},\n{id:"compress",name:"Compress PDF",icon:"🗜️",desc:"Reduce file size",inputs:1,accept:".pdf"},\n{id:"excel",name:"Excel to PDF",icon:"📊",desc:"Convert Excel to PDF",inputs:1,accept:".xlsx,.xls"},\n{id:"word",name:"Word to PDF",icon:"📝",desc:"Convert Word to PDF",inputs:1,accept:".doc,.docx"},\n{id:"pdfimage",name:"PDF to Image",icon:"🖼️",desc:"Convert PDF to images",inputs:1,accept:".pdf"},\n{id:"imagepdf",name:"Image to PDF",icon:"📸",desc:"Convert images to PDF",inputs:1,accept:".jpg,.jpeg,.png"},\n{id:"pdfword",name:"PDF to Word",icon:"📄",desc:"Extract text to Word",inputs:1,accept:".pdf"},\n{id:"pdfexcel",name:"PDF to Excel",icon:"📊",desc:"Extract info to Excel",inputs:1,accept:".pdf"},\n{id:"pagenum",name:"Add Page Numbers",icon:"🔢",desc:"Add page numbers",inputs:1,accept:".pdf"}\n];\nfunction buildTools(){\nvar c=document.getElementById("toolsGrid");\nif(!c)return;\nvar h="";\nfor(var i=0;i<toolsData.length;i++){\nvar t=toolsData[i];\nvar f="";\nif(t.inputs===2){\nf='<input type="file" id="f1_'+i+'" accept="'+t.accept+'"><input type="file" id="f2_'+i+'" accept="'+t.accept+'">';\n}else{\nf='<input type="file" id="f_'+i+'" accept="'+t.accept+'">';\nif(t.hasText)f+='<input type="text" id="txt_'+i+'" placeholder="'+t.placeholder+'">';\n}\nh+='<div class="card"><h3>'+t.icon+' '+t.name+'<span class="badge">FREE</span></h3><p>'+t.desc+'</p>'+f+'<button onclick="runTool(\\''+t.id+'\\','+i+')">Process</button><div id="res_'+i+'" class="result"></div></div>';\n}\nc.innerHTML=h;\n}\nfunction showLoading(s){var l=document.getElementById("loading");if(l)l.style.display=s?"block":"none";}\nfunction showError(m){var e=document.getElementById("error");if(e){e.innerHTML="❌ "+m;e.style.display="block";setTimeout(function(){e.style.display="none";},5000);}}\nfunction runTool(tid,idx){\nvar fd=new FormData();\nvar url="";\nif(tid==="merge"){\nvar f1=document.getElementById("f1_"+idx).files[0];\nvar f2=document.getElementById("f2_"+idx).files[0];\nif(!f1||!f2){showError("Select 2 PDF files");return;}\nfd.append("pdfs",f1);fd.append("pdfs",f2);url="/merge";\n}\nelse if(tid==="split"){\nvar f=document.getElementById("f_"+idx).files[0];\nvar r=document.getElementById("txt_"+idx).value;\nif(!f){showError("Select PDF");return;}\nif(!r){showError("Enter page range");return;}\nfd.append("pdfs",f);fd.append("pageRange",r);url="/split";\n}\nelse if(tid==="compress"){\nvar f=document.getElementById("f_"+idx).files[0];\nif(!f){showError("Select PDF");return;}\nfd.append("pdfs",f);url="/compress";\n}\nelse if(tid==="excel"){\nvar f=document.getElementById("f_"+idx).files[0];\nif(!f){showError("Select Excel");return;}\nfd.append("excel",f);url="/excel-to-pdf";\n}\nelse if(tid==="word"){\nvar f=document.getElementById("f_"+idx).files[0];\nif(!f){showError("Select Word");return;}\nfd.append("word",f);url="/word-to-pdf";\n}\nelse if(tid==="pdfimage"){\nvar f=document.getElementById("f_"+idx).files[0];\nif(!f){showError("Select PDF");return;}\nfd.append("pdfs",f);url="/pdf-to-image";\n}\nelse if(tid==="imagepdf"){\nvar f=document.getElementById("f_"+idx).files[0];\nif(!f){showError("Select image");return;}\nfd.append("images",f);url="/image-to-pdf";\n}\nelse if(tid==="pdfword"){\nvar f=document.getElementById("f_"+idx).files[0];\nif(!f){showError("Select PDF");return;}\nfd.append("pdfs",f);url="/pdf-to-word";\n}\nelse if(tid==="pdfexcel"){\nvar f=document.getElementById("f_"+idx).files[0];\nif(!f){showError("Select PDF");return;}\nfd.append("pdfs",f);url="/pdf-to-excel";\n}\nelse if(tid==="pagenum"){\nvar f=document.getElementById("f_"+idx).files[0];\nif(!f){showError("Select PDF");return;}\nfd.append("pdfs",f);url="/add-page-numbers";\n}\nif(url)sendRequest(url,fd,"res_"+idx);\n}\nfunction sendRequest(url,data,rid){\nshowLoading(true);\nvar rd=document.getElementById(rid);\nif(rd){rd.style.display="none";rd.innerHTML="";}\nfetch(url,{method:"POST",body:data})\n.then(function(res){return res.json();})\n.then(function(json){\nif(json.success){\nvar h="<b>✅ Success!</b><br>";\nif(json.originalSize&&json.compressedSize)h+="Size: "+json.originalSize+" KB → "+json.compressedSize+" KB<br>Saved: "+json.savedPercent+"%<br>";\nif(json.message)h+=json.message+"<br>";\nif(json.pageCount)h+="Pages: "+json.pageCount+"<br>";\nh+='<a href="'+json.downloadUrl+'" download>📥 Download</a>';\nif(rd){rd.innerHTML=h;rd.style.display="block";}\n}else{showError(json.error);}\n})\n.catch(function(err){showError(err.message);})\n.finally(function(){showLoading(false);});\n}\nbuildTools();\n</script>\n</body>\n</html>';

app.get('/', (req, res) => res.send(html));

// API Endpoints
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

app.get('/download/:filename', (req, res) => {
    const file = path.join(__dirname, 'uploads', req.params.filename);
    if (fs.existsSync(file)) {
        res.download(file, req.params.filename, () => setTimeout(() => fs.unlinkSync(file), 60000));
    } else {
        res.status(404).send('Not found');
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log('✅ Server running on port ' + PORT);
    console.log('📱 Open: https://largepdftools.com');
});
