// server.js
const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const cors = require('cors'); //以此避免跨域问题，虽然同源不需要，但加个保险

const app = express();
const PORT = 3000;

//Config: 请确保这里的文件名和你真实的文件名一模一样（区分大小写）
const EXCEL_FILE_NAME = 'N related.xlsx'; 

// 1. 静态文件托管 (使得浏览器能访问 public/index.html)
app.use(express.static('public'));
app.use(cors());
app.use(express.json({ limit: '10mb' }));


// 辅助函数：获取文件路径并检查
function getExcelPath() {
    const filePath = path.join(__dirname, EXCEL_FILE_NAME);
    // console.log(`[检查路径] 正在寻找文件: ${filePath}`);
    
    if (!fs.existsSync(filePath)) {
        console.error(`[错误]文件不存在! 请检查:`);
        console.error(`   1. 文件名是否真的是 "${EXCEL_FILE_NAME}"?`);
        console.error(`   2. 文件是否在 "${__dirname}" 目录下?`);
        return null;
    }
    return filePath;
}

// API 1: 获取所有 Sheet 名称
app.get('/api/sheets', (req, res) => {
    const filePath = getExcelPath();
    
    if (!filePath) return res.status(404).json({ error: '文件未找到' });

    try {
        console.log('--- 正在读取 Excel 文件 ---');
        
        // 修改关键点：去掉 { bookProps: true }
        // 改为直接读取，这样最稳，虽然慢几毫秒，但不会报错
        const workbook = xlsx.readFile(filePath);

        const sheetNames = workbook.SheetNames;
        console.log(`[成功] 读取到 Sheet: ${sheetNames.join(', ')}`);
        
        res.json({ sheets: sheetNames });
    } catch (e) {
        console.error('[读取失败]', e.message);
        res.status(500).json({ error: '读取 Excel 失败，请确认文件未加密且不是损坏的' });
    }
});

// API 2: 获取具体 Sheet 的数据 (你的代码里缺了这个，必须补上！)
app.get('/api/data/:sheetName', (req, res) => {
    const sheetName = req.params.sheetName;
    // console.log(`--- 请求: 获取 Sheet [${sheetName}] 的数据 ---`);
    
    const filePath = getExcelPath();
    if (!filePath) return res.status(404).json({ error: '文件未找到' });

    try {
        const workbook = xlsx.readFile(filePath);
        
        // 检查 Sheet 是否存在
        if (!workbook.Sheets[sheetName]) {
            console.error(`[错误] Sheet "${sheetName}" 不存在`);
            return res.status(404).json({ error: 'Sheet 不存在' });
        }

        // 将 Excel 数据转换为 JSON 数组
        const worksheet = workbook.Sheets[sheetName];
        // defval: '' 保证空单元格也有字段
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: '' }); 

        // console.log(`[成功] 读取到 ${jsonData.length} 条数据`);
        res.json(jsonData);
    } catch (e) {
        console.error('[异常]', e.message);
        res.status(500).json({ error: e.message });
    }
});


// ============================================================
// Layout positions persistence (save to file)
// ============================================================
const LAYOUT_FILE = path.join(__dirname, 'layout_positions.json');
const ORGANELLES_FILE = path.join(__dirname, 'organelles_layout.json');

function writeJsonAtomic(filePath, dataObj) {
    const tmpPath = filePath + '.tmp';
    fs.writeFileSync(tmpPath, JSON.stringify(dataObj, null, 2), 'utf-8');
    fs.renameSync(tmpPath, filePath);
}

function ensureOrganellesFile(){
    if (!fs.existsSync(ORGANELLES_FILE)) {
        // Default organelle layout: world/model coordinates (so it pans/zooms with Cytoscape)
        writeJsonAtomic(ORGANELLES_FILE, { version: 3, coordMode: 'world', opacity: 0.18, organelles: {} });
    }
}

// 获取上一次保存的布局
app.get('/api/layout', (req, res) => {
    try {
        if (!fs.existsSync(LAYOUT_FILE)) {
            return res.status(404).json({ error: 'layout file not found' });
        }
        const raw = fs.readFileSync(LAYOUT_FILE, 'utf-8');
        const obj = JSON.parse(raw || '{}');
        return res.json(obj);
    } catch (e) {
        console.error('[layout read error]', e.message);
        return res.status(500).json({ error: e.message });
    }
});

// 保存布局（positions: { nodeId: {x,y} }）
app.post('/api/layout', (req, res) => {
    try {
        const body = req.body || {};
        const positions = body.positions && typeof body.positions === 'object' ? body.positions : body;
        if (!positions || typeof positions !== 'object') {
            return res.status(400).json({ error: 'invalid positions payload' });
        }
        writeJsonAtomic(LAYOUT_FILE, positions);
        return res.json({ ok: true, count: Object.keys(positions).length });
    } catch (e) {
        console.error('[layout write error]', e.message);
        return res.status(500).json({ error: e.message });
    }
});

// 清除布局
app.delete('/api/layout', (req, res) => {
    try {
        if (fs.existsSync(LAYOUT_FILE)) fs.unlinkSync(LAYOUT_FILE);
        return res.json({ ok: true });
    } catch (e) {
        console.error('[layout delete error]', e.message);
        return res.status(500).json({ error: e.message });
    }
});

// === Organelles layout (background images) ===
app.get('/api/organelles', (req, res) => {
    try{
        ensureOrganellesFile();
        const raw = fs.readFileSync(ORGANELLES_FILE, 'utf-8');
        res.json(JSON.parse(raw));
    }catch(e){
        console.error('Read organelles layout failed:', e);
        res.status(500).json({ error: 'Failed to read organelles layout.' });
    }
});

app.post('/api/organelles', (req, res) => {
    try{
        const payload = req.body || {};
        const tmpPath = ORGANELLES_FILE + '.tmp';
        fs.writeFileSync(tmpPath, JSON.stringify(payload, null, 2), 'utf-8');
        fs.renameSync(tmpPath, ORGANELLES_FILE);
        res.json({ ok: true });
    }catch(e){
        console.error('Save organelles layout failed:', e);
        res.status(500).json({ error: 'Failed to save organelles layout.' });
    }
});

app.delete('/api/organelles', (req, res) => {
    try{
        const tmpPath = ORGANELLES_FILE + '.tmp';
        fs.writeFileSync(tmpPath, JSON.stringify({ version: 3, coordMode: 'world', opacity: 0.18, organelles: {} }, null, 2), 'utf-8');
        fs.renameSync(tmpPath, ORGANELLES_FILE);
        res.json({ ok: true });
    }catch(e){
        console.error('Reset organelles layout failed:', e);
        res.status(500).json({ error: 'Failed to reset organelles layout.' });
    }
});


// 启动服务器
app.listen(PORT, () => {
    console.log(`\n===============================================`);
    console.log(`✅ 服务已启动`);
    console.log(`📂 请确保 Excel 文件 "${EXCEL_FILE_NAME}" 已放在同级目录`);
    console.log(`👉 访问地址: http://localhost:${PORT}`);
    console.log(`===============================================\n`);
});
