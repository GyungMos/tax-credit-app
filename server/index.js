const express = require('express');
const xlsx = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const chokidar = require('chokidar');
const mongoose = require('mongoose');

const app = express();
const PORT = process.env.PORT || 5010;
const WATCH_DIR = process.env.DATA_DIR || path.join(__dirname, '..', 'data');
const MONGODB_URI = process.env.MONGODB_URI;

if (!fs.existsSync(WATCH_DIR)) {
    fs.mkdirSync(WATCH_DIR, { recursive: true });
}

// --- MongoDB Setup ---
const FileSchema = new mongoose.Schema({
    originalName: String,
    path: String, // Relative to WATCH_DIR
    data: Buffer,
    year: Number,
    lastModified: Number
});

const ConfigSchema = new mongoose.Schema({
    key: { type: String, unique: true },
    content: Object
});

const FileModel = mongoose.model('File', FileSchema);
const ConfigModel = mongoose.model('Config', ConfigSchema);

async function syncToDB(filePath, isDelete = false) {
    if (!MONGODB_URI) return;
    try {
        const relativePath = path.relative(WATCH_DIR, filePath).replace(/\\/g, '/');
        if (isDelete) {
            await FileModel.deleteOne({ path: relativePath });
            console.log(`DB: Deleted ${relativePath}`);
        } else if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
            const stats = fs.statSync(filePath);
            const content = fs.readFileSync(filePath);
            await FileModel.findOneAndUpdate(
                { path: relativePath },
                { 
                    originalName: path.basename(filePath),
                    data: content,
                    lastModified: stats.mtimeMs
                },
                { upsert: true }
            );
            console.log(`DB: Saved ${relativePath}`);
        }
    } catch (e) {
        console.error('DB Sync Error:', e);
    }
}

async function restoreFromDB() {
    if (!MONGODB_URI) return;
    console.log('Restoring data from MongoDB...');
    try {
        const configs = await ConfigModel.find();
        for (const cfg of configs) {
            const fileName = cfg.key === 'mapping' ? 'mapping.json' : 'activity.json';
            const filePath = path.join(WATCH_DIR, fileName);
            fs.writeFileSync(filePath, JSON.stringify(cfg.content, null, 2));
            console.log(`Restored Config: ${fileName}`);
        }

        const files = await FileModel.find();
        for (const file of files) {
            const targetPath = path.join(WATCH_DIR, file.path);
            const targetDir = path.dirname(targetPath);
            if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });
            fs.writeFileSync(targetPath, file.data);
            console.log(`Restored File: ${file.path}`);
        }
        console.log('Database restoration complete.');
    } catch (e) {
        console.error('DB Restore Error:', e);
    }
}

if (MONGODB_URI) {
    mongoose.connect(MONGODB_URI)
        .then(() => {
            console.log('Connected to MongoDB Atlas');
            restoreFromDB().then(() => updateAllData());
        })
        .catch(err => console.error('MongoDB Connection Error:', err));
}

function decodeText(text) {
    if (!text) return text;
    try {
        return Buffer.from(text, 'latin1').toString('utf8');
    } catch (e) {
        return text;
    }
}

app.use(cors());
app.use(express.json());

const MAPPING_FILE = path.join(WATCH_DIR, 'mapping.json');
const ACTIVITY_FILE = path.join(WATCH_DIR, 'activity.json');

let recentActivities = [];
function loadActivity() {
    if (fs.existsSync(ACTIVITY_FILE)) {
        try { recentActivities = JSON.parse(fs.readFileSync(ACTIVITY_FILE, 'utf8')); } catch (e) { recentActivities = []; }
    }
}
async function logActivity(type, msg) {
    recentActivities.unshift({ type, msg, timestamp: new Date().toISOString() });
    recentActivities = recentActivities.slice(0, 50);
    try { 
        fs.writeFileSync(ACTIVITY_FILE, JSON.stringify(recentActivities, null, 2)); 
        if (MONGODB_URI) await ConfigModel.findOneAndUpdate({ key: 'activity' }, { content: recentActivities }, { upsert: true });
    } catch (e) {}
}
loadActivity();

function getAvailableYears() {
    try {
        const items = fs.readdirSync(WATCH_DIR);
        let years = items.filter(i => /^\d{4}$/.test(i)).map(i => parseInt(i));
        // Always include 2022~2026 as a baseline
        const baseline = [2022, 2023, 2024, 2025, 2026];
        years = [...years, ...baseline];
        const currentYear = new Date().getFullYear();
        if (!years.includes(currentYear)) years.push(currentYear);
        return Array.from(new Set(years)).sort((a,b) => b - a);
    } catch (e) {
        return [2022, 2023, 2024, 2025, 2026];
    }
}

let SUPPORTED_YEARS = getAvailableYears();
SUPPORTED_YEARS.forEach(yr => {
    const yrDir = path.join(WATCH_DIR, yr.toString());
    if (!fs.existsSync(yrDir)) fs.mkdirSync(yrDir, { recursive: true });
});

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const year = req.query.year || new Date().getFullYear().toString();
        const dest = path.join(WATCH_DIR, year.toString());
        if (!fs.existsSync(dest)) fs.mkdirSync(dest, { recursive: true });
        cb(null, dest);
    },
    filename: (req, file, cb) => {
        const branch = req.query.branch;
        const originalName = decodeText(file.originalname);
        if (branch && branch !== 'auto' && branch !== 'undefined') {
            cb(null, `${branch}_사원등록.xlsx`);
        } else {
            cb(null, originalName);
        }
    }
});
const upload = multer({ storage });

let mapping = { corporations: [] };
function loadMapping() {
    if (fs.existsSync(MAPPING_FILE)) {
        try {
            const raw = fs.readFileSync(MAPPING_FILE, 'utf8');
            const parsed = JSON.parse(raw);
            mapping = (Array.isArray(parsed)) ? { corporations: parsed } : (parsed.corporations ? parsed : { corporations: [] });
        } catch (e) { mapping = { corporations: [] }; }
    } else { mapping = { corporations: [] }; }
}

async function saveMapping() {
    try {
        fs.writeFileSync(MAPPING_FILE, JSON.stringify(mapping, null, 2));
        if (MONGODB_URI) await ConfigModel.findOneAndUpdate({ key: 'mapping' }, { content: mapping }, { upsert: true });
    } catch (e) { console.error('Save mapping failed:', e); }
}

let processedResults = { updatedAt: new Date().toISOString(), corporations: {}, unassigned: {} };

function getBirthDate(rrn) {
    if (!rrn) return null;
    const s = rrn.toString().replace(/-/g, '');
    if (s.length < 7) return null;
    const yy = s.substring(0, 2), mm = s.substring(2, 4), dd = s.substring(4, 6), gender = s.charAt(6);
    let yearPrefix = (gender === '3' || gender === '4' || gender === '7' || gender === '8') ? '20' : '19';
    return `${yearPrefix}${yy}${mm}${dd}`;
}

function calculateAgeAtJoin(birth, joinDate) {
    if (!birth || !joinDate) return 0;
    const bY = parseInt(birth.substring(0, 4)), bM = parseInt(birth.substring(4, 6)), bD = parseInt(birth.substring(6, 8));
    const jStr = joinDate.toString(), jY = parseInt(jStr.substring(0, 4)), jM = parseInt(jStr.substring(4, 6)), jD = parseInt(jStr.substring(6, 8));
    let age = jY - bY;
    if (jM < bM || (jM === bM && jD < bD)) age--;
    return age;
}

function parseBranchName(fileName) {
    const match = fileName.match(/(.+)_사원등록/);
    return match ? match[1] : fileName.replace('.xlsx', '');
}

function processFile(filePath) {
    try {
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = xlsx.utils.sheet_to_json(sheet);
        const validation = [];
        rawData.forEach((emp, idx) => {
            const rowNum = idx + 2, name = (emp.사원명 || '').toString().trim(), rrn = (emp['주민(외국인)등록번호'] || '').toString().replace(/-/g, '').trim(), joinDate = (emp.입사일자 || '').toString().trim();
            if (!name) validation.push({ row: rowNum, type: 'error', field: '사원명', msg: '사원명이 누락되었습니다.' });
            if (!rrn) validation.push({ row: rowNum, type: 'error', field: '주민번호', msg: '주민번호가 누락되었습니다.' });
            if (!joinDate) validation.push({ row: rowNum, type: 'error', field: '입사일자', msg: '입사일자가 누락되었습니다.' });
        });
        const uniqueEmployeesMap = new Map();
        rawData.forEach(emp => {
            const name = (emp.사원명 || '').toString().trim(), rrn = (emp['주민(외국인)등록번호'] || '').toString().replace(/-/g, '').trim(), joinDate = (emp.입사일자 || '').toString().trim(), key = `${name}|${rrn}|${joinDate}`;
            if (key !== '||' && !uniqueEmployeesMap.has(key)) uniqueEmployeesMap.set(key, emp);
        });
        const data = Array.from(uniqueEmployeesMap.values()), years = [2022, 2023, 2024, 2025, 2026];
        const branchDataMap = {};
        years.forEach(year => {
            const deptGroups = {};
            data.forEach(emp => {
                let dept = (emp.부서 || '미지정').toString().trim();
                const originalDept = dept;
                dept = dept.replace(/^\d{4}/, '').replace(/.*평우서비스/, '').trim();
                if (!dept) dept = originalDept;
                if (!deptGroups[dept]) deptGroups[dept] = [];
                deptGroups[dept].push(emp);
            });
            Object.keys(deptGroups).forEach(deptName => {
                const deptData = deptGroups[deptName];
                if (!branchDataMap[deptName]) branchDataMap[deptName] = { branch: deptName, years: {} };
                const monthly = [];
                for (let m = 1; m <= 12; m++) {
                    const lastDay = new Date(year, m, 0).toISOString().split('T')[0].replace(/-/g, '');
                    let tW = 0, yW = 0;
                    deptData.forEach(emp => {
                        if (emp.입사일자 <= lastDay && (emp.퇴사일자 || '99991231') >= lastDay) {
                            const weight = parseFloat(emp.단시간유형 || emp.가중치 || 1.0);
                            tW += weight;
                            const birth = getBirthDate(emp['주민(외국인)등록번호']), age = calculateAgeAtJoin(birth, emp.입사일자.toString());
                            if (age >= 15 && age <= 34) yW += weight;
                        }
                    });
                    monthly.push({ total: tW, youth: yW });
                }
                const sumT = monthly.reduce((s, h) => s + h.total, 0), sumY = monthly.reduce((s, h) => s + h.youth, 0);
                const avgT = parseFloat((sumT / 12).toFixed(2)), avgY = parseFloat((sumY / 12).toFixed(2));
                const details = deptData.filter(emp => emp.입사일자 <= `${year}1231` && (emp.퇴사일자 || '99991231') >= `${year}0101`).map(emp => {
                    let monthsActiveCount = 0, startM = null, endM = null;
                    for (let m = 1; m <= 12; m++) {
                        const lDay = new Date(year, m, 0).toISOString().split('T')[0].replace(/-/g, '');
                        if (emp.입사일자 <= lDay && (emp.퇴사일자 || '99991231') >= lDay) {
                            monthsActiveCount++;
                            if (startM === null) startM = m;
                            endM = m;
                        }
                    }
                    const birth = getBirthDate(emp['주민(외국인)등록번호']), ageAtJoin = calculateAgeAtJoin(birth, emp.입사일자.toString()), isYouthAtJoin = (ageAtJoin >= 15 && ageAtJoin <= 34);
                    let exclusionDate = ''; if (isYouthAtJoin && birth) exclusionDate = `${parseInt(birth.substring(0, 4)) + 35}-${birth.substring(4, 6)}-${birth.substring(6, 8)}`;
                    const address = emp.주소 || emp.근무지 || emp.사업장 || '', isMetro = address.includes('서울') || address.includes('경기');
                    const retiredThisYear = !!(emp.퇴사일자 && emp.퇴사일자.toString().startsWith(year.toString()));
                    return {
                        col14: birth ? `${birth.substring(0, 4)}-${birth.substring(4, 6)}-${birth.substring(6, 8)}` : '',
                        col15: emp.사원명 || '', col16: 'Y', col17: '', col18: emp.단시간유형 || emp.가중치 || '1.0',
                        col19: (emp.내외국인구분 || '').includes('내국인') ? 'Y' : 'N', col20: isMetro ? 'Y' : 'N',
                        col21: emp.입사일자 ? `${emp.입사일자.toString().substring(0, 4)}-${emp.입사일자.toString().substring(4, 6)}-${emp.입사일자.toString().substring(6, 8)}` : '',
                        col22: startM ? `${startM.toString().padStart(2, '0')}~${endM.toString().padStart(2, '0')}` : '',
                        col23: monthsActiveCount, col24: isYouthAtJoin ? monthsActiveCount : 0, col25: isYouthAtJoin ? 'Y' : 'N', col26: isYouthAtJoin ? 'Y' : 'N', col27: exclusionDate, col28: emp.장애인 > 0 ? 'Y' : 'N', col29: (emp.나이 || 0) >= 60 ? 'Y' : 'N', col30: (emp.경력단절 || 0) > 0 ? 'Y' : 'N', col31: (emp.북약 || 0) > 0 ? 'Y' : 'N', col32: 'N', col33: 'N', col34: '', col35: '', _retiredThisYear: retiredThisYear
                    };
                });
                branchDataMap[deptName].years[year] = { monthly, avgTotal: avgT, avgYouth: avgY, avgOther: parseFloat((avgT - avgY).toFixed(2)), sumTotal: parseFloat(sumT.toFixed(2)), sumYouth: parseFloat(sumY.toFixed(2)), summary: { col1: sumT.toFixed(2), col2: '12', col3: avgT.toFixed(2), col5: sumY.toFixed(2), col6: '0', col7: '0', col8: '0', col9: '0', col10: sumY.toFixed(2), col11: '12', col12: avgY.toFixed(2), col13: (avgT - avgY).toFixed(2) }, details };
            });
        });
        return { branchDataMap, validation };
    } catch (e) { console.error(`Error processing file ${filePath}:`, e); return null; }
}

async function updateAllData() {
    try {
        loadMapping(); 
        SUPPORTED_YEARS = getAvailableYears();
        // Ensure baseline folders exist
        SUPPORTED_YEARS.forEach(yr => {
            const yrDir = path.join(WATCH_DIR, yr.toString());
            if (!fs.existsSync(yrDir)) fs.mkdirSync(yrDir, { recursive: true });
        });
        const allProcessed = {}, branchFiles = {}, scanDirs = [WATCH_DIR, ...SUPPORTED_YEARS.map(yr => path.join(WATCH_DIR, yr.toString()))];
        for (const dir of scanDirs) {
            if (!fs.existsSync(dir)) continue;
            const dirYear = dir === WATCH_DIR ? null : parseInt(path.basename(dir));
            const files = fs.readdirSync(dir).filter(f => f.includes('사원등록') && f.endsWith('.xlsx'));
            for (const f of files) {
                const fullPath = path.join(dir, f), stats = fs.statSync(fullPath);
                let yr = dirYear; if (!yr) { const match = f.match(/^(\d{4})/); if (match) yr = parseInt(match[1]); }
                if (yr && SUPPORTED_YEARS.includes(yr)) {
                    const branch = parseBranchName(f);
                    if (!branchFiles[branch]) branchFiles[branch] = {};
                    if (!branchFiles[branch][yr] || branchFiles[branch][yr].mtime < stats.mtime.getTime()) branchFiles[branch][yr] = { name: f, fullPath, branch, mtime: stats.mtime.getTime(), year: yr };
                }
            }
        }
        const validationReports = {};
        for (const branchKey of Object.keys(branchFiles)) {
            const yearsData = branchFiles[branchKey];
            for (const yr of Object.keys(yearsData)) {
                const result = processFile(yearsData[yr].fullPath);
                if (result) {
                    const { branchDataMap, validation } = result;
                    if (validation && validation.length > 0) { const fileName = yearsData[yr].name; if (!validationReports[fileName]) validationReports[fileName] = []; validationReports[fileName].push(...validation); }
                    if (branchDataMap) {
                        Object.keys(branchDataMap).forEach(deptName => {
                            const deptResult = branchDataMap[deptName];
                            if (!allProcessed[deptName]) {
                                allProcessed[deptName] = { branch: deptName, updatedAt: new Date().toISOString(), years: {} };
                                [2022, 2023, 2024, 2025, 2026].forEach(y => { allProcessed[deptName].years[y] = { monthly: Array(12).fill({total:0, youth:0}), avgTotal:0, avgYouth:0, avgOther:0, sumTotal:0, sumYouth:0, summary:{}, details:[] }; });
                            }
                            if (deptResult.years[yr]) allProcessed[deptName].years[yr] = deptResult.years[yr];
                        });
                    }
                }
            }
        }
        const newCorpData = {}, assignedBranches = new Set(), normalizeForMatch = (name) => name.replace(/^\d{4}/, '').replace(/.*평우서비스/, '').trim();
        mapping.corporations.forEach(corp => {
            newCorpData[corp.name] = { branches: {}, total: null };
            (corp.branchNames || []).forEach(bn => {
                const normalizedBN = normalizeForMatch(bn);
                if (allProcessed[normalizedBN]) { newCorpData[corp.name].branches[normalizedBN] = allProcessed[normalizedBN]; assignedBranches.add(normalizedBN); }
                else if (allProcessed[bn]) { newCorpData[corp.name].branches[bn] = allProcessed[bn]; assignedBranches.add(bn); }
            });
            const brs = Object.values(newCorpData[corp.name].branches);
            if (brs.length > 0) {
                const consolidated = { years: {} };
                [2022, 2023, 2024, 2025, 2026].forEach(yr => {
                    let tAvgT = 0, tAvgY = 0, tSumT = 0, tSumY = 0; const combMonthly = Array.from({ length: 12 }, () => ({ total: 0, youth: 0 })), combDetails = [];
                    brs.forEach(b => { const yD = b.years[yr]; if (!yD) return; tAvgT += yD.avgTotal || 0; tAvgY += yD.avgYouth || 0; tSumT += yD.sumTotal || 0; tSumY += yD.sumYouth || 0; if (yD.monthly) yD.monthly.forEach((m, i) => { if (combMonthly[i]) { combMonthly[i].total += m.total || 0; combMonthly[i].youth += m.youth || 0; } }); if (yD.details) combDetails.push(...yD.details); });
                    consolidated.years[yr] = { monthly: combMonthly, avgTotal: parseFloat(tAvgT.toFixed(2)), avgYouth: parseFloat(tAvgY.toFixed(2)), avgOther: parseFloat((tAvgT - tAvgY).toFixed(2)), sumTotal: parseFloat(tSumT.toFixed(2)), sumYouth: parseFloat(tSumY.toFixed(2)), summary: { col1: tSumT.toFixed(2), col2: '12', col3: tAvgT.toFixed(2), col5: tSumY.toFixed(2), col6: '0', col7: '0', col8: '0', col9: '0', col10: tSumY.toFixed(2), col11: '12', col12: tAvgY.toFixed(2), col13: (tAvgT - tAvgY).toFixed(2) }, details: combDetails };
                });
                newCorpData[corp.name].total = consolidated;
            }
        });
        const unassigned = {}; Object.keys(allProcessed).forEach(bn => { if (!assignedBranches.has(bn)) unassigned[bn] = allProcessed[bn]; });
        processedResults = { updatedAt: new Date().toISOString(), corporations: newCorpData, unassigned, mapping: mapping.corporations, validation: validationReports, years: SUPPORTED_YEARS, activities: recentActivities };
    } catch (err) { console.error('CRITICAL ERROR in updateAllData:', err); }
}

try {
    chokidar.watch(WATCH_DIR, { ignored: /(^|[\/\\])\..|node_modules|_초기화_백업/, persistent: true, ignoreInitial: true })
    .on('all', (event, filePath) => {
        if (path.extname(filePath) === '.xlsx' && path.basename(filePath).includes('사원등록')) {
            console.log(`Watcher: ${event} on ${path.basename(filePath)}`);
            updateAllData();
        }
    });
} catch (e) {}

app.get('/api/data', async (req, res) => {
    console.log('[SERVER] GET /api/data');
    if (!processedResults || Object.keys(processedResults.corporations || {}).length === 0) {
        await updateAllData();
    }
    res.json(processedResults);
});

app.post('/api/refresh', async (req, res) => {
    try {
        const { year, branch } = req.body;
        const backupRoot = path.join(WATCH_DIR, '_초기화_백업'); if (!fs.existsSync(backupRoot)) fs.mkdirSync(backupRoot, { recursive: true });
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').split('Z')[0], backupDir = path.join(backupRoot, timestamp); fs.mkdirSync(backupDir, { recursive: true });
        
        if (year && branch) {
            // Contextual reset: Only one file
            const dir = path.join(WATCH_DIR, year.toString());
            const fileName = `${branch}_사원등록.xlsx`;
            const oldPath = path.join(dir, fileName);
            if (fs.existsSync(oldPath)) {
                const newName = `${year}_${fileName}`;
                const newPath = path.join(backupDir, newName);
                fs.renameSync(oldPath, newPath);
                await syncToDB(newPath);
                await syncToDB(oldPath, true);
                console.log(`[SERVER] Contextual reset: ${year} ${branch}`);
                await logActivity('CLEAR', `[${branch}] ${year}년 데이터 초기화`);
            }
        } else {
            // Full reset: Original logic
            const scanDirs = [WATCH_DIR, ...SUPPORTED_YEARS.map(yr => path.join(WATCH_DIR, yr.toString()))];
            for (const dir of scanDirs) {
                if (!fs.existsSync(dir)) continue;
                const files = fs.readdirSync(dir).filter(f => f.includes('사원등록') && f.endsWith('.xlsx'));
                for (const f of files) {
                    const oldPath = path.join(dir, f), dirName = path.basename(dir), newName = SUPPORTED_YEARS.includes(parseInt(dirName)) ? `${dirName}_${f}` : f, newPath = path.join(backupDir, newName);
                    fs.renameSync(oldPath, newPath);
                    await syncToDB(newPath); await syncToDB(oldPath, true);
                }
            }
            console.log('[SERVER] Full Refresh complete.');
            await logActivity('RESET', '전체 데이터 초기화');
        }
        await updateAllData(); 
        res.json({ success: true, message: 'Data archived', data: processedResults });
    } catch (e) {
        console.error('[SERVER] Refresh error:', e);
        res.status(500).json({ error: 'Reset failed' }); 
    }
});

app.post('/api/upload', upload.single('file'), async (req, res) => {
    const { year, branch } = req.query; 
    console.log(`[SERVER] Upload started: year=${year}, branch=${branch}, file=${req.file?.originalname}`);
    try {
        const finalPath = req.file.path; 
        await syncToDB(finalPath); 
        console.log('[SERVER] DB sync complete, updating data...');
        await updateAllData(); 
        await logActivity('UPLOAD', `[${branch}] ${year}년 업로드: ${decodeText(req.file.originalname)}`);
        console.log('[SERVER] Upload process finished successfully.');
        res.json({ success: true, data: processedResults });
    } catch (e) { 
        console.error('[SERVER] Upload error:', e);
        res.status(500).json({ error: 'Upload failed' }); 
    }
});

app.post('/api/rename-branch', async (req, res) => {
    const { oldName, newName } = req.body;
    try {
        const scanDirs = [WATCH_DIR, ...SUPPORTED_YEARS.map(yr => path.join(WATCH_DIR, yr.toString()))];
        for (const dir of scanDirs) {
            if (!fs.existsSync(dir)) continue;
            const oldFile = path.join(dir, `${oldName}_사원등록.xlsx`), newFile = path.join(dir, `${newName}_사원등록.xlsx`);
            if (fs.existsSync(oldFile)) { fs.renameSync(oldFile, newFile); await syncToDB(newFile); await syncToDB(oldFile, true); }
        }
        loadMapping(); mapping.corporations.forEach(corp => { if (corp.branchNames) { const idx = corp.branchNames.indexOf(oldName); if (idx !== -1) corp.branchNames[idx] = newName; } });
        await saveMapping(); 
        await updateAllData(); 
        await logActivity('RENAME', `${oldName} -> ${newName}`);
        console.log(`[SERVER] Rename success: ${oldName} -> ${newName}`);
        res.json({ success: true, data: processedResults });
    } catch (e) { 
        console.error('[SERVER] Rename error:', e);
        res.status(500).json({ error: 'Rename failed' }); 
    }
});

app.post('/api/delete-branch', async (req, res) => {
    const { branchName } = req.body;
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').split('Z')[0], backupDir = path.join(WATCH_DIR, '_초기화_백업', `deleted_${branchName}_${timestamp}`); fs.mkdirSync(backupDir, { recursive: true });
        const scanDirs = [WATCH_DIR, ...SUPPORTED_YEARS.map(yr => path.join(WATCH_DIR, yr.toString()))];
        for (const dir of scanDirs) {
            if (!fs.existsSync(dir)) continue;
            const branchFile = path.join(dir, `${branchName}_사원등록.xlsx`);
            if (fs.existsSync(branchFile)) { const newName = `${path.basename(dir)}_${branchName}_사원등록.xlsx`, archPath = path.join(backupDir, newName); fs.renameSync(branchFile, archPath); await syncToDB(archPath); await syncToDB(branchFile, true); }
        }
        loadMapping(); mapping.corporations.forEach(corp => { if (corp.branchNames) corp.branchNames = corp.branchNames.filter(bn => bn !== branchName); });
        await saveMapping(); 
        await updateAllData(); 
        await logActivity('DELETE', `지점 삭제: ${branchName}`);
        console.log(`[SERVER] Delete success: ${branchName}`);
        res.json({ success: true, data: processedResults });
    } catch (e) { 
        console.error('[SERVER] Delete error:', e);
        res.status(500).json({ error: 'Delete failed' }); 
    }
});

app.post('/api/clear-branch-data', async (req, res) => {
    const { branchName } = req.body;
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').split('Z')[0], backupDir = path.join(WATCH_DIR, '_초기화_백업', `clear_${branchName}_${timestamp}`); fs.mkdirSync(backupDir, { recursive: true });
        const scanDirs = [WATCH_DIR, ...SUPPORTED_YEARS.map(yr => path.join(WATCH_DIR, yr.toString()))];
        for (const dir of scanDirs) {
            if (!fs.existsSync(dir)) continue;
            const branchFile = path.join(dir, `${branchName}_사원등록.xlsx`);
            if (fs.existsSync(branchFile)) { const archPath = path.join(backupDir, `${path.basename(dir)}_${branchName}_xlsx`); fs.renameSync(branchFile, archPath); await syncToDB(archPath); await syncToDB(branchFile, true); }
        }
        await updateAllData(); 
        await logActivity('CLEAR', `지점 비우기: ${branchName}`);
        console.log(`[SERVER] Clear success: ${branchName}`);
        res.json({ success: true, data: processedResults });
    } catch (e) { 
        console.error('[SERVER] Clear error:', e);
        res.status(500).json({ error: 'Clear failed' }); 
    }
});

app.post('/api/add-year', async (req, res) => {
    const { year } = req.body;
    if (!year || isNaN(year)) return res.status(400).json({ error: 'Invalid year' });
    try {
        const yrDir = path.join(WATCH_DIR, year.toString());
        if (!fs.existsSync(yrDir)) {
            fs.mkdirSync(yrDir, { recursive: true });
            console.log(`[SERVER] Created new year folder: ${year}`);
            await logActivity('CONFIG', `${year}년 연도 추가`);
        }
        await updateAllData();
        res.json({ success: true, data: processedResults });
    } catch (e) {
        console.error('[SERVER] Add year error:', e);
        res.status(500).json({ error: 'Failed to add year' });
    }
});

app.get('/api/mapping', (req, res) => { loadMapping(); res.json(mapping); });
app.post('/api/mapping', async (req, res) => { 
    console.log('[SERVER] POST /api/mapping');
    mapping = req.body; 
    await saveMapping(); 
    await updateAllData(); 
    res.json({ success: true, data: processedResults }); 
});

const clientBuildPath = path.join(__dirname, '..', 'client', 'dist');
if (fs.existsSync(clientBuildPath)) {
    app.use(express.static(clientBuildPath));
    app.get('*', (req, res) => res.sendFile(path.join(clientBuildPath, 'index.html')));
}

const server = app.listen(process.env.PORT || PORT, async () => {
    console.log(`Server running on port ${process.env.PORT || PORT}`);
    await updateAllData();
});
server.on('error', (err) => console.error('SERVER ERROR:', err));
setInterval(() => {}, 10000);
