const express = require('express');
const xlsx = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const chokidar = require('chokidar');

const app = express();
const PORT = process.env.PORT || 5010;
const WATCH_DIR = process.env.DATA_DIR || path.join(__dirname, '..', 'data');

if (!fs.existsSync(WATCH_DIR)) {
    fs.mkdirSync(WATCH_DIR, { recursive: true });
}

// Helper to fix garbled Korean text from multer/busboy (Windows/Latin1 issues)
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
function logActivity(type, msg) {
    recentActivities.unshift({ type, msg, timestamp: new Date().toISOString() });
    recentActivities = recentActivities.slice(0, 50);
    try { fs.writeFileSync(ACTIVITY_FILE, JSON.stringify(recentActivities, null, 2)); } catch (e) {}
}
loadActivity();

function getAvailableYears() {
    try {
        const items = fs.readdirSync(WATCH_DIR);
        const years = items.filter(i => /^\d{4}$/.test(i)).map(i => parseInt(i));
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
        // Use query params for immediate availability during stream processing
        const year = req.query.year || new Date().getFullYear().toString();
        const dest = path.join(WATCH_DIR, year.toString());
        if (!fs.existsSync(dest)) fs.mkdirSync(dest, { recursive: true });
        cb(null, dest);
    },
    filename: (req, file, cb) => {
        const branch = req.query.branch; // Query params are already decoded usually
        const originalName = decodeText(file.originalname);
        console.log(`Upload Storage - Branch: [${branch}], Year: [${req.query.year}], File: [${originalName}]`);
        if (branch) {
            cb(null, `${branch}_사원등록.xlsx`);
        } else {
            cb(null, originalName);
        }
    }
});
const upload = multer({ storage });

let mapping = {
    corporations: [] // [ { id, name, branchNames: [] } ]
};

function loadMapping() {
    if (fs.existsSync(MAPPING_FILE)) {
        try {
            const raw = fs.readFileSync(MAPPING_FILE, 'utf8');
            const parsed = JSON.parse(raw);
            if (Array.isArray(parsed)) {
                mapping = { corporations: parsed };
            } else if (parsed && parsed.corporations) {
                mapping = parsed;
            } else {
                mapping = { corporations: [] };
            }
        } catch (e) {
            console.error('Error parsing mapping.json:', e);
            mapping = { corporations: [] };
        }
    } else {
        mapping = { corporations: [] };
    }
}

function saveMapping() {
    fs.writeFileSync(MAPPING_FILE, JSON.stringify(mapping, null, 2));
}

let processedResults = {
    updatedAt: new Date().toISOString(),
    corporations: {},
    unassigned: {} // { branchName: data }
};

function getBirthDate(rrn) {
    if (!rrn) return null;
    const s = rrn.toString().replace(/-/g, '');
    if (s.length < 7) return null;
    const yy = s.substring(0, 2);
    const mm = s.substring(2, 4);
    const dd = s.substring(4, 6);
    const gender = s.charAt(6);
    let yearPrefix = '19';
    if (gender === '3' || gender === '4' || gender === '7' || gender === '8') yearPrefix = '20';
    return `${yearPrefix}${yy}${mm}${dd}`;
}

function calculateAgeAtJoin(birth, joinDate) {
    if (!birth || !joinDate) return 0;
    const bY = parseInt(birth.substring(0, 4));
    const bM = parseInt(birth.substring(4, 6));
    const bD = parseInt(birth.substring(6, 8));

    const jStr = joinDate.toString();
    const jY = parseInt(jStr.substring(0, 4));
    const jM = parseInt(jStr.substring(4, 6));
    const jD = parseInt(jStr.substring(6, 8));

    let age = jY - bY;
    if (jM < bM || (jM === bM && jD < bD)) age--;
    return age;
}

function parseBranchName(fileName) {
    // Just extract the branch name part before _사원등록
    const match = fileName.match(/(.+)_사원등록/);
    return match ? match[1] : fileName.replace('.xlsx', '');
}

function processFile(filePath) {
    // ... (logic remains largely same, just returns the object)
    try {
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = xlsx.utils.sheet_to_json(sheet);
        
        const validation = [];
        rawData.forEach((emp, idx) => {
            const rowNum = idx + 2; 
            const name = (emp.사원명 || '').toString().trim();
            const rrn = (emp['주민(외국인)등록번호'] || '').toString().replace(/-/g, '').trim();
            const joinDate = (emp.입사일자 || '').toString().trim();
            const retireDate = (emp.퇴사일자 || '').toString().trim();

            if (!name) validation.push({ row: rowNum, type: 'error', field: '사원명', msg: '사원명이 누락되었습니다.' });
            if (!rrn) validation.push({ row: rowNum, type: 'error', field: '주민번호', msg: '주민번호가 누락되었습니다.' });
            if (!joinDate) validation.push({ row: rowNum, type: 'error', field: '입사일자', msg: '입사일자가 누락되었습니다.' });
            
            if (joinDate && retireDate && joinDate > retireDate) {
                validation.push({ row: rowNum, type: 'warning', field: '입퇴사일', msg: `입사일(${joinDate})이 퇴사일(${retireDate})보다 늦습니다.` });
            }
        });

        // Deduplication: Use Name + SSN as key
        const uniqueEmployeesMap = new Map();
        rawData.forEach(emp => {
            const name = (emp.사원명 || '').toString().trim();
            const rrn = (emp['주민(외국인)등록번호'] || '').toString().replace(/-/g, '').trim();
            const key = `${name}|${rrn}`;
            if (key !== '|' && !uniqueEmployeesMap.has(key)) {
                uniqueEmployeesMap.set(key, emp);
            }
        });
        const data = Array.from(uniqueEmployeesMap.values());
        const years = [2022, 2023, 2024, 2025, 2026];
        const fileResult = {
            fileName: path.basename(filePath),
            updatedAt: fs.statSync(filePath).mtime.toISOString(),
            years: {}
        };

        const branchDataMap = {}; // departmentName -> { years: { yr: { data } } }

        years.forEach(year => {
            // Group raw data by department first
            const deptGroups = {};
            data.forEach(emp => {
                let dept = (emp.부서 || '미지정').toString().trim();
                const originalDept = dept;
                // Simple and effective normalization:
                // 1. Remove year: "2024 (주)평우서비스(본점)" -> " (주)평우서비스(본점)"
                // 2. Remove corporate name: " (주)평우서비스(본점)" -> "(본점)"
                dept = dept.replace(/^\d{4}/, '').replace(/.*평우서비스/, '').trim();
                
                if (!dept) dept = originalDept;
                // console.log(`DEBUG: Normalized "${originalDept}" -> "${dept}"`);
                
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
                            const birth = getBirthDate(emp['주민(외국인)등록번호']);
                            const age = calculateAgeAtJoin(birth, emp.입사일자.toString());
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
                    const birth = getBirthDate(emp['주민(외국인)등록번호']);
                    const ageAtJoin = calculateAgeAtJoin(birth, emp.입사일자.toString());
                    const isYouthAtJoin = (ageAtJoin >= 15 && ageAtJoin <= 34);
                    let exclusionDate = '';
                    if (isYouthAtJoin && birth) {
                        const bYear = parseInt(birth.substring(0, 4)) + 35;
                        exclusionDate = `${bYear}-${birth.substring(4, 6)}-${birth.substring(6, 8)}`;
                    }
                    const address = emp.주소 || emp.근무지 || emp.사업장 || '';
                    const isMetro = address.includes('서울') || address.includes('경기');
                    
                    // Retirement highlight logic: must be in the CURRENT processing year
                    const retirementDate = emp.퇴사일자 ? emp.퇴사일자.toString().trim() : '';
                    const retiredThisYear = !!(retirementDate && retirementDate.startsWith(year.toString()));
                    
                    if (emp.사원명 === '김명섭' && (year === 2024 || year === 2025)) {
                        console.log(`DEBUG: Employee ${emp.사원명} (${year}) - Retired Date: [${retirementDate}], retiredThisYear: ${retiredThisYear}`);
                    }

                    return {
                        col14: birth ? `${birth.substring(0, 4)}-${birth.substring(4, 6)}-${birth.substring(6, 8)}` : '',
                        col15: emp.사원명 || '', col16: 'Y', col17: '',
                        col18: emp.단시간유형 || emp.가중치 || '1.0',
                        col19: (emp.내외국인구분 || '').includes('내국인') ? 'Y' : 'N',
                        col20: isMetro ? 'Y' : 'N',
                        col21: emp.입사일자 ? `${emp.입사일자.toString().substring(0, 4)}-${emp.입사일자.toString().substring(4, 6)}-${emp.입사일자.toString().substring(6, 8)}` : '',
                        col22: startM ? `${startM.toString().padStart(2, '0')}~${endM.toString().padStart(2, '0')}` : '',
                        col23: monthsActiveCount, col24: isYouthAtJoin ? monthsActiveCount : 0,
                        col25: isYouthAtJoin ? 'Y' : 'N', col26: isYouthAtJoin ? 'Y' : 'N',
                        col27: exclusionDate, col28: emp.장애인 > 0 ? 'Y' : 'N',
                        col29: (emp.나이 || 0) >= 60 ? 'Y' : 'N', col30: (emp.경력단절 || 0) > 0 ? 'Y' : 'N',
                        col31: (emp.북약 || 0) > 0 ? 'Y' : 'N', col32: 'N', col33: 'N', col34: '', col35: '',
                        _retiredThisYear: retiredThisYear
                    };
                });

                branchDataMap[deptName].years[year] = {
                    monthly, avgTotal: avgT, avgYouth: avgY,
                    avgOther: parseFloat((avgT - avgY).toFixed(2)),
                    sumTotal: parseFloat(sumT.toFixed(2)), sumYouth: parseFloat(sumY.toFixed(2)),
                    summary: {
                        col1: sumT.toFixed(2), col2: '12', col3: avgT.toFixed(2),
                        col5: sumY.toFixed(2), col6: '0', col7: '0', col8: '0', col9: '0', col10: sumY.toFixed(2),
                        col11: '12', col12: avgY.toFixed(2), col13: (avgT - avgY).toFixed(2)
                    },
                    details
                };
            });
        });
        return { branchDataMap, validation }; // Returns data and validation results
    } catch (e) {
        console.error(`Error processing file ${filePath}:`, e);
        return null;
    }
}

async function updateAllData() {
    console.log(`[${new Date().toISOString()}] Data update triggered...`);
    try {
        loadMapping();
        SUPPORTED_YEARS = getAvailableYears();
        console.log(`Supported Years: ${SUPPORTED_YEARS}`);
        
        // Scan all year-specific folders
        const allProcessed = {};
        const branchFiles = {}; // branchName -> { year -> { fullPath, mtime } }

    // Scan both root and yearly folders
    const scanDirs = [WATCH_DIR, ...SUPPORTED_YEARS.map(yr => path.join(WATCH_DIR, yr.toString()))];
    
    scanDirs.forEach(dir => {
        if (!fs.existsSync(dir)) return;
        const dirYear = dir === WATCH_DIR ? null : parseInt(path.basename(dir));
        
        fs.readdirSync(dir)
            .filter(f => f.includes('사원등록') && f.endsWith('.xlsx'))
            .forEach(f => {
                const fullPath = path.join(dir, f);
                const stats = fs.statSync(fullPath);
                
                // Determine year: from folder name or filename prefix
                let yr = dirYear;
                if (!yr) {
                    const match = f.match(/^(\d{4})/);
                    if (match) yr = parseInt(match[1]);
                }
                
                if (yr && SUPPORTED_YEARS.includes(yr)) {
                    const branch = parseBranchName(f);
                    if (!branchFiles[branch]) branchFiles[branch] = {};
                    if (!branchFiles[branch][yr] || branchFiles[branch][yr].mtime < stats.mtime.getTime()) {
                        branchFiles[branch][yr] = { name: f, fullPath, branch, mtime: stats.mtime.getTime(), year: yr };
                    }
                }
            });
    });

    const validationReports = {};
    // Process each branch by merging its yearly data
    Object.keys(branchFiles).forEach(branchKey => {
        const yearsData = branchFiles[branchKey];
        
        Object.keys(yearsData).forEach(yr => {
            const result = processFile(yearsData[yr].fullPath);
            if (result) {
                const { branchDataMap, validation } = result;
                if (validation && validation.length > 0) {
                    const fileName = yearsData[yr].name;
                    if (!validationReports[fileName]) validationReports[fileName] = [];
                    validationReports[fileName].push(...validation);
                }
                if (branchDataMap) {
                    Object.keys(branchDataMap).forEach(deptName => {
                        const deptResult = branchDataMap[deptName];
                        if (!allProcessed[deptName]) {
                            allProcessed[deptName] = {
                                branch: deptName,
                                updatedAt: new Date().toISOString(),
                                years: {}
                            };
                            [2022, 2023, 2024, 2025, 2026].forEach(y => {
                                allProcessed[deptName].years[y] = { monthly: Array(12).fill({total:0, youth:0}), avgTotal:0, avgYouth:0, avgOther:0, sumTotal:0, sumYouth:0, summary:{}, details:[] };
                            });
                        }
                        if (deptResult.years[yr]) {
                            allProcessed[deptName].years[yr] = deptResult.years[yr];
                        }
                    });
                }
            }
        });
    });

    // Grouping by mapping
    const newCorpData = {};
    const assignedBranches = new Set();

    const normalizeForMatch = (name) => {
        return name.replace(/^\d{4}/, '').replace(/.*평우서비스/, '').trim();
    };

    mapping.corporations.forEach(corp => {
        // Ensure every corporation exists in newCorpData, even if it has no branches
        newCorpData[corp.name] = { branches: {}, total: null };
        (corp.branchNames || []).forEach(bn => {
            const normalizedBN = normalizeForMatch(bn);
            if (allProcessed[normalizedBN]) {
                newCorpData[corp.name].branches[normalizedBN] = allProcessed[normalizedBN];
                assignedBranches.add(normalizedBN);
            } else if (allProcessed[bn]) {
                // Fallback to exact match if normalized didn't work
                newCorpData[corp.name].branches[bn] = allProcessed[bn];
                assignedBranches.add(bn);
            }
        });

        const brs = Object.values(newCorpData[corp.name].branches);
        if (brs.length > 0) {
            const consolidated = { years: {} };
            [2022, 2023, 2024, 2025, 2026].forEach(yr => {
                let tAvgT = 0, tAvgY = 0, tSumT = 0, tSumY = 0;
                const combMonthly = Array.from({ length: 12 }, () => ({ total: 0, youth: 0 }));
                const combDetails = [];
                brs.forEach(b => {
                    const yD = b.years[yr];
                    if (!yD) return;
                    tAvgT += yD.avgTotal || 0; 
                    tAvgY += yD.avgYouth || 0; 
                    tSumT += yD.sumTotal || 0; 
                    tSumY += yD.sumYouth || 0;
                    if (yD.monthly) {
                        yD.monthly.forEach((m, i) => { 
                            if (combMonthly[i]) {
                                combMonthly[i].total += m.total || 0; 
                                combMonthly[i].youth += m.youth || 0; 
                            }
                        });
                    }
                    if (yD.details) combDetails.push(...yD.details);
                });
                consolidated.years[yr] = {
                    monthly: combMonthly, 
                    avgTotal: parseFloat(tAvgT.toFixed(2)), 
                    avgYouth: parseFloat(tAvgY.toFixed(2)),
                    avgOther: parseFloat((tAvgT - tAvgY).toFixed(2)), 
                    sumTotal: parseFloat(tSumT.toFixed(2)), 
                    sumYouth: parseFloat(tSumY.toFixed(2)),
                    summary: {
                        col1: tSumT.toFixed(2), col2: '12', col3: tAvgT.toFixed(2), 
                        col5: tSumY.toFixed(2), col6: '0', col7: '0', col8: '0', 
                        col9: '0', col10: tSumY.toFixed(2), col11: '12', 
                        col12: tAvgY.toFixed(2), col13: (tAvgT - tAvgY).toFixed(2)
                    },
                    details: combDetails
                };
            });
            newCorpData[corp.name].total = consolidated;
        }
    });

    const unassigned = {};
    Object.keys(allProcessed).forEach(bn => {
        if (!assignedBranches.has(bn)) unassigned[bn] = allProcessed[bn];
    });

        processedResults = {
            updatedAt: new Date().toISOString(),
            corporations: newCorpData,
            unassigned,
            mapping: mapping.corporations,
            validation: validationReports,
            years: SUPPORTED_YEARS,
            activities: recentActivities
        };
        console.log('Processed results updated with mapping.');
    } catch (err) {
        console.error('CRITICAL ERROR in updateAllData:', err);
    }
}

// File Watcher for auto-refresh
try {
    chokidar.watch(WATCH_DIR, { 
        ignored: /(^|[\/\\])\..|node_modules|_초기화_백업/,
        persistent: true,
        ignoreInitial: true 
    }).on('all', (event, filePath) => {
        if (path.extname(filePath) === '.xlsx' && path.basename(filePath).includes('사원등록')) {
            console.log(`Watcher: File change detected [${event}] on ${path.basename(filePath)}`);
            updateAllData();
        }
    });
} catch (e) {
    console.warn('Watcher could not be initialized (might be non-local or permission issue):', e.message);
}

app.get('/api/data', (req, res) => {
    res.json(processedResults);
});

app.post('/api/refresh', (req, res) => {
    try {
        const backupRoot = path.join(WATCH_DIR, '_초기화_백업');
        if (!fs.existsSync(backupRoot)) fs.mkdirSync(backupRoot, { recursive: true });

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').split('Z')[0];
        const backupDir = path.join(backupRoot, timestamp);
        fs.mkdirSync(backupDir, { recursive: true });

        const scanDirs = [WATCH_DIR, ...SUPPORTED_YEARS.map(yr => path.join(WATCH_DIR, yr.toString()))];
        let movedCount = 0;

        scanDirs.forEach(dir => {
            if (fs.existsSync(dir)) {
                const files = fs.readdirSync(dir);
                files.forEach(f => {
                    if (f.includes('사원등록') && f.endsWith('.xlsx')) {
                        const oldPath = path.join(dir, f);
                        // Add year prefix to backup filename if it's from a subfolder
                        const dirName = path.basename(dir);
                        const newName = SUPPORTED_YEARS.includes(parseInt(dirName)) ? `${dirName}_${f}` : f;
                        const newPath = path.join(backupDir, newName);
                        
                        fs.renameSync(oldPath, newPath);
                        movedCount++;
                    }
                });
            }
        });

        console.log(`Reset triggered: ${movedCount} files moved to backup folder [${timestamp}].`);
        updateAllData();
        res.json({ success: true, message: 'Data initialized (Archived)', data: processedResults });
    } catch (e) {
        console.error('Reset failed:', e);
        res.status(500).json({ error: 'Failed to reset data and backup files' });
    }
});

app.post('/api/upload', upload.single('file'), (req, res) => {
    const year = req.query.year || req.body.year;
    const branch = req.query.branch || req.body.branch;
    
    console.log(`API Upload Hit - Year: ${year}, Branch: ${branch}`);

    if (!year) {
        return res.status(400).json({ error: 'Year is required in query or body' });
    }
    // If branch is 'auto' or undefined, we skip the match check but still process.
    const isAuto = !branch || branch === 'auto' || branch === 'undefined';

    try {
        // Basic validation: Check if '부서' in Excel matches 'branch'
        const workbook = xlsx.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet);
        
        let matchFound = false;
        const deptsInFile = new Set();
        data.forEach(emp => {
            let dept = (emp.부서 || '미지정').toString().trim();
            // Same normalization as in processFile
            dept = dept.replace(/^\d{4}/, '').replace(/.*평우서비스/, '').trim();
            deptsInFile.add(dept);
            if (isAuto || dept === branch) matchFound = true;
        });

        updateAllData();

        logActivity('UPLOAD', `[${branch}] 지점 ${year}년 자료 업로드: ${decodeText(req.file.originalname)}`);
        const responseData = { 
            success: true, 
            data: processedResults,
            uploadedFile: decodeText(req.file.originalname),
            targetBranch: decodeText(branch),
            matchFound,
            deptsInFile: Array.from(deptsInFile)
        };
        console.log(`Upload successful - Returning: ${responseData.uploadedFile} for ${responseData.targetBranch}`);
        res.json(responseData);
    } catch (e) {
        console.error('Upload processing failed:', e);
        res.status(500).json({ error: 'Failed to process uploaded file' });
    }
});

app.post('/api/rename-branch', (req, res) => {
    const { oldName, newName } = req.body;
    if (!oldName || !newName) return res.status(400).json({ error: 'Old and New names are required' });

    try {
        const scanDirs = [WATCH_DIR, ...SUPPORTED_YEARS.map(yr => path.join(WATCH_DIR, yr.toString()))];
        let renamedCount = 0;

        scanDirs.forEach(dir => {
            if (fs.existsSync(dir)) {
                const oldFile = path.join(dir, `${oldName}_사원등록.xlsx`);
                const newFile = path.join(dir, `${newName}_사원등록.xlsx`);
                if (fs.existsSync(oldFile)) {
                    fs.renameSync(oldFile, newFile);
                    renamedCount++;
                }
            }
        });

        // Update mapping.json
        loadMapping();
        mapping.corporations.forEach(corp => {
            if (corp.branchNames) {
                const idx = corp.branchNames.indexOf(oldName);
                if (idx !== -1) corp.branchNames[idx] = newName;
            }
        });
        saveMapping();
        updateAllData();

        logActivity('RENAME', `지점명 변경: [${oldName}] -> [${newName}]`);
        console.log(`Rename branch: [${oldName}] -> [${newName}] (${renamedCount} files renamed)`);
        res.json({ success: true, data: processedResults });
    } catch (e) {
        console.error('Rename failed:', e);
        res.status(500).json({ error: 'Failed to rename branch' });
    }
});

app.post('/api/delete-branch', (req, res) => {
    const { branchName } = req.body;
    if (!branchName) return res.status(400).json({ error: 'Branch name is required' });

    try {
        const backupRoot = path.join(WATCH_DIR, '_초기화_백업');
        if (!fs.existsSync(backupRoot)) fs.mkdirSync(backupRoot, { recursive: true });

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').split('Z')[0];
        const backupDir = path.join(backupRoot, `deleted_${branchName}_${timestamp}`);
        fs.mkdirSync(backupDir, { recursive: true });

        const scanDirs = [WATCH_DIR, ...SUPPORTED_YEARS.map(yr => path.join(WATCH_DIR, yr.toString()))];
        let movedCount = 0;

        scanDirs.forEach(dir => {
            if (fs.existsSync(dir)) {
                const branchFile = path.join(dir, `${branchName}_사원등록.xlsx`);
                if (fs.existsSync(branchFile)) {
                    const dirName = path.basename(dir);
                    const newName = SUPPORTED_YEARS.includes(parseInt(dirName)) ? `${dirName}_${branchName}_사원등록.xlsx` : `${branchName}_사원등록.xlsx`;
                    fs.renameSync(branchFile, path.join(backupDir, newName));
                    movedCount++;
                }
            }
        });

        // Update mapping.json
        loadMapping();
        mapping.corporations.forEach(corp => {
            if (corp.branchNames) {
                corp.branchNames = corp.branchNames.filter(bn => bn !== branchName);
            }
        });
        saveMapping();
        updateAllData();

        logActivity('DELETE', `지점 삭제: [${branchName}]`);
        console.log(`Delete branch: [${branchName}] (${movedCount} files archived)`);
        res.json({ success: true, data: processedResults });
    } catch (e) {
        console.error('Delete failed:', e);
        res.status(500).json({ error: 'Failed to delete branch' });
    }
});

app.post('/api/clear-branch-data', (req, res) => {
    const { branchName } = req.body;
    if (!branchName) return res.status(400).json({ error: 'Branch name is required' });

    try {
        const backupRoot = path.join(WATCH_DIR, '_초기화_백업');
        if (!fs.existsSync(backupRoot)) fs.mkdirSync(backupRoot, { recursive: true });

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').split('Z')[0];
        const backupDir = path.join(backupRoot, `clear_${branchName}_${timestamp}`);
        fs.mkdirSync(backupDir, { recursive: true });

        const scanDirs = [WATCH_DIR, ...SUPPORTED_YEARS.map(yr => path.join(WATCH_DIR, yr.toString()))];
        let movedCount = 0;

        scanDirs.forEach(dir => {
            if (fs.existsSync(dir)) {
                const branchFile = path.join(dir, `${branchName}_사원등록.xlsx`);
                if (fs.existsSync(branchFile)) {
                    const dirName = path.basename(dir);
                    const newName = SUPPORTED_YEARS.includes(parseInt(dirName)) ? `${dirName}_${branchName}_사원등록.xlsx` : `${branchName}_사원등록.xlsx`;
                    fs.renameSync(branchFile, path.join(backupDir, newName));
                    movedCount++;
                }
            }
        });

        updateAllData();

        logActivity('CLEAR', `지점 데이터 비우기: [${branchName}]`);
        console.log(`Clear branch data: [${branchName}] (${movedCount} files archived)`);
        res.json({ success: true, data: processedResults });
    } catch (e) {
        console.error('Clear failed:', e);
        res.status(500).json({ error: 'Failed to clear branch data' });
    }
});

app.get('/api/mapping', (req, res) => {
    loadMapping();
    res.json(mapping);
});

app.post('/api/mapping', (req, res) => {
    mapping = req.body;
    saveMapping();
    updateAllData();
    res.json({ success: true, data: processedResults });
});

// Serve static files from the React app
const clientBuildPath = path.join(__dirname, '..', 'client', 'dist');
if (fs.existsSync(clientBuildPath)) {
    app.use(express.static(clientBuildPath));
    app.get('*', (req, res) => {
        res.sendFile(path.join(clientBuildPath, 'index.html'));
    });
}

const server = app.listen(process.env.PORT || PORT, () => {
    console.log(`Server running on port ${process.env.PORT || PORT}`);
    console.log(`Data directory: ${WATCH_DIR}`);
    updateAllData();
});

server.on('error', (err) => {
    console.error('SERVER ERROR:', err);
});

// Keep-alive handle
setInterval(() => {
    // Just keeping the event loop busy
}, 10000);
