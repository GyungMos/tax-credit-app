import React, { useState, useEffect } from 'react';
import { RefreshCw, Download, FileText, ArrowUpRight, ArrowDownRight, Upload, FileSpreadsheet, ArrowUp, ArrowDown, Moon, Sun, Edit2, Trash2, RotateCcw, AlertCircle, AlertTriangle } from 'lucide-react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';

const getBirthDate = (rrn) => {
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
};

const calculateAgeAtJoin = (birth, joinDate) => {
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
};

const App = () => {
  const [data, setData] = useState({ corporations: {}, unassigned: {} });
  const [loading, setLoading] = useState(true);
  const [year, setYear] = useState(2025);
  const [showPicker, setShowPicker] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [selectedCorp, setSelectedCorp] = useState(null);
  const [selectedBranch, setSelectedBranch] = useState(null);
  const [mapping, setMapping] = useState({ corporations: [] });
  const [isManaging, setIsManaging] = useState(false);
  const [dragOverCorpId, setDragOverCorpId] = useState(null);
  const [theme, setTheme] = useState(localStorage.getItem('theme') || 'dark-theme');
  const [uploadYear, setUploadYear] = useState(2025);
  const [validationReports, setValidationReports] = useState({});
  const [showValidation, setShowValidation] = useState(null); // { fileName, items }

  const parseBranchName = (fileName) => {
    const match = fileName.match(/(.+)_사원등록/);
    return match ? match[1] : fileName.replace('.xlsx', '');
  };

  const getValidationForBranch = (branchName) => {
    const allItems = [];
    Object.keys(validationReports || {}).forEach(fileName => {
      if (parseBranchName(fileName) === branchName) {
        (validationReports[fileName] || []).forEach(item => {
          allItems.push({ ...item, fileName });
        });
      }
    });
    return allItems;
  };

  const processWorkbook = async (files) => {
    setLoading(true);
    const newUnassigned = { ...data.unassigned };
    const newCorporations = JSON.parse(JSON.stringify(data.corporations));
    
    for (const file of files) {
      const workbook = await new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(XLSX.read(e.target.result, { type: 'binary' }));
        reader.readAsBinaryString(file);
      });

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(sheet);
      const branchName = parseBranchName(file.name);
      
      const yearsSet = [2024, 2025, 2026];
      const branchProcessed = { fileName: file.name, years: {} };

      yearsSet.forEach(yr => {
        const monthly = [];
        for (let m = 1; m <= 12; m++) {
          const lastDay = new Date(yr, m, 0).toISOString().split('T')[0].replace(/-/g, '');
          let tW = 0, yW = 0;
          rawData.forEach(emp => {
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
        const sumT = monthly.reduce((s, h) => s + h.total, 0);
        const sumY = monthly.reduce((s, h) => s + h.youth, 0);
        const avgT = parseFloat((sumT / 12).toFixed(2));
        const avgY = parseFloat((sumY / 12).toFixed(2));
        const details = rawData.filter(emp => emp.입사일자 <= `${yr}1231` && (emp.퇴사일자 || '99991231') >= `${yr}0101`).map(emp => {
            let monthsActiveCount = 0, startM = null, endM = null;
            for (let m = 1; m <= 12; m++) {
              const lDay = new Date(yr, m, 0).toISOString().split('T')[0].replace(/-/g, '');
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
            };
        });

        branchProcessed.years[yr] = {
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

      // Find if this branch is mapped to any corp
      let assigned = false;
      for (const corpName in newCorporations) {
        const corpMapping = mapping.corporations.find(c => c.name === corpName);
        if (corpMapping && corpMapping.branchNames.includes(branchName)) {
           newCorporations[corpName].branches[branchName] = branchProcessed;
           assigned = true;
           break;
        }
      }
      if (!assigned) {
        newUnassigned[branchName] = branchProcessed;
      }
    }

    // Recalculate Totals
    Object.keys(newCorporations).forEach(corpName => {
      const c = newCorporations[corpName];
      const brs = Object.values(c.branches);
      if (brs.length === 0) return;
      const consolidated = { years: {} };
      [2024, 2025, 2026].forEach(yr => {
        let tAvgT = 0, tAvgY = 0, tSumT = 0, tSumY = 0;
        const combMonthly = Array.from({ length: 12 }, () => ({ total: 0, youth: 0 }));
        const combDetails = [];
        brs.forEach(b => {
          const yData = b.years[yr];
          tAvgT += yData.avgTotal; tAvgY += yData.avgYouth;
          tSumT += yData.sumTotal; tSumY += yData.sumYouth;
          yData.monthly.forEach((m, idx) => {
            combMonthly[idx].total += m.total; combMonthly[idx].youth += m.youth;
          });
          combDetails.push(...yData.details);
        });
        consolidated.years[yr] = {
          monthly: combMonthly, avgTotal: tAvgT, avgYouth: tAvgY,
          avgOther: parseFloat((tAvgT - tAvgY).toFixed(2)),
          sumTotal: parseFloat(tSumT.toFixed(2)), sumYouth: parseFloat(tSumY.toFixed(2)),
          summary: {
            col1: tSumT.toFixed(2), col2: '12', col3: tAvgT.toFixed(2),
            col5: tSumY.toFixed(2), col6: '0', col7: '0', col8: '0', col9: '0', col10: tSumY.toFixed(2),
            col11: '12', col12: tAvgY.toFixed(2), col13: (tAvgT - tAvgY).toFixed(2)
          },
          details: combDetails
        };
      });
      c.total = consolidated;
    });

    setData({ ...data, corporations: newCorporations, unassigned: newUnassigned });
    setLoading(false);
  };

  const handleTargetedUpload = async (file) => {
    const isSystemEmpty = !data.corporations || Object.keys(data.corporations).length === 0;
    if (!selectedBranch && !selectedCorp && !isSystemEmpty) {
        alert('먼저 지점을 선택해주세요.');
        return;
    }
    setLoading(true);
    const formData = new FormData();
    formData.append('file', file);
    formData.append('year', year);
    if (selectedBranch) formData.append('branch', selectedBranch);
    else if (selectedCorp) formData.append('branch', selectedCorp);
    else formData.append('branch', 'auto');
    const targetBranch = selectedBranch || selectedCorp || 'auto';
    try {
      const res = await fetch(`/api/upload?year=${year}&branch=${encodeURIComponent(targetBranch)}`, {
        method: 'POST',
        body: formData
      });
      if (res.ok) {
        const result = await res.json();
        const serverData = result.data;
        setData({ 
          ...serverData, 
          corporations: serverData.corporations || {}, 
          unassigned: serverData.unassigned || {} 
        });
        setMapping({ corporations: serverData.mapping || [] });
        setValidationReports(serverData.validation || {});
        
        const { uploadedFile, matchFound, deptsInFile } = result;
        if (matchFound) {
          alert(`성공: 파일 [${uploadedFile}]이(가) [${targetBranch}] 지점으로 정상 저장되었습니다.`);
        } else {
          alert(`경고: 파일 [${uploadedFile}]을(를) [${targetBranch}]에 저장했으나, 파일 내부에는 해당 지점 데이터가 없는 것으로 보입니다.\n(파일 내 부서: ${deptsInFile.join(', ') || '없음'})`);
        }
      } else {
        const error = await res.json();
        alert(`업로드 실패: ${error.error || '알 수 없는 오류'}`);
      }
    } catch (e) {
      console.error('Upload failed:', e);
      alert('업로드 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  const saveMappingToServer = async (newMapping) => {
    try {
      const res = await fetch('/api/mapping', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(newMapping)
      });
      if (res.ok) {
        const result = await res.json();
        const serverData = result.data;
        const corps = serverData.corporations || {};
        const unassigned = serverData.unassigned || {};
        setData({ ...serverData, corporations: corps, unassigned });
        setMapping({ corporations: serverData.mapping || [] });
        setValidationReports(serverData.validation || {});
        if (Object.keys(corps).length > 0 && !selectedCorp) {
          setSelectedCorp(Object.keys(corps)[0]);
        } else if (Object.keys(unassigned).length > 0 && !selectedBranch && !selectedCorp) {
          setSelectedBranch(Object.keys(unassigned)[0]);
        }
      }
    } catch (e) {
      console.error('Error saving mapping:', e);
    }
  };

  const addCorporation = () => {
    const name = prompt('법인명을 입력하세요:');
    if (name) {
      const newMapping = {
        ...mapping,
        corporations: [...mapping.corporations, { id: Date.now(), name, branchNames: [] }]
      };
      setMapping(newMapping);
      saveMappingToServer(newMapping);
    }
  };

  const deleteCorporation = (id) => {
    if (confirm('이 법인을 삭제하시겠습니까? (연결된 지점은 미분류로 이동합니다)')) {
      const newMapping = {
        ...mapping,
        corporations: mapping.corporations.filter(c => c.id !== id)
      };
      setMapping(newMapping);
      saveMappingToServer(newMapping);
    }
  };

  const renameCorporation = (id, currentName) => {
    const newName = prompt('법인명을 수정하세요:', currentName);
    if (newName && newName !== currentName) {
      const newMapping = {
        ...mapping,
        corporations: mapping.corporations.map(c => c.id === id ? { ...c, name: newName } : c)
      };
      setMapping(newMapping);
      saveMappingToServer(newMapping);
    }
  };

  const renameBranch = async (oldName) => {
    const newName = prompt(`[${oldName}] 지점의 이름을 수정하세요:`, oldName);
    if (newName && newName !== oldName) {
      setLoading(true);
      try {
        const res = await fetch('/api/rename-branch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ oldName, newName })
        });
        if (res.ok) {
          const result = await res.json();
          const serverData = result.data;
          setData({ ...serverData, corporations: serverData.corporations || {}, unassigned: serverData.unassigned || {} });
          setMapping({ corporations: serverData.mapping || [] });
          if (selectedBranch === oldName) setSelectedBranch(newName);
        } else {
          alert('지점 이름 수정에 실패했습니다.');
        }
      } catch (e) {
        console.error(e);
        alert('오류가 발생했습니다.');
      } finally {
        setLoading(false);
      }
    }
  };

  const deleteBranch = async (branchName) => {
    if (confirm(`[${branchName}] 지점과 관련된 모든 데이터를 삭제(백업폴더로 이동)하시겠습니까?`)) {
      setLoading(true);
      try {
        const res = await fetch('/api/delete-branch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ branchName })
        });
        if (res.ok) {
          const result = await res.json();
          const serverData = result.data;
          setData({ ...serverData, corporations: serverData.corporations || {}, unassigned: serverData.unassigned || {} });
          setMapping({ corporations: serverData.mapping || [] });
          if (selectedBranch === branchName) setSelectedBranch(null);
        } else {
          alert('지점 삭제에 실패했습니다.');
        }
      } catch (e) {
        console.error(e);
        alert('오류가 발생했습니다.');
      } finally {
        setLoading(false);
      }
    }
  };

  const clearBranchData = async (branchName) => {
    if (confirm(`[${branchName}] 지점의 데이터를 화면에서 비우시겠습니까?\n(파일은 백업 폴더로 이동됩니다. 새 파일을 바로 업로드할 수 있습니다.)`)) {
      setLoading(true);
      try {
        const res = await fetch('/api/clear-branch-data', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ branchName })
        });
        if (res.ok) {
          const result = await res.json();
          const serverData = result.data;
          setData({ ...serverData, corporations: serverData.corporations || {}, unassigned: serverData.unassigned || {} });
          setMapping({ corporations: serverData.mapping || [] });
          alert(`[${branchName}] 지점 데이터를 비웠습니다. 이제 새 파일을 업로드하세요.`);
        } else {
          alert('데이터 비우기에 실패했습니다.');
        }
      } catch (e) {
        console.error(e);
        alert('오류가 발생했습니다.');
      } finally {
        setLoading(false);
      }
    }
  };

  const moveCorp = (index, direction) => {
    const newCorps = [...mapping.corporations];
    const targetIndex = index + direction;
    if (targetIndex < 0 || targetIndex >= newCorps.length) return;
    [newCorps[index], newCorps[targetIndex]] = [newCorps[targetIndex], newCorps[index]];
    const newMapping = { ...mapping, corporations: newCorps };
    setMapping(newMapping);
    saveMappingToServer(newMapping);
  };

  const moveBranch = (corpId, branchIndex, direction) => {
    const newMapping = {
      ...mapping,
      corporations: mapping.corporations.map(c => {
        if (c.id === corpId) {
          const newBranches = [...(c.branchNames || [])];
          const targetIndex = branchIndex + direction;
          if (targetIndex < 0 || targetIndex >= newBranches.length) return c;
          [newBranches[branchIndex], newBranches[targetIndex]] = [newBranches[targetIndex], newBranches[branchIndex]];
          return { ...c, branchNames: newBranches };
        }
        return c;
      })
    };
    setMapping(newMapping);
    saveMappingToServer(newMapping);
  };

  const onBranchDragStart = (e, branchName) => {
    e.dataTransfer.setData('branchName', branchName);
    e.dataTransfer.effectAllowed = 'move';
  };

  const onBranchDragEnd = () => {
    setDragOverCorpId(null);
  };

  const onCorpDrop = (e, corpId) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverCorpId(null);
    const branchName = e.dataTransfer.getData('branchName');
    if (!branchName) return;

    const newMapping = {
      ...mapping,
      corporations: mapping.corporations.map(c => {
        const filtered = (c.branchNames || []).filter(bn => bn !== branchName);
        if (c.id === corpId || String(c.id) === String(corpId)) {
          return { ...c, branchNames: [...filtered, branchName] };
        }
        return { ...c, branchNames: filtered };
      })
    };

    setMapping(newMapping);
    saveMappingToServer(newMapping);
  };

  const onUnassignedDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOverCorpId('unassigned');
    const branchName = e.dataTransfer.getData('branchName');
    setDragOverCorpId(null);
    if (!branchName) return;

    const newMapping = {
      ...mapping,
      corporations: mapping.corporations.map(c => ({
        ...c,
        branchNames: (c.branchNames || []).filter(bn => bn !== branchName)
      }))
    };

    setMapping(newMapping);
    saveMappingToServer(newMapping);
  };

  const fetchData = async () => {
    setLoading(true);
    try {
      const res = await fetch('/api/data');
      if (res.ok) {
        const json = await res.json();
        const corps = json.corporations || {};
        const unassigned = json.unassigned || {};
        setData({ ...json, corporations: corps, unassigned });
        setMapping({ corporations: json.mapping || [] });
        if (Object.keys(corps).length > 0 && !selectedCorp) {
          setSelectedCorp(Object.keys(corps)[0]);
        } else if (Object.keys(unassigned).length > 0 && !selectedBranch && !selectedCorp) {
          setSelectedBranch(Object.keys(unassigned)[0]);
        }
      }
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  useEffect(() => {
    const poll = setInterval(async () => {
      try {
        const res = await fetch('/api/data');
        if (res.ok) {
          const result = await res.json();
          // Detect changes by comparing updatedAt timestamp
          if (result.updatedAt && result.updatedAt !== data.updatedAt) {
            console.log("Auto-refreshing UI...");
            setData(prev => ({ 
              ...result, 
              corporations: result.corporations || {}, 
              unassigned: result.unassigned || {} 
            }));
            setMapping({ corporations: result.mapping || [] });
            setValidationReports(result.validation || {});
          }
        }
      } catch (e) {}
    }, 5000);
    return () => clearInterval(poll);
  }, [data.updatedAt]);

  useEffect(() => {
    document.body.className = theme;
    localStorage.setItem('theme', theme);
  }, [theme]);

  const toggleTheme = () => {
    setTheme(prev => prev === 'dark-theme' ? 'light-theme' : 'dark-theme');
  };

  const handleRefresh = async () => {
    if (!confirm('모든 데이터를 초기화하고 저장된 파일을 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.')) {
      return;
    }
    setLoading(true);
    try {
      await fetch('/api/refresh', { method: 'POST' });
      await fetchData();
      alert('데이터가 초기화되었습니다.');
    } catch (e) {
      console.error(e);
      alert('초기화 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  // Column definitions for Table 2 (Employee Details)
  const allColumns = [
    { id: 'col14', label: '⑭ 생년월일 (YYYY MM DD)', default: false },
    { id: 'col15', label: '⑮ 성명', default: true },
    { id: 'col16', label: '⑯ 상시 근로자 (N/Y)', default: true },
    { id: 'col17', label: '⑰ 상시 근로자 제외사유', default: false },
    { id: 'col18', label: '⑱ 단시간 근로자 유형 (0.5/0.75)', default: false },
    { id: 'col19', label: '⑲ 내국인 (N/Y)', default: false },
    { id: 'col20', label: '⑳ 수도권 근무지 (N/Y)', default: false },
    { id: 'col21', label: '㉑ 근로계약 체결일 (YYYY MM DD)', default: true },
    { id: 'col22', label: '㉒ 근로기간 (MM-MM)', default: false },
    { id: 'col23', label: '㉓ 상시 근로자 근무 개월수', default: true },
    { id: 'col24', label: '㉔ 청년등 상시 근로자 근무 개월수', default: false },
    { id: 'col25', label: '㉕ 과세연도 기준 청년여부 (N/Y)', default: true },
    { id: 'col26', label: '㉖ 근로계약 체결일 기준 청년여부 (N/Y)', default: false },
    { id: 'col27', label: '㉗ 청년 제외시점 (YYYY MM DD)', default: false },
    { id: 'col28', label: '㉘ 장애인 (N/Y)', default: false },
    { id: 'col29', label: '㉙ 고령자 (N/Y)', default: false },
    { id: 'col30', label: '㉚ 경력 단절 근로자 (N/Y)', default: false },
    { id: 'col31', label: '㉛ 북한 이탈 주민 (N/Y)', default: false },
    { id: 'col32', label: '㉜ 정규직 전환자 (N/Y)', default: false },
    { id: 'col33', label: '㉝ 육아 휴직 복귀자 (N/Y)', default: false },
    { id: 'col34', label: '㉞ 기타', default: false },
    { id: 'col35', label: '㉟ 비고', default: false },
  ];

  // Column definitions for Table 1 (Summary)
  const allSummaryColumns = [
    { id: 'col1', label: '① 해당(직전) 과세연도의 상시근로자 근무개월수의 합계', default: true },
    { id: 'col2', label: '② 과세연도 개월수', default: true },
    { id: 'col3', label: '③ 상시근로자수 (=①÷②)', default: true },
    { id: 'col5', label: '⑤ 청년', default: true },
    { id: 'col6', label: '⑥ 장애인', default: false },
    { id: 'col7', label: '⑦ 고령자', default: false },
    { id: 'col8', label: '⑧ 경력 단절 근로자', default: false },
    { id: 'col9', label: '⑨ 북한 이탈 주민', default: false },
    { id: 'col10', label: '⑩ 합계', default: true },
    { id: 'col11', label: '⑪ 과세연도 개월수', default: true },
    { id: 'col12', label: '⑫ 청년등상시 근로자수 (=⑩÷⑪)', default: true },
    { id: 'col13', label: '⑬ 청년등외 상시 근로자수 (=③-⑫)', default: true },
  ];

  const [visibleColumns, setVisibleColumns] = useState(
    allColumns.filter(c => c.default).map(c => c.id)
  );

  const [visibleSummaryColumns, setVisibleSummaryColumns] = useState(
    allSummaryColumns.filter(c => c.default).map(c => c.id)
  );

  const activeResult = data;
  const currentCorp = activeResult?.corporations?.[selectedCorp];
  const currentCorpName = selectedCorp || (selectedBranch ? "지정되지 않음" : "");
  
  // Prevent flickering by disabling pointer events on children during drag
  const dragStyle = isDragging ? { pointerEvents: 'none' } : {};
  const currentBranchData = selectedBranch 
    ? (currentCorp?.branches?.[selectedBranch] || activeResult?.unassigned?.[selectedBranch])
    : (currentCorp?.total);
  
  const currentYearData = currentBranchData?.years?.[year] || { avgTotal: 0, avgYouth: 0, avgOther: 0, sumTotal: 0, details: [], summary: {} };
  const prevYearData = currentBranchData?.years?.[year - 1] || { avgTotal: 0, avgYouth: 0, avgOther: 0, sumTotal: 0, summary: {} };

  const exportFullExcel = () => {
    if (!currentYearData) return;
    const wb = XLSX.utils.book_new();
    const summaryHeader = allSummaryColumns.map(c => c.label);
    const summaryData = [
      allSummaryColumns.map(c => currentYearData.summary?.[c.id] || ''),
      allSummaryColumns.map(c => prevYearData.summary?.[c.id] || '')
    ];
    const ws1 = XLSX.utils.aoa_to_sheet([['구분', ...summaryHeader], ['해당과세연도', ...summaryData[0]], ['직전과세연도', ...summaryData[1]]]);
    XLSX.utils.book_append_sheet(wb, ws1, '1.상시근로자수 계산');
    const detailHeader = allColumns.map(c => c.label);
    const detailRows = currentYearData.details.map(d => allColumns.map(c => d[c.id] || ''));
    const ws2 = XLSX.utils.aoa_to_sheet([detailHeader, ...detailRows]);
    XLSX.utils.book_append_sheet(wb, ws2, '2.상시근로자별 명세');
    XLSX.writeFile(wb, `통합고용세액공제_${selectedCorp}_${selectedBranch || '통합'}_${year}.xlsx`);
  };

  const exportPDF = () => {
    const input = document.getElementById('report-container');
    html2canvas(input, { scale: 2 }).then((canvas) => {
      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF('p', 'mm', 'a4');
      const imgProps = pdf.getImageProperties(imgData);
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
      pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
      pdf.save(`통합고용세액공제_${selectedCorp}_${selectedBranch || '통합'}_${year}.pdf`);
    });
  };

  const toggleColumn = (id) => {
    setVisibleColumns(prev => 
      prev.includes(id) ? prev.filter(c => c !== id) : [...prev, id]
    );
  };

  const toggleSummaryColumn = (id) => {
    setVisibleSummaryColumns(prev => 
      prev.includes(id) ? prev.filter(c => c !== id) : [...prev, id]
    );
  };

  if (loading && !Object.keys(activeResult?.corporations || {}).length) return <div className="loading-overlay">데이터를 분석하는 중...</div>;

  const totalDiff = (currentYearData.avgTotal - prevYearData.avgTotal).toFixed(2);
  const youthDiff = (currentYearData.avgYouth - prevYearData.avgYouth).toFixed(2);

  const trendData = (data.years || [2022, 2023, 2024, 2025, 2026]).slice().sort((a,b) => a-b).map(yr => {
    const yrD = currentBranchData?.years?.[yr] || {};
    return {
      name: `${yr}년`,
      '전체': yrD.avgTotal || 0,
      '청년': yrD.avgYouth || 0,
    };
  });

  return (
    <div className="app-root">
      <aside className="sidebar no-print">
        <div className="sidebar-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
          <div className="sidebar-title" style={{ margin: 0 }}>Corporations</div>
          <button 
            onClick={() => setIsManaging(!isManaging)}
            style={{ 
              padding: '0.25rem 0.5rem', 
              fontSize: '0.7rem', 
              background: isManaging ? 'var(--primary)' : 'var(--glass)',
              color: isManaging ? 'white' : 'var(--text)',
              border: '1px solid var(--border)'
            }}
          >
            {isManaging ? '완료' : '편집'}
          </button>
        </div>

        {isManaging && (
          <div style={{ marginBottom: '1.5rem', padding: '0 0.5rem' }}>
            <button onClick={addCorporation} style={{ width: '100%', fontSize: '0.8rem', padding: '0.5rem', background: 'var(--primary)', color: 'white', borderRadius: '4px' }}>
              + 법인 추가
            </button>
          </div>
        )}

        {isManaging && (
          <div style={{ padding: '0 0.5rem', marginBottom: '1rem', fontSize: '0.75rem', color: 'var(--text-muted)', fontStyle: 'italic' }}>
            * 지점을 드래그하여 법인에 연결하세요
          </div>
        )}
        {(mapping?.corporations || []).map((corpMapping, cIdx) => {
          const corpName = corpMapping.name;
          const corpData = activeResult?.corporations?.[corpName];
          return (
            <div 
               key={corpMapping.id} 
               className={`corp-group ${dragOverCorpId === corpMapping.id ? 'drag-over' : ''}`}
               onDragOver={(e) => { 
                 e.preventDefault(); 
                 e.stopPropagation();
                 e.dataTransfer.dropEffect = 'move';
                 if (dragOverCorpId !== corpMapping.id) setDragOverCorpId(corpMapping.id); 
               }}
               onDrop={(e) => onCorpDrop(e, corpMapping.id)}
               style={{ 
                 border: isManaging ? '1px dashed rgba(255,255,255,0.2)' : 'none', 
                 padding: isManaging ? '0.5rem' : '0', 
                 marginBottom: '0.5rem', 
                 borderRadius: '8px',
                 background: dragOverCorpId === corpMapping.id ? 'rgba(59, 130, 246, 0.2)' : 'transparent',
                 boxShadow: dragOverCorpId === corpMapping.id ? 'inset 0 0 0 2px #60a5fa, 0 0 15px rgba(59, 130, 246, 0.3)' : 'none',
                 transition: 'all 0.1s ease'
               }}
            >
              <div className="corp-name" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '0.2rem 0' }}>
                <span style={{ display: 'flex', alignItems: 'center', gap: '0.4rem', flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', fontWeight: 700, fontSize: '0.75rem', color: '#60a5fa' }}>
                  <FileSpreadsheet size={12} style={{ flexShrink: 0 }} />
                  {corpName}
                </span>
                {isManaging && (
                  <div style={{ display: 'flex', gap: '1px' }}>
                    <button onClick={() => moveCorp(cIdx, -1)} disabled={cIdx === 0} style={{ padding: '1px 3px', background: 'transparent', color: 'var(--text-muted)', border: 'none', opacity: cIdx === 0 ? 0.25 : 0.6 }}><ArrowUp size={10} /></button>
                    <button onClick={() => moveCorp(cIdx, 1)} disabled={cIdx === mapping.corporations.length - 1} style={{ padding: '1px 3px', background: 'transparent', color: 'var(--text-muted)', border: 'none', opacity: cIdx === mapping.corporations.length - 1 ? 0.25 : 0.6 }}><ArrowDown size={10} /></button>
                  </div>
                )}
              </div>
              {isManaging && (
                <div style={{ display: 'flex', gap: '8px', padding: '2px 2px 6px' }}>
                  <button onClick={() => renameCorporation(corpMapping.id, corpName)} style={{ display: 'flex', alignItems: 'center', gap: '3px', padding: '0', fontSize: '0.65rem', background: 'none', color: '#818cf8', border: 'none', cursor: 'pointer' }}><Edit2 size={9} /> 이름수정</button>
                  <button onClick={() => deleteCorporation(corpMapping.id)} style={{ display: 'flex', alignItems: 'center', gap: '3px', padding: '0', fontSize: '0.65rem', background: 'none', color: '#f87171', border: 'none', cursor: 'pointer' }}><Trash2 size={9} /> 삭제</button>
                </div>
              )}
              <ul className="branch-list" style={dragStyle}>
                <li 
                  className={`branch-item total ${selectedCorp === corpName && selectedBranch === null ? 'active' : ''}`}
                  onClick={() => { setSelectedCorp(corpName); setSelectedBranch(null); }}
                >
                  전체 합계 (Consolidated)
                </li>
                {(corpMapping.branchNames || []).map((brName, bIdx) => (
                  <li
                    key={brName}
                    style={{ padding: isManaging ? '0' : undefined }}
                  >
                    <div
                      className={`branch-item ${selectedCorp === corpName && selectedBranch === brName ? 'active' : ''}`}
                      style={{
                        cursor: isManaging ? 'grab' : 'pointer',
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                        marginBottom: 0
                      }}
                      onClick={() => { setSelectedCorp(corpName); setSelectedBranch(brName); }}
                      draggable={isManaging}
                      onDragStart={(e) => { e.dataTransfer.effectAllowed = 'move'; onBranchDragStart(e, brName); }}
                      onDragEnd={onBranchDragEnd}
                    >
                      <span style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', color: isManaging ? '#34d399' : undefined, fontWeight: isManaging ? 500 : undefined, display: 'flex', alignItems: 'center', gap: '6px' }}>
                        {brName}
                        {!isManaging && (() => {
                          const vItems = getValidationForBranch(brName);
                          if (vItems.length === 0) return null;
                          const hasError = vItems.some(i => i.type === 'error');
                          return (
                            <button 
                              onClick={(e) => { e.stopPropagation(); setShowValidation({ branch: brName, items: vItems }); }}
                              style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                              title={`${vItems.length}개의 데이터 경고가 있습니다.`}
                            >
                              {hasError ? <AlertCircle size={12} style={{ color: '#ef4444' }} /> : <AlertTriangle size={12} style={{ color: '#fbbf24' }} />}
                            </button>
                          );
                        })()}
                      </span>
                      {isManaging && (
                        <div style={{ display: 'flex', gap: '1px' }}>
                          <button onClick={(e) => { e.stopPropagation(); moveBranch(corpMapping.id, bIdx, -1); }} disabled={bIdx === 0} style={{ padding: '1px 3px', background: 'transparent', border: 'none', color: 'var(--text-muted)', opacity: bIdx === 0 ? 0.25 : 0.6 }}><ArrowUp size={9} /></button>
                          <button onClick={(e) => { e.stopPropagation(); moveBranch(corpMapping.id, bIdx, 1); }} disabled={bIdx === corpMapping.branchNames.length - 1} style={{ padding: '1px 3px', background: 'transparent', border: 'none', color: 'var(--text-muted)', opacity: bIdx === corpMapping.branchNames.length - 1 ? 0.25 : 0.6 }}><ArrowDown size={9} /></button>
                        </div>
                      )}
                    </div>
                    {isManaging && (
                      <div style={{ display: 'flex', gap: '10px', padding: '2px 4px 5px' }}>
                        <button onClick={(e) => { e.stopPropagation(); renameBranch(brName); }} style={{ display: 'flex', alignItems: 'center', gap: '3px', padding: '0', fontSize: '0.63rem', background: 'none', color: '#818cf8', border: 'none', cursor: 'pointer' }}><Edit2 size={8} /> 수정</button>
                        <button onClick={(e) => { e.stopPropagation(); clearBranchData(brName); }} style={{ display: 'flex', alignItems: 'center', gap: '3px', padding: '0', fontSize: '0.63rem', background: 'none', color: '#fbbf24', border: 'none', cursor: 'pointer' }}><RotateCcw size={8} /> 비우기</button>
                        <button onClick={(e) => { e.stopPropagation(); deleteBranch(brName); }} style={{ display: 'flex', alignItems: 'center', gap: '3px', padding: '0', fontSize: '0.63rem', background: 'none', color: '#f87171', border: 'none', cursor: 'pointer' }}><Trash2 size={8} /> 삭제</button>
                      </div>
                    )}
                  </li>
                ))}
              </ul>
            </div>
          );
        })}

        {(Object.keys(activeResult?.unassigned || {}).length > 0 || isManaging) && (
           <div 
            className={`corp-group unassigned ${dragOverCorpId === 'unassigned' ? 'drag-over' : ''}`}
            onDragOver={(e) => { 
              e.preventDefault(); 
              e.stopPropagation();
              e.dataTransfer.dropEffect = 'move';
              if (dragOverCorpId !== 'unassigned') setDragOverCorpId('unassigned'); 
            }}
            onDrop={onUnassignedDrop}
            style={{ 
              marginTop: '2rem', 
              borderTop: '1px solid rgba(255,255,255,0.1)', 
              paddingTop: '1rem',
              background: dragOverCorpId === 'unassigned' ? 'rgba(59, 130, 246, 0.1)' : 'transparent',
              padding: dragOverCorpId === 'unassigned' ? '0.5rem' : '0',
              boxShadow: dragOverCorpId === 'unassigned' ? 'inset 0 0 0 2px rgba(59, 130, 246, 0.3)' : 'none',
              borderRadius: '8px'
            }}
          >
            <div className="corp-name" style={{ ...dragStyle, color: '#94a3b8', fontSize: '0.75rem' }}>미분류 지점 (Unassigned)</div>
            <ul className="branch-list" style={dragStyle}>
              {Object.keys(activeResult?.unassigned || {}).map(brName => (
                  <li 
                    key={brName} 
                    className={`branch-item ${selectedBranch === brName ? 'active' : ''}`}
                    style={{ color: '#94a3b8', cursor: isManaging ? 'grab' : 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}
                    onClick={() => { setSelectedCorp(null); setSelectedBranch(brName); }}
                    draggable={isManaging}
                    onDragStart={(e) => onBranchDragStart(e, brName)}
                    onDragEnd={onBranchDragEnd}
                  >
                    <span style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      {brName}
                      {!isManaging && (() => {
                        const vItems = getValidationForBranch(brName);
                        if (vItems.length === 0) return null;
                        const hasError = vItems.some(i => i.type === 'error');
                        return (
                          <AlertTriangle 
                            size={11} 
                            style={{ color: hasError ? '#ef4444' : '#fbbf24', cursor: 'pointer' }} 
                            onClick={(e) => { e.stopPropagation(); setShowValidation({ branch: brName, items: vItems }); }}
                            title="데이터 경고"
                          />
                        );
                      })()}
                    </span>
                  </li>
              ))}
              {Object.keys(activeResult?.unassigned || {}).length === 0 && (
                <li style={{ fontSize: '0.75rem', color: '#64748b', padding: '0.5rem', textAlign: 'center' }}>데이터 없음</li>
              )}
            </ul>
          </div>
        )}

        <div style={{ padding: '1rem 0.75rem', borderTop: '1px solid rgba(255,255,255,0.05)', marginTop: 'auto', background: 'rgba(0,0,0,0.1)' }}>
          <h4 style={{ fontSize: '0.75rem', marginBottom: '0.75rem', color: 'var(--text-muted)', display: 'flex', alignItems: 'center', gap: '0.4rem' }}>
            <RotateCcw size={12} /> 최근 활동 기록
          </h4>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.6rem' }}>
            {(data.activities || []).slice(0, 5).map((act, i) => (
              <div key={i} style={{ fontSize: '0.7rem', borderLeft: '2px solid var(--primary)', paddingLeft: '0.5rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.1rem', opacity: 0.7, fontSize: '0.6rem' }}>
                  <span>{act.type}</span>
                  <span>{new Date(act.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
                </div>
                <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', color: 'var(--text)' }} title={act.msg}>
                  {act.msg}
                </div>
              </div>
            ))}
            {(!data.activities || data.activities.length === 0) && (
              <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', textAlign: 'center', padding: '0.5rem' }}>최근 기록 없음</div>
            )}
          </div>
        </div>
      </aside>

      <div className="container">
        {isDragging && <div className="loading-overlay" style={{ background: 'rgba(59, 130, 246, 0.4)', border: '4px dashed white', zIndex: 100 }}>파일들을 여기에 놓으세요!</div>}
        <header className="no-print">
          <div className="title-area">
            <h1>통합고용세액공제 시스템</h1>
            <p style={{ color: 'var(--text-muted)', fontSize: '0.875rem', display: 'flex', alignItems: 'center', gap: '0.4rem' }}>
              <FileSpreadsheet size={16} /> 
              <span style={{ color: 'var(--primary)', fontWeight: 'bold' }}>{selectedCorp || '지정되지 않음'}</span>
              {selectedBranch ? ` / ${selectedBranch}` : ' / 전체 합계'}
            </p>
          </div>
          <div className="controls">
            <button 
              onClick={toggleTheme}
              style={{ background: 'var(--glass)', color: 'var(--text)', border: '1px solid var(--border)' }}
              title="테마 전환"
            >
              {theme === 'dark-theme' ? <Sun size={18} /> : <Moon size={18} />}
            </button>
            <select value={year} onChange={(e) => setYear(parseInt(e.target.value))}>
              {(data.years || [2026, 2025, 2024, 2023, 2022]).map(y => <option key={y} value={y}>{y}년</option>)}
            </select>
            <div style={{ display: 'flex', alignItems: 'center', background: 'var(--glass)', borderRadius: '0.75rem', border: '1px solid var(--border)', padding: '2px' }}>
              <label 
                className="btn-upload" 
                style={{ 
                  margin: 0, 
                  opacity: 1,
                  cursor: 'pointer'
                }}
              >
                <Upload size={18} /> 자료 업로드
                <input 
                  type="file" 
                  accept=".xlsx" 
                  style={{ display: 'none' }} 
                  disabled={false}
                  onChange={(e) => {
                    const file = e.target.files[0];
                    if (file) handleTargetedUpload(file);
                    e.target.value = ''; // Reset
                  }} 
                />
              </label>
            </div>
            <button onClick={() => setShowPicker(!showPicker)} style={{ background: '#6366f1' }}>항목 선택</button>
            <button onClick={exportFullExcel} style={{ background: '#10b981' }}><Download size={18} /> 엑셀</button>
            <button onClick={exportPDF} style={{ background: '#f59e0b' }}><FileText size={18} /> PDF</button>
            <button onClick={handleRefresh} style={{ marginLeft: 'auto' }}><RefreshCw size={18} /> 갱신(초기화)</button>
          </div>
        </header>

        {showPicker && (
          <div className="glass-card no-print" style={{ marginBottom: '1.5rem' }}>
            <div style={{ marginBottom: '1rem', borderBottom: '1px solid rgba(255,255,255,0.1)', paddingBottom: '0.5rem' }}>
              <h4 style={{ fontSize: '0.9rem', marginBottom: '0.5rem', color: '#6366f1' }}>1. 상시근로자수 계산 항목</h4>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '0.5rem' }}>
                {allSummaryColumns.map(col => (
                  <label key={col.id} style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', fontSize: '0.8rem', cursor: 'pointer' }}>
                    <input type="checkbox" checked={visibleSummaryColumns.includes(col.id)} onChange={() => toggleSummaryColumn(col.id)} />
                    {col.label}
                  </label>
                ))}
              </div>
            </div>
            <div>
              <h4 style={{ fontSize: '0.9rem', marginBottom: '0.5rem', color: '#10b981' }}>2. 상시근로자별 명세 항목</h4>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '0.5rem' }}>
                {allColumns.map(col => (
                  <label key={col.id} style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', fontSize: '0.8rem', cursor: 'pointer' }}>
                    <input type="checkbox" checked={visibleColumns.includes(col.id)} onChange={() => toggleColumn(col.id)} />
                    {col.label}
                  </label>
                ))}
              </div>
            </div>
          </div>
        )}

        <div className="glass-card no-print" style={{ marginBottom: '1.5rem', padding: '1rem' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
            <h3 style={{ fontSize: '0.9rem', margin: 0, color: 'var(--text-muted)' }}>연도별 고용 인원 추세 (상시근로자 수 평균)</h3>
            <div style={{ fontSize: '0.75rem', color: '#10b981' }}>※ 최근 5개년 데이터</div>
          </div>
          <div style={{ width: '100%', height: 200 }}>
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={trendData} margin={{ top: 5, right: 20, left: -20, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
                <XAxis dataKey="name" stroke="var(--text-muted)" fontSize={11} tickLine={false} axisLine={false} />
                <YAxis stroke="var(--text-muted)" fontSize={11} tickLine={false} axisLine={false} />
                <Tooltip 
                  contentStyle={{ background: 'var(--bg-card)', border: '1px solid var(--border)', borderRadius: '8px', fontSize: '11px' }}
                  itemStyle={{ padding: '2px 0' }}
                />
                <Legend iconType="circle" wrapperStyle={{ fontSize: '11px', paddingTop: '10px' }} />
                <Line type="monotone" dataKey="전체" stroke="#6366f1" strokeWidth={3} dot={{ r: 4, fill: '#6366f1' }} activeDot={{ r: 6 }} />
                <Line type="monotone" dataKey="청년" stroke="#10b981" strokeWidth={3} dot={{ r: 4, fill: '#10b981' }} activeDot={{ r: 6 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="stats-grid no-print">
          <div className="glass-card stat-card">
            <div className="stat-label">상시근로자 수 (평균)</div>
            <div className="stat-value">{currentYearData.avgTotal}명</div>
            <div className={`stat-diff ${totalDiff >= 0 ? 'diff-up' : 'diff-down'}`}>
              {totalDiff >= 0 ? <ArrowUpRight size={16} /> : <ArrowDownRight size={16} />}
              전년 대비 {Math.abs(totalDiff)}
            </div>
          </div>
          <div className="glass-card stat-card">
            <div className="stat-label">청년 근로자 수 (평균)</div>
            <div className="stat-value">{currentYearData.avgYouth}명</div>
            <div className={`stat-diff ${youthDiff >= 0 ? 'diff-up' : 'diff-down'}`}>
              {youthDiff >= 0 ? <ArrowUpRight size={16} /> : <ArrowDownRight size={16} />}
              전년 대비 {Math.abs(youthDiff)}
            </div>
          </div>
          <div className="glass-card stat-card">
            <div className="stat-label">청년 외 근로자 수 (평균)</div>
            <div className="stat-value">{currentYearData.avgOther}명</div>
            <div className="stat-diff" style={{ color: 'var(--text-muted)' }}>
              비중 {currentYearData.avgTotal > 0 ? ((currentYearData.avgOther / currentYearData.avgTotal) * 100).toFixed(1) : 0}%
            </div>
          </div>
        </div>

        <div id="report-container" className="table-container">
          <div className="table-header">
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.8rem', marginBottom: '0.5rem' }}>
                <span>※ [ ] 에는 해당하는 곳에 [V] 표를 합니다.</span>
                <span>(앞쪽)</span>
              </div>
              <table className="info-table" style={{ width: '100%', marginBottom: '1rem', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                <tbody>
                  <tr>
                    <td style={{ width: '15%', background: '#f8fafc', fontWeight: '600' }}>과세연도</td>
                    <td style={{ width: '35%', textAlign: 'left' }}>{year}.01.01 ~ {year}.12.31</td>
                    <td style={{ width: '15%', background: '#f8fafc', fontWeight: '600' }}>법인종류별구분</td>
                    <td style={{ width: '35%', textAlign: 'left' }}>중소</td>
                  </tr>
                  <tr>
                    <td style={{ background: '#f8fafc', fontWeight: '600' }}>상호 또는 법인명</td>
                    <td style={{ textAlign: 'left' }}>{selectedCorp || '(주)평우서비스'}</td>
                    <td style={{ background: '#f8fafc', fontWeight: '600' }}>사업자등록번호</td>
                    <td style={{ textAlign: 'left' }}>126-86-22464</td>
                  </tr>
                </tbody>
              </table>
              <h2 style={{ fontSize: '1.25rem', color: '#1e293b', marginBottom: '1.5rem', marginTop: '2rem' }}>통합고용세액공제 상시근로자 명세서</h2>
          </div>

          <h3 style={{ fontSize: '1rem', marginBottom: '0.5rem', color: '#334155', textAlign: 'left' }}>1. 상시근로자수계산</h3>
          <table style={{ marginBottom: '2.5rem' }}>
            <thead>
              <tr>
                <th rowSpan={2}>구 분</th>
                <th rowSpan={2}>① 해당(직전) 과세연도의 상시근로자 근무개월수의 합계</th>
                <th rowSpan={2}>② 과세연도 개월수</th>
                <th rowSpan={2}>③ 상시근로자수 (=①÷②)</th>
                <th colSpan={6}>④ 해당(직전) 과세연도의 청년등 상시근로자 근무개월수의 합계</th>
                <th rowSpan={2}>⑪ 과세연도 개월수</th>
                <th rowSpan={2}>⑫ 청년등상시 근로자수 (=⑩÷⑪)</th>
                <th rowSpan={2}>⑬ 청년등외 상시 근로자수 (=③-⑫)</th>
              </tr>
              <tr>
                <th>⑤ 청년</th>
                <th>⑥ 장애인</th>
                <th>⑦ 고령자</th>
                <th>⑧ 경력 단절 근로자</th>
                <th>⑨ 북한 이탈 주민</th>
                <th>⑩ 합계</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td style={{ background: '#f8fafc', fontWeight: '600' }}>해당과세연도</td>
                <td>{currentYearData.summary?.col1 || '0'}</td>
                <td>{currentYearData.summary?.col2 || '12'}</td>
                <td style={{ fontWeight: '700', color: '#2563eb' }}>{currentYearData.summary?.col3 || '0'}</td>
                <td>{currentYearData.summary?.col5 || '0'}</td>
                <td>{currentYearData.summary?.col6 || '0'}</td>
                <td>{currentYearData.summary?.col7 || '0'}</td>
                <td>{currentYearData.summary?.col8 || '0'}</td>
                <td>{currentYearData.summary?.col9 || '0'}</td>
                <td>{currentYearData.summary?.col10 || '0'}</td>
                <td>{currentYearData.summary?.col11 || '12'}</td>
                <td style={{ fontWeight: '700', color: '#2563eb' }}>{currentYearData.summary?.col12 || '0'}</td>
                <td>{currentYearData.summary?.col13 || '0'}</td>
              </tr>
              <tr>
                <td style={{ background: '#f8fafc', fontWeight: '600' }}>직전과세연도</td>
                <td>{prevYearData.summary?.col1 || '0'}</td>
                <td>{prevYearData.summary?.col2 || '12'}</td>
                <td style={{ fontWeight: '700', color: '#64748b' }}>{prevYearData.summary?.col3 || '0'}</td>
                <td>{prevYearData.summary?.col5 || '0'}</td>
                <td>{prevYearData.summary?.col6 || '0'}</td>
                <td>{prevYearData.summary?.col7 || '0'}</td>
                <td>{prevYearData.summary?.col8 || '0'}</td>
                <td>{prevYearData.summary?.col9 || '0'}</td>
                <td>{prevYearData.summary?.col10 || '0'}</td>
                <td>{prevYearData.summary?.col11 || '12'}</td>
                <td style={{ fontWeight: '700', color: '#64748b' }}>{prevYearData.summary?.col12 || '0'}</td>
                <td>{prevYearData.summary?.col13 || '0'}</td>
              </tr>
            </tbody>
          </table>

          <h3 style={{ fontSize: '1rem', marginBottom: '0.5rem', color: '#334155', textAlign: 'left' }}>2. 해당과세연도상시근로자별명세</h3>
          <table id="statement-table">
            <thead>
              <tr>
                {allColumns.map(col => {
                  if (!visibleColumns.includes(col.id)) return null;
                  if (['col25', 'col26', 'col27'].includes(col.id)) {
                    if (col.id === 'col25' || (!visibleColumns.includes('col25') && col.id === 'col26') || (!visibleColumns.includes('col25') && !visibleColumns.includes('col26') && col.id === 'col27')) {
                      const count = ['col25', 'col26', 'col27'].filter(id => visibleColumns.includes(id)).length;
                      return <th key="group-youth" colSpan={count}>청년</th>;
                    }
                    return null;
                  }
                  return <th key={col.id} rowSpan={2}>{col.label}</th>;
                })}
              </tr>
              <tr>
                {['col25', 'col26', 'col27'].filter(id => visibleColumns.includes(id)).map(id => {
                  const col = allColumns.find(c => c.id === id);
                  return <th key={id}>{col.label}</th>;
                })}
              </tr>
            </thead>
            <tbody>
              {currentYearData.details.map((d, i) => (
                <tr 
                  key={i} 
                  style={{ 
                    backgroundColor: d._retiredThisYear ? 'rgba(239, 68, 68, 0.25)' : 'transparent',
                    borderLeft: d._retiredThisYear ? '4px solid #ef4444' : 'none'
                  }}
                >
                  {allColumns.filter(c => visibleColumns.includes(c.id)).map(col => (
                    <td key={col.id} style={{ paddingLeft: d._retiredThisYear && col.id === visibleColumns[0] ? '0.5rem' : '' }}>
                      {d[col.id]}
                      {col.id === 'col15' && d._retiredThisYear && (
                        <span className="retired-marker" style={{ color: '#ef4444', fontWeight: 'bold', marginLeft: '0.5rem' }}>(퇴사)</span>
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {showValidation && (
        <div className="modal-overlay" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }} onClick={() => setShowValidation(null)}>
          <div className="glass-card" style={{ width: '600px', maxHeight: '80vh', overflowY: 'auto', padding: '1.5rem', position: 'relative' }} onClick={e => e.stopPropagation()}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem', borderBottom: '1px solid var(--border)', paddingBottom: '0.5rem' }}>
              <h3 style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', margin: 0 }}>
                {showValidation.items.some(i => i.type === 'error') ? <AlertCircle style={{ color: '#ef4444' }} /> : <AlertTriangle style={{ color: '#fbbf24' }} />}
                [{showValidation.branch}] 데이터 검증 결과
              </h3>
              <button onClick={() => setShowValidation(null)} style={{ background: 'none', border: 'none', color: 'var(--text)', fontSize: '1.5rem', cursor: 'pointer' }}>&times;</button>
            </div>
            
            <div style={{ fontSize: '0.9rem', color: 'var(--text-muted)', marginBottom: '1rem' }}>
              총 {showValidation.items.length}개의 어색한 데이터가 발견되었습니다. 파일을 확인해 주세요.
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
              {showValidation.items.map((item, idx) => (
                <div key={idx} style={{ 
                  background: item.type === 'error' ? 'rgba(239, 68, 68, 0.1)' : 'rgba(234, 179, 8, 0.1)', 
                  padding: '0.75rem', 
                  borderRadius: '8px',
                  border: `1px solid ${item.type === 'error' ? 'rgba(239, 68, 68, 0.2)' : 'rgba(234, 179, 8, 0.2)'}`
                }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.3rem' }}>
                    <span style={{ fontWeight: 'bold', color: item.type === 'error' ? '#f87171' : '#fbbf24', fontSize: '0.8rem' }}>
                      {item.type === 'error' ? '[심각 오류]' : '[주의]'} {item.field}
                    </span>
                    <span style={{ fontSize: '0.75rem', opacity: 0.7 }}>{item.fileName} / {item.row}행</span>
                  </div>
                  <div style={{ fontSize: '0.85rem' }}>{item.msg}</div>
                </div>
              ))}
            </div>

            <button 
              onClick={() => setShowValidation(null)} 
              style={{ width: '100%', marginTop: '1.5rem', padding: '0.75rem', borderRadius: '8px', cursor: 'pointer', background: 'var(--primary)', border: 'none', color: 'white', fontWeight: 'bold' }}
            >
              닫기
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
