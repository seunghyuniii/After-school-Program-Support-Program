/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { StudentRow, CourseInfo, StudentSettlement, CourseAttendance } from './types';

/**
 * Robust normalization of Student ID values to prevent trailing float decimal mismatches (.0) from Excel reading
 */
export function normalizeStudentId(id: any): string {
  if (id === null || id === undefined) return '';
  let str = String(id).trim();
  if (str.endsWith('.0')) {
    str = str.slice(0, -2);
  }
  return str;
}

/**
 * Parses school Student ID to extract Grade and Class dynamically
 */
export function parseStudentClass(studentId: string): { grade: string, classNum: string, key: string, name: string } {
  const idStr = String(studentId).trim();
  if (idStr.length === 5 && /^\d+$/.test(idStr)) {
    const grade = idStr.substring(0, 1);
    const classNum = parseInt(idStr.substring(1, 3), 10).toString();
    return { grade, classNum, key: `${grade}학년 ${classNum}반`, name: `${classNum}반` };
  } else if (idStr.length === 4 && /^\d+$/.test(idStr)) {
    const grade = idStr.substring(0, 1);
    const classNum = parseInt(idStr.substring(1, 2), 10).toString();
    return { grade, classNum, key: `${grade}학년 ${classNum}반`, name: `${classNum}반` };
  }
  return { grade: '공통', classNum: '공통', key: '공통반', name: '공통반' };
}

/**
 * Generates a mock standard Liro School afterschool enrollment spreadsheet
 * so users can test the application instantly without needing real files.
 */
export function generateSampleExcel() {
  const headers = [
    '순번',
    '학번',
    '이름',
    '강좌수',
    '유료강좌',
    '수리논술 심화(A반)',
    '수능 실전 최고급 영어회화',
    '인공지능 웹서비스 개발 실습',
    '문학적 상상력과 논술 입문',
    '실험 중심 고급 화학'
  ];
  
  const sampleData = [
    { '순번': 1, '학번': 10101, '이름': '강유민', '강좌수': 2, '유료강좌': 2, '수리논술 심화(A반)': '수리논술 심화(A반)', '수능 실전 최고급 영어회화': '', '인공지능 웹서비스 개발 실습': '인공지능 웹서비스 개발 실습', '문학적 상상력과 논술 입문': '', '실험 중심 고급 화학': '' },
    { '순번': 2, '학번': 10105, '이름': '고태환', '강좌수': 1, '유료강좌': 1, '수리논술 심화(A반)': '', '수능 실전 최고급 영어회화': '수능 실전 최고급 영어회화', '인공지능 웹서비스 개발 실습': '', '문학적 상상력과 논술 입문': '', '실험 중심 고급 화학': '' },
    { '순번': 3, '학번': 10112, '이름': '김지민', '강좌수': 3, '유료강좌': 3, '수리논술 심화(A반)': '수리논술 심화(A반)', '수능 실전 최고급 영어회화': '수능 실전 최고급 영어회화', '인공지능 웹서비스 개발 실습': '', '문학적 상상력과 논술 입문': '문학적 상상력과 논술 입문', '실험 중심 고급 화학': '' },
    { '순번': 4, '학번': 10203, '이름': '박건우', '강좌수': 2, '유료강좌': 2, '수리논술 심화(A반)': '', '수능 실전 최고급 영어회화': '', '인공지능 웹서비스 개발 실습': '인공지능 웹서비스 개발 실습', '문학적 상상력과 논술 입문': '문학적 상상력과 논술 입문', '실험 중심 고급 화학': '' },
    { '순번': 5, '학번': 10214, '이름': '배성우', '강좌수': 2, '유료강좌': 2, '수리논술 심화(A반)': '수리논술 심화(A반)', '수능 실전 최고급 영어회화': '', '인공지능 웹서비스 개발 실습': '', '문학적 상상력과 논술 입문': '', '실험 중심 고급 화학': '실험 중심 고급 화학' },
    { '순번': 6, '학번': 10302, '이름': '신은선', '강좌수': 1, '유료강좌': 1, '수리논술 심화(A반)': '', '수능 실전 최고급 영어회화': '', '인공지능 웹서비스 개발 실습': '', '문학적 상상력과 논술 입문': '', '실험 중심 고급 화학': '실험 중심 고급 화학' },
    { '순번': 7, '학번': 10309, '이름': '안재현', '강좌수': 2, '유료강좌': 2, '수리논술 심화(A반)': '', '수능 실전 최고급 영어회화': '수능 실전 최고급 영어회화', '인공지능 웹서비스 개발 실습': '인공지능 웹서비스 개발 실습', '문학적 상상력과 논술 입문': '', '실험 중심 고급 화학': '' },
    { '순번': 8, '학번': 10411, '이름': '오진우', '강좌수': 4, '유료강좌': 4, '수리논술 심화(A반)': '수리논술 심화(A반)', '수능 실전 최고급 영어회화': '수능 실전 최고급 영어회화', '인공지능 웹서비스 개발 실습': '인공지능 웹서비스 개발 실습', '문학적 상상력과 논술 입문': '문학적 상상력과 논술 입문', '실험 중심 고급 화학': '' },
    { '순번': 9, '학번': 10425, '이름': '임태경', '강좌수': 1, '유료강좌': 1, '수리논술 심화(A반)': '', '수능 실전 최고급 영어회화': '', '인공지능 웹서비스 개발 실습': '', '문학적 상상력과 논술 입문': '문학적 상상력과 논술 입문', '실험 중심 고급 화학': '' },
    { '순번': 10, '학번': 10507, '이름': '장하진', '강좌수': 2, '유료강좌': 2, '수리논술 심화(A반)': '', '수능 실전 최고급 영어회화': '', '인공지능 웹서비스 개발 실습': '', '문학적 상상력과 논술 입문': '문학적 상상력과 논술 입문', '실험 중심 고급 화학': '실험 중심 고급 화학' },
    { '순번': 11, '학번': 10515, '이름': '최수아', '강좌수': 3, '유료강좌': 3, '수리논술 심화(A반)': '수리논술 심화(A반)', '수능 실전 최고급 영어회화': '', '인공지능 웹서비스 개발 실습': '인공지능 웹서비스 개발 실습', '문학적 상상력과 논술 입문': '', '실험 중심 고급 화학': '실험 중심 고급 화학' },
    { '순번': 12, '학번': 10603, '이름': '한지원', '강좌수': 2, '유료강좌': 2, '수리논술 심화(A반)': '', '수능 실전 최고급 영어회화': '수능 실전 최고급 영어회화', '인공지능 웹서비스 개발 실습': '', '문학적 상상력과 논술 입문': '', '실험 중심 고급 화학': '실험 중심 고급 화학' },
    { '순번': 13, '학번': 10620, '이름': '황정민', '강좌수': 0, '유료강좌': 0, '수리논술 심화(A반)': '', '수능 실전 최고급 영어회화': '', '인공지능 웹서비스 개발 실습': '', '문학적 상상력과 논술 입문': '', '실험 중심 고급 화학': '' }
  ];

  const ws = XLSX.utils.json_to_sheet(sampleData, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "방과후신청현황_리로스쿨");
  
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  
  const a = document.createElement('a');
  a.href = url;
  a.download = '리로스쿨_방과후_샘플데이터.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Parses the raw workbook sheet into standard structured JSON
 * while filtering headers and redundant empty records.
 */
export function parseWorkbook(workbook: XLSX.WorkBook): {
  rawData: StudentRow[];
  courseColumns: string[];
} {
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const jsonData = XLSX.utils.sheet_to_json<any>(worksheet, { defval: "" });

  if (jsonData.length === 0) {
    throw new Error("엑셀 파일에 분석 가능한 데이터가 존재하지 않습니다.");
  }

  // Filter out headers or rows where '순번' is the literal title '순번' or '학번' is missing
  const filteredData = jsonData.filter((row: any) => {
    const sId = normalizeStudentId(row['학번']);
    const sNum = String(row['순번'] || '').trim();
    return sNum !== '순번' && sId !== '학번' && sId !== "";
  }) as StudentRow[];

  if (filteredData.length === 0) {
    throw new Error("유효한 학생 데이터 행을 찾지 못했습니다. 학번이나 이름이 포함되어 있는지 양식을 재확인 하십시오.");
  }

  // First row serves as representative key mapping
  const headers = Object.keys(filteredData[0]);
  const baseCols = ['순번', '학번', '이름', '강좌수', '유료강좌'];
  const courseColumns = headers.filter(col => !baseCols.includes(col) && !col.startsWith('__EMPTY'));

  return {
    rawData: filteredData,
    courseColumns
  };
}

/**
 * Generates a mock standard Subsidy spreadsheet
 * so users can test uploading support details instantly.
 */
export function generateSampleSubsidyExcel() {
  const headers = ['학번', '이름', '지원금 종류', '지원금액'];
  const sampleData = [
    { '학번': '10101', '이름': '강유민', '지원금 종류': '자유수강권', '지원금액': 120000 },
    { '학번': '10204', '이름': '김도윤', '지원금 종류': '학교자체지원금', '지원금액': 50000 },
    { '학번': '10212', '이름': '남궁혜원', '지원금 종류': '교육청지원금', '지원금액': 80000 }
  ];

  const ws = XLSX.utils.json_to_sheet(sampleData, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "지원금_명부");
  
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  
  const a = document.createElement('a');
  a.href = url;
  a.download = '방과후_지원금_업로드_양식.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function checkLowIncome(str: string): boolean {
  const s = str.toLowerCase();
  return s.includes('저소득') || 
         s.includes('자유수강') || 
         s.includes('바우처') || 
         s.includes('기초수급') || 
         s.includes('기초생활') || 
         s.includes('차상위') || 
         s.includes('한부모') || 
         s.includes('법정') ||
         s.includes('수급자');
}

function checkMultiChild(str: string): boolean {
  const s = str.toLowerCase();
  return s.includes('다자녀') || 
         s.includes('다자') || 
         s.includes('세자녀') || 
         s.includes('삼자녀') || 
         s.includes('셋째') || 
         s.includes('둘째') || 
         s.includes('형제') || 
         (s.includes('자녀') && !s.includes('자유수강'));
}

/**
 * Parses the subsidy spreadsheet workbook into structured key-value map by Student ID.
 */
export function parseSubsidyWorkbook(workbook: XLSX.WorkBook): Record<string, { type: string; amount: number }> {
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const jsonData = XLSX.utils.sheet_to_json<any>(worksheet, { defval: "" });

  if (jsonData.length === 0) {
    throw new Error("지원금 엑셀 파일에 분석 가능한 데이터가 존재하지 않습니다.");
  }

  const result: Record<string, { type: string; amount: number }> = {};

  jsonData.forEach((row: any) => {
    let studentId = "";
    let name = "";
    let isLowIncome = false;
    let isMultiChild = false;
    let amount = 0;

    for (const key of Object.keys(row)) {
      const cleanKey = key.trim();
      const val = String(row[key] !== undefined && row[key] !== null ? row[key] : "").trim();
      if (!val) continue;

      const isIdCol = cleanKey === '학번' || cleanKey.includes('학번') || cleanKey.toLowerCase() === 'studentid' || cleanKey === '학생번호';
      const isNameCol = cleanKey === '이름' || cleanKey === '학생명' || cleanKey === '성명' || cleanKey.includes('이름') || cleanKey.includes('학생명') || cleanKey.includes('성명') || cleanKey.toLowerCase() === 'name';

      if (isIdCol) {
        studentId = normalizeStudentId(val);
      } else if (isNameCol) {
        name = val;
      } else {
        // Check if the key itself indicates a subsidy
        if (checkLowIncome(cleanKey)) {
          if (val !== '0' && val.toLowerCase() !== 'n' && val.toLowerCase() !== 'no' && val !== 'X' && val !== 'x' && val !== 'N') {
            isLowIncome = true;
          }
        } else if (checkMultiChild(cleanKey)) {
          if (val !== '0' && val.toLowerCase() !== 'n' && val.toLowerCase() !== 'no' && val !== 'X' && val !== 'x' && val !== 'N') {
            isMultiChild = true;
          }
        }

        // Check the value too
        if (checkLowIncome(val)) {
          isLowIncome = true;
        } else if (checkMultiChild(val)) {
          isMultiChild = true;
        }

        // Parse amount if any
        if (cleanKey.includes('금액') || cleanKey.includes('수혜액') || cleanKey.includes('보조금') || cleanKey.includes('수혜금액')) {
          const num = parseFloat(val.replace(/,/g, ''));
          if (!isNaN(num)) {
            amount = num;
          }
        }
      }
    }

    if (studentId) {
      // Determine this row's type
      let rowType: '저소득층' | '다자녀' | null = null;
      if (isLowIncome) {
        rowType = '저소득층';
      } else if (isMultiChild) {
        rowType = '다자녀';
      } else {
        // Default to 저소득층 if listed but unspecified
        rowType = '저소득층';
      }

      if (rowType) {
        const prev = result[studentId];
        let finalType: '저소득층' | '다자녀' = '저소득층';

        if (prev) {
          // "지원금 2개 모두에 해당하는 친구들은 저소득층 지원금 우선이므로 중복해서 지원금이 나가지 않도록 해줘"
          if (prev.type === '저소득층' || rowType === '저소득층') {
            finalType = '저소득층';
          } else {
            finalType = '다자녀';
          }
        } else {
          finalType = rowType;
        }

        result[studentId] = {
          type: finalType,
          amount: amount || (prev ? prev.amount : 0)
        };
      }
    }
  });

  return result;
}

/**
 * Calculates current analytics: course cost matrix, student balances, and attendance databases.
 */
export function calculateSettlement(
  rawData: StudentRow[],
  courseColumns: string[],
  courseHours: Record<string, number>,
  feePerHour: number,
  useTruncate10: boolean = false,
  beneficiaryCourses: Record<string, boolean> = {},
  subsidies: Record<string, { type: string; amount: number }> = {}
): {
  courses: CourseInfo[];
  students: StudentSettlement[];
  attendance: CourseAttendance[];
} {
  const courseFeeDict: Record<string, number> = {};
  
  // 1. Calculate course items
  const courses: CourseInfo[] = courseColumns.map(col => {
    // A student is signed up if the cell under the course column has some non-empty value
    const studentCount = rawData.filter(row => {
      const val = row[col];
      return val !== null && val !== undefined && String(val).trim() !== "";
    }).length;

    const hours = courseHours[col] !== undefined ? courseHours[col] : 20;
    const totalLecturerCost = hours * feePerHour;
    
    // Check if the course is beneficiary-pays (수익자 부담)
    const isBeneficiary = beneficiaryCourses[col] !== false;

    // Per student tuition calculation
    let perStudentFee = 0;
    if (studentCount > 0) {
      if (isBeneficiary) {
        if (useTruncate10) {
          // Truncation/round down to nearest 10 won
          perStudentFee = Math.floor((totalLecturerCost / studentCount) / 10) * 10;
        } else {
          // Exact 1 won unit calculation (round to nearest integer)
          perStudentFee = Math.round(totalLecturerCost / studentCount);
        }
      } else {
        perStudentFee = 0;
      }
    }
    courseFeeDict[col] = perStudentFee;

    return {
      name: col,
      studentCount,
      hours,
      totalLecturerCost,
      perStudentFee
    };
  });

  // 2. Compute individual student balances
  const students: StudentSettlement[] = rawData.map(row => {
    const studentId = normalizeStudentId(row['학번']);
    const name = String(row['이름'] || '').trim();
    const enrolledCourses: string[] = [];
    
    let originalFee = 0;
    let courseCount = 0;

    courseColumns.forEach(col => {
      const cellVal = row[col];
      if (cellVal !== null && cellVal !== undefined && String(cellVal).trim() !== "") {
        enrolledCourses.push(col);
        originalFee += courseFeeDict[col] || 0;
        courseCount++;
      }
    });

    const sub = subsidies[studentId] || { type: '없음', amount: 0 };
    const subsidyType = sub.type || '없음';
    
    // If the student has a valid subsidy, their fee is fully covered (user fee is 0).
    const isSubsidized = subsidyType !== '없음' && subsidyType !== '';
    const subsidyAmount = isSubsidized ? originalFee : 0;
    const totalFee = isSubsidized ? 0 : originalFee;

    return {
      studentId,
      name,
      courseCount,
      enrolledCourses,
      originalFee,
      subsidyType,
      subsidyAmount,
      totalFee
    };
  });

  // 3. Extract attendance rosters per course
  const attendance: CourseAttendance[] = courseColumns.map(col => {
    const enrolledRows = rawData.filter(row => {
      const val = row[col];
      return val !== null && val !== undefined && String(val).trim() !== "";
    });

    const studentsList = enrolledRows.map((row, index) => ({
      index: index + 1,
      studentId: normalizeStudentId(row['학번']),
      name: String(row['이름'] || '').trim()
    }));

    return {
      courseName: col,
      students: studentsList
    };
  });

  return {
    courses,
    students,
    attendance
  };
}

/**
 * Generates and downloads the final processed settlement workbook containing separate worksheets with high-readability styles.
 */
export async function downloadSettlementExcel(
  courses: CourseInfo[],
  students: StudentSettlement[],
  attendance: CourseAttendance[],
  feePerHour: number,
  excludedList?: {
    studentId: string;
    name: string;
    courseCount: number;
    enrolledCourses: string[];
  }[]
) {
  const wb = new ExcelJS.Workbook();
  wb.creator = '방과후학교 정산 자동화 시스템';
  wb.created = new Date();

  // Reusable Style Constants
  const fontDefault: Partial<ExcelJS.Font> = { name: '맑은 고딕', size: 10, color: { argb: 'FF1E293B' } };
  const fontBold: Partial<ExcelJS.Font> = { name: '맑은 고딕', size: 10, bold: true, color: { argb: 'FF1E293B' } };
  const fontTitle: Partial<ExcelJS.Font> = { name: '맑은 고딕', size: 14, bold: true, color: { argb: 'FF0F172A' } };
  const fontSection: Partial<ExcelJS.Font> = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FF1E3A8A' } };
  const fontDiff: Partial<ExcelJS.Font> = { name: '맑은 고딕', size: 10, bold: true, color: { argb: 'FFB45309' } };
  const fontTotal: Partial<ExcelJS.Font> = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FF065F46' } };

  const fillHeader: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
  const fillDarkHeader: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
  const fillTotal: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } };
  const fillDiff: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFBEB' } };
  const fillZebra: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };

  const borderThin: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
  };

  const borderHeader: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF94A3B8' } },
    bottom: { style: 'medium', color: { argb: 'FF64748B' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
  };

  const borderTotal: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF94A3B8' } },
    bottom: { style: 'double', color: { argb: 'FF475569' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
  };

  const alignCenter: Partial<ExcelJS.Alignment> = { vertical: 'middle', horizontal: 'center' };
  const alignRight: Partial<ExcelJS.Alignment> = { vertical: 'middle', horizontal: 'right' };
  const alignLeft: Partial<ExcelJS.Alignment> = { vertical: 'middle', horizontal: 'left' };

  // 1. Compute totals for funding-source summaries
  const totalStudentPaid = students.reduce((sum, s) => sum + s.totalFee, 0);
  const totalLowIncome = students
    .filter(s => s.subsidyType !== '없음' && s.subsidyType !== '다자녀')
    .reduce((sum, s) => sum + s.subsidyAmount, 0);
  const totalMultiChild = students
    .filter(s => s.subsidyType === '다자녀')
    .reduce((sum, s) => sum + s.subsidyAmount, 0);

  const uniqueStudentsCount = students.filter(s => s.courseCount > 0).length;
  const totalEnrollments = students.reduce((sum, s) => sum + s.courseCount, 0);
  const totalInstructorCost = courses.reduce((sum, c) => sum + c.totalLecturerCost, 0);

  const lowIncomeCount = students.filter(s => s.subsidyType !== '없음' && s.subsidyType !== '다자녀' && s.courseCount > 0).length;
  const multiChildCount = students.filter(s => s.subsidyType === '다자녀' && s.courseCount > 0).length;
  const selfPaidCount = students.filter(s => s.totalFee > 0).length;

  // ==========================================
  // Sheet 1: 정산 총괄 요약
  // ==========================================
  const wsSummary = wb.addWorksheet('정산 총괄 요약', {
    views: [{ showGridLines: true }]
  });

  wsSummary.columns = [
    { width: 28 }, // 구분
    { width: 24 }, // 대상인원 / 수치
    { width: 24 }, // 금액 / 단위
    { width: 48 }  // 비고
  ];

  // Title
  wsSummary.addRow(['방과후학교 정산결과 총괄 요약 및 재원별 정산 합산']);
  const titleCell = wsSummary.getCell('A1');
  titleCell.font = fontTitle;
  wsSummary.getRow(1).height = 28;

  wsSummary.addRow([`(출력일자: ${new Date().toLocaleDateString('ko-KR')})`]);
  const dateCell = wsSummary.getCell('A2');
  dateCell.font = { name: '맑은 고딕', size: 9, italic: true, color: { argb: 'FF64748B' } };

  wsSummary.addRow([]); // Row 3 Blank

  // Section 1
  wsSummary.addRow(['[1. 재원에 따른 수강료 정산 합산]']);
  wsSummary.getCell('A4').font = fontSection;

  const summaryT1Headers = ['구분', '대상 학생 수(실인원)', '총 정산 금액(원)', '비고'];
  const r5 = wsSummary.addRow(summaryT1Headers);
  r5.height = 24;
  r5.eachCell((cell) => {
    cell.font = fontBold;
    cell.fill = fillHeader;
    cell.border = borderHeader;
    cell.alignment = alignCenter;
  });

  const summaryT1Data = [
    ['수강료 자부담 소계', `${selfPaidCount}명`, totalStudentPaid, '학생/학부모 자부담 총수납액'],
    ['저소득층 지원금 소계', `${lowIncomeCount}명`, totalLowIncome, '저소득층 복지 지원금 (전액 면제)'],
    ['다자녀 지원금 소계', `${multiChildCount}명`, totalMultiChild, '다자녀 학생 교육복지 지원 (전액 면제)'],
    ['합계 (정산 수강료 총액)', `${uniqueStudentsCount}명`, totalStudentPaid + totalLowIncome + totalMultiChild, '실수납 및 전체 지원금 총 집계 수강료']
  ];

  summaryT1Data.forEach((row, idx) => {
    const isTotal = idx === summaryT1Data.length - 1;
    const r = wsSummary.addRow(row);
    r.height = 22;
    r.getCell(1).alignment = alignCenter;
    r.getCell(2).alignment = alignCenter;
    r.getCell(3).alignment = alignRight;
    r.getCell(3).numFmt = '#,##0';
    r.getCell(4).alignment = alignLeft;

    r.eachCell((cell) => {
      cell.font = isTotal ? fontTotal : fontDefault;
      cell.border = isTotal ? borderTotal : borderThin;
      if (isTotal) cell.fill = fillTotal;
    });
  });

  wsSummary.addRow([]); // Blank

  // Section 2
  wsSummary.addRow(['[2. 방과후학교 개설 강좌 운영 총괄]']);
  const s2TitleRow = wsSummary.lastRow?.number || 11;
  wsSummary.getCell(`A${s2TitleRow}`).font = fontSection;

  const summaryT2Headers = ['구분', '수치 및 통계', '단위', '비고'];
  const rT2Header = wsSummary.addRow(summaryT2Headers);
  rT2Header.height = 24;
  rT2Header.eachCell((cell) => {
    cell.font = fontBold;
    cell.fill = fillHeader;
    cell.border = borderHeader;
    cell.alignment = alignCenter;
  });

  const summaryT2Data = [
    ['총 개설 강좌 수', courses.length, '개', '전체 개설된 방과후학교 강좌 수'],
    ['총 수강신청 건수', totalEnrollments, '건', '중복 신청건수를 합계한 총 수강 건수'],
    ['총 수강신청 학생 수', uniqueStudentsCount, '명', '수강 신청에 들어간 실제 고유 학생 수'],
    ['총 강사수당 소요 예산', totalInstructorCost, '원', '시수 일괄 합산에 의해 산정된 전체 강사료']
  ];

  summaryT2Data.forEach((row) => {
    const r = wsSummary.addRow(row);
    r.height = 22;
    r.getCell(1).alignment = alignCenter;
    r.getCell(2).alignment = typeof row[1] === 'number' ? alignRight : alignCenter;
    if (typeof row[1] === 'number') r.getCell(2).numFmt = '#,##0';
    r.getCell(3).alignment = alignCenter;
    r.getCell(4).alignment = alignLeft;
    r.eachCell((cell) => {
      cell.font = fontDefault;
      cell.border = borderThin;
    });
  });

  // Section 3: 수강 조정 및 정산 실제 삭제 학생 명단
  wsSummary.addRow([]);
  wsSummary.addRow(['[3. 수강 조정 및 정산 실제 삭제 학생 명단]']);
  const s3TitleRow = wsSummary.lastRow?.number || 18;
  wsSummary.getCell(`A${s3TitleRow}`).font = fontSection;

  wsSummary.addRow(["※ 해당 명단의 학생들은 '개별 정산 고지' 화면 상에서 수동으로 삭제처리된 학생들입니다."]);
  wsSummary.getCell(`A${s3TitleRow + 1}`).font = { name: '맑은 고딕', size: 9, color: { argb: 'FF64748B' } };
  wsSummary.addRow(["※ 원래 수강신청 정보에 있었으나 삭제됨으로써, 강좌별 단가/수강 학생 목록에서 완벽히 수동 차감 및 재산출되었습니다."]);
  wsSummary.getCell(`A${s3TitleRow + 2}`).font = { name: '맑은 고딕', size: 9, color: { argb: 'FF64748B' } };

  const s3Headers = ['학번', '이름', '신청했던 강좌 개수', '기존 수강 강좌명', '정산 영향 및 비고'];
  const rS3Header = wsSummary.addRow(s3Headers);
  rS3Header.height = 24;
  rS3Header.eachCell((cell) => {
    cell.font = fontBold;
    cell.fill = fillHeader;
    cell.border = borderHeader;
    cell.alignment = alignCenter;
  });

  if (excludedList && excludedList.length > 0) {
    excludedList.forEach(e => {
      const r = wsSummary.addRow([
        e.studentId,
        e.name,
        `${e.courseCount}개`,
        e.enrolledCourses.join(', '),
        '정산 대상에서 완벽히 수동 삭제 완료 (잔여 학생 수 기준 수강료 단가 자동 상승 재정산)'
      ]);
      r.height = 22;
      r.getCell(1).alignment = alignCenter;
      r.getCell(2).alignment = alignCenter;
      r.getCell(3).alignment = alignCenter;
      r.getCell(4).alignment = alignLeft;
      r.getCell(5).alignment = alignLeft;
      r.eachCell(cell => {
        cell.font = fontDefault;
        cell.border = borderThin;
      });
    });
  } else {
    const r = wsSummary.addRow([
      '-',
      '삭제된 학생 없음',
      '0개',
      '없음',
      '최초 업로드된 원본 수강신청 명단 시트 내용과 완전히 일치하게 정산되었습니다.'
    ]);
    r.height = 22;
    r.getCell(1).alignment = alignCenter;
    r.getCell(2).alignment = alignCenter;
    r.getCell(3).alignment = alignCenter;
    r.getCell(4).alignment = alignLeft;
    r.getCell(5).alignment = alignLeft;
    r.eachCell(cell => {
      cell.font = fontDefault;
      cell.border = borderThin;
    });
  }

  // ==========================================
  // Sheet 2: 학생별 수강료 정산 (통합)
  // ==========================================
  const wsFee = wb.addWorksheet('학생별 수강료 정산 (통합)', {
    views: [{ showGridLines: true }]
  });

  // Top Summary Box
  wsFee.addRow(['[재원별 수강료 정산 핵심요약]']);
  wsFee.getCell('A1').font = fontSection;

  const feeSummaryHeaders = ['구분', '총 금액 (원)', '대상 인원 수 (실인원)', '비고'];
  const rFeeSumHeader = wsFee.addRow(feeSummaryHeaders);
  rFeeSumHeader.height = 24;
  rFeeSumHeader.eachCell(cell => {
    cell.font = fontBold;
    cell.fill = fillHeader;
    cell.border = borderHeader;
    cell.alignment = alignCenter;
  });

  const feeSummaryRows = [
    ['수강료 자부담 소계', totalStudentPaid, `${selfPaidCount}명`, '학생 개인 실정산 수강료'],
    ['저소득층 지원금 소계', totalLowIncome, `${lowIncomeCount}명`, '저소득층/자유수강 지원 소계 (전액지원)'],
    ['다자녀 지원금 소계', totalMultiChild, `${multiChildCount}명`, '다자녀 복지 지원 소계 (전액지원)'],
    ['정산 수강료 합계', totalStudentPaid + totalLowIncome + totalMultiChild, `${uniqueStudentsCount}명`, '전체 수강료 실정산 대상']
  ];

  feeSummaryRows.forEach((row, idx) => {
    const isTotal = idx === feeSummaryRows.length - 1;
    const r = wsFee.addRow(row);
    r.height = 22;
    r.getCell(1).alignment = alignCenter;
    r.getCell(2).alignment = alignRight;
    r.getCell(2).numFmt = '#,##0';
    r.getCell(3).alignment = alignCenter;
    r.getCell(4).alignment = alignLeft;
    r.eachCell(cell => {
      cell.font = isTotal ? fontTotal : fontDefault;
      cell.border = isTotal ? borderTotal : borderThin;
      if (isTotal) cell.fill = fillTotal;
    });
  });

  wsFee.addRow([]); // Row 7 Blank

  // Main Settlement Table (Starting at Row 8)
  const mainFeeHeaders = [
    '학번',
    '이름',
    '지원비 대상(저소득층)',
    '지원비 대상(다자녀)',
    ...courses.map(c => c.name),
    '원래 수강료(원)',
    '지원금 수혜액(원)',
    '최종 납부액(원)',
    '비고'
  ];

  const rMainFeeHeader = wsFee.addRow(mainFeeHeaders);
  rMainFeeHeader.height = 26;
  rMainFeeHeader.eachCell(cell => {
    cell.font = fontBold;
    cell.fill = fillDarkHeader;
    cell.border = borderHeader;
    cell.alignment = alignCenter;
  });

  students.forEach((s, sIdx) => {
    const isMultiChild = s.subsidyType === '다자녀';
    const isLowIncome = s.subsidyType !== '없음' && s.subsidyType !== '다자녀';

    const rowData: any[] = [
      s.studentId,
      s.name,
      isLowIncome ? 1 : '',
      isMultiChild ? 1 : '',
      ...courses.map(c => s.enrolledCourses.includes(c.name) ? 1 : 0),
      s.originalFee,
      s.subsidyAmount,
      s.totalFee,
      s.subsidyType !== '없음' ? `${s.subsidyType} 전액지원` : ''
    ];

    const r = wsFee.addRow(rowData);
    r.height = 21;
    const isEven = sIdx % 2 === 1;

    r.eachCell((cell, colNumber) => {
      cell.font = fontDefault;
      cell.border = borderThin;
      if (isEven) cell.fill = fillZebra;

      // Formatting based on column
      if (colNumber === 1 || colNumber === 2 || colNumber === 3 || colNumber === 4) {
        cell.alignment = alignCenter;
      } else if (colNumber > 4 && colNumber <= 4 + courses.length) {
        cell.alignment = alignCenter;
      } else if (colNumber === 5 + courses.length || colNumber === 6 + courses.length || colNumber === 7 + courses.length) {
        cell.alignment = alignRight;
        cell.numFmt = '#,##0';
      } else {
        cell.alignment = alignCenter;
      }
    });
  });

  // Dynamic Column Widths for wsFee
  const feeColWidths: number[] = [12, 12, 22, 20];
  courses.forEach(c => feeColWidths.push(Math.max(14, c.name.length * 2.2)));
  feeColWidths.push(16, 16, 16, 24);

  feeColWidths.forEach((w, idx) => {
    wsFee.getColumn(idx + 1).width = w;
  });

  // Enable AutoFilter on Main Table
  if (students.length > 0) {
    const totalCols = mainFeeHeaders.length;
    const lastRow = 8 + students.length;
    wsFee.autoFilter = {
      from: { row: 8, column: 1 },
      to: { row: lastRow, column: totalCols }
    };
  }

  // ==========================================
  // Sheet 3: 학생별 징수금액
  // ==========================================
  const sortedStudents = [...students].sort((a, b) => a.studentId.localeCompare(b.studentId));
  const studentCount = sortedStudents.length;
  const lastDataRow = 5 + studentCount;

  const totalPreservation = students
    .filter(s => s.subsidyType !== '없음' && s.subsidyType !== '다자녀' && !s.subsidyType.includes('자유수강') && !s.subsidyType.includes('저소득') && !s.subsidyType.includes('수급자') && !s.subsidyType.includes('한부모') && !s.subsidyType.includes('법정') && s.subsidyType !== '자유수강권')
    .reduce((sum, s) => sum + s.subsidyAmount, 0);

  const wsCollection = wb.addWorksheet('학생별 징수금액', {
    views: [{ showGridLines: true }]
  });

  wsCollection.columns = [
    { width: 14 }, // A: 학번
    { width: 18 }, // B: 이름 / 차액명
    { width: 24 }, // C: 자유수강권(저소득층)
    { width: 22 }, // D: 자유수강권(다자녀)
    { width: 24 }, // E: 수익자 대상 징수금액
    { width: 34 }, // F: 수익자 대상 징수금액(10원 단위 절하)
    { width: 20 }  // G: 합계(상단용)
  ];

  // Top Summary Table (Rows 1-2)
  const r1 = wsCollection.getRow(1);
  r1.values = ['', '', '자유수강권(저소득층) 총액', '자유수강권(다자녀) 총액', '수익자 대상 징수금액', '방과후학교 보전금', '합계'];
  r1.height = 24;
  [3, 4, 5, 6, 7].forEach(col => {
    const cell = r1.getCell(col);
    cell.font = fontBold;
    cell.fill = fillHeader;
    cell.border = borderHeader;
    cell.alignment = alignCenter;
  });

  const r2 = wsCollection.getRow(2);
  r2.height = 26;
  r2.getCell(3).value = studentCount > 0 ? { formula: `SUM(C6:C${lastDataRow})`, result: totalLowIncome } : 0;
  r2.getCell(4).value = studentCount > 0 ? { formula: `SUM(D6:D${lastDataRow})`, result: totalMultiChild } : 0;
  r2.getCell(5).value = studentCount > 0 ? { formula: `SUM(E6:E${lastDataRow})`, result: totalStudentPaid } : 0;
  r2.getCell(6).value = totalPreservation;
  r2.getCell(7).value = { formula: `SUM(C2:F2)`, result: totalLowIncome + totalMultiChild + totalStudentPaid + totalPreservation };

  [3, 4, 5, 6, 7].forEach(col => {
    const cell = r2.getCell(col);
    cell.font = col === 7 ? fontTotal : fontBold;
    cell.alignment = alignRight;
    cell.numFmt = '#,##0';
    cell.border = col === 7 ? borderTotal : borderThin;
    if (col === 7) cell.fill = fillTotal;
  });

  // Title at Row 3
  const r3 = wsCollection.getRow(3);
  r3.height = 32;
  wsCollection.mergeCells('A3:F3');
  const collTitleCell = wsCollection.getCell('A3');
  collTitleCell.value = '방과후학교 및 교과보충 집중 프로그램 학생별 징수 금액';
  collTitleCell.font = fontTitle;
  collTitleCell.alignment = alignCenter;

  // Row 4: Blank
  wsCollection.getRow(4).height = 12;

  // Row 5: Table Header
  const r5Coll = wsCollection.getRow(5);
  r5Coll.values = ['학번', '이름', '자유수강권(저소득층)', '자유수강권(다자녀)', '수익자 대상 징수금액', '수익자 대상 징수금액(10원 단위 절하)'];
  r5Coll.height = 26;
  [1, 2, 3, 4, 5, 6].forEach(col => {
    const cell = r5Coll.getCell(col);
    cell.font = fontBold;
    cell.fill = fillDarkHeader;
    cell.border = borderHeader;
    cell.alignment = alignCenter;
  });

  // Student Data Rows
  sortedStudents.forEach((s, idx) => {
    const rowNum = 6 + idx;
    const isMultiChild = s.subsidyType === '다자녀';
    const isLowIncome = s.subsidyType !== '없음' && s.subsidyType !== '다자녀' && (
      s.subsidyType.includes('자유수강') || 
      s.subsidyType.includes('저소득') || 
      s.subsidyType.includes('수급자') || 
      s.subsidyType.includes('한부모') || 
      s.subsidyType.includes('법정') ||
      s.subsidyType === '자유수강권'
    );

    const lowIncomeVal = isLowIncome ? s.subsidyAmount : null;
    const multiChildVal = isMultiChild ? s.subsidyAmount : null;
    const payerVal = s.totalFee > 0 ? s.totalFee : null;
    const payerValTruncated = s.totalFee > 0 ? Math.floor(s.totalFee / 10) * 10 : null;

    const r = wsCollection.getRow(rowNum);
    r.height = 21;
    r.values = [
      s.studentId,
      s.name,
      lowIncomeVal,
      multiChildVal,
      payerVal,
      payerValTruncated
    ];

    const isEven = idx % 2 === 1;
    [1, 2, 3, 4, 5, 6].forEach(col => {
      const cell = r.getCell(col);
      cell.font = fontDefault;
      cell.border = borderThin;
      if (isEven) cell.fill = fillZebra;

      if (col === 1 || col === 2) {
        cell.alignment = alignCenter;
      } else {
        cell.alignment = alignRight;
        cell.numFmt = '#,##0';
      }
    });
  });

  // 10원 단위 절하 차액 Row
  const diffRowNumber = 6 + studentCount;
  const rDiff = wsCollection.getRow(diffRowNumber);
  rDiff.height = 24;
  rDiff.getCell(2).value = '10원 단위 절하 차액';
  rDiff.getCell(2).font = fontDiff;
  rDiff.getCell(2).alignment = alignCenter;

  const truncatedSum = sortedStudents.reduce((sum, s) => sum + (s.totalFee > 0 ? Math.floor(s.totalFee / 10) * 10 : 0), 0);
  const diffAmount = totalStudentPaid - truncatedSum;

  rDiff.getCell(6).value = studentCount > 0 ? { formula: `SUM(E6:E${lastDataRow})-SUM(F6:F${lastDataRow})`, result: diffAmount } : 0;
  rDiff.getCell(6).font = fontDiff;
  rDiff.getCell(6).alignment = alignRight;
  rDiff.getCell(6).numFmt = '#,##0';
  rDiff.getCell(6).fill = fillDiff;

  [1, 2, 3, 4, 5, 6].forEach(col => {
    rDiff.getCell(col).border = borderThin;
  });

  // Total Row at Bottom
  const totalRowNumber = 7 + studentCount;
  const rTotal = wsCollection.getRow(totalRowNumber);
  rTotal.height = 26;
  rTotal.getCell(2).value = '합계';
  rTotal.getCell(2).alignment = alignCenter;

  rTotal.getCell(3).value = studentCount > 0 ? { formula: `SUM(C6:C${lastDataRow})`, result: totalLowIncome } : 0;
  rTotal.getCell(4).value = studentCount > 0 ? { formula: `SUM(D6:D${lastDataRow})`, result: totalMultiChild } : 0;
  rTotal.getCell(5).value = studentCount > 0 ? { formula: `SUM(E6:E${lastDataRow})`, result: totalStudentPaid } : 0;
  rTotal.getCell(6).value = studentCount > 0 ? { formula: `SUM(F6:F${diffRowNumber})`, result: totalStudentPaid } : 0;

  [1, 2, 3, 4, 5, 6].forEach(col => {
    const cell = rTotal.getCell(col);
    cell.font = fontTotal;
    cell.border = borderTotal;
    cell.fill = fillTotal;
    if (col >= 3) {
      cell.alignment = alignRight;
      cell.numFmt = '#,##0';
    }
  });

  // Autofilter for student rows in wsCollection
  if (studentCount > 0) {
    wsCollection.autoFilter = {
      from: { row: 5, column: 1 },
      to: { row: lastDataRow, column: 6 }
    };
  }

  // Export and download workbook
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  
  const datestr = new Date().toISOString().substring(0, 10).replace(/-/g, '');
  const a = document.createElement('a');
  a.href = url;
  a.download = `방과후학교_정산결과_${datestr}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
