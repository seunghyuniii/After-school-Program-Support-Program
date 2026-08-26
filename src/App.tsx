/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Building, 
  Upload, 
  Download, 
  Search, 
  Settings, 
  Users, 
  CheckSquare, 
  RefreshCw, 
  AlertCircle, 
  FileSpreadsheet, 
  BookOpen, 
  Coins, 
  FileUp, 
  Info, 
  Trash2, 
  CheckCircle,
  Clock,
  HelpCircle,
  ChevronRight
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { StudentRow } from './types';
import { 
  generateSampleExcel, 
  parseWorkbook, 
  calculateSettlement, 
  downloadSettlementExcel,
  generateSampleSubsidyExcel,
  parseSubsidyWorkbook,
  normalizeStudentId
} from './utils';

export default function App() {
  // File upload state
  const [rawData, setRawData] = useState<StudentRow[]>([]);
  const [courseColumns, setCourseColumns] = useState<string[]>([]);
  const [fileName, setFileName] = useState<string | null>(null);
  const [subsidyFileName, setSubsidyFileName] = useState<string | null>(null);
  
  // Interactive configurations
  const [feePerHour, setFeePerHour] = useState<number>(33000);
  const [courseHours, setCourseHours] = useState<Record<string, number>>({});
  const [useTruncate10, setUseTruncate10] = useState<boolean>(false);
  const [beneficiaryCourses, setBeneficiaryCourses] = useState<Record<string, boolean>>({});
  const [subsidies, setSubsidies] = useState<Record<string, { type: string; amount: number }>>({});
  const [excludedStudentIds, setExcludedStudentIds] = useState<string[]>([]);
  const [editingStudentId, setEditingStudentId] = useState<string | null>(null);
  const [editingStudentName, setEditingStudentName] = useState<string>('');
  const [editingStudentIdNumber, setEditingStudentIdNumber] = useState<string>('');
  
  // Dialog input states for subsidy modal
  const [dialogType, setDialogType] = useState<string>('없음');
  const [dialogAmount, setDialogAmount] = useState<number>(0);
  
  // UI states
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [isDraggingSubsidy, setIsDraggingSubsidy] = useState<boolean>(false);
  const [activePreviewTab, setActivePreviewTab] = useState<'courses' | 'students' | 'attendance'>('courses');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [selectedAttendanceCourse, setSelectedAttendanceCourse] = useState<string>('');
  const [batchHoursVal, setBatchHoursVal] = useState<number>(20);

  // Drag and Drop handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0]);
    }
  };

  const handleFile = (file: File) => {
    if (!file) return;
    setError(null);
    setLoading(true);
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellText: false, cellDates: true });
        
        const parsed = parseWorkbook(workbook);
        setRawData(parsed.rawData);
        setCourseColumns(parsed.courseColumns);
        
        // Initialize default hours and beneficiary-pays boolean per course
        const initialHours: Record<string, number> = {};
        const initialBeneficiary: Record<string, boolean> = {};
        parsed.courseColumns.forEach(col => {
          initialHours[col] = 20; // Default standard hours duration
          initialBeneficiary[col] = true; // 수익자 부담이 기본값
        });
        setCourseHours(initialHours);
        setBeneficiaryCourses(initialBeneficiary);
        setSubsidies({}); // Reset subsidies on new file upload
        setExcludedStudentIds([]); // Reset excluded students on new file upload
        setSubsidyFileName(null);
        
        if (parsed.courseColumns.length > 0) {
          setSelectedAttendanceCourse(parsed.courseColumns[0]);
        }
      } catch (err: any) {
        console.error(err);
        setError(err.message || '엑셀 파일을 파싱하는 중 오류가 발생했습니다. 올바른 방과후 신청 양식인지 확인해 주세요.');
        setRawData([]);
        setCourseColumns([]);
        setFileName(null);
      } finally {
        setLoading(false);
      }
    };
    reader.onerror = () => {
      setError('파일을 읽는 도중 오류가 발생했습니다.');
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  // Subsidy Drag and Drop handlers
  const handleSubsidyDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingSubsidy(true);
  };

  const handleSubsidyDragLeave = () => {
    setIsDraggingSubsidy(false);
  };

  const handleSubsidyDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingSubsidy(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleSubsidyFile(e.dataTransfer.files[0]);
    }
  };

  const handleSubsidyFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      handleSubsidyFile(e.target.files[0]);
    }
  };

  const handleSubsidyFile = (file: File) => {
    if (!file) return;
    setError(null);
    setLoading(true);
    setSubsidyFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellText: false, cellDates: true });
        
        const parsedSubsidies = parseSubsidyWorkbook(workbook);
        
        setSubsidies((prev) => {
          // Merge with mock or exist, or fully replace with newly uploaded Excel
          return {
            ...prev,
            ...parsedSubsidies
          };
        });
      } catch (err: any) {
        console.error(err);
        setError(err.message || '지원금 엑셀 파일을 파싱하는 중 오류가 발생했습니다. 올바른 지원금 양식인지 확인해 주세요.');
        setSubsidyFileName(null);
      } finally {
        setLoading(false);
      }
    };
    reader.onerror = () => {
      setError('지원금 파일을 읽는 도중 오류가 발생했습니다.');
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  // 1-Click test data loading
  const handleLoadDemo = () => {
    setError(null);
    setLoading(true);
    setTimeout(() => {
      try {
        setFileName('리로스쿨_방과후_정산용_데모파일_2026.xlsx');
        
        // Setup raw mock Liro school student matrix
        const mockRows: StudentRow[] = [
          { '순번': 1, '학번': 10101, '이름': '강유민', '강좌수': 2, '유료강좌': 2, 'AI 웹서비스 풀스택 실습': 'AI 웹서비스 풀스택 실습', '독서와 논구술 창의융합': '독서와 논구술 창의융합', '수리논술 핵심 정리': '', '원어민 수능 고급영어': '', '화학 반응 탐구 교실': '' },
          { '순번': 2, '학번': 10115, '이름': '고상우', '강좌수': 1, '유료강좌': 1, 'AI 웹서비스 풀스택 실습': '', '독서와 논구술 창의융합': '', '수리논술 핵심 정리': '수리논술 핵심 정리', '원어민 수능 고급영어': '', '화학 반응 탐구 교실': '' },
          { '순번': 3, '학번': 10204, '이름': '김도윤', '강좌수': 2, '유료강좌': 2, 'AI 웹서비스 풀스택 실습': '', '독서와 논구술 창의융합': '독서와 논구술 창의융합', '수리논술 핵심 정리': '', '원어민 수능 고급영어': '원어민 수능 고급영어', '화학 반응 탐구 교실': '' },
          { '순번': 4, '학번': 10212, '이름': '남궁혜원', '강좌수': 3, '유료강좌': 3, 'AI 웹서비스 풀스택 실습': 'AI 웹서비스 풀스택 실습', '독서와 논구술 창의융합': '', '수리논술 핵심 정리': '수리논술 핵심 정리', '원어민 수능 고급영어': '', '화학 반응 탐구 교실': '화학 반응 탐구 교실' },
          { '순번': 5, '학번': 10303, '이름': '박건율', '강좌수': 1, '유료강좌': 1, 'AI 웹서비스 풀스택 실습': 'AI 웹서비스 풀스택 실습', '독서와 논구술 창의융합': '', '수리논술 핵심 정리': '', '원어민 수능 고급영어': '', '화학 반응 탐구 교실': '' },
          { '순번': 6, '학번': 10318, '이름': '서영주', '강좌수': 2, '유료강좌': 2, 'AI 웹서비스 풀스택 실습': '', '독서와 논구술 창의융합': '독서와 논구술 창의융합', '수리논술 핵심 정리': '', '원어민 수능 고급영어': '', '화학 반응 탐구 교실': '화학 반응 탐구 교실' },
          { '순번': 7, '학번': 10402, '이름': '윤다솜', '강좌수': 1, '유료강좌': 1, 'AI 웹서비스 풀스택 실습': '', '독서와 논구술 창의융합': '', '수리논술 핵심 정리': '', '원어민 수능 고급영어': '원어민 수능 고급영어', '화학 반응 탐구 교실': '' },
          { '순번': 8, '학번': 10421, '이름': '이하경', '강좌수': 4, '유료강좌': 4, 'AI 웹서비스 풀스택 실습': 'AI 웹서비스 풀스택 실습', '독서와 논구술 창의융합': '독서와 논구술 창의융합', '수리논술 핵심 정리': '수리논술 핵심 정리', '원어민 수능 고급영어': '', '화학 반응 탐구 교실': '화학 반응 탐구 교실' },
          { '순번': 9, '학번': 10505, '이름': '정우진', '강좌수': 2, '유료강좌': 2, 'AI 웹서비스 풀스택 실습': '', '독서와 논구술 창의융합': '', '수리논술 핵심 정리': '수리논술 핵심 정리', '원어민 수능 고급영어': '원어민 수능 고급영어', '화학 반응 탐구 교실': '' },
          { '순번': 10, '학번': 10519, '이름': '최유은', '강좌수': 2, '유료강좌': 2, 'AI 웹서비스 풀스택 실습': 'AI 웹서비스 풀스택 실습', '독서와 논구술 창의융합': '', '수리논술 핵심 정리': '', '원어민 수능 고급영어': '', '화학 반응 탐구 교실': '화학 반응 탐구 교실' },
          { '순번': 11, '학번': 10611, '이름': '한지원', '강좌수': 2, '유료강좌': 2, 'AI 웹서비스 풀스택 실습': '', '독서와 논구술 창의융합': '독서와 논구술 창의융합', '수리논술 핵심 정리': '', '원어민 수능 고급영어': '원어민 수능 고급영어', '화학 반응 탐구 교실': '' },
          { '순번': 12, '학번': 10624, '이름': '황보명', '강좌수': 0, '유료강좌': 0, 'AI 웹서비스 풀스택 실습': '', '독서와 논구술 창의융합': '', '수리논술 핵심 정리': '', '원어민 수능 고급영어': '', '화학 반응 탐구 교실': '' }
        ];

        const mockCols = [
          'AI 웹서비스 풀스택 실습',
          '독서와 논구술 창의융합',
          '수리논술 핵심 정리',
          '원어민 수능 고급영어',
          '화학 반응 탐구 교실'
        ];

        setRawData(mockRows);
        setCourseColumns(mockCols);

        const initialHours: Record<string, number> = {
          'AI 웹서비스 풀스택 실습': 32,
          '독서와 논구술 창의융합': 16,
          '수리논술 핵심 정리': 20,
          '원어민 수능 고급영어': 20,
          '화학 반응 탐구 교실': 16
        };
        
        const initialBeneficiary: Record<string, boolean> = {};
        mockCols.forEach(col => {
          initialBeneficiary[col] = true; // Default all as beneficiary
        });

        // Pre-populate some demo student subsidies to showcase the input feature
        const initialSubsidies: Record<string, { type: string; amount: number }> = {
          '10101': { type: '저소득층', amount: 0 },
          '10204': { type: '다자녀', amount: 0 }
        };

        setCourseHours(initialHours);
        setBeneficiaryCourses(initialBeneficiary);
        setSubsidies(initialSubsidies);
        setExcludedStudentIds([]); // Reset excluded students on demo load
        setSelectedAttendanceCourse(mockCols[0]);
      } catch (err) {
        setError('데모 데이터를 준비 중 에러가 발생했습니다.');
      } finally {
        setLoading(false);
      }
    }, 500);
  };

  // Exclude students selected by user
  const activeRawData = useMemo(() => {
    return rawData.filter((row) => {
      const id = normalizeStudentId(row['학번']);
      return !excludedStudentIds.includes(id);
    });
  }, [rawData, excludedStudentIds]);

  // Memoized settlement calculations based on configurations
  const calculated = useMemo(() => {
    return calculateSettlement(
      activeRawData, 
      courseColumns, 
      courseHours, 
      feePerHour, 
      useTruncate10, 
      beneficiaryCourses, 
      subsidies
    );
  }, [activeRawData, courseColumns, courseHours, feePerHour, useTruncate10, beneficiaryCourses, subsidies]);

  // Overall calculations budget summaries split into all course types & beneficiary-pays types
  const summaries = useMemo(() => {
    // 1. Overall summary
    // Unique students registered for at least one course
    const allUniqueStudents = calculated.students.filter(s => s.courseCount > 0).length;
    const allCoursesCount = courseColumns.length;
    const allInstructorCost = calculated.courses.reduce((sum, c) => sum + c.totalLecturerCost, 0);
    const allStudentPaid = calculated.students.reduce((sum, s) => sum + s.totalFee, 0);

    // 2. Beneficiary (수익자 부담) summary
    const beneficiaryUniqueStudents = calculated.students.filter(s => 
      s.enrolledCourses.some(course => beneficiaryCourses[course] !== false)
    ).length;
    const beneficiaryCoursesCount = courseColumns.filter(col => beneficiaryCourses[col] !== false).length;
    const beneficiaryInstructorCost = calculated.courses
      .filter(c => beneficiaryCourses[c.name] !== false)
      .reduce((sum, c) => sum + c.totalLecturerCost, 0);

    // Calculate Low-income and Multiple-children subsidies totals
    const lowIncomeSubsidyTotal = calculated.students
      .filter(s => s.subsidyType !== '없음' && s.subsidyType !== '' && s.subsidyType !== '다자녀')
      .reduce((sum, s) => sum + s.subsidyAmount, 0);

    const multiChildSubsidyTotal = calculated.students
      .filter(s => s.subsidyType === '다자녀')
      .reduce((sum, s) => sum + s.subsidyAmount, 0);

    return {
      all: {
        uniqueStudents: allUniqueStudents,
        coursesCount: allCoursesCount,
        instructorCost: allInstructorCost,
        studentPaid: allStudentPaid
      },
      beneficiary: {
        uniqueStudents: beneficiaryUniqueStudents,
        coursesCount: beneficiaryCoursesCount,
        instructorCost: beneficiaryInstructorCost,
        studentPaid: allStudentPaid,
        lowIncomeSubsidyTotal,
        multiChildSubsidyTotal
      }
    };
  }, [courseColumns, calculated, beneficiaryCourses]);

  // Handler to adjust specific class hours
  const handleHoursChange = (courseName: string, hours: number) => {
    setCourseHours((prev) => ({
      ...prev,
      [courseName]: hours,
    }));
  };

  // Delete non-beneficiary or unwanted course from settlement
  const handleDeleteCourse = (courseName: string) => {
    setCourseColumns((prev) => prev.filter((col) => col !== courseName));
    setSelectedAttendanceCourse((prev) => 
      prev === courseName ? (courseColumns.find((c) => c !== courseName) || '') : prev
    );
  };

  // Apply batch hours to all classes
  const handleApplyBatchHours = () => {
    const updated: Record<string, number> = {};
    courseColumns.forEach((col) => {
      updated[col] = batchHoursVal;
    });
    setCourseHours(updated);
  };

  // Open subsidy management modal
  const openSubsidyModal = (stdId: string, name: string, originalFee: number) => {
    setEditingStudentId(stdId);
    setEditingStudentName(name);
    setEditingStudentIdNumber(stdId);
    
    const existing = subsidies[stdId] || { type: '없음', amount: 0 };
    setDialogType(existing.type || '없음');
    setDialogAmount(existing.amount || 0);
  };

  // Save subsidy from modal values
  const saveSubsidy = () => {
    if (!editingStudentId) return;
    setSubsidies((prev) => ({
      ...prev,
      [editingStudentId]: {
        type: dialogType,
        amount: dialogType === '없음' ? 0 : Math.max(0, dialogAmount)
      }
    }));
    setEditingStudentId(null);
  };

  // Reset to initial clean state
  const handleClear = () => {
    setRawData([]);
    setCourseColumns([]);
    setFileName(null);
    setSubsidyFileName(null);
    setCourseHours({});
    setBeneficiaryCourses({});
    setSubsidies({});
    setExcludedStudentIds([]);
    setSearchQuery('');
  };

  // Trigger spreadsheet file download
  const handleExport = async () => {
    if (calculated.courses.length === 0) return;

    // Construct details of excluded students
    const excludedList = excludedStudentIds.map(id => {
      const row = rawData.find(r => normalizeStudentId(r['학번']) === id);
      const name = row ? String(row['이름'] || '').trim() : '알 수 없음';

      const enrolledCourses: string[] = [];
      if (row) {
        courseColumns.forEach(col => {
          const val = row[col];
          if (val !== null && val !== undefined && String(val).trim() !== "") {
            enrolledCourses.push(col);
          }
        });
      }

      return {
        studentId: id,
        name,
        courseCount: enrolledCourses.length,
        enrolledCourses,
      };
    });

    await downloadSettlementExcel(
      calculated.courses,
      calculated.students,
      calculated.attendance,
      feePerHour,
      excludedList
    );
  };

  // Compute filtered students search lists
  const filteredStudents = useMemo(() => {
    if (!searchQuery.trim()) return calculated.students;
    const query = searchQuery.toLowerCase().trim();
    return calculated.students.filter(
      (std) =>
        std.name.toLowerCase().includes(query) ||
        std.studentId.includes(query)
    );
  }, [calculated.students, searchQuery]);

  // Current active attendance rosters selector
  const currentAttendanceRecords = useMemo(() => {
    if (!selectedAttendanceCourse) return [];
    const record = calculated.attendance.find(
      (a) => a.courseName === selectedAttendanceCourse
    );
    return record ? record.students : [];
  }, [calculated.attendance, selectedAttendanceCourse]);

  return (
    <div className="min-h-screen lg:h-screen bg-[#F8FAFC] font-sans text-[#1E293B] antialiased flex flex-col selection:bg-blue-500 selection:text-white lg:overflow-hidden">
      
      {/* Premium Elegant Header */}
      <header className="bg-white border-b border-[#E2E8F0] sticky top-0 z-30 shadow-[0_1px_2px_rgba(0,0,0,0.02)] flex-shrink-0">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-[#2563EB] rounded-lg text-white shadow-sm shadow-[#2563EB]/10">
              <Building className="w-5.5 h-5.5" id="app-logo-icon" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-medium px-2.5 py-0.5 bg-slate-50 text-slate-600 border border-[#E2E8F0] rounded-full flex items-center gap-1">
                  <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full inline-block animate-pulse"></span>
                  안전한 오프라인 전용 계산 (서버 전송 없음)
                </span>
              </div>
              <h1 className="text-lg sm:text-xl font-bold text-[#1E293B] tracking-tight mt-0.5">
                방과후학교정산 자동화 시스템
              </h1>
            </div>
          </div>
          
          <div className="flex items-center gap-2 w-full sm:w-auto">
            <button
              onClick={generateSampleExcel}
              className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-3 py-2 text-xs font-semibold text-slate-700 bg-white border border-[#E2E8F0] rounded-lg hover:bg-slate-50 transition shadow-xs cursor-pointer"
              title="리로스쿨 로드용 빈 예제 파일을 다운로드하여 구조를 확인하세요."
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-[#2563EB]" />
              양식 예제 다운로드
            </button>
            
            {rawData.length > 0 && (
              <button
                onClick={handleClear}
                className="flex items-center gap-1.5 px-3 py-2 text-xs font-semibold text-rose-600 bg-rose-50/50 border border-rose-100 rounded-lg hover:bg-rose-100/80 transition active:scale-95 cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                목록 초기화
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="flex-1 min-h-0 lg:h-[calc(100vh-74px)] max-w-7xl mx-auto w-full px-4 sm:px-6 lg:px-8 py-4 flex flex-col lg:flex-row gap-5 lg:overflow-hidden">
        
        {/* Left Side: Upload & Core Constant Settings */}
        <section className="w-full lg:w-1/3 flex flex-col gap-4 lg:overflow-y-auto lg:pr-1.5 scrollbar-thin shrink-0 pb-6 text-left">
          
          {/* 📂 데이터 파일 등록 센터 */}
          <div className="bg-white p-4 rounded-xl border border-[#E2E8F0] shadow-[0_1px_3px_rgba(0,0,0,0.02)]">
            <div className="flex justify-between items-center mb-2.5">
              <h3 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2">
                <span className="w-5 h-5 rounded-full bg-slate-100 text-[#2563EB] flex items-center justify-center text-[10px] font-black">1</span>
                엑셀 파일 입력 대장
              </h3>
              <button
                type="button"
                onClick={generateSampleSubsidyExcel}
                className="text-[10px] text-blue-600 font-extrabold flex items-center gap-0.5 hover:underline cursor-pointer"
                title="지원금 업로드용 템플릿 양식을 다운로드하세요."
              >
                📥 지원금 양식
              </button>
            </div>

            <div className="grid grid-cols-2 gap-2.5">
              {/* Box 1: Liro School Core Enrollment */}
              <div
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                className={`border border-dashed rounded-lg p-2.5 text-center transition cursor-pointer flex flex-col items-center justify-center min-h-[92px] ${
                  isDragging 
                    ? 'border-[#2563EB] bg-[#EFF6FF] text-[#2563EB]' 
                    : rawData.length > 0 
                      ? 'border-emerald-300 bg-emerald-50/10 text-slate-700 hover:border-emerald-400' 
                      : 'border-[#E2E8F0] bg-[#F1F5F9]/30 text-slate-500 hover:border-[#2563EB] hover:bg-[#EFF6FF]/40'
                }`}
                onClick={() => document.getElementById('excel-file-hidden')?.click()}
              >
                <input 
                  type="file" 
                  id="excel-file-hidden" 
                  className="hidden" 
                  accept=".xlsx, .xls, .csv" 
                  onChange={handleFileChange}
                />
                
                {rawData.length > 0 ? (
                  <div className="flex flex-col items-center leading-normal">
                    <CheckCircle className="w-4 h-4 text-emerald-600 mb-1 animate-pulse" />
                    <p className="text-[10px] font-black text-emerald-800 uppercase tracking-tight">수강신청 로드 완료</p>
                    <p className="text-[9px] text-slate-500 break-all max-w-[110px] truncate mt-0.5">
                      {fileName}
                    </p>
                  </div>
                ) : (
                  <div className="flex flex-col items-center leading-none">
                    <FileUp className="w-4.5 h-4.5 text-[#2563EB] mb-1.5" />
                    <p className="text-[10px] font-black text-[#1E293B]">수강대장 엑셀</p>
                    <p className="text-[8.5px] text-[#64748B] mt-1.5">드래그 또는 클릭</p>
                  </div>
                )}
              </div>

              {/* Box 2: Subsidy Setup */}
              <div
                onDragOver={handleSubsidyDragOver}
                onDragLeave={handleSubsidyDragLeave}
                onDrop={handleSubsidyDrop}
                className={`border border-dashed rounded-lg p-2.5 text-center transition cursor-pointer flex flex-col items-center justify-center min-h-[92px] ${
                  isDraggingSubsidy
                    ? 'border-[#2563EB] bg-[#EFF6FF] text-[#2563EB]'
                    : subsidyFileName
                      ? 'border-emerald-300 bg-emerald-50/10 text-slate-700 hover:border-emerald-400'
                      : 'border-[#E2E8F0] bg-[#F1F5F9]/30 text-slate-500 hover:border-[#2563EB] hover:bg-[#EFF6FF]/40'
                }`}
                onClick={() => document.getElementById('subsidy-file-hidden')?.click()}
              >
                <input 
                  type="file" 
                  id="subsidy-file-hidden" 
                  className="hidden" 
                  accept=".xlsx, .xls, .csv" 
                  onChange={handleSubsidyFileChange}
                />
                
                {subsidyFileName ? (
                  <div className="flex flex-col items-center leading-normal">
                    <CheckCircle className="w-4 h-4 text-emerald-600 mb-1" />
                    <p className="text-[10px] font-black text-emerald-800 uppercase tracking-tight">지원명부 로드 완료</p>
                    <p className="text-[9px] text-slate-500 break-all max-w-[110px] truncate mt-0.5">
                      {subsidyFileName}
                    </p>
                  </div>
                ) : (
                  <div className="flex flex-col items-center leading-none">
                    <Coins className="w-4.5 h-4.5 text-amber-500 mb-1.5" />
                    <p className="text-[10px] font-black text-[#1E293B]">지원금 명부 엑셀</p>
                    <p className="text-[8.5px] text-[#64748B] mt-1.5">드래그 또는 선택</p>
                  </div>
                )}
              </div>
            </div>

            {/* Quick Experience sandbox trigger */}
            {rawData.length === 0 && (
              <div className="mt-2.5">
                <button
                  type="button"
                  onClick={handleLoadDemo}
                  className="w-full py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-lg text-xs font-semibold transition shadow-xs cursor-pointer flex items-center justify-center gap-1.5"
                >
                  <span className="relative flex h-1.5 w-1.5">
                    <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-blue-400 opacity-75"></span>
                    <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-blue-500"></span>
                  </span>
                  가상 수강 데모 데이터 로드하기
                </button>
              </div>
            )}
          </div>

          {/* Step 2: Constant values and base rates */}
          <div className="bg-white p-4 rounded-xl border border-[#E2E8F0] shadow-[0_1px_3px_rgba(0,0,0,0.02)] flex flex-col gap-3.5 text-left">
            <h3 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2">
              <span className="w-5 h-5 rounded-full bg-slate-100 text-[#2563EB] flex items-center justify-center text-[10px] font-black">2</span>
              정산 기본 요율 구성
            </h3>

            {/* Price-per-hour instructor salary input */}
            <div className="space-y-1.5">
              <label className="text-xs font-semibold text-[#64748B] flex justify-between items-center">
                <span>👨‍🏫 시간당 기본 강사료 단가</span>
                <span className="text-[10px] font-normal text-[#64748B]">
                  (단위: 원)
                </span>
              </label>
              <div className="flex items-center gap-1">
                <span className="p-2 px-3 bg-[#F1F5F9] border border-r-0 border-[#E2E8F0] rounded-l-lg text-[#64748B] text-xs font-medium">
                  ₩
                </span>
                <input
                  type="number"
                  value={feePerHour}
                  onChange={(e) => setFeePerHour(Math.max(0, parseInt(e.target.value) || 0))}
                  className="w-full p-1.5 px-3 text-xs font-bold border border-[#E2E8F0] rounded-r-lg focus:outline-none focus:border-[#2563EB] text-[#1E293B]"
                  step="1000"
                  placeholder="예: 40000"
                />
              </div>
              
              {/* Preset buttons */}
              <div className="grid grid-cols-4 gap-1 mt-1.5">
                {[33000, 35000, 43000, 45000].map((preset) => (
                  <button
                    key={preset}
                    onClick={() => setFeePerHour(preset)}
                    className={`py-1 rounded text-[10px] font-medium transition cursor-pointer ${
                      feePerHour === preset 
                        ? 'bg-[#2563EB] text-white font-bold' 
                        : 'bg-[#F1F5F9] text-slate-600 border border-[#E2E8F0] hover:bg-slate-200/50'
                    }`}
                  >
                    {preset.toLocaleString()}원
                  </button>
                ))}
              </div>
            </div>

            {/* 환경설정: 절하 단위 선택 */}
            <div className="border-t border-[#E2E8F0] pt-3 flex flex-col gap-1.5">
              <label className="text-[11px] font-semibold text-[#64748B] flex justify-between items-center">
                <span>⚙️ 단수 계산 방식 환경설정</span>
              </label>
              <div 
                className="flex items-start gap-2.5 p-2.5 bg-slate-50 border border-[#E2E8F0] rounded-lg cursor-pointer hover:bg-slate-100/70 transition"
                onClick={() => setUseTruncate10(!useTruncate10)}
              >
                <input
                  type="checkbox"
                  checked={useTruncate10}
                  onChange={(e) => setUseTruncate10(e.target.checked)}
                  className="mt-0.5 w-4 h-4 text-[#2563EB] border-slate-300 rounded focus:ring-blue-500 cursor-pointer"
                  onClick={(e) => e.stopPropagation()} // Prevent double trigger
                />
                <div className="text-left select-none">
                  <span className="text-[11px] font-bold text-slate-800 block">
                    {useTruncate10 ? '10원 단위 절하(버림) 적용' : '1원 단위 최종 세부 산정'}
                  </span>
                  <span className="text-[9.5px] text-slate-500 block leading-tight mt-0.5">
                    {useTruncate10 
                      ? '수강생 단가에서 10원 미만 원단위 금액을 절사(버림) 정산합니다.' 
                      : '원단위 금액을 버리지 않고 1원 단위까지 정산 고지서에 상세 반영합니다.'}
                  </span>
                </div>
              </div>
            </div>

            {/* Mass default hour assignment for all courses */}
            {rawData.length > 0 && (
              <div className="border-t border-[#E2E8F0] pt-3.5 space-y-2">
                <label className="text-xs font-semibold text-[#64748B] flex justify-between items-center">
                  <span>⏰ 모든 개설 강좌 기본 시수 일괄 지정</span>
                </label>
                <div className="flex gap-1.5">
                  <input
                    type="number"
                    value={batchHoursVal}
                    onChange={(e) => setBatchHoursVal(Math.max(0, parseInt(e.target.value) || 0))}
                    className="w-16 p-1.5 text-xs border border-[#E2E8F0] rounded-lg focus:outline-none focus:border-[#2563EB] text-center font-bold text-slate-800 bg-[#F1F5F9]/50"
                  />
                  <button
                    onClick={handleApplyBatchHours}
                    className="flex-1 p-1.5 bg-slate-100 hover:bg-slate-200/70 border border-[#E2E8F0] text-[#1E293B] rounded-lg text-xs font-bold transition active:scale-98 cursor-pointer flex items-center justify-center gap-1"
                  >
                    <RefreshCw className="w-3.5 h-3.5 text-slate-500" />
                    시수 일괄 지정
                  </button>
                </div>
              </div>
            )}
          </div>

          {/* Quick Stats banner (Summary) */}
          {rawData.length > 0 && (
            <div className="bg-white p-4 rounded-xl border border-[#E2E8F0] shadow-[0_1px_3px_rgba(0,0,0,0.02)] flex flex-col gap-3.5 text-left">
              <h4 className="text-xs font-bold text-[#1E293B] uppercase tracking-wider flex items-center justify-between border-b border-slate-100 pb-2">
                <span>📊 실시간 수강 정산 요약</span>
                <span className="text-[10px] font-semibold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-100">계산 정상</span>
              </h4>
              
              <div>
                {/* 전체 강좌 요약 */}
                <div className="bg-emerald-50/40 p-4 rounded-xl border border-emerald-100/60 flex flex-col justify-between">
                  <div>
                    <span className="text-xs text-emerald-800 font-extrabold block border-b border-emerald-100 pb-1.5 mb-2.5 flex items-center justify-between">
                      <span className="flex items-center gap-1">🏛️ 전체 강좌 요약</span>
                      <span className="text-[10px] bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded-md font-semibold">전체</span>
                    </span>
                    <div className="space-y-2 text-xs text-slate-600">
                      <div className="flex justify-between items-center">
                        <span>신청 학생 수</span>
                        <span className="font-bold text-emerald-700">{summaries.all.uniqueStudents}명</span>
                      </div>
                      <div className="flex justify-between items-center">
                        <span>개설 강좌 수</span>
                        <span className="font-bold text-emerald-700">{summaries.all.coursesCount}개</span>
                      </div>
                      <div className="flex justify-between border-t border-emerald-100/40 pt-2.5 mt-2.5 font-semibold text-slate-700">
                        <span className="text-[11px] text-slate-500 font-medium">저소득층 지원금 소계</span>
                        <span className="font-black text-amber-700 font-mono text-xs tabular-nums">{summaries.beneficiary.lowIncomeSubsidyTotal.toLocaleString()}원</span>
                      </div>
                      <div className="flex justify-between font-semibold text-slate-700">
                        <span className="text-[11px] text-slate-500 font-medium">다자녀 지원금 소계</span>
                        <span className="font-black text-amber-700 font-mono text-xs tabular-nums">{summaries.beneficiary.multiChildSubsidyTotal.toLocaleString()}원</span>
                      </div>
                    </div>
                  </div>
                  <div className="bg-emerald-500/10 p-2 rounded-lg mt-3 text-xs flex justify-between items-center">
                    <span className="text-emerald-600 font-bold">수강료 자부담 소계</span>
                    <span className="font-extrabold text-emerald-700 font-mono text-xs tabular-nums">{summaries.beneficiary.studentPaid.toLocaleString()}원</span>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Desktop small inline info footer */}
          <div className="hidden lg:block mt-auto pt-4 border-t border-slate-200/50">
          </div>

        </section>

        {/* Right Side: Calculation details or waiting state */}
        <section className="flex-1 min-w-0 flex flex-col gap-3.5 lg:h-full lg:overflow-hidden text-left">
          
          {error && (
            <div className="p-4 bg-rose-50 border border-rose-200 text-rose-900 rounded-xl flex gap-3 items-start shadow-xs">
              <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
              <div>
                <p className="text-sm font-semibold text-rose-900">처리 오류가 검출되었습니다</p>
                <p className="text-xs text-rose-800 mt-0.5 whitespace-pre-wrap">{error}</p>
              </div>
            </div>
          )}

          {loading ? (
            <div className="bg-white rounded-2xl border border-slate-200 p-12 text-center flex-1 flex flex-col justify-center items-center gap-4 min-h-[450px]">
              <div className="relative">
                <div className="w-12 h-12 border-4 border-sky-200 border-t-sky-600 rounded-full animate-spin"></div>
                <RefreshCw className="w-5 h-5 text-sky-600 absolute inset-0 m-auto animate-pulse" />
              </div>
              <p className="text-base font-semibold text-slate-950 mt-2">
                학사 행정 데이터 분석 및 정산 처리 연산 중...
              </p>
              <p className="text-xs text-slate-400">
                중복 명부 필터링 및 학생별 매핑 계산을 안전하게 진행하고 있습니다.
              </p>
            </div>
          ) : rawData.length === 0 ? (
            
            // Welcome Stage: Instructions for Uploading (Scrollable wrapper on small desktop screens)
            <div className="flex-1 overflow-y-auto pr-1.5 flex flex-col lg:justify-center min-h-0 scrollbar-thin">
              <div className="bg-white rounded-2xl border border-slate-200 p-8 sm:p-10 text-center flex-1 flex flex-col justify-center items-center">
                <div className="w-14 h-14 bg-gradient-to-tr from-sky-100 to-indigo-100 text-sky-600 rounded-2xl flex items-center justify-center mb-5 shadow-sm">
                  <FileSpreadsheet className="w-7 h-7" />
                </div>
                
                <h2 className="text-lg sm:text-2xl font-black text-slate-900 tracking-tight max-w-md block">
                  리로스쿨 학생 수강 내역을 등록하여 정산을 순식간에 끝내세요
                </h2>
                <p className="text-slate-500 text-xs sm:text-sm mt-3 max-w-lg leading-relaxed text-center">
                  방과후 행정 처리를 더는 복잡한 서식 수작업으로 하지 마세요. 좌측에 리로스쿨 방과후 수강 대장 엑셀을 끌어다 놓으시면, <strong>시수별 1인당 단가, 최종 기수 수강료 정산 대장, 분반별 개별 출석 명단</strong>까지 원클릭으로 일괄 완성됩니다.
                </p>

                {/* Step checklist guide */}
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 w-full max-w-2xl mt-6">
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-150 text-left">
                    <div className="text-sky-600 font-extrabold text-xs mb-1 flex justify-between items-center">
                      <span>1. 엑셀 업로드</span>
                      <span className="text-[9px] px-1.5 py-0.2 bg-sky-100 text-sky-700 rounded font-semibold font-sans">리로스쿨형</span>
                    </div>
                    <p className="text-[10px] text-slate-500 leading-normal">
                      리로스쿨에서 내려받은 신청 현황 행렬 목록 파일을 수정 없이 그대로 끌어서 넣으세요.
                    </p>
                  </div>
                  
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-150 text-left">
                    <div className="text-sky-600 font-extrabold text-xs mb-1 flex justify-between items-center">
                      <span>2. 강좌별 시수 조정</span>
                      <span className="text-[9px] px-1.5 py-0.2 bg-teal-100 text-teal-800 rounded font-semibold font-sans">실시간</span>
                    </div>
                    <p className="text-[10px] text-slate-500 leading-normal">
                      기수마다 다른 실제 운영 시간(시수)을 한 기수씩 수동 입력하거나 일괄 지정할 수 있습니다.
                    </p>
                  </div>

                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-150 text-left">
                    <div className="text-sky-600 font-extrabold text-xs mb-1 flex justify-between items-center">
                      <span>3. 다운로드 및 완성</span>
                      <span className="text-[9px] px-1.5 py-0.2 bg-amber-100 text-amber-800 rounded font-semibold font-sans">XLSX</span>
                    </div>
                    <p className="text-[10px] text-slate-500 leading-normal">
                      각 정산액과 출석 일지를 모두 담은 최종 다중 시트 마스터 정산 엑셀 파일을 다운로드합니다.
                    </p>
                  </div>
                </div>

                {/* Sample file reference */}
                <div className="mt-6 flex items-center gap-3 px-3.5 py-2 bg-amber-50 rounded-xl border border-amber-200/50 max-w-lg text-left">
                  <HelpCircle className="w-4.5 h-4.5 text-amber-600 flex-shrink-0" />
                  <span className="text-[10.5px] text-amber-800 leading-normal">
                    당장 테스트해 볼 파일이 없으시다면 상단 우측의 <strong>[양식 예제 다운로드]</strong> 혹은 좌측의 <strong>[데모 데이터]</strong> 버튼을 사용하여 가상 조작을 시뮬레이션 하실 수 있습니다.
                  </span>
                </div>
              </div>
            </div>

          ) : (
            
            // Main Output Area when file is populated
            <div className="flex-1 flex flex-col gap-3.5 min-h-0 overflow-hidden">
              
              {/* Workspace CTA Action bar (Highly Compressed & Professional) */}
              <div className="bg-emerald-50/60 p-2.5 px-4 rounded-xl border border-emerald-100/80 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3.5 shrink-0">
                <div className="flex items-center gap-2">
                  <span className="relative flex h-2 w-2 grow-0 shrink-0">
                    <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                    <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
                  </span>
                  <div>
                    <h3 className="text-xs font-black text-emerald-900 flex items-center gap-1.5 leading-none">
                      정산 분석 명부 가공 완료
                    </h3>
                  </div>
                </div>
                
                <button
                  onClick={handleExport}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold text-[11px] px-4 py-1.5 rounded-lg transition-all shadow-xs duration-200 hover:shadow-emerald-600/10 cursor-pointer active:scale-98 flex items-center justify-center gap-1.5"
                >
                  <Download className="w-3.5 h-3.5" />
                  정산 엑셀 다운로드 (.xlsx)
                </button>
              </div>

              {/* Dynamic Tabs selectors */}
              <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden flex-1 flex flex-col">
                <div className="flex border-b border-slate-200 bg-slate-50/70 p-1.5 gap-1">
                  
                  <button
                    onClick={() => setActivePreviewTab('courses')}
                    className={`flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2.5 text-xs sm:text-sm font-semibold rounded-lg transition-all cursor-pointer ${
                      activePreviewTab === 'courses'
                        ? 'bg-slate-900 border border-slate-900 text-white'
                        : 'text-slate-600 hover:text-slate-950 hover:bg-slate-200/50'
                    }`}
                  >
                    <Coins className="w-4 h-4 text-emerald-500" />
                    1. 강좌별 시수 및 정산 단가
                  </button>

                  <button
                    onClick={() => setActivePreviewTab('students')}
                    className={`flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2.5 text-xs sm:text-sm font-semibold rounded-lg transition-all cursor-pointer ${
                      activePreviewTab === 'students'
                        ? 'bg-slate-900 border border-slate-900 text-white'
                        : 'text-slate-600 hover:text-slate-950 hover:bg-slate-200/50'
                    }`}
                  >
                    <Users className="w-4 h-4 text-sky-500" />
                    2. 학생별 개별 정산 고지
                  </button>

                  <button
                    onClick={() => setActivePreviewTab('attendance')}
                    className={`flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2.5 text-xs sm:text-sm font-semibold rounded-lg transition-all cursor-pointer ${
                      activePreviewTab === 'attendance'
                        ? 'bg-slate-900 border border-slate-900 text-white'
                        : 'text-slate-600 hover:text-slate-950 hover:bg-slate-200/50'
                    }`}
                  >
                    <CheckSquare className="w-4 h-4 text-indigo-500" />
                    3. 강좌별 임시 출석부 명단
                  </button>
                </div>

                {/* Tab Content Display Container */}
                <div className="p-5 flex-1 overflow-y-auto">
                  <AnimatePresence mode="wait">
                    
                    {/* Tab 1: Course matrix with manual override inputs */}
                    {activePreviewTab === 'courses' && (
                      <motion.div
                        key="tab-courses"
                        initial={{ opacity: 0, y: 5 }}
                        animate={{ opacity: 1, y: 0 }}
                        exit={{ opacity: 0, y: -5 }}
                        className="space-y-4"
                      >
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-slate-50 p-3.5 rounded-xl border border-slate-100 text-slate-700">
                          <span className="text-xs leading-normal font-semibold text-slate-800">
                            💡 강좌별 실제 운영일치에 따라서 운영 시수(시간값)를 각각 맞게 입력하세요. 비수익자 강좌 등 정산에서 제외할 강좌는 [삭제] 버튼으로 제거할 수 있습니다.
                          </span>
                          <span className="text-[10px] text-slate-400">
                            *수강 인원이 없으면 정산 불가
                          </span>
                        </div>

                        <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white shadow-xs">
                          <table className="min-w-full divide-y divide-slate-200 text-left text-xs">
                            <thead className="bg-slate-50 text-slate-700">
                              <tr>
                                <th scope="col" className="px-4 py-3 font-semibold md:w-1/3">강좌명</th>
                                <th scope="col" className="px-4 py-3 font-semibold text-center w-28">신청 학생 수</th>
                                <th scope="col" className="px-4 py-3 font-semibold text-center w-24">강좌 삭제</th>
                                <th scope="col" className="px-4 py-3 font-semibold text-center w-36">운영 시수 (수정 가능)</th>
                                <th scope="col" className="px-4 py-3 font-semibold text-right">총 강사료 (₩)</th>
                                <th scope="col" className="px-4 py-3 font-semibold text-right text-emerald-700 bg-emerald-50/50">1인당 수강료 (₩)</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-150">
                              {calculated.courses.map((course) => (
                                <tr key={course.name} className="hover:bg-slate-50/50 transition">
                                  <td className="px-4 py-3 font-medium text-slate-900 break-words max-w-[180px]">
                                    {course.name}
                                  </td>
                                  <td className="px-4 py-3 text-center">
                                    <span className="px-2 py-1 bg-slate-100/85 text-slate-800 font-bold rounded-md">
                                      {course.studentCount}명
                                    </span>
                                  </td>
                                  <td className="px-4 py-3 text-center">
                                    <button
                                      type="button"
                                      onClick={() => handleDeleteCourse(course.name)}
                                      className="inline-flex items-center justify-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-rose-600 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-md transition cursor-pointer active:scale-95"
                                      title="비수익자 강좌 등 정산에서 제외 및 강좌 삭제"
                                    >
                                      <Trash2 className="w-3 h-3" />
                                      삭제
                                    </button>
                                  </td>
                                  <td className="px-4 py-3 text-center">
                                    <div className="flex items-center justify-center gap-1.5">
                                      <input
                                        type="number"
                                        min="0"
                                        value={course.hours}
                                        onChange={(e) => handleHoursChange(course.name, parseInt(e.target.value) || 0)}
                                        className="w-16 p-1 border border-slate-300 rounded focus:outline-none focus:border-sky-500 text-center font-bold text-slate-950"
                                      />
                                      <span className="text-slate-500 text-[11px]">시간</span>
                                    </div>
                                  </td>
                                  <td className="px-4 py-3 text-right">
                                    <span className="tabular-nums text-slate-600 font-mono">
                                      {course.totalLecturerCost.toLocaleString()}
                                    </span>
                                  </td>
                                  <td className="px-4 py-3 text-right text-emerald-800 bg-emerald-50/30 font-bold">
                                    <span className="tabular-nums">
                                      {course.perStudentFee.toLocaleString()}
                                    </span>
                                    <span className="text-[10px] text-emerald-500 font-normal ml-0.5">원</span>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </motion.div>
                    )}

                    {/* Tab 2: Individual billing with filtering */}
                    {activePreviewTab === 'students' && (
                      <motion.div
                        key="tab-students"
                        initial={{ opacity: 0, y: 5 }}
                        animate={{ opacity: 1, y: 0 }}
                        exit={{ opacity: 0, y: -5 }}
                        className="space-y-4"
                      >
                        {/* Search Control */}
                        <div className="relative">
                          <Search className="w-4 h-4 text-slate-400 absolute left-3 top-3" />
                          <input
                            type="text"
                            placeholder="찾으려는 학생명 혹은 학번을 입력하세요 (예: 강유민, 10101)..."
                            value={searchQuery}
                            onChange={(e) => setSearchQuery(e.target.value)}
                            className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-1 focus:ring-sky-500 focus:border-sky-500 text-slate-800 placeholder-slate-400 bg-slate-50/50"
                          />
                        </div>

                        {/* Deleted Students Banner & Mini-List */}
                        {excludedStudentIds.length > 0 && (
                          <motion.div
                            initial={{ opacity: 0, y: -5 }}
                            animate={{ opacity: 1, y: 0 }}
                            className="bg-rose-50 border border-rose-200/60 rounded-xl p-3.5 text-xs text-rose-950 space-y-2 shadow-xs"
                          >
                            <div className="flex items-center justify-between">
                              <div className="flex items-center gap-1.5 font-extrabold text-rose-800">
                                <AlertCircle className="w-4 h-4 text-rose-600 animate-pulse" />
                                <span>정산 및 출석부 수동 실제 삭제 학생 명단 ({excludedStudentIds.length}명)</span>
                              </div>
                              <button
                                type="button"
                                onClick={() => {
                                  setExcludedStudentIds([]);
                                }}
                                className="text-[11px] text-rose-700 underline font-extrabold hover:text-rose-950 cursor-pointer"
                              >
                                전체 학생 복구하기
                              </button>
                            </div>
                            <p className="text-[11px] text-rose-800 leading-normal">
                              학번 왼쪽에 있는 🗑️ 삭제 버튼을 누른 명단입니다. 이 학생들은 <strong>수강 등록정보, 강좌 인원수 집계, 최종 정산 엑셀, 통합 출석부</strong>에서 완벽하게 실제 삭제(차감)되었으며, 남은 학생 수를 기준으로 강좌 단가가 실시간 자동 상승 재계산되었습니다.
                            </p>
                            <div className="flex flex-wrap gap-1.5 pt-1">
                              {excludedStudentIds.map(id => {
                                const row = rawData.find(r => normalizeStudentId(r['학번']) === id);
                                const name = row ? String(row['이름'] || '').trim() : '알 수 없음';
                                return (
                                  <div key={id} className="inline-flex items-center gap-1.5 px-2.5 py-1 bg-rose-100 border border-rose-200 rounded-lg text-[11px] font-bold text-rose-950">
                                    <span>{name} ({id})</span>
                                    <button
                                      type="button"
                                      onClick={() => setExcludedStudentIds(prev => prev.filter(x => x !== id))}
                                      className="text-rose-500 hover:text-rose-900 ml-1 font-extrabold cursor-pointer text-xs"
                                      title="정산에 다시 복구"
                                    >
                                      복구(취소)
                                    </button>
                                  </div>
                                );
                              })}
                            </div>
                          </motion.div>
                        )}

                        <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white shadow-xs">
                          <table className="min-w-full divide-y divide-slate-200 text-left text-xs">
                            <thead className="bg-slate-50 text-slate-700">
                              <tr>
                                <th scope="col" className="px-3 py-3 font-semibold text-center w-14">삭제</th>
                                <th scope="col" className="px-4 py-3 font-semibold">학번</th>
                                <th scope="col" className="px-4 py-3 font-semibold">이름</th>
                                <th scope="col" className="px-4 py-3 font-semibold text-center w-20">강좌 개수</th>
                                <th scope="col" className="px-4 py-3 font-semibold text-right">기본 수강료</th>
                                <th scope="col" className="px-4 py-3 font-semibold text-center">수혜 지원금 정보</th>
                                <th scope="col" className="px-4 py-3 font-semibold text-right bg-sky-50/50 text-sky-900">최종 납부액 (₩)</th>
                                <th scope="col" className="px-4 py-3 font-semibold text-center w-28">지원금 관리</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-150">
                              {filteredStudents.length === 0 ? (
                                <tr>
                                  <td colSpan={8} className="text-center py-8 text-slate-400 text-xs">
                                    검색어와 관련된 학생이 목록에 존재하지 않거나 모두 정산에서 삭제되었습니다.
                                  </td>
                                </tr>
                              ) : (
                                filteredStudents.map((std) => (
                                  <tr key={std.studentId} className="hover:bg-slate-50/50 transition">
                                    <td className="px-3 py-3 text-center">
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setExcludedStudentIds(prev => [...prev, std.studentId]);
                                        }}
                                        className="text-slate-400 hover:text-rose-600 hover:bg-rose-50 p-1.5 rounded-lg transition-colors cursor-pointer inline-flex items-center justify-center"
                                        title="정산 및 출석부에서 실제 삭제 (상단 복구 가능)"
                                      >
                                        <Trash2 className="w-3.5 h-3.5" />
                                      </button>
                                    </td>
                                    <td className="px-4 py-3 font-mono text-slate-600 tabular-nums">
                                      {std.studentId}
                                    </td>
                                    <td className="px-4 py-3 font-bold text-slate-900">
                                      {std.name}
                                    </td>
                                    <td className="px-4 py-3 text-center">
                                      <span className="inline-block px-2 py-0.5 bg-slate-100 text-slate-700 rounded-full font-semibold">
                                        {std.courseCount}개
                                      </span>
                                    </td>
                                    <td className="px-4 py-3 text-right text-slate-500 tabular-nums">
                                      {std.originalFee.toLocaleString()}원
                                    </td>
                                    <td className="px-4 py-3 text-center">
                                      {std.subsidyAmount > 0 ? (
                                        <div className="flex flex-col items-center gap-0.5">
                                          <span className="px-2 py-0.5 rounded text-[10px] font-black bg-amber-50 text-amber-700 border border-amber-200 whitespace-nowrap">
                                            {std.subsidyType}
                                          </span>
                                          <span className="text-[10px] font-bold text-amber-600 font-mono">
                                            -{std.subsidyAmount.toLocaleString()}원
                                          </span>
                                        </div>
                                      ) : (
                                        <span className="text-slate-400 text-[10px]">지정 없음</span>
                                      )}
                                    </td>
                                    <td className="px-4 py-3 text-right bg-sky-50/20 font-extrabold text-[#2563EB] tabular-nums text-sm">
                                      {std.totalFee.toLocaleString()}원
                                    </td>
                                    <td className="px-4 py-3 text-center">
                                      <button
                                        type="button"
                                        onClick={() => openSubsidyModal(std.studentId, std.name, std.originalFee)}
                                        className="px-2 py-1 bg-white hover:bg-slate-50 border border-slate-200 text-slate-700 rounded text-[11px] font-semibold shadow-xs transition active:scale-98 cursor-pointer inline-flex items-center gap-1"
                                      >
                                        ⚙️ 지원금 설정
                                      </button>
                                    </td>
                                  </tr>
                                ))
                              )}
                            </tbody>
                          </table>
                        </div>
                      </motion.div>
                    )}

                    {/* Tab 3: Attendance rosters rendering with dynamic filtering selector */}
                    {activePreviewTab === 'attendance' && (
                      <motion.div
                        key="tab-attendance"
                        initial={{ opacity: 0, y: 5 }}
                        animate={{ opacity: 1, y: 0 }}
                        exit={{ opacity: 0, y: -5 }}
                        className="space-y-4"
                      >
                        <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 bg-slate-50 p-4 rounded-xl border border-slate-150">
                          <label className="text-xs font-black text-slate-800 whitespace-nowrap flex items-center gap-1">
                            <BookOpen className="w-4 h-4 text-indigo-500" />
                            출석부 미리볼 강좌 선택:
                          </label>
                          <select
                            value={selectedAttendanceCourse}
                            onChange={(e) => setSelectedAttendanceCourse(e.target.value)}
                            className="bg-white border border-slate-300 rounded-lg p-2 text-xs focus:outline-none focus:ring-1 focus:ring-indigo-500 font-semibold text-slate-900 flex-1 sm:max-w-md cursor-pointer"
                          >
                            {courseColumns.map(c => (
                              <option key={c} value={c}>{c}</option>
                            ))}
                          </select>
                        </div>

                        {selectedAttendanceCourse && (
                          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                            <div className="bg-indigo-900 text-white p-4 flex justify-between items-center">
                              <span className="text-xs font-bold uppercase tracking-wider bg-indigo-800 text-indigo-200 px-2.5 py-1 rounded">
                                명부 탭 생성 가공
                              </span>
                              <h4 className="text-sm font-black truncate max-w-sm">
                                {selectedAttendanceCourse} 출석부 미리보기
                              </h4>
                              <p className="text-xs text-indigo-200 text-right font-medium">
                                총 {currentAttendanceRecords.length}명 수강
                              </p>
                            </div>
                            
                            <table className="min-w-full divide-y divide-slate-150 text-left text-xs">
                              <thead className="bg-slate-50 text-slate-600">
                                <tr>
                                  <th scope="col" className="px-6 py-3 font-semibold text-center w-20">연번</th>
                                  <th scope="col" className="px-6 py-3 font-semibold text-center w-36">학번</th>
                                  <th scope="col" className="px-6 py-3 font-semibold">학생 이릅</th>
                                  <th scope="col" className="px-6 py-3 font-semibold text-slate-400">교사 확인 서명식</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                {currentAttendanceRecords.length === 0 ? (
                                  <tr>
                                    <td colSpan={4} className="text-center py-8 text-slate-400">
                                      수강을 신청한 학생들이 전무합니다.
                                    </td>
                                  </tr>
                                ) : (
                                  currentAttendanceRecords.map((std) => (
                                    <tr key={std.index} className="hover:bg-slate-50/50 transition duration-100">
                                      <td className="px-6 py-2.5 font-mono text-center text-slate-400 font-medium">
                                        {std.index}
                                      </td>
                                      <td className="px-6 py-2.5 font-mono text-center text-slate-550 tabular-nums">
                                        {std.studentId}
                                      </td>
                                      <td className="px-6 py-2.5 font-bold text-slate-900">
                                        {std.name}
                                      </td>
                                      <td className="px-6 py-2.5 text-slate-300 italic text-[11px] font-normal">
                                        [인쇄시 서명 서식란 출력됨]
                                      </td>
                                    </tr>
                                  ))
                                )}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>
              </div>
            </div>
          )}

        </section>

      </main>

      {/* Persistent Footer */}
      <footer className="lg:hidden bg-white border-t border-slate-200 py-5 text-center text-xs text-slate-400 mt-6">
        <div className="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row justify-between items-center gap-2.5">
          <p>
            © 2026 방과후학교 정산 자동화 시스템. 리로스쿨 출력이식형.
          </p>
          <p className="flex items-center gap-1 justify-center">
            <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full inline-block animate-pulse"></span>
            웹 브라우저의 전용 컴퓨팅 파워를 활용한 100% 무서버 오프라인 전용 연산 설계
          </p>
        </div>
      </footer>

      {/* 지원금 입력 모달(Modal) */}
      <AnimatePresence>
        {editingStudentId !== null && (
          <div className="fixed inset-0 bg-slate-900/60 flex items-center justify-center p-4 z-50 backdrop-blur-xs">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white p-6 rounded-2xl border border-slate-200 shadow-2xl max-w-md w-full relative"
            >
              <h3 className="text-sm font-black text-slate-900 flex items-center gap-1.5 mb-4">
                💎 지원 및 보조 교육비 설정 ({editingStudentName})
              </h3>
              
              <div className="space-y-4">
                {/* 원래 총 수강 정산금 */}
                <div className="bg-slate-50 p-3 rounded-lg border border-slate-200 flex justify-between items-center">
                  <span className="text-xs font-bold text-slate-600">수강 등록 원래 수강료</span>
                  <span className="text-xs font-extrabold text-slate-900">
                    {(calculated.students.find(s => s.studentId === editingStudentId)?.originalFee || 0).toLocaleString()}원
                  </span>
                </div>

                {/* 지원 분류 종류 */}
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-700">학교 보조/지원금 종류 선택</label>
                  <select
                    value={dialogType}
                    onChange={(e) => {
                      setDialogType(e.target.value);
                      if (e.target.value === '없음') {
                        setDialogAmount(0);
                      }
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg p-2 text-xs focus:outline-none focus:ring-1 focus:ring-blue-500 font-bold text-slate-800"
                  >
                    <option value="없음">지정 없음 (전액 학생/학부모 자부담)</option>
                    <option value="저소득층">저소득층 지원금 (전액 무료)</option>
                    <option value="다자녀">다자녀 지원금 (전액 무료)</option>
                    <option value="자유수강권">자유수강권 (방과후학교 바우처)</option>
                    <option value="교육청지원금">교육청 복지 지원금</option>
                    <option value="학교자체지원금">학교 자체 교육보조비</option>
                  </select>
                </div>

                {/* 지원 금액 입력 필드 */}
                {dialogType !== '없음' && (
                  <div className="space-y-1">
                    <div className="flex justify-between items-center">
                      <label className="text-xs font-bold text-slate-700">적용 지원금 금액 (원)</label>
                      <button
                        type="button"
                        onClick={() => {
                          const originalFee = calculated.students.find(s => s.studentId === editingStudentId)?.originalFee || 0;
                          setDialogAmount(originalFee);
                        }}
                        className="text-[10px] text-blue-600 font-extrabold hover:underline cursor-pointer"
                      >
                        [수강료 전액 지원으로 자동 기입]
                      </button>
                    </div>
                    <div className="flex items-center gap-1">
                      <span className="p-2.5 bg-slate-150 border border-r-0 border-slate-300 rounded-l-lg text-slate-500 text-xs font-bold">
                        ₩
                      </span>
                      <input
                        type="number"
                        min="0"
                        value={dialogAmount}
                        onChange={(e) => setDialogAmount(Math.max(0, parseInt(e.target.value) || 0))}
                        className="w-full p-1.5 border border-slate-300 rounded-r-lg focus:outline-none focus:border-blue-500 text-xs font-bold text-slate-900"
                        placeholder="지원 금액을 작성하세요"
                      />
                    </div>
                  </div>
                )}
              </div>

              {/* Action buttons */}
              <div className="mt-6 flex justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setEditingStudentId(null)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-bold rounded-lg transition"
                >
                  취소
                </button>
                <button
                  type="button"
                  onClick={saveSubsidy}
                  className="px-4 py-2 bg-[#2563EB] hover:bg-[#1D4ED8] text-white text-xs font-bold rounded-lg transition"
                >
                  저장 및 명세서 즉시 갱신
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

    </div>
  );
}

