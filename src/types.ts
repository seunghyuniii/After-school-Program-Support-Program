/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface StudentRow {
  [key: string]: string | number;
}

export interface CourseInfo {
  name: string;
  studentCount: number;
  hours: number;
  totalLecturerCost: number;
  perStudentFee: number;
}

export interface StudentSettlement {
  studentId: string;
  name: string;
  courseCount: number;
  enrolledCourses: string[];
  originalFee: number;
  subsidyType: string;
  subsidyAmount: number;
  totalFee: number;
}

export interface AttendanceRecord {
  index: number;
  studentId: string;
  name: string;
}

export interface CourseAttendance {
  courseName: string;
  students: AttendanceRecord[];
}
