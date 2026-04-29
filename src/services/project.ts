import type {
  AppSettings,
  ColumnMapping,
  LabProject,
  MatchedStudentIssue,
  ParsedTimeSlot,
  StudentRecord,
  StudentImportRow,
  TimeSlotRecord,
} from '@/types/domain';

const defaultSigninLayout = {
  orientation: 'portrait' as const,
  blocksPerPage: 2,
  maxStudentsPerBlock: 30,
  blockColumnWidth: 12,
  rowHeight: 20,
  fontSize: 11,
  showGrade: true,
  showClassName: true,
  showGroupName: true,
  showAttendance: true,
  sortByGroupName: true,
  sortByTimeSlot: false,
  sortByStudentId: true,
  separatePageByTimeSlot: false,
  showHeaderMeta: true,
};

export function normalizeText(value: unknown): string {
  return String(value ?? '').trim();
}

function canonical(value: string): string {
  return normalizeText(value).replace(/\s+/g, '').toLowerCase();
}

function matchesHeader(header: string, candidates: string[], keywords: string[]) {
  const normalized = canonical(header);
  return candidates.some((candidate) => normalized === canonical(candidate))
    || keywords.some((keyword) => normalized.includes(canonical(keyword)));
}

export function inferStudentColumnMapping(headers: string[]): ColumnMapping {
  const mapping: ColumnMapping = {};
  const exact = {
    studentId: ['学号', '学生学号', 'student id', 'studentid', 'sid'],
    name: ['姓名', '学生姓名', 'name'],
    grade: ['年级', 'grade'],
    className: ['班级', '行政班', 'class', 'class name'],
    department: ['所属院系', '院系', 'department', '学院'],
    originalGroupInfo: ['已选分组信息', '分组信息', '选课分组', 'group', 'group info'],
    teacherName: ['教师名', '教师名称', 'teacher', 'teacher name'],
  } satisfies Record<keyof ColumnMapping, string[]>;

  const keywords = {
    studentId: ['学号', 'id', '编号'],
    name: ['姓名', '名字'],
    grade: ['年级'],
    className: ['班级', '行政班'],
    department: ['院系', '学院'],
    originalGroupInfo: ['分组', '选课', '组'],
    teacherName: ['教师'],
  } satisfies Record<keyof ColumnMapping, string[]>;

  (Object.keys(exact) as Array<keyof ColumnMapping>).forEach((key) => {
    const found = headers.find((header) => matchesHeader(header, exact[key], keywords[key]));
    if (found) {
      mapping[key] = found;
    }
  });

  return mapping;
}

export function buildStudentRecords(rows: StudentImportRow[], mapping: ColumnMapping): StudentRecord[] {
  return rows.map((row) => ({
    studentId: normalizeText(row[mapping.studentId ?? ''] ?? ''),
    name: normalizeText(row[mapping.name ?? ''] ?? ''),
    grade: normalizeText(row[mapping.grade ?? ''] ?? ''),
    className: normalizeText(row[mapping.className ?? ''] ?? ''),
    department: normalizeText(row[mapping.department ?? ''] ?? ''),
    originalGroupInfo: normalizeText(row[mapping.originalGroupInfo ?? ''] ?? ''),
    teacherName: normalizeText(row[mapping.teacherName ?? ''] ?? ''),
    timeSlotId: '',
    timeSlotLabel: '',
    // if the imported data contains group info, use it; otherwise default to className when available
    groupName: normalizeText(row[mapping.originalGroupInfo ?? ''] ?? '') || normalizeText(row[mapping.className ?? ''] ?? ''),
    attendance: false,
    submissionFile: '',
    submissionStatus: 'missing',
    score: null,
    remark: '',
  }));
}

export function buildTimeSlotRecords(slots: ParsedTimeSlot[]): TimeSlotRecord[] {
  return slots.map((slot, index) => ({
    id: `slot_${String(index + 1).padStart(3, '0')}`,
    label: slot.label,
    limit: slot.limit,
    registered: slot.registered,
    groupName: '',
    students: slot.students,
  }));
}

export function generateGroupNames(timeSlots: TimeSlotRecord[], startGroupName: string): TimeSlotRecord[] {
  const match = startGroupName.match(/^(.*?)(\d+)$/);
  if (!match) {
    return timeSlots.map((slot, index) => ({
      ...slot,
      groupName: index === 0 ? startGroupName : `${startGroupName}_${index + 1}`,
    }));
  }

  const prefix = match[1];
  const startNumber = Number.parseInt(match[2], 10);
  const digitCount = match[2].length;
  return timeSlots.map((slot, index) => ({
    ...slot,
    groupName: `${prefix}${String(startNumber + index).padStart(digitCount, '0')}`,
  }));
}

export function findDuplicateGroupNames(timeSlots: TimeSlotRecord[]): string[] {
  const counts = new Map<string, number>();
  timeSlots.forEach((slot) => {
    const groupName = normalizeText(slot.groupName);
    if (!groupName) {
      return;
    }
    counts.set(groupName, (counts.get(groupName) ?? 0) + 1);
  });
  return [...counts.entries()].filter(([, count]) => count > 1).map(([groupName]) => groupName);
}

export function getSortedStudentsByLayout(students: StudentRecord[], sortByGroupName: boolean, sortByTimeSlot: boolean, sortByStudentId: boolean) {
  const rows = [...students];
  if (sortByGroupName) {
    rows.sort((left, right) => left.groupName.localeCompare(right.groupName, 'zh-Hans-CN') || left.studentId.localeCompare(right.studentId, 'zh-Hans-CN'));
    return rows;
  }
  if (sortByTimeSlot) {
    rows.sort((left, right) => left.timeSlotLabel.localeCompare(right.timeSlotLabel, 'zh-Hans-CN') || left.studentId.localeCompare(right.studentId, 'zh-Hans-CN'));
    return rows;
  }
  if (sortByStudentId) {
    rows.sort((left, right) => left.studentId.localeCompare(right.studentId, 'zh-Hans-CN'));
  }
  return rows;
}

export function matchStudents(project: LabProject): { students: StudentRecord[]; timeSlots: TimeSlotRecord[]; issues: MatchedStudentIssue[] } {
  const students = project.students.map((student) => ({ ...student, timeSlotId: '', timeSlotLabel: '', groupName: '' }));
  const timeSlots = project.timeSlots.map((slot) => ({ ...slot }));
  const issues: MatchedStudentIssue[] = [];
  const byId = new Map<string, StudentRecord[]>();
  const byName = new Map<string, StudentRecord[]>();

  students.forEach((student) => {
    const studentId = normalizeText(student.studentId);
    const name = normalizeText(student.name);
    if (studentId) {
      byId.set(studentId, [...(byId.get(studentId) ?? []), student]);
    }
    if (name) {
      byName.set(name, [...(byName.get(name) ?? []), student]);
    }
  });

  students.forEach((student) => {
    if (!student.studentId) {
      issues.push({ type: 'missing-student-id', message: `缺失学号：${student.name || '未知学生'}` , name: student.name });
    }
    if (!student.name) {
      issues.push({ type: 'missing-name', message: `缺失姓名：${student.studentId || '未知学号'}`, studentId: student.studentId });
    }
  });

  byId.forEach((items, studentId) => {
    if (items.length > 1) {
      issues.push({ type: 'duplicate-student-id', message: `重复学号：${studentId}` , studentId });
    }
  });

  byName.forEach((items, name) => {
    if (items.length > 1) {
      issues.push({ type: 'duplicate-name', message: `重复姓名：${name}` , name });
    }
  });

  const usedStudentIds = new Set<string>();
  timeSlots.forEach((slot) => {
    const linked: StudentRecord[] = [];
    slot.students.forEach((rawValue) => {
      const value = normalizeText(rawValue);
      const idMatch = byId.get(value)?.[0];
      const nameMatch = byName.get(value)?.[0];
      const matched = idMatch ?? nameMatch;
      if (!matched) {
        issues.push({ type: 'unmatched-timeslot', message: `时间段 ${slot.label} 中未匹配到学生：${value}` });
        return;
      }
      if (matched.timeSlotId && matched.timeSlotId !== slot.id) {
        issues.push({ type: 'multiple-timeslot', message: `学生 ${matched.name || matched.studentId} 匹配到多个时间段`, studentId: matched.studentId, name: matched.name });
        return;
      }
      matched.timeSlotId = slot.id;
      matched.timeSlotLabel = slot.label;
      matched.groupName = slot.groupName;
      linked.push(matched);
      usedStudentIds.add(matched.studentId);
    });
    slot.registered = linked.length;
    slot.students = linked.map((student) => student.name || student.studentId);
  });

  students.forEach((student) => {
    if (!student.timeSlotId && student.name && student.studentId) {
      issues.push({ type: 'unmatched-timeslot', message: `学生未匹配时间段：${student.name}（${student.studentId}）`, studentId: student.studentId, name: student.name });
    }
  });

  findDuplicateGroupNames(timeSlots).forEach((groupName) => {
    issues.push({ type: 'duplicate-group-name', message: `重复组名：${groupName}`, groupName });
  });

  return { students, timeSlots, issues };
}

export function validateProject(project: LabProject): MatchedStudentIssue[] {
  return matchStudents(project).issues;
}

export function buildDefaultProject(): LabProject {
  const settings: AppSettings = {
    autoSave: true,
    projectRootDir: '',
    defaultExportDir: '',
    emailImportConfig: undefined,
  };

  return {
    courseName: '实验课签到与成绩管理助手',
    term: '2025 秋',
    teacher: '',
    students: [],
    timeSlots: [],
    columnMappings: {},
    columnMappingPresets: [],
    signinLayout: defaultSigninLayout,
    settings,
    sourceFiles: {},
  };
}
