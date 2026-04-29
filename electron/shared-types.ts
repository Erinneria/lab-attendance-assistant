export interface LabProject {
  courseName: string;
  term: string;
  teacher: string;
  students: Array<{
    studentId: string;
    name: string;
    grade: string;
    className: string;
    department: string;
    originalGroupInfo: string;
    teacherName: string;
    timeSlotId: string;
    timeSlotLabel: string;
    groupName: string;
    attendance: boolean;
    submissionFile: string;
    submissionStatus: 'missing' | 'submitted' | 'manual';
    score: number | null;
    remark: string;
  }>;
  timeSlots: Array<{
    id: string;
    label: string;
    limit: number;
    registered: number;
    groupName: string;
    students: string[];
  }>;
  columnMappings: Record<string, string | undefined>;
  signinLayout: SigninLayoutSettings;
  settings: {
    autoSave: boolean;
    projectRootDir?: string;
    defaultExportDir: string;
    emailImportConfig?: EmailImportConfig;
  };
  sourceFiles: {
    studentList?: string;
    timeSlots?: string;
  };
  importedEmailAttachmentKeys?: string[];
}

export interface ParsedTimeSlot {
  label: string;
  limit: number;
  registered: number;
  students: string[];
  rawRowIndex: number;
}

export interface SigninLayoutSettings {
  orientation: 'portrait' | 'landscape';
  blocksPerPage: number;
  maxStudentsPerBlock: number;
  blockColumnWidth: number;
  rowHeight: number;
  fontSize: number;
  showGrade: boolean;
  showClassName: boolean;
  showGroupName: boolean;
  showAttendance: boolean;
  sortByGroupName: boolean;
  sortByTimeSlot: boolean;
  sortByStudentId: boolean;
  separatePageByTimeSlot: boolean;
  showHeaderMeta: boolean;
}

export interface EmailImportConfig {
  host: string;
  port: number;
  secure: boolean;
  user: string;
  password: string;
  mailbox: string;
  since?: string;
  before?: string;
  subjectKeyword?: string;
  limit: number;
  downloadDir: string;
  knownAttachmentKeys?: string[];
  matchHints?: Array<{ studentId: string; name: string }>;
}

export interface ImportedEmailAttachment {
  uid: number;
  messageId: string;
  fromName: string;
  fromAddress: string;
  subject: string;
  date: string;
  originalFileName: string;
  savedFileName: string;
  savedPath: string;
  contentType: string;
  size: number;
}

export interface AutoReviewSubmissionInput {
  studentId: string;
  name: string;
  submissionFile: string;
  submissionStatus: 'missing' | 'submitted' | 'manual';
}

export interface AutoReviewResult {
  studentId: string;
  score: number;
  remark: string;
  details: string[];
}
