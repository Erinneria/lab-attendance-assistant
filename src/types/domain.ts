export interface StudentRecord {
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
  suggestedScore?: number | null;
  remark: string;
}

export interface TimeSlotRecord {
  id: string;
  label: string;
  limit: number;
  registered: number;
  groupName: string;
  students: string[];
}

export interface ColumnMapping {
  studentId?: string;
  name?: string;
  grade?: string;
  className?: string;
  department?: string;
  originalGroupInfo?: string;
  teacherName?: string;
}

export interface ColumnMappingPreset {
  name: string;
  mapping: ColumnMapping;
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

export interface AppSettings {
  autoSave: boolean;
  projectRootDir: string;
  defaultExportDir: string;
  emailImportConfig?: EmailImportConfig;
}

export interface LabProject {
  courseName: string;
  term: string;
  teacher: string;
  students: StudentRecord[];
  timeSlots: TimeSlotRecord[];
  columnMappings: ColumnMapping;
  columnMappingPresets?: ColumnMappingPreset[];
  signinLayout: SigninLayoutSettings;
  settings: AppSettings;
  sourceFiles: {
    studentList?: string;
    timeSlots?: string;
    reportDir?: string;
  };
  importedEmailAttachmentKeys?: string[];
}

export interface StudentImportRow {
  [key: string]: string;
}

export interface ParsedTimeSlot {
  label: string;
  limit: number;
  registered: number;
  students: string[];
  rawRowIndex: number;
}

export interface MatchedStudentIssue {
  type:
    | 'missing-student-id'
    | 'missing-name'
    | 'duplicate-student-id'
    | 'duplicate-name'
    | 'unmatched-timeslot'
    | 'multiple-timeslot'
    | 'duplicate-group-name';
  message: string;
  studentId?: string;
  name?: string;
  groupName?: string;
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

export interface EmailAttachmentMatch {
  attachment: ImportedEmailAttachment;
  status: 'matched' | 'unmatched' | 'duplicate-name' | 'conflict';
  matchedStudentId?: string;
  matchedStudentName?: string;
  matchBasis?: 'sender-name' | 'attachment-student-id' | 'attachment-name' | 'subject-student-id' | 'subject-name';
  reason: string;
}

export interface AutoReviewSubmissionInput {
  studentId: string;
  name: string;
  submissionFile: string;
  submissionStatus: StudentRecord['submissionStatus'];
}

export interface AutoReviewResult {
  studentId: string;
  score: number;
  remark: string;
  details: string[];
}
