import { app, BrowserWindow, dialog, ipcMain, shell } from 'electron';
import { access, copyFile, cp, mkdir, readFile, rename, readdir, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import puppeteer from 'puppeteer';
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import mammoth from 'mammoth';
import { PDFParse } from 'pdf-parse';
import JSZip from 'jszip';
import type { AutoReviewResult, AutoReviewSubmissionInput, EmailImportConfig, ImportedEmailAttachment, LabProject, ParsedTimeSlot, SigninLayoutSettings } from './shared-types';

const defaultSigninLayout: SigninLayoutSettings = {
  orientation: 'portrait',
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

function buildDefaultProject(): LabProject {
  return {
    courseName: '实验课签到与成绩管理助手',
    term: '2025 秋',
    teacher: '',
    students: [],
    timeSlots: [],
    columnMappings: {},
    signinLayout: defaultSigninLayout,
    settings: {
      autoSave: true,
      projectRootDir: '',
      defaultExportDir: '',
      emailImportConfig: undefined,
    },
    sourceFiles: {},
  };
}

function normalizeText(value: unknown): string {
  return String(value ?? '').trim();
}

function sanitizeFileNamePart(value: string) {
  return String(value ?? '')
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, ' ')
    .replace(/^\.+/, '')
    .replace(/\.+$/, '') || 'attachment';
}

function toNumber(value: unknown): number {
  const parsed = Number.parseInt(normalizeText(value), 10);
  return Number.isFinite(parsed) ? parsed : 0;
}

function clamp(value: number, min: number, max: number) {
  return Math.max(min, Math.min(max, value));
}

function isDevServerRunning(): boolean {
  return Boolean(process.env.VITE_DEV_SERVER_URL);
}

function getRendererEntry(): string {
  return path.join(__dirname, '..', 'dist-renderer', 'index.html');
}

function getAppWritableBaseDir(): string {
  return app.isPackaged ? path.dirname(app.getPath('exe')) : path.resolve(__dirname, '..');
}

function getPreferredProjectLibraryRoot(): string {
  return path.join(getAppWritableBaseDir(), '.lab-attendance-assistant', 'projects');
}

function getLegacyProjectLibraryRoot(): string {
  return path.join(app.getPath('userData'), 'projects');
}

async function exists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function ensureProjectLibraryRoot(): Promise<string> {
  const preferredDirPath = getPreferredProjectLibraryRoot();
  let dirPath = preferredDirPath;
  try {
    await mkdir(preferredDirPath, { recursive: true });
  } catch {
    dirPath = getLegacyProjectLibraryRoot();
    await mkdir(dirPath, { recursive: true });
  }

  const legacyDirPath = getLegacyProjectLibraryRoot();
  if (path.resolve(legacyDirPath) === path.resolve(dirPath) || !(await exists(legacyDirPath))) {
    return dirPath;
  }

  const entries = await readdir(legacyDirPath, { withFileTypes: true }).catch(() => []);
  await Promise.all(entries.filter((entry) => entry.isDirectory()).map(async (entry) => {
    const fromPath = path.join(legacyDirPath, entry.name);
    const toPath = path.join(dirPath, entry.name);
    if (await exists(toPath)) {
      return;
    }
    await cp(fromPath, toPath, { recursive: true, force: false, errorOnExist: false });
  }));

  return dirPath;
}

function sortStudentsForExport(students: LabProject['students'], projectLayout: LabProject['signinLayout']) {
  const rows = [...students];
  if (projectLayout.sortByGroupName) {
    rows.sort((left, right) => left.groupName.localeCompare(right.groupName, 'zh-Hans-CN') || left.studentId.localeCompare(right.studentId, 'zh-Hans-CN'));
    return rows;
  }
  if (projectLayout.sortByTimeSlot) {
    rows.sort((left, right) => left.timeSlotLabel.localeCompare(right.timeSlotLabel, 'zh-Hans-CN') || left.studentId.localeCompare(right.studentId, 'zh-Hans-CN'));
    return rows;
  }
  if (projectLayout.sortByStudentId) {
    rows.sort((left, right) => left.studentId.localeCompare(right.studentId, 'zh-Hans-CN'));
  }
  return rows;
}

function getSigninVisibleColumns(layout: SigninLayoutSettings) {
  const columns = ['姓名'];
  if (layout.showGrade) columns.push('年级');
  if (layout.showClassName) columns.push('班级');
  if (layout.showGroupName) columns.push('分组');
  if (layout.showAttendance) columns.push('签到');
  return columns;
}

function getSigninColumnWidths(layout: SigninLayoutSettings, visibleColumns: string[]) {
  const baseWidthPreset: Record<string, number> = {
    姓名: 13,
    年级: 7,
    班级: 10,
    分组: 9,
    签到: 9,
  };
  const scale = Math.max(0.8, Math.min(1.5, (layout.blockColumnWidth || 12) / 12));
  return visibleColumns.map((label) => Number(((baseWidthPreset[label] ?? (layout.blockColumnWidth || 12)) * scale).toFixed(2)));
}

async function loadWorkbook(filePath: string) {
  return XLSX.readFile(filePath, { cellDates: false });
}

async function parseStudentWorkbook(filePath: string) {
  const workbook = await loadWorkbook(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: '',
    raw: false,
    blankrows: false,
  }) as unknown[][];

  // try to detect header row by looking for explicit column names or rows with multiple non-empty cells
  const exactColumnKeywords = ['学号', '学生学号', '姓名', '学生姓名', 'name', 'student id', 'studentid', '班级', '年级', '院系', '所属院系', '已选分组信息', '分组', '教师名', '教师'];
  let headerRowIndex = -1;
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i] ?? [];
    const nonEmptyCount = (row as unknown[]).filter((c) => normalizeText(c) !== '').length;
    const hasExact = (row as unknown[]).some((cell) => {
      const t = normalizeText(cell).toLowerCase();
      return exactColumnKeywords.some((kw) => t === kw.toLowerCase() || t.includes(kw.toLowerCase()));
    });
    // prefer rows that explicitly contain column names (and have at least 2 non-empty cells),
    // or rows with several non-empty cells
    if ((hasExact && nonEmptyCount >= 2) || nonEmptyCount >= 4) {
      headerRowIndex = i;
      break;
    }
  }
  // fallback to first non-empty row
  if (headerRowIndex < 0) {
    headerRowIndex = rows.findIndex((row) => row.some((cell) => normalizeText(cell) !== ''));
  }
  const headers = (rows[headerRowIndex] ?? []).map((cell) => normalizeText(cell));
  const dataRows = rows.slice(headerRowIndex + 1).map((row) => {
    const record: Record<string, string> = {};
    headers.forEach((header, index) => {
      if (!header) {
        return;
      }
      record[header] = normalizeText(row[index]);
    });
    return record;
  }).filter((record) => Object.values(record).some((value) => value !== ''));

  return { headers, rows: dataRows };
}

async function parseTimeSlotWorkbook(filePath: string): Promise<ParsedTimeSlot[]> {
  const workbook = await loadWorkbook(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: '',
    raw: false,
    blankrows: false,
  }) as unknown[][];

  const result: ParsedTimeSlot[] = [];
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index] ?? [];
    const label = normalizeText(row[0]);
    if (!label || label === '选项' || label === '时间段') {
      continue;
    }
    const limit = toNumber(row[1]);
    const registered = toNumber(row[2]);
    const students = row.slice(3).map((cell) => normalizeText(cell)).filter(Boolean);
    if (!students.length && !limit && !registered) {
      continue;
    }
    result.push({
      label,
      limit,
      registered,
      students,
      rawRowIndex: index + 1,
    });
  }
  return result;
}

async function openProject(filePath: string): Promise<LabProject> {
  const content = await readFile(filePath, 'utf-8');
  const parsed = JSON.parse(content) as Partial<LabProject>;
  return {
    ...buildDefaultProject(),
    ...parsed,
    columnMappings: parsed.columnMappings ?? {},
    signinLayout: parsed.signinLayout ?? defaultSigninLayout,
    settings: parsed.settings ?? buildDefaultProject().settings,
    sourceFiles: parsed.sourceFiles ?? {},
  };
}

async function saveProject(filePath: string, project: LabProject): Promise<void> {
  await writeFile(filePath, JSON.stringify(project, null, 2), 'utf-8');
}

async function buildUniqueFilePath(dirPath: string, fileName: string) {
  const parsed = path.parse(sanitizeFileNamePart(fileName));
  const baseName = parsed.name || 'attachment';
  const ext = parsed.ext;
  let candidate = path.join(dirPath, `${baseName}${ext}`);
  let index = 2;
  while (await exists(candidate)) {
    candidate = path.join(dirPath, `${baseName}_${index}${ext}`);
    index += 1;
  }
  return candidate;
}

function toImapDate(value?: string) {
  if (!value) return undefined;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? undefined : date;
}

function shouldSaveEmailAttachment(config: EmailImportConfig, attachment: {
  fromName: string;
  subject: string;
  originalFileName: string;
}) {
  const students = config.matchHints ?? [];
  if (!students.length) return true;

  const byName = new Map<string, Array<{ studentId: string; name: string }>>();
  students.forEach((student) => {
    const name = normalizeText(student.name);
    if (!name) return;
    byName.set(name, [...(byName.get(name) ?? []), student]);
  });

  const senderMatches = byName.get(normalizeText(attachment.fromName)) ?? [];
  if (senderMatches.length > 0) return true;

  const fileName = normalizeText(attachment.originalFileName);
  const subject = normalizeText(attachment.subject);
  return students.some((student) => (
    (student.studentId && (fileName.includes(student.studentId) || subject.includes(student.studentId)))
    || (student.name && (fileName.includes(student.name) || subject.includes(student.name)))
  ));
}

async function importEmailAttachments(config: EmailImportConfig): Promise<ImportedEmailAttachment[]> {
  if (!config.host || !config.user || !config.password) {
    throw new Error('请填写 IMAP 服务器、邮箱账号和授权码');
  }
  if (!config.downloadDir) {
    throw new Error('缺少附件保存目录');
  }

  await mkdir(config.downloadDir, { recursive: true });
  const client = new ImapFlow({
    host: config.host,
    port: config.port || 993,
    secure: config.secure !== false,
    auth: {
      user: config.user,
      pass: config.password,
    },
    logger: false,
  });

  const attachments: ImportedEmailAttachment[] = [];
  const knownAttachmentKeys = new Set(config.knownAttachmentKeys ?? []);
  try {
    await client.connect();
    const lock = await client.getMailboxLock(config.mailbox || 'INBOX');
    try {
      const searchQuery: Record<string, unknown> = { all: true };
      const since = toImapDate(config.since);
      const before = toImapDate(config.before);
      if (since) searchQuery.since = since;
      if (before) searchQuery.before = before;
      if (normalizeText(config.subjectKeyword)) searchQuery.subject = normalizeText(config.subjectKeyword);

      const found = await client.search(searchQuery as any, { uid: true });
      const uids = Array.isArray(found) ? found.slice(-Math.max(1, config.limit || 50)).reverse() : [];

      for await (const message of client.fetch(uids, { uid: true, source: true }, { uid: true })) {
        if (!message.source) continue;
        const parsed = await simpleParser(message.source);
        const from = parsed.from?.value?.[0];
        const fromName = normalizeText(from?.name);
        const fromAddress = normalizeText(from?.address);
        const subject = normalizeText(parsed.subject);
        const date = parsed.date?.toISOString() ?? '';
        const messageId = normalizeText(parsed.messageId);

        for (const attachment of parsed.attachments) {
          if (attachment.related) continue;
          const originalFileName = sanitizeFileNamePart(attachment.filename || 'attachment');
          const attachmentKey = [
            messageId || `uid:${Number(message.uid ?? 0)}`,
            originalFileName.toLowerCase(),
            String(attachment.size || 0),
          ].join('|');
          if (knownAttachmentKeys.has(attachmentKey)) {
            continue;
          }
          const shouldSave = shouldSaveEmailAttachment(config, { fromName, subject, originalFileName });
          if (!shouldSave) {
            attachments.push({
              uid: Number(message.uid ?? 0),
              messageId,
              fromName,
              fromAddress,
              subject,
              date,
              originalFileName,
              savedFileName: '',
              savedPath: '',
              contentType: attachment.contentType,
              size: attachment.size,
            });
            continue;
          }
          const prefix = [
            date ? date.slice(0, 10).replace(/-/g, '') : '',
            fromName,
          ].filter(Boolean).map(sanitizeFileNamePart).join('_');
          const savedPath = await buildUniqueFilePath(
            config.downloadDir,
            `${prefix ? `${prefix}_` : ''}${originalFileName}`,
          );
          await writeFile(savedPath, attachment.content);
          attachments.push({
            uid: Number(message.uid ?? 0),
            messageId,
            fromName,
            fromAddress,
            subject,
            date,
            originalFileName,
            savedFileName: path.basename(savedPath),
            savedPath,
            contentType: attachment.contentType,
            size: attachment.size,
          });
        }
      }
    } finally {
      lock.release();
    }
  } finally {
    await client.logout().catch(() => undefined);
  }

  return attachments;
}

async function extractSubmissionText(filePath: string) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.docx') {
    const result = await mammoth.extractRawText({ path: filePath });
    const mammothText = normalizeText(result.value);
    if (compactReviewText(mammothText).length >= 200) {
      return mammothText;
    }
    const fallbackText = await extractDocxXmlText(filePath);
    return compactReviewText(fallbackText).length > compactReviewText(mammothText).length ? fallbackText : mammothText;
  }
  if (ext === '.pdf') {
    const buffer = await readFile(filePath);
    const parser = new PDFParse({ data: buffer });
    try {
      const result = await parser.getText();
      return normalizeText(result.text);
    } finally {
      await parser.destroy().catch(() => undefined);
    }
  }
  if (['.txt', '.md', '.csv'].includes(ext)) {
    return normalizeText(await readFile(filePath, 'utf-8'));
  }
  return '';
}

function decodeXmlText(value: string) {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

async function extractDocxXmlText(filePath: string) {
  const zip = await JSZip.loadAsync(await readFile(filePath));
  const xmlFiles = Object.values(zip.files)
    .filter((file) => !file.dir && /^word\/.+\.xml$/i.test(file.name));
  const parts: string[] = [];
  for (const file of xmlFiles) {
    const xml = await file.async('string');
    const matches = xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g);
    for (const match of matches) {
      parts.push(decodeXmlText(match[1]));
    }
  }
  return normalizeText(parts.join('\n'));
}

function compactReviewText(text: string) {
  return text
    .replace(/\s+/g, '')
    .replace(/[：:]/g, '')
    .replace(/[（(]\d+[）)]/g, '')
    .replace(/[^\u4e00-\u9fa5a-zA-Z0-9？?。.，,、；;：:]/g, '');
}

function mergeIntervals(intervals: Array<[number, number]>) {
  const sorted = intervals
    .map(([start, end]) => [Math.max(0, start), Math.max(start, end)] as [number, number])
    .sort((left, right) => left[0] - right[0]);
  const merged: Array<[number, number]> = [];
  sorted.forEach((current) => {
    const previous = merged[merged.length - 1];
    if (!previous || current[0] > previous[1]) {
      merged.push(current);
    } else {
      previous[1] = Math.max(previous[1], current[1]);
    }
  });
  return merged;
}

function estimateQuestionAnswerAmount(text: string) {
  const compact = compactReviewText(text);
  if (!compact) return 0;

  const sectionMatch = compact.match(/(预习)?思考题(解答)?|问题解答|课后思考|讨论题|思考与讨论/);
  const intervals: Array<[number, number]> = [];
  if (sectionMatch?.index !== undefined) {
    const start = sectionMatch.index;
    const tail = compact.slice(start + 8);
    const endMatch = tail.match(/参考文献|附录|实验总结|总结|结论|心得/);
    const end = endMatch?.index !== undefined ? start + 8 + endMatch.index : compact.length;
    intervals.push([start, end]);
  }

  const markerPattern = /[？?]|为什么|为何|如何|怎样|怎么|是否|能否|原因|解释|说明|证明|推导|分析|解答|答案|答/g;
  let match: RegExpExecArray | null;
  while ((match = markerPattern.exec(compact)) !== null) {
    intervals.push([match.index - 80, match.index + 260]);
  }

  const merged = mergeIntervals(intervals);
  const coveredLength = merged.reduce((sum, [start, end]) => sum + Math.max(0, end - start), 0);
  const answerMarkers = compact.match(/解答|答案|根据|因为|所以|因此|可知|说明|证明|推导/g)?.length ?? 0;
  const questionMarkers = compact.match(/[1-9][.、．]|[？?]|问题[一二三四五六七八九十\d]/g)?.length ?? 0;
  return Math.min(compact.length, coveredLength) + answerMarkers * 35 + questionMarkers * 25;
}

type ReviewMetric = {
  input: AutoReviewSubmissionInput;
  totalAmount: number;
  qaAmount: number;
  raw: number;
  readable: boolean;
  details: string[];
};

function normalizedLog(value: number, max: number) {
  return clamp(Math.log1p(Math.max(0, value)) / Math.log1p(max), 0, 1);
}

async function measureSubmission(input: AutoReviewSubmissionInput): Promise<ReviewMetric> {
  const details: string[] = [];
  const ext = path.extname(input.submissionFile).toLowerCase();
  let text = '';
  try {
    await stat(input.submissionFile);
    text = await extractSubmissionText(input.submissionFile);
  } catch {
    details.push('文件无法读取或文本提取失败');
    return { input, totalAmount: 0, qaAmount: 0, raw: 0, readable: false, details };
  }

  const compact = compactReviewText(text);
  const totalAmount = compact.length;
  const qaAmount = estimateQuestionAnswerAmount(text);
  const readable = totalAmount > 0;
  const raw = normalizedLog(totalAmount, 4200) * 0.45 + normalizedLog(qaAmount, 2200) * 0.55;
  details.push(`总有效内容量：${totalAmount}`);
  details.push(`问答/思考类内容量：${qaAmount}`);
  if (!['.docx', '.pdf', '.txt', '.md', '.csv'].includes(ext)) {
    details.push(`文件类型 ${ext || '未知'} 不适合自动文本评阅`);
  }
  return { input, totalAmount, qaAmount, raw, readable, details };
}

const referenceSubmittedScores = [
  77, 84, 81, 81, 84, 86, 85, 85, 82, 83, 79, 80, 83, 83, 80, 86, 82, 85,
  85, 84, 78, 84, 85, 80, 87, 86, 78, 86, 84, 80, 85, 75, 83, 85, 88, 87,
  80, 82, 87, 81, 85, 86, 82, 78, 83, 83, 85, 76, 83, 90, 90, 82, 85,
].sort((left, right) => left - right);

function quantile(values: number[], ratio: number) {
  if (!values.length) return 82;
  const position = clamp(ratio, 0, 1) * (values.length - 1);
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) return values[lower];
  return values[lower] + (values[upper] - values[lower]) * (position - lower);
}

function mapReferenceScoreToTargetRange(score: number) {
  const minReference = referenceSubmittedScores[0] ?? 75;
  const maxReference = referenceSubmittedScores[referenceSubmittedScores.length - 1] ?? 90;
  if (maxReference <= minReference) return score;
  return 72 + ((score - minReference) / (maxReference - minReference)) * 18;
}

async function autoReviewSubmissions(inputs: AutoReviewSubmissionInput[]): Promise<AutoReviewResult[]> {
  const missing = inputs.filter((input) => !input.submissionFile || input.submissionStatus === 'missing');
  const submitted = inputs.filter((input) => input.submissionFile && input.submissionStatus !== 'missing');
  const metrics: ReviewMetric[] = [];
  for (const input of submitted) {
    metrics.push(await measureSubmission(input));
  }

  const ranked = [...metrics].sort((left, right) => left.raw - right.raw);

  const results: AutoReviewResult[] = missing.map((input) => ({
    studentId: input.studentId,
    score: 0,
    remark: '自动评阅：未检测到提交文件，记 0 分。',
    details: ['未检测到提交文件'],
  }));

  ranked.forEach((metric, index) => {
    const percentile = ranked.length <= 1 ? 0.5 : index / (ranked.length - 1);
    let mappedScore = mapReferenceScoreToTargetRange(quantile(referenceSubmittedScores, percentile));
    if (!metric.readable) {
      mappedScore = Math.min(mappedScore, 75);
    }
    const finalScore = Math.round(clamp(mappedScore, 72, 90));
    results.push({
      studentId: metric.input.studentId,
      score: finalScore,
      remark: `自动评阅：总有效内容量 ${metric.totalAmount}；问答/思考类内容量 ${metric.qaAmount}；按参考成绩分布映射为 ${finalScore} 分。`,
      details: [
        ...metric.details,
        `全班相对位置：${Math.round(percentile * 100)}%`,
        `参考成绩分布映射：${mappedScore.toFixed(1)}`,
        `最终成绩：${finalScore}`,
      ],
    });
  });

  return results.sort((left, right) => inputs.findIndex((input) => input.studentId === left.studentId) - inputs.findIndex((input) => input.studentId === right.studentId));
}

function getSortedStudents(project: LabProject, fullRecord: boolean) {
  const rows = sortStudentsForExport(project.students, project.signinLayout);
  if (fullRecord) {
    return rows.map((student) => ({
      studentId: student.studentId,
      name: student.name,
      grade: student.grade,
      className: student.className,
      department: student.department,
      groupName: student.groupName,
      timeSlotLabel: student.timeSlotLabel,
      submissionStatus: student.submissionStatus,
      score: student.score,
      remark: student.remark,
    }));
  }
  return rows.map((student) => ({
    studentId: student.studentId,
    name: student.name,
    score: student.score ?? '',
  }));
}

async function exportGradesWorkbook(filePath: string, payload: { project: LabProject; fullRecord: boolean }) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('成绩单');
  const rows = getSortedStudents(payload.project, payload.fullRecord);
  const columns = payload.fullRecord
    ? [
      { header: '学号', key: 'studentId', width: 16 },
      { header: '姓名', key: 'name', width: 12 },
      { header: '年级', key: 'grade', width: 10 },
      { header: '班级', key: 'className', width: 14 },
      { header: '院系', key: 'department', width: 18 },
      { header: '组名', key: 'groupName', width: 12 },
      { header: '实验时间', key: 'timeSlotLabel', width: 22 },
      { header: '提交状态', key: 'submissionStatus', width: 12 },
      { header: '分数', key: 'score', width: 10 },
      { header: '备注', key: 'remark', width: 22 },
    ]
    : [
      { header: '学号', key: 'studentId', width: 16 },
      { header: '姓名', key: 'name', width: 12 },
      { header: '分数', key: 'score', width: 10 },
    ];

  worksheet.columns = columns;
  rows.forEach((row) => {
    worksheet.addRow(row);
  });

  worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
  worksheet.autoFilter = { from: 'A1', to: worksheet.getCell(1, columns.length).address };
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  worksheet.eachRow((row, rowNumber) => {
    row.alignment = { vertical: 'middle', horizontal: 'center' };
    row.height = 22;
    if (rowNumber > 1 && payload.fullRecord) {
      const scoreCell = row.getCell(9);
      if (scoreCell.value === null || scoreCell.value === '') {
        scoreCell.font = { color: { argb: 'FFB91C1C' } };
      }
    }
  });

  await workbook.xlsx.writeFile(filePath);
}

async function exportSigninWorkbook(filePath: string, payload: { project: LabProject; layout: SigninLayoutSettings }) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('签到表');
  const students = sortStudentsForExport(payload.project.students, payload.layout);
  const effectiveFontSize = Math.max(10, Math.min(12, payload.layout.fontSize || 11));
  const effectiveRowHeight = Math.max(18, Math.min(24, payload.layout.rowHeight || 20));

  worksheet.pageSetup.orientation = payload.layout.orientation;
  worksheet.pageSetup.paperSize = 9;
  worksheet.pageSetup.fitToPage = false;
  worksheet.pageSetup.scale = 100;
  worksheet.pageSetup.horizontalCentered = true;
  worksheet.pageSetup.margins = { left: 0.2, right: 0.2, top: 0.28, bottom: 0.28, header: 0.15, footer: 0.15 };
  worksheet.properties.defaultRowHeight = effectiveRowHeight;

  const visibleColumns = getSigninVisibleColumns(payload.layout);
  const columnWidths = getSigninColumnWidths(payload.layout, visibleColumns);

  const colsPerBlock = visibleColumns.length; // columns per block
  const blocksPerRow = Math.max(1, payload.layout.blocksPerPage);
  const blockWidth = colsPerBlock + 1; // add spacing
  let maxUsedRow = 1;

  const writeBlock = (startRow: number, startCol: number, _title: string, rowsToWrite: typeof students) => {
    const headerRow = worksheet.getRow(startRow);
    maxUsedRow = Math.max(maxUsedRow, startRow);
    visibleColumns.forEach((label, index) => {
      const cell = headerRow.getCell(startCol + index);
      cell.value = label;
      cell.font = { bold: true, size: effectiveFontSize };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF2F8' } };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' },
      };
    });

    rowsToWrite.slice(0, payload.layout.maxStudentsPerBlock).forEach((student, rowOffset) => {
      const row = worksheet.getRow(startRow + 1 + rowOffset);
      const values: Array<string | number> = [];
      values.push(student.name);
      if (payload.layout.showGrade) values.push(student.grade);
      if (payload.layout.showClassName) values.push(student.className);
      if (payload.layout.showGroupName) values.push(student.groupName);
      if (payload.layout.showAttendance) values.push('');

      values.forEach((value, index) => {
        const cell = row.getCell(startCol + index);
        cell.value = value;
        cell.font = { size: effectiveFontSize };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' },
        };
      });
      row.height = effectiveRowHeight;
      maxUsedRow = Math.max(maxUsedRow, startRow + 1 + rowOffset);
    });
  };

  const groups = payload.layout.separatePageByTimeSlot
    ? payload.project.timeSlots.map((slot) => ({ title: `${slot.label}${slot.groupName ? ` · ${slot.groupName}` : ''}`, students: students.filter((s) => s.timeSlotId === slot.id) }))
    : [{ title: `${payload.project.courseName} 签到表`, students }];

  let rowCursor = 1;
  for (const group of groups) {
    const list = group.students ?? [];
    const totalBlocks = Math.max(1, Math.ceil(list.length / payload.layout.maxStudentsPerBlock));
    let blockIndex = 0;
    while (blockIndex < totalBlocks) {
      // each page contains up to blocksPerRow blocks horizontally
      const blocksOnPage = Math.min(blocksPerRow, totalBlocks - blockIndex);
      for (let b = 0; b < blocksOnPage; b += 1) {
        const startCol = b * blockWidth + 1;
        const start = blockIndex * payload.layout.maxStudentsPerBlock + b * payload.layout.maxStudentsPerBlock;
        const blockStudents = list.slice(start, start + payload.layout.maxStudentsPerBlock);
        const title = payload.layout.showHeaderMeta ? `${group.title}${blockStudents.length ? ` · 第${blockIndex + b + 1}块` : ''}` : group.title;
        writeBlock(rowCursor, startCol, title, blockStudents);
      }
      // advance row cursor by rows occupied plus spacing（无标题行）
      rowCursor += payload.layout.maxStudentsPerBlock + 2;
      blockIndex += blocksOnPage;
      // add page break after each printed page when not the last
      if (blockIndex < totalBlocks || groups.indexOf(group) < groups.length - 1) {
        (worksheet as any).addPageBreak(rowCursor);
        rowCursor += 2; // margin after page break
      }
    }
    // small gap between groups if no page break was added
    rowCursor += 1;
  }

  for (let b = 0; b < blocksPerRow; b += 1) {
    const startCol = b * blockWidth + 1;
    visibleColumns.forEach((_, idx) => {
      worksheet.getColumn(startCol + idx).width = columnWidths[idx];
    });
    worksheet.getColumn(startCol + colsPerBlock).width = 1.2;
  }

  const lastCol = blocksPerRow * blockWidth;
  worksheet.pageSetup.printArea = `A1:${worksheet.getCell(maxUsedRow, lastCol).address}`;

  await workbook.xlsx.writeFile(filePath);
}

async function exportSigninPdf(filePath: string, payload: { project: LabProject; layout: SigninLayoutSettings }) {
  const { project, layout } = payload;
  const students = sortStudentsForExport(project.students, layout);

  const groups = layout.separatePageByTimeSlot
    ? project.timeSlots.map((slot) => ({ title: `${slot.label}${slot.groupName ? ` · ${slot.groupName}` : ''}`, students: students.filter((s) => s.timeSlotId === slot.id) }))
    : [{ title: `${project.courseName || '签到表'}`, students }];

  const pageSize = layout.orientation === 'landscape' ? 'landscape' : 'portrait';
  const visibleColumns = getSigninVisibleColumns(layout);
  const columnWidths = getSigninColumnWidths(layout, visibleColumns);
  const widthTotal = columnWidths.reduce((sum, value) => sum + value, 0) || 1;
  const columnPercents = columnWidths.map((value) => (value / widthTotal) * 100);

  const escapeHtml = (s: any) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  let html = `<!doctype html><html><head><meta charset="utf-8"><style>
    @page { size: A4 ${pageSize}; margin: 6mm; }
    body{font-family: Arial, Helvetica, "Microsoft Yahei", sans-serif; color:#111;margin:0;padding:0}
    .page{page-break-after: always;padding:2px}
    .blocks{display:flex;flex-wrap:wrap;gap:6px}
    .block{border:1px solid #e0e0e0;padding:3px;box-sizing:border-box;width:calc(${100 / Math.max(1, payload.layout.blocksPerPage)}% - 6px)}
    table{width:100%;border-collapse:collapse;margin:0;table-layout:fixed}
    thead th{background:#f3f6f9;font-weight:700;padding:3px;border:1px solid #eee;font-size:${Math.max(10, Math.min(12, layout.fontSize || 11))}px}
    tbody tr{height:7.1mm}
    td{border:1px solid #f0f0f0;padding:3px;font-size:${Math.max(10, Math.min(12, layout.fontSize || 11))}px;text-align:left;line-height:1.2}
  </style></head><body>`;

  for (const group of groups) {
    const list = group.students ?? [];
    const maxPer = layout.maxStudentsPerBlock || 30;
    const blocks: any[] = [];
    for (let i = 0; i < list.length; i += maxPer) blocks.push(list.slice(i, i + maxPer));

    // render pages where each page contains up to layout.blocksPerPage blocks
    const blocksPerPage = Math.max(1, layout.blocksPerPage || 1);
    for (let p = 0; p < blocks.length; p += blocksPerPage) {
      html += `<div class="page"><div class="blocks">`;
      const slice = blocks.slice(p, p + blocksPerPage);
      for (let bi = 0; bi < slice.length; bi++) {
        const block = slice[bi];
        html += `<div class="block"><table><colgroup>${columnPercents.map((percent) => `<col style="width:${percent}%">`).join('')}</colgroup><thead><tr>`;
        for (const col of visibleColumns) html += `<th>${escapeHtml(col)}</th>`;
        html += `</tr></thead><tbody>`;
        for (let r = 0; r < maxPer; r++) {
          const s = block[r];
          html += `<tr>`;
          if (s) {
            html += `<td>${escapeHtml(s.name)}</td>`;
            if (layout.showGrade) html += `<td>${escapeHtml(s.grade)}</td>`;
            if (layout.showClassName) html += `<td>${escapeHtml(s.className)}</td>`;
            if (layout.showGroupName) html += `<td>${escapeHtml(s.groupName)}</td>`;
            if (layout.showAttendance) html += `<td></td>`;
          } else {
            for (let c = 0; c < visibleColumns.length; c++) html += `<td>&nbsp;</td>`;
          }
          html += `</tr>`;
        }
        html += `</tbody></table></div>`;
      }
      html += `</div></div>`;
    }
  }

  html += `</body></html>`;
  try {
    const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] });
    try {
      const page = await browser.newPage();
      await page.setViewport({ width: 1400, height: 2000, deviceScaleFactor: 1 });
      await page.setContent(html, { waitUntil: 'networkidle0' });
      await page.evaluate(() => (document.fonts ? document.fonts.ready : Promise.resolve()));
      const data = await page.pdf({
        path: filePath,
        format: 'A4',
        printBackground: true,
        landscape: layout.orientation === 'landscape',
        preferCSSPageSize: true,
        margin: { top: '6mm', right: '6mm', bottom: '6mm', left: '6mm' },
      });
      if (!data || !data.length) {
        throw new Error('PDF output is empty');
      }
    } finally {
      await browser.close();
    }
  } finally {
    // no-op
  }
}

app.whenReady().then(() => {
  ipcMain.handle('dialog:selectOpenFile', async (_event, options) => {
    const result = await dialog.showOpenDialog({
      title: options?.title,
      properties: ['openFile'],
      filters: options?.filters ?? [],
    });
    return result.canceled ? null : result.filePaths[0] ?? null;
  });

  ipcMain.handle('dialog:selectSaveFile', async (_event, options) => {
    const result = await dialog.showSaveDialog({
      title: options?.title,
      defaultPath: options?.defaultPath,
      filters: options?.filters ?? [],
    });
    return result.canceled ? null : result.filePath ?? null;
  });

  ipcMain.handle('dialog:selectDirectory', async (_event, options) => {
    const result = await dialog.showOpenDialog({
      title: options?.title,
      defaultPath: options?.defaultPath,
      properties: ['openDirectory', 'createDirectory'],
    });
    return result.canceled ? null : result.filePaths[0] ?? null;
  });

  ipcMain.handle('file:readText', async (_event, filePath: string) => readFile(filePath, 'utf-8'));
  ipcMain.handle('file:writeText', async (_event, filePath: string, content: string) => {
    await writeFile(filePath, content, 'utf-8');
  });
  ipcMain.handle('file:copy', async (_event, fromPath: string, toPath: string) => {
    await mkdir(path.dirname(toPath), { recursive: true });
    await copyFile(fromPath, toPath);
  });
  ipcMain.handle('file:openExternal', async (_event, filePath: string) => {
    await shell.openPath(filePath);
  });

  ipcMain.handle('file:exists', async (_event, filePath: string) => {
    try {
      await access(filePath);
      return true;
    } catch {
      return false;
    }
  });

  ipcMain.handle('dir:createTree', async (_event, dirPath: string) => {
    await mkdir(dirPath, { recursive: true });
    return dirPath;
  });

  ipcMain.handle('dir:rename', async (_event, fromPath: string, toPath: string) => {
    await mkdir(path.dirname(toPath), { recursive: true });
    await rename(fromPath, toPath);
    return toPath;
  });

  ipcMain.handle('dir:listDirectories', async (_event, dirPath: string) => {
    try {
      const entries = await readdir(dirPath, { withFileTypes: true });
      return entries.filter((d) => d.isDirectory()).map((d) => ({ name: d.name, path: path.join(dirPath, d.name) }));
    } catch {
      return [];
    }
  });

  ipcMain.handle('project:getLibraryRoot', async () => {
    return ensureProjectLibraryRoot();
  });

  ipcMain.handle('project:open', async (_event, filePath: string) => openProject(filePath));
  ipcMain.handle('project:save', async (_event, filePath: string, project: LabProject) => saveProject(filePath, project));
  ipcMain.handle('excel:parseStudents', async (_event, filePath: string) => parseStudentWorkbook(filePath));
  ipcMain.handle('excel:parseTimeSlots', async (_event, filePath: string) => parseTimeSlotWorkbook(filePath));
  ipcMain.handle('excel:exportSignin', async (_event, filePath: string, payload) => exportSigninWorkbook(filePath, payload));
  ipcMain.handle('excel:exportSigninPdf', async (_event, filePath: string, payload) => exportSigninPdf(filePath, payload));
  ipcMain.handle('excel:exportGrades', async (_event, filePath: string, payload) => exportGradesWorkbook(filePath, payload));
  ipcMain.handle('email:importAttachments', async (_event, config: EmailImportConfig) => importEmailAttachments(config));
  ipcMain.handle('grading:autoReview', async (_event, inputs: AutoReviewSubmissionInput[]) => autoReviewSubmissions(inputs));
  ipcMain.handle('dir:listFiles', async (_event, dirPath: string) => {
    try {
      const entries = await readdir(dirPath, { withFileTypes: true });
      const files = entries
        .filter((d: any) => d.isFile())
        .map((d: any) => ({ name: d.name, path: path.join(dirPath, d.name) }));
      return files;
    } catch (err) {
      return [];
    }
  });

  const window = new BrowserWindow({
    width: 1480,
    height: 920,
    minWidth: 1240,
    minHeight: 760,
    title: '实验课签到与成绩管理助手',
    backgroundColor: '#f5f7fb',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  if (isDevServerRunning()) {
    window.loadURL(process.env.VITE_DEV_SERVER_URL as string);
    if (process.env.LAB_OPEN_DEVTOOLS === '1') {
      window.webContents.openDevTools({ mode: 'detach' });
    }
  } else {
    window.loadFile(getRendererEntry());
  }
  // 自动测试模式：若设置 AUTOTEST_EXPORT_SIGNIN=1，则在启动后自动生成 PDF 并退出
  if (process.env.AUTOTEST_EXPORT_SIGNIN === '1') {
    (async () => {
      try {
        const samplePath = path.join(__dirname, '..', 'lab-project-import.json');
        let projectData: any = null;
        try {
          projectData = JSON.parse(await readFile(samplePath, 'utf-8')) as Partial<LabProject>;
        } catch (e) {
          projectData = null;
        }
        const mergedProject: LabProject = {
          ...buildDefaultProject(),
          ...(projectData ?? {}),
          signinLayout: (projectData && projectData.signinLayout) ? projectData.signinLayout : defaultSigninLayout,
        };
        const outPath = path.join(__dirname, '..', 'output', 'signin_export_electron.pdf');
        await exportSigninPdf(outPath, { project: mergedProject, layout: mergedProject.signinLayout });
        console.log('Auto export finished:', outPath);
      } catch (err) {
        console.error('Auto export failed:', err);
      } finally {
        app.quit();
      }
    })();
  }

  // 开发交互自动点击导出（模拟真实用户在 Electron 窗口点击）
  if (process.env.AUTOTEST_MANUAL === '1') {
    window.webContents.once('did-finish-load', async () => {
      try {
        const savePath = path.join(__dirname, '..', 'output', 'signin_manual.pdf');
        const script = `(async ()=>{\n          try{\n            // 尝试从开发服务器获取示例项目文件\n            let project = null;\n            try{ const r = await fetch('/lab-project-import.json'); if(r.ok) project = await r.json(); }catch(e){}\n            if(!project) project = { courseName: '自动测试项目', signinLayout: ${JSON.stringify(defaultSigninLayout)}, students: [] };\n            const layout = project.signinLayout || ${JSON.stringify(defaultSigninLayout)};\n            if(window.labApi && typeof window.labApi.exportSigninPdf === 'function'){
              await window.labApi.exportSigninPdf('${savePath.replace(/\\/g, '\\\\')}', { project, layout });\n              return 'exported';\n            }\n            return 'no-api';\n          }catch(e){ return 'error:'+String(e); }\n        })()`;
        const result = await window.webContents.executeJavaScript(script, true);
        // 如果上面的直接调用未能触发（例如没有暴露 API），尝试导航到 /signin 并自动点击页面上的导出按钮
        let clickResult = null;
        try {
          const clickScript = `(async ()=>{\n            try{ location.href='/signin'; const wait=(ms)=>new Promise(r=>setTimeout(r,ms)); await wait(500); for(let i=0;i<10;i++){ const btn = Array.from(document.querySelectorAll('button')).find(b=>/导出为 PDF|导出签到表（PDF）/.test(b.innerText)); if(btn){ btn.click(); return 'clicked'; } await wait(300);} return 'not-found'; }catch(e){ return 'error:'+String(e);} })()`;
          clickResult = await window.webContents.executeJavaScript(clickScript, true);
        } catch (e) {
          clickResult = `error:${String(e)}`;
        }
        console.log('Auto-export (manual) result:', result, 'clickAttempt:', clickResult);
      } catch (e) {
        console.error('Auto-manual failed:', e);
      }
      setTimeout(() => app.quit(), 3000);
    });
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
