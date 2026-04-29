import { computed, reactive } from 'vue';
import { defineStore } from 'pinia';
import type { AutoReviewResult, ColumnMapping, EmailAttachmentMatch, EmailImportConfig, ImportedEmailAttachment, LabProject, MatchedStudentIssue, StudentImportRow, StudentRecord } from '@/types/domain';
import {
  buildDefaultProject,
  buildStudentRecords,
  buildTimeSlotRecords,
  findDuplicateGroupNames,
  generateGroupNames,
  inferStudentColumnMapping,
  matchStudents,
  getSortedStudentsByLayout,
  normalizeText,
} from '@/services/project';
import {
  autoReviewSubmissions,
  exportGradesWorkbook,
  exportSigninWorkbook,
  copyFile,
  createDirectory,
  fileExists,
  getProjectLibraryRoot,
  importEmailAttachments,
  listDirectoriesInDirectory,
  openProject,
  parseStudentWorkbook,
  parseTimeSlotWorkbook,
  renameDirectory,
  saveProject,
  selectOpenFile,
  selectSaveFile,
  selectDirectory,
  listFilesInDirectory,
} from '@/services/api';

function cloneProject(project: LabProject): LabProject {
  return JSON.parse(JSON.stringify(project)) as LabProject;
}

function sanitizeFileNamePart(value: string) {
  return String(value ?? '')
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, ' ')
    .replace(/^\.+/, '')
    .replace(/\.+$/, '');
}

function buildDefaultProjectFileName(project: LabProject) {
  const parts = [sanitizeFileNamePart(project.courseName), sanitizeFileNamePart(project.term)].filter(Boolean);
  return `${parts.length ? parts.join(' - ') : 'lab-project'}.json`;
}

function buildDefaultProjectFolderName(project: LabProject) {
  const parts = [sanitizeFileNamePart(project.courseName), sanitizeFileNamePart(project.term)].filter(Boolean);
  return `${parts.length ? parts.join(' - ') : 'lab-project'}`;
}

function normalizeDirLike(fullPath: string) {
  return String(fullPath ?? '').trim().replace(/[\\/]+$/, '');
}

function normalizePathLike(fullPath: string) {
  return String(fullPath ?? '').trim().replace(/\\/g, '/');
}

function dirnameLike(fullPath: string) {
  const value = String(fullPath ?? '').trim();
  if (!value) return '';
  const normalized = value.replace(/\\/g, '/');
  const index = normalized.lastIndexOf('/');
  if (index < 0) return '';
  return normalized.slice(0, index).replace(/\//g, '\\');
}

function joinLike(baseDir: string, fileName: string) {
  const dir = String(baseDir ?? '').trim();
  const name = String(fileName ?? '').trim();
  if (!dir) return name;
  const separator = dir.includes('\\') || /^[A-Za-z]:/.test(dir) ? '\\' : '/';
  const normalizedDir = dir.replace(/[\\/]+$/, '');
  return `${normalizedDir}${separator}${name}`;
}

function joinPathLike(baseDir: string, ...segments: string[]) {
  return [baseDir, ...segments].filter(Boolean).reduce((acc, segment) => joinLike(acc, segment), '');
}

function basenameLike(fullPath: string) {
  const value = String(fullPath ?? '').trim().replace(/\\/g, '/');
  const index = value.lastIndexOf('/');
  return index >= 0 ? value.slice(index + 1) : value;
}

function stripExtLike(fileName: string) {
  const value = String(fileName ?? '').trim();
  const index = value.lastIndexOf('.');
  return index > 0 ? value.slice(0, index) : value;
}

function getProjectHierarchyRoot(project: LabProject) {
  return normalizeDirLike(project.settings.projectRootDir || project.settings.defaultExportDir || '');
}

function buildProjectSavePath(project: LabProject, currentPath: string) {
  const fileName = buildDefaultProjectFileName(project);
  const baseDir = currentPath
    ? dirnameLike(currentPath)
    : joinPathLike(getProjectHierarchyRoot(project), 'project');
  return joinLike(baseDir, fileName);
}

  function buildProjectSubmissionDir(project: LabProject) {
  return joinPathLike(getProjectHierarchyRoot(project), 'submissions', 'email');
}

function buildProjectExportPath(project: LabProject, category: 'signin' | 'grades', extension: 'pdf' | 'xlsx') {
  return joinLike(
    joinPathLike(getProjectHierarchyRoot(project), 'exports', category),
    `${buildDefaultProjectFileName(project).replace(/\.json$/i, '')}.${extension}`,
  );
}

type ProjectDirectoryItem = { name: string; path: string };

function buildDefaultEmailImportConfig(): EmailImportConfig {
  return {
    host: '',
    port: 993,
    secure: true,
    user: '',
    password: '',
    mailbox: 'INBOX',
    since: '',
    before: '',
    subjectKeyword: '',
    limit: 50,
    downloadDir: '',
  };
}

function buildBlankProject(): LabProject {
  const project = buildDefaultProject();
  project.courseName = '';
  project.term = '';
  project.teacher = '';
  project.students = [];
  project.timeSlots = [];
  project.columnMappings = {};
  project.columnMappingPresets = [];
  project.settings.projectRootDir = '';
  project.settings.defaultExportDir = '';
  project.sourceFiles = {};
  project.importedEmailAttachmentKeys = [];
  return project;
}

export const useProjectStore = defineStore('project', () => {
  const project = reactive<LabProject>(buildBlankProject());
  const projectDir = reactive({ value: '' });
  const studentImport = reactive({
    filePath: '',
    headers: [] as string[],
    rows: [] as StudentImportRow[],
    mapping: {} as ColumnMapping,
    visible: false,
    errorMessage: '',
  });
  const timeSlotImport = reactive({
    filePath: '',
    visible: false,
  });
  const reports = reactive({
    dirPath: '',
    files: [] as Array<{ name: string; path: string }>,
    visible: false,
  });
  const emailImport = reactive({
    visible: false,
    loading: false,
    config: buildDefaultEmailImportConfig(),
    results: [] as EmailAttachmentMatch[],
    errorMessage: '',
  });
  const autoReview = reactive({
    loading: false,
    results: [] as AutoReviewResult[],
    errorMessage: '',
  });
  const issues = reactive<MatchedStudentIssue[]>([]);

  const stats = computed(() => {
    const matchedCount = project.students.filter((student) => Boolean(student.timeSlotId)).length;
    const submittedCount = project.students.filter((student) => student.submissionStatus !== 'missing').length;
    const scoredCount = project.students.filter((student) => student.score !== null && student.score !== undefined).length;
    return {
      studentCount: project.students.length,
      matchedCount,
      submittedCount,
      scoredCount,
      timeSlotCount: project.timeSlots.length,
      issueCount: issues.length,
    };
  });

  function applyProject(nextProject: LabProject) {
    Object.assign(project, cloneProject(nextProject));
    const matched = matchStudents(project);
    project.students = matched.students;
    project.timeSlots = matched.timeSlots;
    issues.splice(0, issues.length, ...matched.issues);
    // restore report dir/files if present in project
    reports.dirPath = resolveProjectPath(project.sourceFiles.reportDir ?? 'submissions/email');
    reports.files.splice(0, reports.files.length);
    emailImport.results.splice(0, emailImport.results.length);
    emailImport.errorMessage = '';
    autoReview.results.splice(0, autoReview.results.length);
    autoReview.errorMessage = '';
    Object.assign(emailImport.config, buildDefaultEmailImportConfig(), project.settings.emailImportConfig ?? {});
  }

  function snapshot(): LabProject {
    return cloneProject(project);
  }

  function currentProjectFolderName() {
    return buildDefaultProjectFolderName(project);
  }

  function currentProjectRoot() {
    return normalizeDirLike(project.settings.projectRootDir || project.settings.defaultExportDir || '');
  }

  async function ensureProjectLibraryRoot() {
    const root = normalizeDirLike(await getProjectLibraryRoot());
    project.settings.projectRootDir = root;
    return root;
  }

  function currentProjectDir() {
    return projectDir.value;
  }

  function projectJsonPath(dirPath: string) {
    return joinLike(dirPath, 'project.json');
  }

  function resolveProjectPath(storedPath: string) {
    const value = String(storedPath ?? '').trim();
    if (!value) return '';
    if (/^[A-Za-z]:[\\/]/.test(value) || value.startsWith('/') || value.startsWith('\\')) {
      return value;
    }
    return joinLike(currentProjectDir(), value);
  }

  function toProjectRelativePath(absPath: string) {
    const projectRoot = normalizePathLike(currentProjectDir());
    const value = normalizePathLike(absPath);
    if (!projectRoot || !value) return absPath;
    const prefix = `${projectRoot}/`;
    if (value.toLowerCase().startsWith(prefix.toLowerCase())) {
      return value.slice(prefix.length).replace(/\\/g, '/');
    }
    return absPath;
  }

  function hydrateProjectSourceFiles(baseDir: string, sourceFiles: LabProject['sourceFiles']) {
    const resolveIfRelative = (value?: string) => {
      if (!value) return '';
      if (/^[A-Za-z]:[\\/]/.test(value) || value.startsWith('/') || value.startsWith('\\')) {
        return value;
      }
      return joinLike(baseDir, value);
    };
    return {
      studentList: resolveIfRelative(sourceFiles.studentList),
      timeSlots: resolveIfRelative(sourceFiles.timeSlots),
      reportDir: resolveIfRelative(sourceFiles.reportDir),
    };
  }

  function makeProjectTree(dirPath: string) {
    const importsDir = joinLike(dirPath, 'imports');
    const submissionsDir = joinLike(dirPath, 'submissions');
    const emailSubmissionsDir = joinLike(submissionsDir, 'email');
    const exportsDir = joinLike(dirPath, 'exports');
    const signinDir = joinLike(exportsDir, 'signin');
    const gradesDir = joinLike(exportsDir, 'grades');
    const projectDirPath = joinLike(dirPath, 'project');
    return { importsDir, submissionsDir, emailSubmissionsDir, exportsDir, signinDir, gradesDir, projectDirPath };
  }

  async function persistProject() {
    if (!projectDir.value) {
      return;
    }
    const exportProject = snapshot();
    exportProject.settings.emailImportConfig = {
      ...emailImport.config,
      downloadDir: '',
      knownAttachmentKeys: [],
      matchHints: [],
    };
    exportProject.sourceFiles = {
      ...exportProject.sourceFiles,
      studentList: toProjectRelativePath(exportProject.sourceFiles.studentList ?? ''),
      timeSlots: toProjectRelativePath(exportProject.sourceFiles.timeSlots ?? ''),
      reportDir: toProjectRelativePath(exportProject.sourceFiles.reportDir ?? ''),
    };
    await saveProject(projectJsonPath(projectDir.value), exportProject);
  }

  async function createNewProject(): Promise<boolean> {
    return false;
  }

  async function createProjectFromMetadata(courseName: string, term: string): Promise<boolean> {
    const nextProject = buildBlankProject();
    nextProject.courseName = normalizeText(courseName);
    nextProject.term = normalizeText(term);

    if (!nextProject.courseName || !nextProject.term) {
      throw new Error('课程名称和时间不能为空');
    }

    const baseRoot = currentProjectRoot() || await ensureProjectLibraryRoot();
    if (!baseRoot) {
      return false;
    }

    const targetDir = joinLike(baseRoot, buildDefaultProjectFolderName(nextProject));
    if (await fileExists(targetDir)) {
      throw new Error(`项目文件夹已存在：${basenameLike(targetDir)}`);
    }

    await createDirectory(targetDir);
    const tree = makeProjectTree(targetDir);
    await createDirectory(tree.importsDir);
    await createDirectory(tree.submissionsDir);
    await createDirectory(tree.emailSubmissionsDir);
    await createDirectory(tree.signinDir);
    await createDirectory(tree.gradesDir);

    nextProject.settings.projectRootDir = baseRoot;
    nextProject.settings.defaultExportDir = '';
    nextProject.sourceFiles = {
      studentList: '',
      timeSlots: '',
      reportDir: 'submissions/email',
    };
    nextProject.importedEmailAttachmentKeys = [];

    projectDir.value = targetDir;
    applyProject(nextProject);
    await persistProject();
    return true;
  }

  async function openProjectFromDirectory(dirPath: string): Promise<boolean> {
    if (!dirPath) {
      return false;
    }
    const filePath = projectJsonPath(dirPath);
    const loaded = await openProject(filePath);
    projectDir.value = dirPath;
    loaded.sourceFiles = hydrateProjectSourceFiles(dirPath, loaded.sourceFiles);
    if (!loaded.sourceFiles.reportDir || normalizePathLike(loaded.sourceFiles.reportDir).replace(/\/+$/, '').toLowerCase().endsWith('/submissions')) {
      loaded.sourceFiles.reportDir = joinLike(dirPath, 'submissions/email');
    }
    applyProject(loaded);
    project.settings.projectRootDir = dirnameLike(dirPath);
    await createDirectory(makeProjectTree(dirPath).emailSubmissionsDir);
    return true;
  }

  async function listAvailableProjects(): Promise<ProjectDirectoryItem[]> {
    const root = await ensureProjectLibraryRoot();
    const dirs = await listDirectoriesInDirectory(root);
    const available = await Promise.all(dirs.map(async (item) => {
      const exists = await fileExists(projectJsonPath(item.path));
      return exists ? item : null;
    }));
    return available.filter((item): item is ProjectDirectoryItem => Boolean(item)).sort((a, b) => a.name.localeCompare(b.name, 'zh-Hans-CN'));
  }

  async function saveCurrentProject(): Promise<boolean> {
    if (!projectDir.value) {
      return saveCurrentProjectAs();
    }

    const targetFolderName = currentProjectFolderName();
    const currentDir = projectDir.value;
    const currentFolderName = basenameLike(currentDir);
    const rootDir = dirnameLike(currentDir);
    const targetDir = joinLike(rootDir, targetFolderName);

    if (targetDir && normalizeDirLike(targetDir) !== normalizeDirLike(currentDir)) {
      if (await fileExists(targetDir)) {
        throw new Error(`项目文件夹已存在：${targetFolderName}`);
      }
      await renameDirectory(currentDir, targetDir);
      projectDir.value = targetDir;
    }

    if (!currentFolderName && !targetDir) {
      throw new Error('无法确定项目保存位置');
    }

    await persistProject();
    return true;
  }

  async function saveCurrentProjectAs(): Promise<boolean> {
    const rootDir = currentProjectRoot();
    const baseRoot = rootDir || await ensureProjectLibraryRoot();
    if (!baseRoot) {
      return false;
    }
    const targetDir = joinLike(baseRoot, currentProjectFolderName());
    if (await fileExists(targetDir)) {
      throw new Error(`项目文件夹已存在：${basenameLike(targetDir)}`);
    }
    await createDirectory(targetDir);
    const tree = makeProjectTree(targetDir);
    await createDirectory(tree.importsDir);
    await createDirectory(tree.submissionsDir);
    await createDirectory(tree.emailSubmissionsDir);
    await createDirectory(tree.signinDir);
    await createDirectory(tree.gradesDir);
    projectDir.value = targetDir;
    project.settings.projectRootDir = baseRoot;
    await persistProject();
    return true;
  }

  async function importStudents() {
    const filePath = await selectOpenFile({
      title: '导入学生名单 Excel',
      filters: [{ name: 'Excel 文件', extensions: ['xlsx', 'xls'] }],
    });
    if (!filePath) {
      return;
    }
    const workbook = await parseStudentWorkbook(filePath);
    const sourceFileName = basenameLike(filePath);
    if (projectDir.value) {
      const targetPath = joinLike(makeProjectTree(projectDir.value).importsDir, sourceFileName);
      await copyFile(filePath, targetPath);
      studentImport.filePath = joinLike('imports', sourceFileName);
    } else {
      studentImport.filePath = filePath;
    }
    studentImport.headers = workbook.headers;
    studentImport.rows = workbook.rows;
    studentImport.mapping = inferStudentColumnMapping(workbook.headers);
    studentImport.visible = true;
  }

  // import students from a given file path and auto-apply mapping and matching
  async function importStudentsFromPath(filePath: string) {
    if (!filePath) return;
    const workbook = await parseStudentWorkbook(filePath);
    // infer mapping
    const mapping = inferStudentColumnMapping(workbook.headers);
    // build records and set project state
    project.columnMappings = mapping;
    project.students = buildStudentRecords(workbook.rows, mapping);
    project.sourceFiles.studentList = filePath;
    // run matching
    const matched = matchStudents(project);
    project.students = matched.students;
    project.timeSlots = matched.timeSlots;
    issues.splice(0, issues.length, ...matched.issues);
    // note: this API persists by default; if you want a transient (non-persistent) load,
    // use `importStudentsTransient` which updates in-memory state but does not call persistProject.
  }

  // import students into memory only (do not persist or set projectPath)
  async function importStudentsTransient(filePath: string) {
    if (!filePath) return;
    const workbook = await parseStudentWorkbook(filePath);
    const mapping = inferStudentColumnMapping(workbook.headers);
    project.columnMappings = mapping;
    project.students = buildStudentRecords(workbook.rows, mapping);
    // do not modify projectPath or persist
    const matched = matchStudents(project);
    project.students = matched.students;
    project.timeSlots = matched.timeSlots;
    issues.splice(0, issues.length, ...matched.issues);
  }

  function confirmStudentImport(mapping: ColumnMapping) {
    project.columnMappings = mapping;
    project.students = buildStudentRecords(studentImport.rows, mapping);
    studentImport.visible = false;
    const matched = matchStudents(project);
    project.students = matched.students;
    project.timeSlots = matched.timeSlots;
    issues.splice(0, issues.length, ...matched.issues);
    project.sourceFiles.studentList = studentImport.filePath;
    void persistProject();
  }

  async function importTimeSlots() {
    const filePath = await selectOpenFile({
      title: '导入时间分组 Excel',
      filters: [{ name: 'Excel 文件', extensions: ['xlsx', 'xls'] }],
    });
    if (!filePath) {
      return;
    }
    const slots = await parseTimeSlotWorkbook(filePath);
    const sourceFileName = basenameLike(filePath);
    if (projectDir.value) {
      const targetPath = joinLike(makeProjectTree(projectDir.value).importsDir, sourceFileName);
      await copyFile(filePath, targetPath);
      project.sourceFiles.timeSlots = joinLike('imports', sourceFileName);
    } else {
      project.sourceFiles.timeSlots = filePath;
    }
    project.timeSlots = buildTimeSlotRecords(slots);
    timeSlotImport.filePath = filePath;
    timeSlotImport.visible = true;
    const matched = matchStudents(project);
    project.students = matched.students;
    project.timeSlots = matched.timeSlots;
    issues.splice(0, issues.length, ...matched.issues);
    void persistProject();
  }

  async function selectReportsDirectory() {
    const dir = await selectDirectory({
      title: '选择成绩报告文件夹',
      defaultPath: buildProjectSubmissionDir(project),
    });
    if (!dir) return;
    reports.dirPath = dir;
    project.sourceFiles.reportDir = toProjectRelativePath(dir);
    await scanReportsDirectory();
    void persistProject();
  }

  async function scanReportsDirectory() {
    if (!reports.dirPath) return;
    const files = await listFilesInDirectory(reports.dirPath);
    reports.files.splice(0, reports.files.length, ...files);
    // attempt to auto-match files to students by studentId or name in filename
    const byId = new Map<string, any[]>();
    project.students.forEach((s) => {
      const id = normalizeText(s.studentId);
      const name = normalizeText(s.name);
      if (id) byId.set(id, [...(byId.get(id) ?? []), s]);
      if (name) byId.set(name, [...(byId.get(name) ?? []), s]);
    });
    reports.files.forEach((file) => {
      const basename = normalizeText(stripExtLike(basenameLike(file.path)));
      const keyCandidates = [basename, basename.replace(/\s+/g, ''), basename.toLowerCase()];
      for (const [key, students] of byId.entries()) {
        if (!key) continue;
        const k = key.replace(/\s+/g, '').toLowerCase();
        if (keyCandidates.some((c) => c.includes(k))) {
          const student = students[0];
          student.submissionFile = file.path;
          student.submissionStatus = 'submitted';
          // attempt to extract a suggested score from filename
          const nums = (basename.match(/\d{1,3}/g) ?? []).map((s) => Number(s)).filter((n) => n >= 0 && n <= 100);
          if (nums.length) {
            student.suggestedScore = nums[0];
          }
        }
      }
    });
    const matched = matchStudents(project);
    project.students = matched.students;
    project.timeSlots = matched.timeSlots;
    issues.splice(0, issues.length, ...matched.issues);
    void persistProject();
  }

  function findUniqueStudentByName(name: string) {
    const value = normalizeText(name);
    if (!value) return { status: 'none' as const, student: null as StudentRecord | null };
    const matches = project.students.filter((student) => normalizeText(student.name) === value);
    if (matches.length === 1) return { status: 'unique' as const, student: matches[0] };
    if (matches.length > 1) return { status: 'duplicate' as const, student: null as StudentRecord | null };
    return { status: 'none' as const, student: null as StudentRecord | null };
  }

  function findStudentByContainedStudentId(text: string) {
    const value = normalizeText(text);
    if (!value) return null;
    return project.students.find((student) => student.studentId && value.includes(student.studentId)) ?? null;
  }

  function findStudentByContainedName(text: string) {
    const value = normalizeText(text);
    if (!value) return { status: 'none' as const, student: null as StudentRecord | null };
    const matches = project.students.filter((student) => student.name && value.includes(student.name));
    const unique = [...new Map(matches.map((student) => [student.studentId, student])).values()];
    if (unique.length === 1) return { status: 'unique' as const, student: unique[0] };
    if (unique.length > 1) return { status: 'duplicate' as const, student: null as StudentRecord | null };
    return { status: 'none' as const, student: null as StudentRecord | null };
  }

  function matchEmailAttachment(attachment: ImportedEmailAttachment): EmailAttachmentMatch {
    const senderNameMatch = findUniqueStudentByName(attachment.fromName);
    const attachmentIdMatch = findStudentByContainedStudentId(attachment.originalFileName);
    const attachmentNameMatch = findStudentByContainedName(attachment.originalFileName);
    const subjectIdMatch = findStudentByContainedStudentId(attachment.subject);
    const subjectNameMatch = findStudentByContainedName(attachment.subject);

    if (senderNameMatch.status === 'duplicate') {
      return {
        attachment,
        status: 'duplicate-name',
        reason: `发件人姓名重复：${attachment.fromName}`,
        matchBasis: 'sender-name',
      };
    }

    if (senderNameMatch.student) {
      if (attachmentIdMatch && attachmentIdMatch.studentId !== senderNameMatch.student.studentId) {
        return {
          attachment,
          status: 'conflict',
          matchedStudentId: senderNameMatch.student.studentId,
          matchedStudentName: senderNameMatch.student.name,
          matchBasis: 'sender-name',
          reason: `发件人姓名匹配 ${senderNameMatch.student.name}，但附件名学号指向 ${attachmentIdMatch.name}`,
        };
      }
      return {
        attachment,
        status: 'matched',
        matchedStudentId: senderNameMatch.student.studentId,
        matchedStudentName: senderNameMatch.student.name,
        matchBasis: 'sender-name',
        reason: '发件人显示名匹配学生姓名',
      };
    }

    if (attachmentIdMatch) {
      return {
        attachment,
        status: 'matched',
        matchedStudentId: attachmentIdMatch.studentId,
        matchedStudentName: attachmentIdMatch.name,
        matchBasis: 'attachment-student-id',
        reason: '附件名包含学号',
      };
    }

    if (attachmentNameMatch.status === 'duplicate') {
      return { attachment, status: 'duplicate-name', matchBasis: 'attachment-name', reason: '附件名包含多个学生姓名' };
    }
    if (attachmentNameMatch.student) {
      return {
        attachment,
        status: 'matched',
        matchedStudentId: attachmentNameMatch.student.studentId,
        matchedStudentName: attachmentNameMatch.student.name,
        matchBasis: 'attachment-name',
        reason: '附件名包含学生姓名',
      };
    }

    if (subjectIdMatch) {
      return {
        attachment,
        status: 'matched',
        matchedStudentId: subjectIdMatch.studentId,
        matchedStudentName: subjectIdMatch.name,
        matchBasis: 'subject-student-id',
        reason: '邮件标题包含学号',
      };
    }

    if (subjectNameMatch.status === 'duplicate') {
      return { attachment, status: 'duplicate-name', matchBasis: 'subject-name', reason: '邮件标题包含多个学生姓名' };
    }
    if (subjectNameMatch.student) {
      return {
        attachment,
        status: 'matched',
        matchedStudentId: subjectNameMatch.student.studentId,
        matchedStudentName: subjectNameMatch.student.name,
        matchBasis: 'subject-name',
        reason: '邮件标题包含学生姓名',
      };
    }

    return { attachment, status: 'unmatched', reason: '发件人、附件名和标题都未匹配到学生' };
  }

  function applyEmailMatches(matches: EmailAttachmentMatch[]) {
    matches.forEach((match) => {
      const key = buildEmailAttachmentKey(match.attachment);
      if (match.attachment.savedPath && key && !project.importedEmailAttachmentKeys?.includes(key)) {
        if (!project.importedEmailAttachmentKeys) project.importedEmailAttachmentKeys = [];
        project.importedEmailAttachmentKeys.push(key);
      }
      if (match.status !== 'matched' || !match.matchedStudentId) return;
      const student = project.students.find((item) => item.studentId === match.matchedStudentId);
      if (!student) return;
      student.submissionFile = match.attachment.savedPath;
      student.submissionStatus = 'submitted';
    });
    void persistProject();
  }

  function buildEmailDownloadDir() {
    const baseDir = currentProjectDir() ? makeProjectTree(currentProjectDir()).submissionsDir : buildProjectSubmissionDir(project);
    return joinPathLike(baseDir, 'email');
  }

  function buildEmailAttachmentKey(attachment: ImportedEmailAttachment) {
    const stableMessageKey = normalizeText(attachment.messageId) || `uid:${attachment.uid}`;
    return [
      stableMessageKey,
      normalizeText(attachment.originalFileName).toLowerCase(),
      String(attachment.size || 0),
    ].join('|');
  }

  function toPlainEmailImportConfig(config: EmailImportConfig): EmailImportConfig {
    return {
      host: normalizeText(config.host),
      port: Number(config.port || 993),
      secure: Boolean(config.secure),
      user: normalizeText(config.user),
      password: String(config.password ?? ''),
      mailbox: normalizeText(config.mailbox) || 'INBOX',
      since: normalizeText(config.since),
      before: normalizeText(config.before),
      subjectKeyword: normalizeText(config.subjectKeyword),
      limit: Number(config.limit || 50),
      downloadDir: normalizeText(config.downloadDir),
      knownAttachmentKeys: [...(config.knownAttachmentKeys ?? [])],
      matchHints: [...(config.matchHints ?? [])],
    };
  }

  async function importSubmissionsFromEmail(config: EmailImportConfig) {
    emailImport.loading = true;
    emailImport.errorMessage = '';
    try {
      const plainConfig = toPlainEmailImportConfig(config);
      Object.assign(emailImport.config, plainConfig);
      project.settings.emailImportConfig = {
        ...emailImport.config,
        downloadDir: '',
        knownAttachmentKeys: [],
        matchHints: [],
      };
      const downloadDir = plainConfig.downloadDir || buildEmailDownloadDir();
      const attachments = await importEmailAttachments({
        ...plainConfig,
        downloadDir,
        knownAttachmentKeys: [...(project.importedEmailAttachmentKeys ?? [])],
        matchHints: project.students.map((student) => ({ studentId: student.studentId, name: student.name })),
      });
      const knownKeys = new Set(project.importedEmailAttachmentKeys ?? []);
      const freshAttachments = attachments.filter((attachment) => !knownKeys.has(buildEmailAttachmentKey(attachment)));
      const matches = freshAttachments.map(matchEmailAttachment);
      emailImport.results.splice(0, emailImport.results.length, ...matches);
      reports.dirPath = downloadDir;
      applyEmailMatches(matches);
      return matches;
    } catch (error) {
      emailImport.errorMessage = error instanceof Error ? error.message : String(error);
      throw error;
    } finally {
      emailImport.loading = false;
    }
  }

  function updateStudentMappingField(field: keyof ColumnMapping, header: string) {
    studentImport.mapping[field] = header;
  }

  function applyStudentMapping() {
    // validate required fields
    const required: Array<keyof ColumnMapping> = ['studentId', 'name'];
    const missing = required.filter((k) => !studentImport.mapping[k]);
    if (missing.length) {
      studentImport.errorMessage = `必须选择列：${missing.map((m) => String(m)).join(', ')}`;
      return;
    }
    studentImport.errorMessage = '';
    confirmStudentImport(studentImport.mapping);
  }

  function saveStudentImportMappingPreset(name: string) {
    if (!name) return;
    if (!project.columnMappingPresets) project.columnMappingPresets = [];
    const preset = { name, mapping: { ...(studentImport.mapping ?? {}) } };
    const idx = project.columnMappingPresets.findIndex((p) => p.name === name);
    if (idx >= 0) project.columnMappingPresets[idx] = preset; else project.columnMappingPresets.push(preset);
    void persistProject();
  }

  function saveColumnMappingPreset(name: string) {
    if (!name) return;
    if (!project.columnMappingPresets) {
      // initialize if missing (backwards compatibility)
      project.columnMappingPresets = [];
    }
    const preset = { name, mapping: { ...(project.columnMappings ?? {}) } };
    const idx = project.columnMappingPresets.findIndex((p) => p.name === name);
    if (idx >= 0) {
      project.columnMappingPresets[idx] = preset;
    } else {
      project.columnMappingPresets.push(preset);
    }
    void persistProject();
  }

  function applyColumnMappingPreset(name: string) {
    if (!project.columnMappingPresets) return;
    const preset = project.columnMappingPresets.find((p) => p.name === name);
    if (!preset) return;
    project.columnMappings = { ...preset.mapping };
    if (studentImport.visible) {
      studentImport.mapping = { ...preset.mapping };
    }
  }

  function deleteColumnMappingPreset(name: string) {
    if (!project.columnMappingPresets) return;
    const idx = project.columnMappingPresets.findIndex((p) => p.name === name);
    if (idx >= 0) {
      project.columnMappingPresets.splice(idx, 1);
      void persistProject();
    }
  }

  function closeStudentMappingDialog() {
    studentImport.visible = false;
  }

  function closeTimeSlotDialog() {
    timeSlotImport.visible = false;
  }

  function updateTimeSlotGroupName(slotId: string, value: string) {
    const slot = project.timeSlots.find((item) => item.id === slotId);
    if (!slot) {
      return;
    }
    slot.groupName = normalizeText(value);
    const matched = matchStudents(project);
    project.students = matched.students;
    project.timeSlots = matched.timeSlots;
    issues.splice(0, issues.length, ...matched.issues);
    void persistProject();
  }

  function clearGroupNames() {
    project.timeSlots.forEach((slot) => {
      slot.groupName = '';
    });
    const matched = matchStudents(project);
    project.students = matched.students;
    project.timeSlots = matched.timeSlots;
    issues.splice(0, issues.length, ...matched.issues);
    void persistProject();
  }

  function batchGenerateGroupNames(startName: string) {
    project.timeSlots = generateGroupNames(project.timeSlots, startName);
    const matched = matchStudents(project);
    project.students = matched.students;
    project.timeSlots = matched.timeSlots;
    issues.splice(0, issues.length, ...matched.issues);
    void persistProject();
  }

  function refreshMatching() {
    const matched = matchStudents(project);
    project.students = matched.students;
    project.timeSlots = matched.timeSlots;
    issues.splice(0, issues.length, ...matched.issues);
  }

  function updateStudentScore(studentId: string, score: number | null) {
    const student = project.students.find((item) => item.studentId === studentId);
    if (!student) {
      return;
    }
    student.score = score;
    if (score === null) {
      student.submissionStatus = 'missing';
    } else if (student.submissionStatus === 'missing') {
      student.submissionStatus = 'manual';
    }
    void persistProject();
  }

  function updateStudentRemark(studentId: string, remark: string) {
    const student = project.students.find((item) => item.studentId === studentId);
    if (!student) {
      return;
    }
    student.remark = normalizeText(remark);
    void persistProject();
  }

  function updateStudentAttendance(studentId: string, attendance: boolean) {
    const student = project.students.find((item) => item.studentId === studentId);
    if (!student) {
      return;
    }
    student.attendance = attendance;
    void persistProject();
  }

  function updateStudentSubmissionStatus(studentId: string, status: StudentRecord['submissionStatus']) {
    const student = project.students.find((item) => item.studentId === studentId);
    if (!student) {
      return;
    }
    student.submissionStatus = status;
    if (status === 'missing' && student.score === null) {
      student.score = 0;
    }
    void persistProject();
  }

  function updateStudentManualScore(studentId: string, score: number | null, remark?: string) {
    updateStudentScore(studentId, score);
    if (remark !== undefined) {
      updateStudentRemark(studentId, remark);
    }
  }

  async function autoReviewSubmittedStudents() {
    autoReview.loading = true;
    autoReview.errorMessage = '';
    try {
      const inputs = project.students
        .filter((student) => student.submissionFile || student.submissionStatus === 'submitted')
        .map((student) => ({
          studentId: student.studentId,
          name: student.name,
          submissionFile: student.submissionFile,
          submissionStatus: student.submissionStatus,
        }));
      const results = await autoReviewSubmissions(inputs);
      autoReview.results.splice(0, autoReview.results.length, ...results);
      results.forEach((result) => {
        const student = project.students.find((item) => item.studentId === result.studentId);
        if (!student) return;
        student.score = result.score;
        student.suggestedScore = result.score;
        student.remark = result.remark;
        if (student.submissionStatus === 'missing') {
          student.submissionStatus = 'manual';
        }
      });
      await persistProject();
      return results;
    } catch (error) {
      autoReview.errorMessage = error instanceof Error ? error.message : String(error);
      throw error;
    } finally {
      autoReview.loading = false;
    }
  }

  async function exportSigninSheet(): Promise<boolean> {
    const filePath = await selectSaveFile({
      title: '导出签到表',
      defaultPath: buildProjectExportPath(project, 'signin', 'xlsx'),
      filters: [{ name: 'Excel 文件', extensions: ['xlsx'] }],
    });
    if (!filePath) {
      return false;
    }
    const exportProject = snapshot();
    await exportSigninWorkbook(filePath, { project: exportProject, layout: exportProject.signinLayout });
    return true;
  }

  async function exportSigninPdf(): Promise<boolean> {
    const filePath = await selectSaveFile({
      title: '导出签到表（PDF）',
      defaultPath: buildProjectExportPath(project, 'signin', 'pdf'),
      filters: [{ name: 'PDF 文件', extensions: ['pdf'] }],
    });
    if (!filePath) return false;
    // call API that delegates to main process
    const exportProject = snapshot();
    await (await import('@/services/api')).exportSigninPdf(filePath, { project: exportProject, layout: exportProject.signinLayout });
    return true;
  }

  async function exportGradesSheet(fullRecord: boolean) {
    const filePath = await selectSaveFile({
      title: '导出成绩单',
      defaultPath: buildProjectExportPath(project, 'grades', 'xlsx'),
      filters: [{ name: 'Excel 文件', extensions: ['xlsx'] }],
    });
    if (!filePath) {
      return false;
    }
    await exportGradesWorkbook(filePath, { project: snapshot(), fullRecord });
    return true;
  }

  async function selectDefaultExportDir() {
    const dirPath = await selectDirectory({ title: '选择默认导出目录' });
    if (!dirPath) {
      return;
    }
    project.settings.defaultExportDir = dirPath;
    void persistProject();
  }

  async function selectProjectRootDir() {
    const dirPath = await selectDirectory({ title: '选择项目根目录' });
    if (!dirPath) {
      return;
    }
    project.settings.projectRootDir = dirPath;
    void persistProject();
  }

  function markMissing(studentId: string) {
    const student = project.students.find((item) => item.studentId === studentId);
    if (!student) {
      return;
    }
    student.submissionStatus = 'missing';
    student.score = 0;
    void persistProject();
  }

  function sortedIssues() {
    return [...issues];
  }

  function duplicateGroupNames() {
    return findDuplicateGroupNames(project.timeSlots);
  }

  function studentsForDisplay() {
    return getSortedStudentsByLayout(
      project.students,
      project.signinLayout.sortByGroupName,
      project.signinLayout.sortByTimeSlot,
      project.signinLayout.sortByStudentId,
    );
  }

  return {
    project,
    projectDir,
    studentImport,
    timeSlotImport,
    issues,
    stats,
    createNewProject,
    createProjectFromMetadata,
    openProjectFromDirectory,
    listAvailableProjects,
    saveCurrentProject,
    saveCurrentProjectAs,
    importStudents,
    confirmStudentImport,
    applyStudentMapping,
    saveStudentImportMappingPreset,
    closeStudentMappingDialog,
    closeTimeSlotDialog,
    importTimeSlots,
    updateStudentMappingField,
    saveColumnMappingPreset,
    applyColumnMappingPreset,
    deleteColumnMappingPreset,
    updateTimeSlotGroupName,
    clearGroupNames,
    batchGenerateGroupNames,
    refreshMatching,
    updateStudentScore,
    updateStudentRemark,
    updateStudentAttendance,
    updateStudentSubmissionStatus,
    updateStudentManualScore,
    exportSigninSheet,
    exportSigninPdf,
    exportGradesSheet,
    selectDefaultExportDir,
    selectProjectRootDir,
    markMissing,
    sortedIssues,
    duplicateGroupNames,
    studentsForDisplay,
    reports,
    emailImport,
    autoReview,
    selectReportsDirectory,
    scanReportsDirectory,
    importSubmissionsFromEmail,
    autoReviewSubmittedStudents,
    importStudentsTransient,
  };
});
