import type { AutoReviewResult, AutoReviewSubmissionInput, EmailImportConfig, ImportedEmailAttachment, LabProject, ParsedTimeSlot, SigninLayoutSettings } from '@/types/domain';

export async function selectOpenFile(options?: { title?: string; filters?: Array<{ name: string; extensions: string[] }> }) {
  return window.labApi.selectOpenFile(options);
}

export async function selectSaveFile(options?: { title?: string; defaultPath?: string; filters?: Array<{ name: string; extensions: string[] }> }) {
  return window.labApi.selectSaveFile(options);
}

export async function selectDirectory(options?: { title?: string; defaultPath?: string }) {
  return window.labApi.selectDirectory(options);
}

export async function listDirectoriesInDirectory(dirPath: string) {
  return window.labApi.listDirectoriesInDirectory(dirPath) as Promise<Array<{ name: string; path: string }>>;
}

export async function listFilesInDirectory(dirPath: string) {
  return window.labApi.listFilesInDirectory(dirPath) as Promise<Array<{ name: string; path: string }>>;
}

export async function createDirectory(dirPath: string) {
  return window.labApi.createDirectory(dirPath);
}

export async function renameDirectory(fromPath: string, toPath: string) {
  return window.labApi.renameDirectory(fromPath, toPath);
}

export async function copyFile(fromPath: string, toPath: string) {
  return window.labApi.copyFile(fromPath, toPath);
}

export async function fileExists(filePath: string) {
  return window.labApi.fileExists(filePath);
}

export async function getProjectLibraryRoot() {
  return window.labApi.getProjectLibraryRoot() as Promise<string>;
}

export async function openProject(filePath: string): Promise<LabProject> {
  return window.labApi.openProject(filePath);
}

export async function saveProject(filePath: string, project: LabProject) {
  return window.labApi.saveProject(filePath, project);
}

export async function parseStudentWorkbook(filePath: string) {
  return window.labApi.parseStudentWorkbook(filePath) as Promise<{ headers: string[]; rows: Record<string, string>[] }>;
}

export async function parseTimeSlotWorkbook(filePath: string): Promise<ParsedTimeSlot[]> {
  return window.labApi.parseTimeSlotWorkbook(filePath);
}

export async function exportSigninWorkbook(filePath: string, payload: { project: LabProject; layout: SigninLayoutSettings }) {
  return window.labApi.exportSigninWorkbook(filePath, payload);
}

export async function exportSigninPdf(filePath: string, payload: { project: LabProject; layout: SigninLayoutSettings }) {
  // delegate to main process to render HTML and print to PDF
  return (window.labApi as any).exportSigninPdf(filePath, payload);
}

export async function exportGradesWorkbook(filePath: string, payload: { project: LabProject; fullRecord: boolean }) {
  return window.labApi.exportGradesWorkbook(filePath, payload);
}

export async function importEmailAttachments(config: EmailImportConfig): Promise<ImportedEmailAttachment[]> {
  return window.labApi.importEmailAttachments(config);
}

export async function autoReviewSubmissions(inputs: AutoReviewSubmissionInput[]): Promise<AutoReviewResult[]> {
  return window.labApi.autoReviewSubmissions(inputs);
}

export async function openExternal(filePath: string) {
  return window.labApi.openExternal(filePath);
}

export async function readTextFile(filePath: string) {
  return window.labApi.readTextFile(filePath) as Promise<string>;
}
