import { contextBridge, ipcRenderer } from 'electron';
import type { AutoReviewResult, AutoReviewSubmissionInput, EmailImportConfig, ImportedEmailAttachment, LabProject, ParsedTimeSlot, SigninLayoutSettings } from './shared-types';

contextBridge.exposeInMainWorld('labApi', {
  selectOpenFile: (options?: { title?: string; filters?: Array<{ name: string; extensions: string[] }> }) =>
    ipcRenderer.invoke('dialog:selectOpenFile', options),
  selectSaveFile: (options?: { title?: string; defaultPath?: string; filters?: Array<{ name: string; extensions: string[] }> }) =>
    ipcRenderer.invoke('dialog:selectSaveFile', options),
  selectDirectory: (options?: { title?: string; defaultPath?: string }) =>
    ipcRenderer.invoke('dialog:selectDirectory', options),
  listFilesInDirectory: (dirPath: string) => ipcRenderer.invoke('dir:listFiles', dirPath),
  listDirectoriesInDirectory: (dirPath: string) => ipcRenderer.invoke('dir:listDirectories', dirPath),
  readTextFile: (filePath: string) => ipcRenderer.invoke('file:readText', filePath),
  writeTextFile: (filePath: string, content: string) => ipcRenderer.invoke('file:writeText', filePath, content),
  copyFile: (fromPath: string, toPath: string) => ipcRenderer.invoke('file:copy', fromPath, toPath),
  fileExists: (filePath: string) => ipcRenderer.invoke('file:exists', filePath) as Promise<boolean>,
  createDirectory: (dirPath: string) => ipcRenderer.invoke('dir:createTree', dirPath),
  renameDirectory: (fromPath: string, toPath: string) => ipcRenderer.invoke('dir:rename', fromPath, toPath),
  getProjectLibraryRoot: () => ipcRenderer.invoke('project:getLibraryRoot') as Promise<string>,
  openExternal: (filePath: string) => ipcRenderer.invoke('file:openExternal', filePath),
  openProject: (filePath: string) => ipcRenderer.invoke('project:open', filePath) as Promise<LabProject>,
  saveProject: (filePath: string, project: LabProject) => ipcRenderer.invoke('project:save', filePath, project),
  parseStudentWorkbook: (filePath: string) => ipcRenderer.invoke('excel:parseStudents', filePath),
  parseTimeSlotWorkbook: (filePath: string) => ipcRenderer.invoke('excel:parseTimeSlots', filePath) as Promise<ParsedTimeSlot[]>,
  exportSigninWorkbook: (filePath: string, payload: { project: LabProject; layout: SigninLayoutSettings }) =>
    ipcRenderer.invoke('excel:exportSignin', filePath, payload),
  exportSigninPdf: (filePath: string, payload: { project: LabProject; layout: SigninLayoutSettings }) =>
    ipcRenderer.invoke('excel:exportSigninPdf', filePath, payload),
  exportGradesWorkbook: (filePath: string, payload: { project: LabProject; fullRecord: boolean }) =>
    ipcRenderer.invoke('excel:exportGrades', filePath, payload),
  importEmailAttachments: (config: EmailImportConfig) =>
    ipcRenderer.invoke('email:importAttachments', config) as Promise<ImportedEmailAttachment[]>,
  autoReviewSubmissions: (inputs: AutoReviewSubmissionInput[]) =>
    ipcRenderer.invoke('grading:autoReview', inputs) as Promise<AutoReviewResult[]>,
  exportReport: (filePath: string, content: string) => ipcRenderer.invoke('file:writeText', filePath, content),
});
