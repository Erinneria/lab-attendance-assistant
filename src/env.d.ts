/// <reference types="vite/client" />

import type { AppSettings, AutoReviewResult, AutoReviewSubmissionInput, ColumnMapping, EmailImportConfig, ImportedEmailAttachment, LabProject, ParsedTimeSlot, SigninLayoutSettings } from '@/types/domain';

declare global {
  interface Window {
    labApi: {
      selectOpenFile(options?: { title?: string; filters?: Array<{ name: string; extensions: string[] }> }): Promise<string | null>;
      selectSaveFile(options?: { title?: string; defaultPath?: string; filters?: Array<{ name: string; extensions: string[] }> }): Promise<string | null>;
      selectDirectory(options?: { title?: string; defaultPath?: string }): Promise<string | null>;
      listDirectoriesInDirectory(dirPath: string): Promise<Array<{ name: string; path: string }>>;
      readTextFile(filePath: string): Promise<string>;
      writeTextFile(filePath: string, content: string): Promise<void>;
      copyFile(fromPath: string, toPath: string): Promise<void>;
      fileExists(filePath: string): Promise<boolean>;
      createDirectory(dirPath: string): Promise<void>;
      renameDirectory(fromPath: string, toPath: string): Promise<string>;
      getProjectLibraryRoot(): Promise<string>;
      openExternal(filePath: string): Promise<void>;
      openProject(filePath: string): Promise<LabProject>;
      saveProject(filePath: string, project: LabProject): Promise<void>;
      parseStudentWorkbook(filePath: string): Promise<{ headers: string[]; rows: Record<string, string>[] }>;
      parseTimeSlotWorkbook(filePath: string): Promise<ParsedTimeSlot[]>;
      exportSigninWorkbook(filePath: string, payload: { project: LabProject; layout: SigninLayoutSettings }): Promise<void>;
      exportSigninPdf(filePath: string, payload: { project: LabProject; layout: SigninLayoutSettings }): Promise<void>;
      exportGradesWorkbook(filePath: string, payload: { project: LabProject; fullRecord: boolean }): Promise<void>;
      importEmailAttachments(config: EmailImportConfig): Promise<ImportedEmailAttachment[]>;
      autoReviewSubmissions(inputs: AutoReviewSubmissionInput[]): Promise<AutoReviewResult[]>;
      listFilesInDirectory(dirPath: string): Promise<Array<{ name: string; path: string }>>;
      exportReport(filePath: string, content: string): Promise<void>;
    };
  }
}

export {};
