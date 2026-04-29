# Lab Attendance Assistant

面向实验课教学的本地桌面管理工具，用于完成学生名单导入、选课时间整理、签到表生成、作业附件收集、报告批改辅助与成绩单导出。

本项目优先服务小班到中等规模实验课的真实教学流程：数据主要来自 Excel、邮件附件和本地文件夹，处理过程默认在本机完成，减少对在线平台和外部服务的依赖。

## Features

- **项目化课程管理**：按课程和学期保存学生名单、时间分组、提交记录、评分结果和导出配置。
- **Excel 名单导入**：支持从 Excel 导入学生名单，并自动识别学号、姓名、班级等常见字段。
- **选课时间整理**：支持导入选课时间表，将学生和实验时间段进行匹配。
- **签到表导出**：按时间段、组名和学生信息生成可打印的 Excel 签到表。
- **邮箱附件导入**：通过 IMAP 扫描邮箱附件，根据发件人显示名、附件名和邮件标题匹配学生。
- **本地作业扫描**：可直接选择本地作业文件夹，批量匹配学生提交文件。
- **自动评阅辅助**：读取 docx、pdf 等报告内容，根据文本规模与思考/问答类内容进行自动评分，并保留评价说明。
- **成绩单导出**：将学生成绩和提交状态导出为 Excel，方便后续登记或归档。

## Tech Stack

- [Electron](https://www.electronjs.org/) for the desktop shell and local file access
- [Vue 3](https://vuejs.org/) for the renderer UI
- [TypeScript](https://www.typescriptlang.org/) for typed application code
- [Pinia](https://pinia.vuejs.org/) for state management
- [Element Plus](https://element-plus.org/) for UI components
- ExcelJS / xlsx for Excel import and export
- imapflow / mailparser for IMAP email attachment import
- mammoth / pdf-parse for report text extraction

## Getting Started

Requirements:

- Node.js 20 or later
- npm

Install dependencies:

```bash
npm install
```

Start the desktop app in development mode:

```bash
npm run dev
```

Run type checks:

```bash
npm run typecheck
```

Build renderer and Electron main process output:

```bash
npm run build
```

The build command verifies that the project can be compiled and emits local build artifacts. It does not currently produce a packaged installer.

## Data Model

The app stores course workspaces locally. A project contains:

- course metadata
- student records
- imported time slots
- attendance sheet layout settings
- submission matching results
- grading results
- export history and local file paths

By default, local project data is stored under:

```text
.lab-attendance-assistant/
  projects/
    <course - term>/
      project.json
      imports/
      submissions/
        email/
      exports/
```

This directory is ignored by Git because it may contain student information, grades, report files, email configuration, and exported spreadsheets.

## Email Import

Email import uses IMAP and supports common mailbox settings:

- host
- port
- SSL/TLS
- username
- password
- mailbox name, usually `INBOX`
- date range
- optional subject keyword
- import limit

Attachment matching is performed in this order:

1. sender display name matches a student name
2. attachment filename contains a student ID
3. attachment filename contains a student name
4. email subject contains a student ID
5. email subject contains a student name

Imported attachments are deduplicated, so repeated scans do not keep downloading the same file. Attachments that cannot be matched to a student are not saved into the submissions folder.

## Auto Grading

Automatic grading is implemented as a local heuristic pipeline. It does not call an online AI service.

The current scorer extracts report text and compares submissions across the class using two primary signals:

- total effective report content
- amount of thinking, question-answer, explanation, and summary style content

The result is mapped to a stable score range and displayed with expandable review details. The feature is intended to reduce repetitive first-pass grading work, while still allowing instructors to open files and manually adjust individual scores.

## Privacy

This application is designed as a local-first tool. Student data and report files are processed on the local machine unless the user explicitly configures an external mailbox connection.

Important notes:

- Local project data should not be committed to a public repository.
- If an email password is saved in a project file, it is currently stored locally in plaintext.
- Public releases should avoid bundling real course data, exported grades, report files, or mailbox credentials.

## Project Structure

```text
electron/                 Electron main process, preload API, local services
scripts/                  development helper scripts
src/                      Vue renderer application
src/views/                page-level views
src/stores/               Pinia stores
src/services/             renderer-side service wrappers
src/types/                shared domain types
index.html                Vite entry HTML
package.json              scripts and dependencies
vite.config.ts            Vite configuration
tsconfig.json             TypeScript configuration
```

## Git Ignore Policy

The repository intentionally ignores:

- dependency folders such as `node_modules/`
- build outputs such as `dist-electron/` and `dist-renderer/`
- local project workspaces under `.lab-attendance-assistant/`
- generated output folders
- environment files and editor noise

`package-lock.json` is committed to keep dependency installation reproducible.

## Roadmap

- Packaged installer generation
- More robust report parsing for scanned PDFs and image-heavy submissions
- Configurable grading rubrics
- Better manual review workflow for borderline submissions
- Optional encrypted storage for mailbox credentials

