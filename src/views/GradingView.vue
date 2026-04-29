<template>
  <div class="page-grid">
    <div class="panel-card">
      <div class="panel-header">
        <div class="panel-title">成绩批量匹配与评分</div>
        <div class="panel-actions">
          <el-button type="primary" @click="onSelectDir">选择报告文件夹</el-button>
          <el-button type="success" @click="openEmailImportDialog">邮箱导入</el-button>
          <el-button @click="onScan" :disabled="!reports.dirPath">扫描并匹配</el-button>
          <el-button type="warning" :loading="autoReview.loading" :disabled="!studentsWithFiles.length" @click="runAutoReview">自动评阅</el-button>
        </div>
      </div>

      <div style="margin-top:12px">
        <el-tag v-if="reports.dirPath">当前文件夹：{{ reports.dirPath }}</el-tag>
      </div>

      <el-table :data="studentsWithFiles" stripe style="margin-top:16px">
        <el-table-column type="expand">
          <template #default="{ row }">
            <div class="review-detail">
              <div class="review-detail-title">自动评阅说明</div>
              <ul v-if="reviewDetails(row.studentId).length">
                <li v-for="item in reviewDetails(row.studentId)" :key="item">{{ item }}</li>
              </ul>
              <div v-else>{{ row.remark || '暂无自动评阅说明' }}</div>
            </div>
          </template>
        </el-table-column>
        <el-table-column prop="studentId" label="学号" width="120" />
        <el-table-column prop="name" label="姓名" width="140" />
        <el-table-column prop="submissionFile" label="报告文件" min-width="240">
          <template #default="{ row }">
            <el-link v-if="row.submissionFile" @click="openExternalFile(row.submissionFile)">{{ trimPath(row.submissionFile) }}</el-link>
            <span v-else>—</span>
          </template>
        </el-table-column>
        <el-table-column label="分数" width="180">
          <template #default="{ row }">
            <div style="display:flex;align-items:center;gap:8px">
              <el-input-number v-model="row.score" :min="0" :max="100" size="small" @change="onScoreChange(row.studentId, $event)" />
            </div>
          </template>
        </el-table-column>
        <el-table-column label="操作" width="160">
          <template #default="{ row }">
            <el-button size="small" @click="applyScore(row.studentId, row.score)">保存</el-button>
            <el-button size="small" type="text" @click="clearMatch(row.studentId)">清除匹配</el-button>
          </template>
        </el-table-column>
      </el-table>

      <div v-if="emailImport.results.length" style="margin-top:16px">
        <div class="panel-header">
          <div class="panel-title">邮箱导入结果</div>
          <el-tag>{{ emailImport.results.length }} 个附件</el-tag>
        </div>
        <el-table :data="emailImport.results" stripe border style="margin-top:8px">
          <el-table-column label="状态" width="110">
            <template #default="{ row }">
              <el-tag :type="emailStatusType(row.status)">{{ emailStatusText(row.status) }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column label="发件人" width="120">
            <template #default="{ row }">{{ row.attachment.fromName }}</template>
          </el-table-column>
          <el-table-column label="附件" min-width="220">
            <template #default="{ row }">
              <el-link v-if="row.attachment.savedPath" @click="showPreview(row.attachment.savedPath)">{{ row.attachment.originalFileName }}</el-link>
              <span v-else>{{ row.attachment.originalFileName }}</span>
            </template>
          </el-table-column>
          <el-table-column label="匹配学生" width="160">
            <template #default="{ row }">{{ row.matchedStudentName || '—' }}</template>
          </el-table-column>
          <el-table-column prop="reason" label="依据/原因" min-width="240" />
        </el-table>
      </div>
    </div>
  </div>

  <el-dialog v-model="preview.visible" width="80%" :close-on-click-modal="false">
  <template #title>文件预览</template>
  <div style="min-height:400px; display:flex; justify-content:center; align-items:center;">
    <template v-if="preview.type === 'image'">
      <img :src="'file://' + preview.path" style="max-width:100%; max-height:80vh;" />
    </template>
    <template v-else-if="preview.type === 'pdf'">
      <iframe :src="'file://' + preview.path" style="width:100%; height:80vh; border:none"></iframe>
    </template>
    <template v-else-if="preview.type === 'text'">
      <el-scrollbar style="width:100%; height:80vh;">
        <pre style="white-space:pre-wrap">{{ preview.text }}</pre>
      </el-scrollbar>
    </template>
    <template v-else>
      <div>无法预览该文件类型，可点击下方“在外部打开”使用系统应用查看。</div>
    </template>
  </div>
  <template #footer>
    <el-button @click="closePreview">关闭</el-button>
    <el-button type="primary" @click="openExternalFile(preview.path)">在外部打开</el-button>
  </template>
  </el-dialog>

  <el-dialog v-model="emailImport.visible" title="邮箱导入作业附件" width="720px">
    <div class="form-grid">
      <el-input v-model="emailImport.config.host" placeholder="IMAP 服务器，例如 imap.example.edu.cn" />
      <el-input-number v-model="emailImport.config.port" :min="1" :max="65535" controls-position="right" />
      <el-switch v-model="emailImport.config.secure" active-text="SSL/TLS" />
      <el-input v-model="emailImport.config.user" placeholder="邮箱账号" />
      <el-input v-model="emailImport.config.password" placeholder="授权码/应用密码" type="password" show-password />
      <el-input v-model="emailImport.config.mailbox" placeholder="邮箱文件夹，默认 INBOX" />
      <el-date-picker v-model="emailImport.config.since" type="date" value-format="YYYY-MM-DD" placeholder="开始日期" />
      <el-date-picker v-model="emailImport.config.before" type="date" value-format="YYYY-MM-DD" placeholder="结束日期" />
      <el-input v-model="emailImport.config.subjectKeyword" placeholder="标题关键词，可留空" />
      <el-input-number v-model="emailImport.config.limit" :min="1" :max="500" controls-position="right" />
    </div>
    <el-alert
      v-if="emailImport.errorMessage"
      :title="emailImport.errorMessage"
      type="error"
      show-icon
      style="margin-top:12px"
    />
    <template #footer>
      <el-button @click="emailImport.visible = false">取消</el-button>
      <el-button type="primary" :loading="emailImport.loading" @click="runEmailImport">开始导入</el-button>
    </template>
  </el-dialog>
</template>

<script setup lang="ts">
import { computed } from 'vue';
import { ElMessage } from 'element-plus';
import type { StudentRecord } from '@/types/domain';
import { useProjectStore } from '@/stores/project';
import { openExternal } from '@/services/api';
import { ref, onMounted, onBeforeUnmount } from 'vue';
import { readTextFile } from '@/services/api';

const store = useProjectStore();

const reports = (store as any).reports as any;
const emailImport = (store as any).emailImport as any;
const autoReview = (store as any).autoReview as any;
const activeIndex = ref<number | null>(null);

function currentChange(row: any) {
  const idx = store.project.students.findIndex((s: StudentRecord) => s.studentId === row?.studentId);
  activeIndex.value = idx >= 0 ? idx : null;
}

function onSelectDir() {
  store.selectReportsDirectory();
}

function onScan() {
  store.scanReportsDirectory();
}

function openEmailImportDialog() {
  emailImport.visible = true;
}

async function runAutoReview() {
  try {
    const results = await store.autoReviewSubmittedStudents();
    ElMessage.success(`已自动评阅 ${results.length} 份作业`);
  } catch (error) {
    ElMessage.error(`自动评阅失败：${error instanceof Error ? error.message : String(error)}`);
  }
}

function reviewDetails(studentId: string) {
  const result = autoReview.results.find((item: any) => item.studentId === studentId);
  return result?.details ?? [];
}

async function runEmailImport() {
  try {
    const matches = await store.importSubmissionsFromEmail(emailImport.config);
    const matchedCount = matches.filter((item: any) => item.status === 'matched').length;
    ElMessage.success(`已下载 ${matches.length} 个附件，自动匹配 ${matchedCount} 个`);
    emailImport.visible = false;
  } catch (error) {
    ElMessage.error(`邮箱导入失败：${error instanceof Error ? error.message : String(error)}`);
  }
}

function emailStatusType(status: string) {
  if (status === 'matched') return 'success';
  if (status === 'conflict') return 'danger';
  if (status === 'duplicate-name') return 'warning';
  return 'info';
}

function emailStatusText(status: string) {
  if (status === 'matched') return '已匹配';
  if (status === 'conflict') return '冲突';
  if (status === 'duplicate-name') return '姓名重复';
  return '未匹配';
}

function trimPath(p: string) {
  if (!p) return '';
  return p.replace(store.project.settings.defaultExportDir || '', '').replace(/^\\|\//, '');
}

function openExternalFile(p: string) {
  if (!p) return;
  openExternal(p);
}

const preview = ref({ visible: false, path: '', type: '' as 'pdf' | 'image' | 'text' | 'other', text: '' });

async function showPreview(filePath: string) {
  if (!filePath) return;
  const ext = (filePath.split('.').pop() || '').toLowerCase();
  if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'].includes(ext)) {
    preview.value.type = 'image';
    preview.value.path = filePath;
  } else if (ext === 'pdf') {
    preview.value.type = 'pdf';
    preview.value.path = filePath;
  } else if (['txt', 'md', 'csv'].includes(ext)) {
    preview.value.type = 'text';
    preview.value.text = await readTextFile(filePath).catch(() => '无法读取文件');
  } else {
    preview.value.type = 'other';
    preview.value.path = filePath;
  }
  preview.value.visible = true;
}

function closePreview() {
  preview.value.visible = false;
  preview.value.path = '';
  preview.value.text = '';
  preview.value.type = 'other';
}

function onScoreChange(studentId: string, value: number) {
  store.updateStudentScore(studentId, value === undefined || value === null ? null : Number(value));
}

function applyScore(studentId: string, score: number | null) {
  store.updateStudentManualScore(studentId, score ?? null);
}

function clearMatch(studentId: string) {
  const s = store.project.students.find((x: StudentRecord) => x.studentId === studentId);
  if (s) {
    s.submissionFile = '';
    s.submissionStatus = 'missing';
    store.saveCurrentProject();
  }
}

const studentsWithFiles = computed(() => store.project.students.filter((s: StudentRecord) => s.submissionFile || s.submissionStatus === 'submitted'));

function saveCurrent() {
  if (activeIndex.value === null) return;
  const s = studentsWithFiles.value[activeIndex.value];
  if (!s) return;
  store.updateStudentManualScore(s.studentId, s.score ?? null);
}

function onKeydown(e: KeyboardEvent) {
  if (!studentsWithFiles.value.length) return;
  if (e.key === 'ArrowDown') {
    activeIndex.value = Math.min((activeIndex.value ?? -1) + 1, studentsWithFiles.value.length - 1);
  } else if (e.key === 'ArrowUp') {
    activeIndex.value = Math.max((activeIndex.value ?? studentsWithFiles.value.length) - 1, 0);
  } else if (e.key.toLowerCase() === 's') {
    saveCurrent();
  }
}

onMounted(() => window.addEventListener('keydown', onKeydown));
onBeforeUnmount(() => window.removeEventListener('keydown', onKeydown));
</script>

<style scoped>
.page-grid { display:flex; gap:16px; }
.panel-card { flex:1 }
.review-detail {
  padding: 10px 18px;
  color: #455468;
  font-size: 13px;
  line-height: 1.7;
}

.review-detail-title {
  font-weight: 600;
  margin-bottom: 6px;
  color: #1f2a44;
}

.review-detail ul {
  margin: 0;
  padding-left: 18px;
}
</style>

