<template>
  <div class="page-grid">
    <div class="panel-card">
      <div class="panel-header">
        <div class="panel-title">签到表生成</div>
        <div class="panel-actions">
          <el-button type="primary" @click="handleExportExcel">导出签到表 Excel</el-button>
          <el-button type="success" style="margin-left:8px" @click="handleExportPdf">导出为 PDF</el-button>
        </div>
      </div>
      <div class="form-grid">
        <el-select v-model="store.project.signinLayout.orientation" placeholder="纸张方向">
          <el-option label="纵向" value="portrait" />
          <el-option label="横向" value="landscape" />
        </el-select>
        <div class="slider-item">
          <label>每页块数：{{ signinLayout.blocksPerPage }}</label>
          <el-slider v-model="signinLayout.blocksPerPage" :min="1" :max="3" :step="1" />
        </div>
        <div class="slider-item">
          <label>每块人数：{{ signinLayout.maxStudentsPerBlock }}</label>
          <el-slider v-model="signinLayout.maxStudentsPerBlock" :min="10" :max="40" :step="1" />
        </div>
        <div class="slider-item">
          <label>字号：{{ signinLayout.fontSize }}</label>
          <el-slider v-model="signinLayout.fontSize" :min="9" :max="14" :step="1" />
        </div>
        <div class="slider-item">
          <label>行高：{{ signinLayout.rowHeight }}</label>
          <el-slider v-model="signinLayout.rowHeight" :min="14" :max="24" :step="1" />
        </div>
        <div class="slider-item">
          <label>列宽参数：{{ signinLayout.blockColumnWidth }}</label>
          <el-slider v-model="signinLayout.blockColumnWidth" :min="8" :max="18" :step="1" />
        </div>
        <div class="slider-item">
          <label>预览缩放：{{ previewZoom }}%</label>
          <el-slider v-model="previewZoom" :min="55" :max="130" :step="5" />
        </div>
        <el-switch v-model="store.project.signinLayout.showGrade" active-text="显示年级" />
        <el-switch v-model="store.project.signinLayout.showClassName" active-text="显示班级" />
        <el-switch v-model="store.project.signinLayout.showGroupName" active-text="显示分组" />
        <el-switch v-model="store.project.signinLayout.showAttendance" active-text="显示签到列" />
        <el-switch v-model="store.project.signinLayout.sortByGroupName" active-text="按组名排序" />
        <el-switch v-model="store.project.signinLayout.sortByTimeSlot" active-text="按时间段排序" />
        <el-switch v-model="store.project.signinLayout.sortByStudentId" active-text="按学号排序" />
        <el-switch v-model="store.project.signinLayout.separatePageByTimeSlot" active-text="每个时间段单独分页" />
      </div>
      <el-divider />
      <div class="preview-wrap">
        <div class="preview-head">
          <strong>打印预览（A4 比例）</strong>
          <div>
            <el-button size="small" @click="prevPage" :disabled="currentPage <= 1">上一页</el-button>
            <el-button size="small" @click="nextPage" :disabled="currentPage >= totalPages">下一页</el-button>
            <span class="page-indicator">第 {{ currentPage }} / {{ totalPages }} 页</span>
          </div>
        </div>

        <div class="a4-stage">
          <div class="a4-page" :style="a4PageStyle">
            <div class="preview-blocks" :style="previewBlocksStyle">
              <div class="preview-block" v-for="(block, idx) in pageBlocksFilled" :key="idx">
                <table class="preview-table" :style="previewTableStyle">
                  <colgroup>
                    <col v-for="(width, colIdx) in previewColumnPercents" :key="colIdx" :style="{ width: `${width}%` }" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th>姓名</th>
                      <th v-if="signinLayout.showGrade">年级</th>
                      <th v-if="signinLayout.showClassName">班级</th>
                      <th v-if="signinLayout.showGroupName">分组</th>
                      <th v-if="signinLayout.showAttendance">签到</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr v-for="(s, i) in block" :key="s.studentId || `empty-${i}`" :style="previewRowStyle">
                      <td>{{ s.name || '' }}</td>
                      <td v-if="signinLayout.showGrade">{{ s.grade || '' }}</td>
                      <td v-if="signinLayout.showClassName">{{ s.className || '' }}</td>
                      <td v-if="signinLayout.showGroupName">{{ s.groupName || '' }}</td>
                      <td v-if="signinLayout.showAttendance"></td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ElMessage } from 'element-plus';
import { useProjectStore } from '@/stores/project';
import { computed, reactive } from 'vue';
import type { StudentRecord } from '@/types/domain';

const store = useProjectStore();
const signinLayout = store.project.signinLayout;

const state = reactive({ currentPage: 1, previewZoom: 100 });

function chunkArray<T>(arr: T[], size: number) {
  const res: T[][] = [];
  for (let i = 0; i < arr.length; i += size) res.push(arr.slice(i, i + size));
  return res;
}

const sortedStudents = computed<StudentRecord[]>(() => store.studentsForDisplay() as StudentRecord[]);

const blocks = computed<StudentRecord[][]>(() => chunkArray(sortedStudents.value, signinLayout.maxStudentsPerBlock));

const visibleColumns = computed(() => {
  const columns = ['姓名'];
  if (signinLayout.showGrade) columns.push('年级');
  if (signinLayout.showClassName) columns.push('班级');
  if (signinLayout.showGroupName) columns.push('分组');
  if (signinLayout.showAttendance) columns.push('签到');
  return columns;
});

const previewColumnWidths = computed(() => {
  const baseWidthPreset: Record<string, number> = {
    姓名: 13,
    年级: 7,
    班级: 10,
    分组: 9,
    签到: 9,
  };
  const scale = Math.max(0.8, Math.min(1.5, (signinLayout.blockColumnWidth || 12) / 12));
  return visibleColumns.value.map((label) => Number(((baseWidthPreset[label] ?? (signinLayout.blockColumnWidth || 12)) * scale).toFixed(2)));
});

const previewColumnPercents = computed(() => {
  const total = previewColumnWidths.value.reduce((sum, value) => sum + value, 0) || 1;
  return previewColumnWidths.value.map((value) => (value / total) * 100);
});

const totalPages = computed(() => Math.max(1, Math.ceil(blocks.value.length / signinLayout.blocksPerPage)));

const pageBlocks = computed<StudentRecord[][]>(() => {
  const p = Math.min(Math.max(1, state.currentPage), totalPages.value);
  const startBlock = (p - 1) * signinLayout.blocksPerPage;
  return blocks.value.slice(startBlock, startBlock + signinLayout.blocksPerPage);
});

const pageBlocksFilled = computed<StudentRecord[][]>(() => {
  const needed = signinLayout.blocksPerPage;
  const base = [...pageBlocks.value];
  const emptyStudent: StudentRecord = {
    studentId: '',
    name: '',
    grade: '',
    className: '',
    department: '',
    originalGroupInfo: '',
    teacherName: '',
    timeSlotId: '',
    timeSlotLabel: '',
    groupName: '',
    attendance: false,
    submissionFile: '',
    submissionStatus: 'missing',
    score: null,
    remark: '',
  };
  while (base.length < needed) base.push([]);
  return base.map((block) => {
    const rows = [...block];
    while (rows.length < signinLayout.maxStudentsPerBlock) {
      rows.push({ ...emptyStudent });
    }
    return rows;
  });
});

const previewZoom = computed({
  get: () => state.previewZoom,
  set: (value: number) => {
    state.previewZoom = value;
  },
});

const a4PageStyle = computed(() => {
  const ratio = signinLayout.orientation === 'portrait' ? 210 / 297 : 297 / 210;
  const baseHeight = 820;
  const width = Math.round(baseHeight * ratio);
  const scale = state.previewZoom / 100;
  return {
    width: `${width}px`,
    height: `${baseHeight}px`,
    transform: `scale(${scale})`,
    transformOrigin: 'top left',
  };
});

const previewBlocksStyle = computed(() => ({
  gridTemplateColumns: `repeat(${signinLayout.blocksPerPage}, minmax(0, 1fr))`,
  gap: `${Math.max(6, Math.round(signinLayout.blockColumnWidth / 2))}px`,
}));

const previewTableStyle = computed(() => ({
  fontSize: `${signinLayout.fontSize}px`,
  tableLayout: 'fixed' as const,
}));

const previewRowStyle = computed(() => ({
  height: `${signinLayout.rowHeight}px`,
}));

function nextPage() { if (state.currentPage < totalPages.value) state.currentPage += 1; }
function prevPage() { if (state.currentPage > 1) state.currentPage -= 1; }

async function handleExportExcel() {
  try {
    const exported = await store.exportSigninSheet();
    if (exported) {
      ElMessage.success('签到表 Excel 已导出');
    } else {
      ElMessage.info('已取消导出');
    }
  } catch (error) {
    ElMessage.error(`导出失败：${error instanceof Error ? error.message : String(error)}`);
  }
}

async function handleExportPdf() {
  try {
    const exported = await store.exportSigninPdf();
    if (exported) {
      ElMessage.success('签到表 PDF 已导出');
    } else {
      ElMessage.info('已取消导出');
    }
  } catch (error) {
    ElMessage.error(`导出失败：${error instanceof Error ? error.message : String(error)}`);
  }
}

const currentPage = computed(() => state.currentPage);

</script>

<style scoped>
.slider-item {
  padding: 6px 8px;
  border: 1px solid #eef2f7;
  border-radius: 8px;
  background: #fafcff;
}

.slider-item label {
  display: block;
  margin-bottom: 4px;
  font-size: 12px;
  color: #425466;
}

.preview-wrap {
  margin-top: 12px;
}

.preview-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 10px;
}

.page-indicator {
  margin-left: 8px;
  color: #5b6b7a;
}

.a4-stage {
  background: #eef2f7;
  border: 1px solid #d9e1ea;
  border-radius: 8px;
  padding: 12px;
  overflow: auto;
  min-height: 520px;
}

.a4-page {
  background: #fff;
  border: 1px solid #cfd8e3;
  box-shadow: 0 2px 8px rgb(15 23 42 / 8%);
  padding: 12px;
  box-sizing: border-box;
}

.preview-blocks {
  display: grid;
  width: 100%;
}

.preview-block {
  border: 1px solid #dbe3ee;
  padding: 4px;
  box-sizing: border-box;
}

.preview-table {
  width: 100%;
  border-collapse: collapse;
}

.preview-table th,
.preview-table td {
  border: 1px solid #e5ebf2;
  padding: 0 4px;
  vertical-align: middle;
  text-align: left;
}

.preview-table th {
  background: #f5f8fc;
  font-weight: 700;
}
</style>
