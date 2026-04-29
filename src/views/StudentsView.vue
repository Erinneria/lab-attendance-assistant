<template>
  <div class="page-grid">
    <div class="panel-card">
      <div class="panel-header">
        <div class="panel-title">学生名单</div>
        <div class="panel-actions">
            <el-button type="primary" @click="store.importStudents">导入学生名单</el-button>
            <el-button @click="store.refreshMatching">重新匹配</el-button>
            <el-select v-model="selectedPreset" clearable placeholder="映射预设" size="small" style="width:200px; margin-left:8px">
              <el-option
                v-for="preset in mappingPresets"
                :key="preset.name"
                :label="preset.name"
                :value="preset.name"
              />
            </el-select>
            <el-button size="small" @click="onApplyPreset" :disabled="!selectedPreset">应用预设</el-button>
            <el-button size="small" @click="onSavePreset">保存预设</el-button>
            <el-button size="small" type="danger" @click="onDeletePreset" :disabled="!selectedPreset">删除预设</el-button>
        </div>
      </div>
      <div class="table-toolbar">
        <el-input v-model="keyword" placeholder="按学号、姓名、班级、分组筛选" clearable style="max-width: 360px" />
        <div class="table-toolbar-right">
          <el-tag type="danger">异常 {{ store.issues.length }}</el-tag>
          <el-tag type="warning">未评分 {{ missingScoreCount }}</el-tag>
        </div>
      </div>
      <el-table :data="filteredStudents" height="560" stripe border style="margin-top: 16px">
        <el-table-column prop="studentId" label="学号" width="120" />
        <el-table-column prop="name" label="姓名" width="100" />
        <el-table-column prop="grade" label="年级" width="90" />
        <el-table-column prop="className" label="班级" width="110" />
        <el-table-column prop="department" label="院系" width="120" />
        <el-table-column prop="originalGroupInfo" label="原始分组信息" min-width="130" />
        <el-table-column prop="timeSlotLabel" label="实验时间" min-width="160" />
        <el-table-column prop="groupName" label="实验组名" width="120" />
        <el-table-column label="签到" width="90">
          <template #default="scope">
            <el-switch :model-value="scope.row.attendance" @change="onAttendanceChange(scope.row.studentId, $event)" />
          </template>
        </el-table-column>
        <el-table-column label="提交状态" width="120">
          <template #default="scope">
            <el-select :model-value="scope.row.submissionStatus" size="small" @change="onSubmissionStatusChange(scope.row.studentId, $event)">
              <el-option label="未提交" value="missing" />
              <el-option label="已提交" value="submitted" />
              <el-option label="手动记录" value="manual" />
            </el-select>
          </template>
        </el-table-column>
        <el-table-column label="分数" width="120">
          <template #default="scope">
            <el-input-number
              :model-value="scope.row.score"
              :min="0"
              :max="100"
              size="small"
              controls-position="right"
              @change="onScoreChange(scope.row.studentId, $event)"
            />
          </template>
        </el-table-column>
        <el-table-column label="备注" min-width="180">
          <template #default="scope">
            <el-input :model-value="scope.row.remark" size="small" placeholder="备注" @change="onRemarkChange(scope.row.studentId, $event)" />
          </template>
        </el-table-column>
      </el-table>
    </div>

    <div class="panel-card">
      <div class="panel-header">
        <div class="panel-title">异常列表</div>
        <el-button size="small" @click="store.refreshMatching">刷新异常</el-button>
      </div>
      <el-scrollbar height="260px">
        <el-empty v-if="!store.issues.length" description="没有异常" />
        <el-alert
          v-for="issue in store.issues"
          :key="issue.message"
          :title="issue.message"
          type="warning"
          :closable="false"
          show-icon
          class="issue-item"
        />
      </el-scrollbar>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from 'vue';
import type { StudentRecord } from '@/types/domain';
import { useProjectStore } from '@/stores/project';

const store = useProjectStore();
const keyword = ref('');
const selectedPreset = ref<string | null>(null);

const mappingPresets = computed(() => store.project.columnMappingPresets ?? []);

const filteredStudents = computed(() => {
  const text = keyword.value.trim();
  if (!text) {
    return store.project.students;
  }
  return store.project.students.filter((student: StudentRecord) => [student.studentId, student.name, student.className, student.groupName]
    .some((field) => String(field).includes(text)));
});

const missingScoreCount = computed(() => store.project.students.filter((student: StudentRecord) => student.score === null).length);

function onAttendanceChange(studentId: string, value: boolean | string | number) {
  store.updateStudentAttendance(studentId, Boolean(value));
}

function onSubmissionStatusChange(studentId: string, value: string | number | boolean) {
  store.updateStudentSubmissionStatus(studentId, String(value) as 'missing' | 'submitted' | 'manual');
}

function onScoreChange(studentId: string, value: number | string | null | undefined) {
  store.updateStudentScore(studentId, value === undefined || value === null ? null : Number(value));
}

function onRemarkChange(studentId: string, value: string | number | null | undefined) {
  store.updateStudentRemark(studentId, String(value ?? ''));
}

function onApplyPreset() {
  if (!selectedPreset.value) return;
  store.applyColumnMappingPreset(selectedPreset.value);
}

function onSavePreset() {
  const name = window.prompt('请输入保存的映射预设名称：', '默认映射');
  if (!name) return;
  store.saveColumnMappingPreset(name.trim());
}

function onDeletePreset() {
  if (!selectedPreset.value) return;
  if (!window.confirm(`确认删除映射预设：${selectedPreset.value} ?`)) return;
  store.deleteColumnMappingPreset(selectedPreset.value);
  selectedPreset.value = null;
}
</script>
