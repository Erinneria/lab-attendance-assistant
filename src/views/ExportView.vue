<template>
  <div class="page-grid">
    <div class="panel-card">
      <div class="panel-header">
        <div class="panel-title">成绩导出</div>
      </div>
      <el-descriptions :column="2" border style="margin-bottom: 16px">
        <el-descriptions-item label="学生人数">{{ store.stats.studentCount }}</el-descriptions-item>
        <el-descriptions-item label="未评分人数">{{ store.project.students.filter((student: import('@/types/domain').StudentRecord) => student.score === null).length }}</el-descriptions-item>
        <el-descriptions-item label="缺交人数">{{ store.project.students.filter((student: import('@/types/domain').StudentRecord) => student.submissionStatus === 'missing').length }}</el-descriptions-item>
        <el-descriptions-item label="有分数人数">{{ store.project.students.filter((student: import('@/types/domain').StudentRecord) => student.score !== null).length }}</el-descriptions-item>
      </el-descriptions>
      <div class="panel-actions">
        <el-button type="primary" @click="onExport(false)">导出三列表格</el-button>
        <el-button @click="onExport(true)">导出完整记录</el-button>
      </div>
      <el-divider />
      <el-alert title="三列表格默认导出 学号、姓名、分数；完整记录会附加班级、组名、实验时间、提交状态、备注。" type="info" :closable="false" />
    </div>
  </div>
</template>

<script setup lang="ts">
import { ElMessage } from 'element-plus';
import { useProjectStore } from '@/stores/project';

const store = useProjectStore();

async function onExport(fullRecord: boolean) {
  try {
    const exported = await store.exportGradesSheet(fullRecord);
    if (exported) {
      ElMessage.success('成绩单已导出');
    } else {
      ElMessage.info('已取消导出');
    }
  } catch (error) {
    ElMessage.error(`导出失败：${error instanceof Error ? error.message : String(error)}`);
  }
}
</script>
