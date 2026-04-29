<template>
  <div class="page-grid">
    <div class="panel-card">
      <div class="panel-header">
        <div class="panel-title">设置</div>
      </div>
      <div class="form-grid">
        <el-switch v-model="store.project.settings.autoSave" active-text="自动保存" />
        <el-input v-model="store.project.settings.defaultExportDir" placeholder="默认导出目录" />
        <el-button @click="store.selectDefaultExportDir">选择默认导出目录</el-button>
      </div>

      <div class="hierarchy-card">
        <div class="hierarchy-title">推荐项目结构</div>
        <ul class="hierarchy-list">
          <li><strong>project/</strong>：项目 JSON，默认按课程名称 + 时间命名</li>
          <li><strong>submissions/</strong>：学生提交作业</li>
          <li><strong>exports/signin/</strong>：签到表导出（PDF / XLSX）</li>
          <li><strong>exports/grades/</strong>：成绩单导出（XLSX）</li>
          <li><strong>imports/</strong>：导入的名单、时间段等原始文件</li>
        </ul>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { useProjectStore } from '@/stores/project';

const store = useProjectStore();
</script>

<style scoped>
.hierarchy-card {
  margin-top: 16px;
  padding: 14px;
  border: 1px solid #e7edf5;
  border-radius: 10px;
  background: #fbfdff;
}

.hierarchy-title {
  font-weight: 600;
  margin-bottom: 6px;
}

.hierarchy-list {
  margin: 0;
  padding-left: 18px;
  color: #455468;
  font-size: 13px;
  line-height: 1.7;
}
</style>
