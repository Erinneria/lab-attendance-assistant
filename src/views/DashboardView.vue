<template>
  <div class="page-grid">
    <div class="metric-row">
      <div class="metric-card">
        <div class="metric-label">学生人数</div>
        <div class="metric-value">{{ store.stats.studentCount }}</div>
      </div>
      <div class="metric-card">
        <div class="metric-label">已分组人数</div>
        <div class="metric-value">{{ store.stats.matchedCount }}</div>
      </div>
      <div class="metric-card">
        <div class="metric-label">已提交人数</div>
        <div class="metric-value">{{ store.stats.submittedCount }}</div>
      </div>
      <div class="metric-card">
        <div class="metric-label">已评分人数</div>
        <div class="metric-value">{{ store.stats.scoredCount }}</div>
      </div>
    </div>

    <div v-if="hasProject" class="split-layout">
      <div class="panel-card">
        <div class="panel-header">
          <div class="panel-title">项目概览</div>
          <div class="panel-actions">
            <el-button size="small" @click="store.importStudents">导入学生名单</el-button>
            <el-button size="small" @click="store.importTimeSlots">导入时间分组</el-button>
            <el-button size="small" type="primary" @click="store.refreshMatching">重新匹配</el-button>
          </div>
        </div>
        <div class="project-meta-form">
          <label>
            <span>课程名称</span>
            <el-input v-model="store.project.courseName" placeholder="请输入课程名称" />
          </label>
          <label>
            <span>时间</span>
            <el-input v-model="store.project.term" placeholder="例如 2025 秋 / 第 1 周" />
          </label>
        </div>
        <div class="project-meta-hint">
          保存项目时会默认使用“课程名称 - 时间.json”作为文件名。
        </div>
      </div>

      <div class="panel-card">
        <div class="panel-header">
          <div class="panel-title">异常检查</div>
          <el-tag type="danger">{{ store.stats.issueCount }} 条</el-tag>
        </div>
        <el-scrollbar height="260px">
          <el-empty v-if="!issues.length" description="当前没有异常" />
          <el-alert v-for="issue in issues" :key="issue.message" :title="issue.message" type="warning" :closable="false" show-icon class="issue-item" />
        </el-scrollbar>
      </div>
    </div>

    <div v-else class="panel-card empty-project-card">
      <el-empty description="当前没有打开任何项目">
        <div style="display:flex;gap:12px;justify-content:center;flex-wrap:wrap;color:#6b7280;font-size:13px">
          请使用顶部“打开”按钮从内部项目库选择项目
        </div>
      </el-empty>
    </div>

    <div v-if="hasProject" class="panel-card">
      <div class="panel-header">
        <div class="panel-title">快捷入口</div>
      </div>
      <div class="panel-actions">
        <el-button type="primary" @click="$router.push('/students')">学生名单</el-button>
        <el-button @click="$router.push('/timeslots')">时间分组</el-button>
        <el-button @click="$router.push('/signin')">签到表生成</el-button>
        <el-button @click="$router.push('/grading')">报告批改</el-button>
        <el-button @click="$router.push('/export')">成绩导出</el-button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue';
import { useProjectStore } from '@/stores/project';

const store = useProjectStore();
const hasProject = computed(() => Boolean(store.projectDir.value));
const issues = computed(() => store.sortedIssues());
</script>

<style scoped>
.project-meta-form {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 12px;
  margin-top: 4px;
}

.project-meta-form label {
  display: flex;
  flex-direction: column;
  gap: 6px;
  font-size: 13px;
  color: #455468;
}

.project-meta-hint {
  margin-top: 10px;
  font-size: 12px;
  color: #6b7280;
}
</style>
