<template>
  <el-container class="app-shell">
    <el-aside class="sidebar" width="280px">
      <div class="brand-card">
        <div class="brand-badge">L</div>
        <div>
          <h1>实验课签到与成绩管理助手</h1>
          <p>本地离线 · Excel 驱动 · 可保存项目</p>
        </div>
      </div>
      <el-menu :default-active="route.path" class="nav-menu" router>
        <el-menu-item index="/">项目首页</el-menu-item>
        <el-menu-item index="/students">学生名单</el-menu-item>
        <el-menu-item index="/timeslots">时间分组</el-menu-item>
        <el-menu-item index="/signin">签到表生成</el-menu-item>
        <el-menu-item index="/grading">报告批改</el-menu-item>
        <el-menu-item index="/export">成绩导出</el-menu-item>
        <el-menu-item index="/settings">设置</el-menu-item>
      </el-menu>
    </el-aside>

    <el-container class="content-shell">
      <el-header class="topbar">
        <div>
          <div class="topbar-title">{{ hasProject ? store.project.courseName : '没有项目' }}</div>
          <div class="topbar-subtitle">
            <span v-if="hasProject">
              {{ store.project.term || '未设置时间' }}
              <span v-if="projectFileNamePreview"> · 建议保存名：{{ projectFileNamePreview }}</span>
            </span>
            <span v-else>请先新建或打开一个项目目录</span>
          </div>
        </div>
        <div class="topbar-actions">
          <el-button type="primary" plain @click="handleCreateNewProject">新建</el-button>
          <el-button @click="handleOpenProject">打开</el-button>
          <el-button type="success" @click="handleSaveProject">保存</el-button>
        </div>
      </el-header>

      <el-main class="main-area">
        <router-view />
      </el-main>
    </el-container>
  </el-container>

  <el-dialog v-model="createProjectDialogVisible" title="新建项目" width="640px">
    <div class="create-project-form">
      <label>
        <span>课程名称</span>
        <el-input v-model="createProjectForm.courseName" placeholder="请输入课程名称" />
      </label>
      <label>
        <span>时间</span>
        <el-input v-model="createProjectForm.term" placeholder="例如 2025 秋 / 第 1 周" />
      </label>
      <div class="create-project-hint">
        项目会创建为独立文件夹，默认名称由“课程名称 + 时间”组成。
      </div>
    </div>
    <template #footer>
      <el-button @click="createProjectDialogVisible = false">取消</el-button>
      <el-button type="primary" @click="confirmCreateProject">创建</el-button>
    </template>
  </el-dialog>

  <el-dialog v-model="openProjectDialogVisible" title="打开项目" width="640px">
    <div class="create-project-form">
      <label>
        <span>已有项目</span>
        <el-select v-model="selectedProjectDir" placeholder="请选择项目" style="width:100%">
          <el-option
            v-for="item in availableProjects"
            :key="item.path"
            :label="item.name"
            :value="item.path"
          />
        </el-select>
      </label>
      <div class="create-project-hint">
        项目列表来自应用内部项目库目录。
      </div>
    </div>
    <template #footer>
      <el-button @click="openProjectDialogVisible = false">取消</el-button>
      <el-button type="primary" :disabled="!selectedProjectDir" @click="confirmOpenProject">打开</el-button>
    </template>
  </el-dialog>

  <el-dialog v-model="store.studentImport.visible" title="确认学生列映射" width="720px">
    <div class="mapping-grid">
      <label v-for="field in mappingFields" :key="field.key">
        <span>{{ field.label }}</span>
        <el-select v-model="store.studentImport.mapping[field.key]" filterable placeholder="选择列名" @change="store.updateStudentMappingField(field.key, store.studentImport.mapping[field.key] || '')">
          <el-option v-for="header in store.studentImport.headers" :key="header" :label="header" :value="header" />
        </el-select>
      </label>
    </div>

    <el-alert v-if="store.studentImport.errorMessage" :title="store.studentImport.errorMessage" type="warning" show-icon style="margin-top:8px" />

    <div style="margin-top:12px">
      <div style="font-weight:600">导入预览（前 5 行）</div>
      <el-table :data="previewRows" size="mini" style="margin-top:8px">
        <el-table-column prop="studentId" label="学号" width="120" />
        <el-table-column prop="name" label="姓名" width="140" />
        <el-table-column prop="className" label="班级" width="120" />
      </el-table>
    </div>
    <template #footer>
      <el-button @click="store.closeStudentMappingDialog">取消</el-button>
      <el-button @click="onSavePreset">保存为预设</el-button>
      <el-button type="primary" @click="store.applyStudentMapping">确认导入</el-button>
    </template>
  </el-dialog>

  <el-dialog v-model="store.timeSlotImport.visible" title="时间分组已导入" width="640px">
    <div class="dialog-body">
      <p>时间分组已解析完成，请前往「时间分组」页面设置组名并完成匹配。</p>
      <p class="muted">文件：{{ store.timeSlotImport.filePath }}</p>
    </div>
    <template #footer>
      <el-button type="primary" @click="store.closeTimeSlotDialog">知道了</el-button>
    </template>
  </el-dialog>
</template>

<script setup lang="ts">
import { computed, reactive, ref } from 'vue';
import { ElMessage } from 'element-plus';
import { useRoute } from 'vue-router';
import { useProjectStore } from '@/stores/project';
import { buildStudentRecords } from '@/services/project';

const route = useRoute();
const store = useProjectStore();
const hasProject = computed(() => Boolean(store.projectDir.value));
const createProjectDialogVisible = ref(false);
const createProjectForm = reactive({ courseName: '', term: '' });
const openProjectDialogVisible = ref(false);
const selectedProjectDir = ref('');
const availableProjects = ref<Array<{ name: string; path: string }>>([]);

const mappingFields = computed(() => [
  { key: 'studentId', label: '学号' },
  { key: 'name', label: '姓名' },
  { key: 'grade', label: '年级' },
  { key: 'className', label: '班级' },
  { key: 'department', label: '所属院系' },
  { key: 'originalGroupInfo', label: '已选分组信息' },
  { key: 'teacherName', label: '教师名' },
] as const);

const previewRows = computed(() => {
  try {
    const rows = buildStudentRecords(store.studentImport.rows, store.studentImport.mapping);
    return rows.slice(0, 5);
  } catch (e) {
    return [];
  }
});

function sanitizeFileNamePart(value: string) {
  return String(value ?? '')
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, ' ')
    .replace(/^\.+/, '')
    .replace(/\.+$/, '');
}

const projectFileNamePreview = computed(() => {
  const parts = [sanitizeFileNamePart(store.project.courseName), sanitizeFileNamePart(store.project.term)].filter(Boolean);
  return `${parts.length ? parts.join(' - ') : 'lab-project'}.json`;
});

function openCreateProjectDialog() {
  createProjectForm.courseName = store.project.courseName || '';
  createProjectForm.term = store.project.term || '';
  createProjectDialogVisible.value = true;
}

function onSavePreset() {
  const name = window.prompt('请输入映射预设名称：', '默认映射');
  if (!name) return;
  store.saveStudentImportMappingPreset(name.trim());
}

async function handleCreateNewProject() {
  openCreateProjectDialog();
}

async function handleOpenProject() {
  try {
    const projects = await store.listAvailableProjects();
    if (!projects.length) {
      ElMessage.info('项目库里还没有可打开的项目');
      return;
    }
    if (projects.length === 1) {
      const opened = await store.openProjectFromDirectory(projects[0].path);
      if (opened) {
        ElMessage.success('项目已打开');
      }
      return;
    }
    availableProjects.value = projects;
    selectedProjectDir.value = projects[0]?.path || '';
    openProjectDialogVisible.value = true;
  } catch (error) {
    ElMessage.error(`打开失败：${error instanceof Error ? error.message : String(error)}`);
  }
}

async function confirmOpenProject() {
  try {
    const opened = await store.openProjectFromDirectory(selectedProjectDir.value);
    if (opened) {
      openProjectDialogVisible.value = false;
      ElMessage.success('项目已打开');
    } else {
      ElMessage.info('已取消打开');
    }
  } catch (error) {
    ElMessage.error(`打开失败：${error instanceof Error ? error.message : String(error)}`);
  }
}

async function confirmCreateProject() {
  try {
    const created = await store.createProjectFromMetadata(createProjectForm.courseName, createProjectForm.term);
    if (created) {
      createProjectDialogVisible.value = false;
      ElMessage.success('已创建项目');
    }
  } catch (error) {
    ElMessage.error(`新建失败：${error instanceof Error ? error.message : String(error)}`);
  }
}

async function handleSaveProject() {
  try {
    const saved = await store.saveCurrentProject();
    if (saved) {
      ElMessage.success('项目已保存');
    } else {
      ElMessage.info('已取消保存');
    }
  } catch (error) {
    ElMessage.error(`保存失败：${error instanceof Error ? error.message : String(error)}`);
  }
}

</script>
