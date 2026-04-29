<template>
  <div class="page-grid">
    <div class="panel-card">
      <div class="panel-header">
        <div class="panel-title">时间分组管理</div>
        <div class="panel-actions">
          <el-button type="primary" @click="store.importTimeSlots">导入时间名单</el-button>
          <el-button @click="store.clearGroupNames">清空组名</el-button>
          <el-button @click="store.refreshMatching">重新匹配</el-button>
        </div>
      </div>
      <div class="table-toolbar">
        <el-tag type="warning">重复组名 {{ duplicateGroupNames.length }}</el-tag>
        <el-tag type="danger">未匹配学生 {{ unmatchedStudentCount }}</el-tag>
      </div>
      <div class="form-grid">
        <el-input v-model="batchStart" placeholder="批量组名起始值，例如 12_51" />
        <el-button type="success" @click="store.batchGenerateGroupNames(batchStart)">批量生成组名</el-button>
      </div>
      <el-table :data="store.project.timeSlots" height="560" stripe border style="margin-top: 16px">
        <el-table-column prop="label" label="时间段" min-width="260" />
        <el-table-column prop="limit" label="限额" width="90" />
        <el-table-column prop="registered" label="已报" width="90" />
        <el-table-column label="状态" width="120">
          <template #default="scope">
            <el-tag v-if="scope.row.groupName && duplicateGroupNames.includes(scope.row.groupName)" type="danger">重复</el-tag>
            <el-tag v-else-if="scope.row.groupName" type="success">已命名</el-tag>
            <el-tag v-else type="info">未命名</el-tag>
          </template>
        </el-table-column>
        <el-table-column label="组名" width="160">
          <template #default="scope">
            <el-input
              :model-value="scope.row.groupName"
              size="small"
              placeholder="输入组名"
              @update:model-value="(value: string | number | null) => store.updateTimeSlotGroupName(scope.row.id, String(value ?? ''))"
            />
          </template>
        </el-table-column>
        <el-table-column label="学生名单" min-width="260">
          <template #default="scope">
            {{ scope.row.students.join('、') }}
          </template>
        </el-table-column>
      </el-table>

      <el-alert
        v-if="duplicateGroupNames.length"
        :title="`检测到重复组名：${duplicateGroupNames.join('、')}`"
        type="error"
        :closable="false"
        style="margin-top: 16px"
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from 'vue';
import type { StudentRecord } from '@/types/domain';
import { useProjectStore } from '@/stores/project';

const store = useProjectStore();
const batchStart = ref('12_51');

const duplicateGroupNames = computed(() => store.duplicateGroupNames());
const unmatchedStudentCount = computed(() => store.project.students.filter((student: StudentRecord) => !student.timeSlotId && student.studentId && student.name).length);
</script>
