import { createRouter, createWebHashHistory } from 'vue-router';

const routes = [
  { path: '/', name: 'dashboard', component: () => import('@/views/DashboardView.vue') },
  { path: '/students', name: 'students', component: () => import('@/views/StudentsView.vue') },
  { path: '/timeslots', name: 'timeslots', component: () => import('@/views/TimeSlotsView.vue') },
  { path: '/signin', name: 'signin', component: () => import('@/views/SigninSheetView.vue') },
  { path: '/grading', name: 'grading', component: () => import('@/views/GradingView.vue') },
  { path: '/export', name: 'export', component: () => import('@/views/ExportView.vue') },
  { path: '/settings', name: 'settings', component: () => import('@/views/SettingsView.vue') },
];

export const router = createRouter({
  history: createWebHashHistory(),
  routes,
});
