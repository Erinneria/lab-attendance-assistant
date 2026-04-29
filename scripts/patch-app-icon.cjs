const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');

const rootDir = path.resolve(__dirname, '..');
const appExePath = path.join(rootDir, 'release', 'win-unpacked', 'Lab Attendance Assistant.exe');
const iconPath = path.join(rootDir, 'build', 'app-icon.ico');
const rceditPath = path.join(rootDir, 'node_modules', 'electron-winstaller', 'vendor', 'rcedit.exe');

for (const filePath of [appExePath, iconPath, rceditPath]) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Missing required file: ${filePath}`);
  }
}

const result = spawnSync(rceditPath, [appExePath, '--set-icon', iconPath], {
  cwd: rootDir,
  stdio: 'inherit',
});

if (result.error) {
  throw result.error;
}

if (result.status !== 0) {
  throw new Error(`rcedit failed with exit code ${result.status}`);
}

console.log(`Patched app icon: ${appExePath}`);
