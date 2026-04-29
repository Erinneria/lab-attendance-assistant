const { spawn } = require('node:child_process');

delete process.env.ELECTRON_RUN_AS_NODE;
delete process.env.AUTOTEST_EXPORT_SIGNIN;
delete process.env.AUTOTEST_MANUAL;
process.env.VITE_DEV_SERVER_URL = 'http://127.0.0.1:5173';

const child = spawn(process.execPath, [require.resolve('electron/cli'), '.'], {
  stdio: 'inherit',
  env: process.env,
  shell: false,
});

child.on('exit', (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  process.exit(code ?? 0);
});

child.on('error', (error) => {
  console.error(error);
  process.exit(1);
});
