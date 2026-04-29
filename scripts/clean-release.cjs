const fs = require('node:fs/promises');
const path = require('node:path');

async function main() {
  const releaseDir = path.resolve(__dirname, '..', 'release');
  await fs.rm(releaseDir, { recursive: true, force: true });
  await fs.mkdir(releaseDir, { recursive: true });
  console.log(`Cleaned ${releaseDir}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
