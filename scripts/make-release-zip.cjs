const crypto = require('node:crypto');
const fs = require('node:fs/promises');
const path = require('node:path');
const JSZip = require('jszip');

async function sha256(filePath) {
  const hash = crypto.createHash('sha256');
  hash.update(await fs.readFile(filePath));
  return hash.digest('hex').toUpperCase();
}

async function addDirectoryToZip(zip, sourceDir, targetDir) {
  const entries = await fs.readdir(sourceDir, { withFileTypes: true });
  for (const entry of entries) {
    const sourcePath = path.join(sourceDir, entry.name);
    const targetPath = `${targetDir}/${entry.name}`;
    if (entry.isDirectory()) {
      await addDirectoryToZip(zip, sourcePath, targetPath);
      continue;
    }
    if (entry.isFile()) {
      zip.file(targetPath, await fs.readFile(sourcePath));
    }
  }
}

async function main() {
  const rootDir = path.resolve(__dirname, '..');
  const releaseDir = path.join(rootDir, 'release');
  const docsDir = path.join(rootDir, 'docs');
  const packageJson = JSON.parse(await fs.readFile(path.join(rootDir, 'package.json'), 'utf-8'));
  const version = packageJson.version;

  const setupName = `LabAttendanceAssistant-Setup-${version}-x64.exe`;
  const unpackedDir = path.join(releaseDir, 'win-unpacked');
  const bundleDirName = `LabAttendanceAssistant-v${version}-Windows`;
  await fs.access(path.join(releaseDir, setupName));
  await fs.access(unpackedDir);

  const zip = new JSZip();
  await addDirectoryToZip(zip, unpackedDir, bundleDirName);
  zip.file(
    `${bundleDirName}/USER_MANUAL.md`,
    await fs.readFile(path.join(docsDir, 'USER_MANUAL.md'), 'utf-8'),
  );

  const zipName = `${bundleDirName}.zip`;
  const zipContent = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
  });
  await fs.writeFile(path.join(releaseDir, zipName), zipContent);

  const checksums = await Promise.all([setupName, zipName].map(async (fileName) => {
    const checksum = await sha256(path.join(releaseDir, fileName));
    return `${checksum}  ${fileName}`;
  }));
  await fs.writeFile(path.join(releaseDir, 'SHA256SUMS.txt'), `${checksums.join('\n')}\n`, 'utf-8');
  await fs.rm(unpackedDir, { recursive: true, force: true });
  await fs.rm(path.join(releaseDir, `${setupName}.blockmap`), { force: true });
  await fs.rm(path.join(releaseDir, 'latest.yml'), { force: true });
  console.log(`Created ${path.join(releaseDir, zipName)}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
