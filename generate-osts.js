const fs = require('fs');
const path = require('path');

// Edit this path if your Office Scripts folder is different
const officeScriptsDir = path.resolve(process.env.HOME, 'Library/CloudStorage/OneDrive-Curtin/Documents/Office Scripts/research-log-scripts');

const baseFilename = 'Base Modules.osts';
const actions = [
  { name: 'Add Entry', file: 'Add Entry.osts' },
  { name: 'Shift Entry', file: 'Shift Entry.osts' },
  { name: 'Pop Entry', file: 'Pop Entry.osts' },
  { name: 'Clock In', file: 'Clock In.osts' },
  { name: 'Clock Out', file: 'Clock Out.osts' },
];

function loadBase() {
  const basePath = path.join(officeScriptsDir, baseFilename);
  if (!fs.existsSync(basePath)) {
    console.error('Base Modules file not found at', basePath);
    process.exit(1);
  }
  return fs.readFileSync(basePath, 'utf8');
}

function makeWrapper(actionName) {
  return `function main(workbook) {
  ResearchLog.runFunction(workbook, "${actionName}");
  return;
}`;
}

function writeScripts() {
  const base = loadBase();
  actions.forEach(a => {
    const content = base + '\n\n' + makeWrapper(a.name) + '\n';
    const outPath = path.join(officeScriptsDir, a.file);
    fs.writeFileSync(outPath, content, 'utf8');
    console.log('Wrote', outPath);
  });
}

writeScripts();
console.log('All scripts generated.');
