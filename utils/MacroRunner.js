const { spawn } = require('child_process');

const scriptPath = 'utils\\setValsAndRunMacro.vbs';
const AValue = '9';
const BValue = '90';
const CValue = '1';

const scriptProcess = spawn('cscript.exe', [scriptPath, AValue, BValue, CValue]);

scriptProcess.stdout.on('data', (data) => {  
  console.log(`Output: ${data}`);
});

scriptProcess.stderr.on('data', (data) => {
  console.error(`Error: ${data}`);
});

scriptProcess.on('close', (code) => {
  console.log(`Process exited with code ${code}`);
});
