const { spawn } = require('child_process');

//const scriptPath = 'utils\\setValsAndRunMacro.vbs';
//const scriptPath = 'utils\\openAndSaveMacroSheet.vbs';
const scriptPath = 'utils\\getValueFromCSV.vbs';

const AValue = '9999';
const BValue = '99990';
const CValue = '1999';

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
