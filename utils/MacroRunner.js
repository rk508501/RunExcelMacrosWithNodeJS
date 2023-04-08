const { spawn } = require('child_process');

const scriptPath = 'C:\\Users\\kulka\\Downloads\\runMacro.vbs';
const AValue = '70';
const BValue = '90';
const CValue = '10';

const scriptProcess = spawn('cscript.exe', [scriptPath, AValue, BValue, CValue]);

scriptProcess.stdout.on('data', (data) => {
  // This event handler will be called when the VBScript writes output to stdout.
  // In this case, the output is the value of cell D1, so we'll log it to the console.
  console.log(`Output: ${data}`);
});

scriptProcess.stderr.on('data', (data) => {
  // This event handler will be called if there's an error running the VBScript.
  console.error(`Error: ${data}`);
});

scriptProcess.on('close', (code) => {
  // This event handler will be called when the VBScript process exits.
  console.log(`Process exited with code ${code}`);
});
