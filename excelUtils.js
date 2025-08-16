// utils/excelUtils.js
const { execFile } = require("child_process");
const path = require("path");

function runPythonExcelUpdate() {
  return new Promise((resolve, reject) => {
    const exePath = path.join(__dirname, "python", "excel.exe"); // python executable name
    console.log(`🔄 Attempting to run: ${exePath}`); // Debug log

    execFile(exePath, (error, stdout, stderr) => {
      console.log(`Python stdout: ${stdout}`); // Debug log
      console.log(`Python stderr: ${stderr}`); // Debug log
      if (error) {
        return reject(`Python error: ${stderr || error.message}`);
      }
      resolve(stdout.trim());
    });
  });
}

module.exports = { runPythonExcelUpdate };
