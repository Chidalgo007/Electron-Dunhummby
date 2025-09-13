const { execFile } = require("child_process");
const path = require("path");

function runPythonExcelUpdate(filePath) {
  return new Promise((resolve, reject) => {
    const exePath = path.join(__dirname, "python", "excel.exe"); // your Python exe
    console.log(`🔄 Attempting to run: ${exePath} with arg: ${filePath}`);

    // Pass filePath as an argument to python.exe
    execFile(exePath, [filePath], (error, stdout, stderr) => {
      console.log(`Python stdout: ${stdout}`);
      console.log(`Python stderr: ${stderr}`);
      if (error) {
        return reject(`Python error: ${false}`);
      }
      resolve(true);
    });
  });
}

module.exports = { runPythonExcelUpdate };
