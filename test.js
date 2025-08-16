const { runPythonExcelUpdate } = require("./excelUtils");

async function testPythonExecution() {
  console.log("🚀 Starting Python EXE test...");

  try {
    const result = await runPythonExcelUpdate();
    console.log("✅ Test succeeded! Output:", result);
  } catch (error) {
    console.error("❌ Test failed! Error:", error);
  }
}

testPythonExecution();
