const fs = require("fs");
const path = require("path");

/**
 * Move a file to a destination folder, replacing any existing file
 * @param {string} sourceFilePath - full path of the file to move
 * @param {string} destinationFolder - full path of the folder to move the file into
 */
function moveFileToDestination(sourceFilePath, destinationFolder) {
  try {
    if (!fs.existsSync(sourceFilePath)) {
      throw new Error(`Source file does not exist: ${sourceFilePath}`);
    }

    if (!fs.existsSync(destinationFolder)) {
      fs.mkdirSync(destinationFolder, { recursive: true });
    }

    const fileName = path.basename(sourceFilePath);
    const destinationFilePath = path.join(destinationFolder, fileName);

    // Move file (will replace if exists)
    fs.renameSync(sourceFilePath, destinationFilePath);

    console.log(`✅ File moved to ${destinationFilePath}`);
    return destinationFilePath;
  } catch (err) {
    console.error(`❌ Failed to move file: ${err.message}`);
    throw err;
  }
}

module.exports = { moveFileToDestination };
