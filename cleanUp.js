const fs = require("fs");
const path = require("path");

/**
 * Cleans up extracted files - finds the main file, renames it to "Export",
 * and removes the temporary extraction folder
 * @param {string} extractedFolder - Path to the temporary extracted folder
 * @param {string} downloadsFolder - Path to the downloads folder where final file should go
 * @returns {Promise<string|null>} - Path to the final "Export" file or null if failed
 */
async function cleanupExtractedFiles(extractedFolder, downloadsFolder) {
  try {
    console.log(`🧹 Starting cleanup of: ${extractedFolder}`);
    console.log(`📁 Downloads folder: ${downloadsFolder}`);

    // Read all files/folders in the extracted directory
    const items = fs.readdirSync(extractedFolder);
    console.log(`📁 Found ${items.length} items:`, items);

    let mainFile = null;

    // Look for the main file (usually a CSV, XLS, XLSX, etc.)
    for (const item of items) {
      const itemPath = path.join(extractedFolder, item);
      const stat = fs.statSync(itemPath);

      if (stat.isFile()) {
        const ext = path.extname(item).toLowerCase();

        // Look for common data file extensions
        if ([".csv", ".xlsx", ".xls", ".txt"].includes(ext)) {
          mainFile = itemPath;
          console.log(`📄 Found main file: ${item} at ${mainFile}`);
          break;
        }
      } else if (stat.isDirectory()) {
        // If there's a subdirectory, look inside it for the main file
        const subItems = fs.readdirSync(itemPath);
        for (const subItem of subItems) {
          const subItemPath = path.join(itemPath, subItem);
          const subStat = fs.statSync(subItemPath);

          if (subStat.isFile()) {
            const ext = path.extname(subItem).toLowerCase();
            if ([".csv", ".xlsx", ".xls", ".txt"].includes(ext)) {
              mainFile = subItemPath;
              console.log(
                `📄 Found main file in subdirectory: ${subItem} at ${mainFile}`
              );
              break;
            }
          }
        }
        if (mainFile) break;
      }
    }

    if (!mainFile) {
      throw new Error("No suitable main file found in extracted contents");
    }

    // Get the file extension for the final file
    const originalExt = path.extname(mainFile);
    const finalFileName = `Export${originalExt}`;
    const finalFilePath = path.join(downloadsFolder, finalFileName);

    console.log(`🎯 Target final file path: ${finalFilePath}`);

    // Remove existing "Export" file if it exists
    if (fs.existsSync(finalFilePath)) {
      fs.unlinkSync(finalFilePath);
      console.log(`🗑️ Removed existing Export file to replace with new one`);
    }

    // MOVE (not copy) the main file to "Export" - this ensures it's out of the temp folder
    console.log(`📋 Moving from: ${mainFile}`);
    console.log(`📋 Moving to: ${finalFilePath}`);

    try {
      // Use rename to MOVE the file (not copy it)
      fs.renameSync(mainFile, finalFilePath);
      console.log(`✅ File moved successfully to downloads folder`);
    } catch (moveError) {
      console.error(`❌ Failed to move file: ${moveError.message}`);
      // Fallback to copy if rename fails (cross-drive issues)
      console.log(`📋 Trying copy as fallback...`);
      fs.copyFileSync(mainFile, finalFilePath);
      console.log(`✅ File copied successfully as fallback`);
    }

    // Verify the final file exists and get its size
    if (fs.existsSync(finalFilePath)) {
      const stats = fs.statSync(finalFilePath);
      console.log(
        `✅ Final Export file confirmed in downloads folder: ${finalFilePath} (${stats.size} bytes)`
      );
    } else {
      throw new Error(
        `Failed to create final file in downloads folder: ${finalFilePath}`
      );
    }

    // Verify the original file is gone (if we used rename/move)
    if (!fs.existsSync(mainFile)) {
      console.log(`✅ Original file successfully moved out of temp folder`);
    } else {
      console.log(
        `ℹ️ Original file still exists (copy was used instead of move)`
      );
    }

    console.log(
      `🗑️ About to remove temporary extraction folder: ${extractedFolder}`
    );

    // Remove the entire temporary extraction folder
    fs.rmSync(extractedFolder, { recursive: true, force: true });
    console.log(`🗑️ Removed temporary extraction folder successfully`);

    // Final verification
    if (fs.existsSync(finalFilePath)) {
      const finalStats = fs.statSync(finalFilePath);
      console.log(
        `🎉 SUCCESS: Final file still exists after cleanup (${finalStats.size} bytes)`
      );
      return finalFilePath;
    } else {
      throw new Error(`Final file was deleted during cleanup process!`);
    }
  } catch (error) {
    console.error(`❌ Error during cleanup: ${error.message}`);

    // Try to clean up the temp folder even if other operations failed
    try {
      if (fs.existsSync(extractedFolder)) {
        fs.rmSync(extractedFolder, { recursive: true, force: true });
        console.log(`🗑️ Cleaned up temp folder after error`);
      }
    } catch (cleanupError) {
      console.error(
        `❌ Failed to cleanup temp folder: ${cleanupError.message}`
      );
    }

    return null;
  }
}

// Export the function for use in other files
module.exports = { cleanupExtractedFiles };
