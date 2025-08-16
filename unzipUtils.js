// unzipUtils.js
const AdmZip = require("adm-zip");
const path = require("node:path"); // Needed for path.join and path.basename
const fs = require("fs"); // Needed for fs.existsSync and fs.mkdirSync

/**
 * Extracts a zip file to a specified directory.
 * @param {string} zipPath - The full path to the zip file.
 * @param {string} extractTo - The directory where the contents should be extracted.
 * @returns {Promise<string|null>} - A Promise that resolves with the extraction path if successful, or null if an error occurs.
 */
async function unzipFile(zipPath, extractTo) {
  console.log(`Attempting to unzip ${zipPath} to ${extractTo}`);
  try {
    // Ensure the extraction directory exists
    if (!fs.existsSync(extractTo)) {
      fs.mkdirSync(extractTo, { recursive: true });
    }

    const zip = new AdmZip(zipPath);
    zip.extractAllTo(extractTo, true); // true overwrites existing files

    console.log(`✅ Successfully unzipped: ${zipPath} to ${extractTo}`);
    return extractTo;
  } catch (error) {
    console.error(`❌ Error unzipping file ${zipPath}: ${error.message}`);
    return null;
  }
}

// Export the function so it can be used in other files
module.exports = unzipFile;
