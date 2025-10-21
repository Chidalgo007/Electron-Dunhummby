// main.js
const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const fs = require("fs");
const Store = require("electron-store").default;
const store = new Store();

// keytar credentials
const keytar = require("keytar");
const SERVICE_NAME = "DunnhumbyAutomation"; // unique name for Windows Vault
const LOGIN_URL = store.get("url") || "";
// NOW import playwright and other dependencies
const playwright = require("playwright");
const { chromium, PlaywrightTimeoutError } = playwright;

// Other imports after playwright
const user_data_dir = path.join(app.getPath("userData"), "user-data");

// Import your utility modules last
const unzipFile = require("./unzipUtils");
const { cleanupExtractedFiles } = require("./cleanUp.js");
const { runPythonExcelUpdate } = require("./excelUtils");
const { moveFileToDestination } = require("./moveToDestination.js");

// for sales file download iteration
const CHECK_INTERVAL_MS = 1 * 30 * 60 * 1000; // 30 minutes in milliseconds
const MAX_DOWNLOAD_ATTEMPTS = 5; // number of attempts

let mainWindow;
let browserContext = null;
let currentPage = null;

// Cross-platform Chrome detection function
function findChromeExecutable() {
  const platform = process.platform;
  let chromePaths = [];

  if (platform === "win32") {
    chromePaths = [
      "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
      "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
      path.join(
        process.env.LOCALAPPDATA || "",
        "Google\\Chrome\\Application\\chrome.exe"
      ),
      "C:\\Program Files\\Chromium\\Application\\chrome.exe",
      "C:\\Program Files (x86)\\Chromium\\Application\\chrome.exe",
    ];
  } else if (platform === "darwin") {
    chromePaths = [
      "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
      "/Applications/Chromium.app/Contents/MacOS/Chromium",
    ];
  } else {
    chromePaths = [
      "/usr/bin/google-chrome",
      "/usr/bin/chromium-browser",
      "/usr/bin/chromium",
      "/snap/bin/chromium",
      "/opt/google/chrome/chrome",
    ];
  }

  for (const chromePath of chromePaths) {
    try {
      if (fs.existsSync(chromePath)) {
        return chromePath;
      }
    } catch (error) {
      console.warn(`Error checking path ${chromePath}:`, error.message);
    }
  }
  return null;
}

const createWindow = () => {
  mainWindow = new BrowserWindow({
    width: 450,
    height: 420,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      nodeIntegration: true,
      contextIsolation: false,
    },
    resizable: false,
    autoHideMenuBar: true,
  });

  mainWindow.loadFile("index.html");
};

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", async () => {
  if (browserContext) {
    await browserContext.close();
    console.log("Playwright browser context closed.");
  }
  if (process.platform !== "darwin") {
    app.quit();
  }
});

// --- Core Automation Functions ---

async function loginHumby() {
  const creds = {
    username: await keytar.getPassword(SERVICE_NAME, "username"),
    password: await keytar.getPassword(SERVICE_NAME, "password"),
  };
  const loginUrl = LOGIN_URL;

  if (!creds.username || !creds.password || !loginUrl) {
    const msg = "Credentials or login URL missing. Please set them.";
    console.error(msg);
    mainWindow.webContents.send("automation-error", msg, "Configuration Error");
    return null;
  }

  try {
    if (!browserContext) {
      let executablePath = null;

      // Try to find system Chrome/Chromium first
      executablePath = findChromeExecutable();

      if (!executablePath) {
        const errorMsg = `Chrome or Chromium browser not found on this system.
        Please install one of the following browsers:
        • Google Chrome: https://www.google.com/chrome/
        • Chromium: https://www.chromium.org/getting-involved/download-chromium/

        After installation, restart this application.`;

        mainWindow.webContents.send(
          "automation-error",
          errorMsg,
          "Browser Not Found"
        );
        throw new Error("Chrome or Chromium browser not found.");
      }

      console.log(`🚀 Using system browser: ${executablePath}`);
      browserContext = await chromium.launchPersistentContext(
        path.join(app.getPath("userData"), "playwright_user_data"),
        {
          headless: false,
          executablePath: executablePath,
          args: [
            "--disable-web-security",
            "--disable-features=IsolateOrigins,site-per-process",
          ],
        }
      );
    }

    currentPage = await browserContext.pages()[0];
    console.log(`Navigating to login URL: ${loginUrl}`);
    mainWindow.webContents.send("update-status", "Navigating to login page...");

    await currentPage.goto(loginUrl, { waitUntil: "load", timeout: 30000 });

    await currentPage.waitForSelector("#userNameInput", { timeout: 10000 });
    console.log("Login page detected. Proceeding with login...");
    mainWindow.webContents.send("update-status", "Entering credentials...");

    await currentPage.fill("#userNameInput", "");
    await currentPage.fill("#userNameInput", creds.username, { timeout: 5000 });
    await currentPage.fill("#passwordInput", creds.password, { timeout: 5000 });
    await currentPage.click("#submitButton");

    await currentPage.waitForSelector("text=Reports", { timeout: 20000 });
    console.log("Login successful. Reached landing page.");
    mainWindow.webContents.send("update-status", "Login successful!");

    return currentPage;
  } catch (error) {
    const msg = `Login failed during Playwright automation: ${error.message}`;
    console.error(msg, error);
    mainWindow.webContents.send("automation-error", msg, "Login Error");
    if (currentPage) {
      await currentPage.close();
      currentPage = null;
    }
    return null;
  }
}

/**
 * Translates download_file from exportFile.py.
 * Waits for and handles a single file download within a frame.
 * @param {import('playwright').Page} page - The Playwright page object.
 * @returns {Promise<string|null>} The path of the downloaded file, or null if failed.
 */
async function downloadFile(page) {
  mainWindow.webContents.send(
    "update-status",
    "Waiting for download to finish..."
  );
  let downloadPath = null;
  try {
    const messagesFrame = page.frame({ name: "messages-frame" });
    if (!messagesFrame) {
      throw new Error("Messages frame not found for download.");
    }

    const [download] = await Promise.all([
      page.waitForEvent("download", { timeout: 60000 }),
      messagesFrame.locator("a:has(span.icon-download)").first().click(),
    ]);

    const suggestedFilename = download.suggestedFilename();
    const downloadsFolder = store.get("downloadFolder");
    const finalPath = path.join(downloadsFolder, suggestedFilename);
    await download.saveAs(finalPath);
    downloadPath = finalPath;
    console.log(`✅ Download finished. File saved to: ${finalPath}`);
    mainWindow.webContents.send(
      "update-status",
      `Download finished: ${suggestedFilename}`
    );

    // Use the more precise locator
    const closeButton = messagesFrame.locator("a#message-modal_modalCloseX");

    await closeButton.waitFor({ state: "visible", timeout: 10000 });
    await closeButton.click();
    console.log("✅ Closed message modal using 'a#message-modal_modalCloseX'.");
  } catch (error) {
    const msg = `Error during file download: ${error.message}`;
    console.error(msg, error);
    // It's good practice to re-throw if this function is part of a larger chain
    throw new Error(msg);
  }
  return downloadPath;
}

/**
 * Translates select_group_selection_for_power_bi from exportFile.py.
 * Selects specific group for Power BI export.
 * @param {import('playwright').Page} page - The Playwright page object.
 */
async function selectGroupSelectionForPowerBi(page) {
  console.log("--- Executing select_group_selection_for_power_bi ---");
  mainWindow.webContents.send("update-status", "Performing group selection...");

  try {
    await page
      .locator("button.form-control:has-text('Merch Category by Brand')")
      .click();
    await page.locator("li.ng-binding:has-text('Category Hierarchy')").click();
    await page
      .locator("span.dynatree-node:has(a.dynatree-title:text('Custom Groups'))")
      .click();
    await page
      .locator("span.dynatree-node:has(a.dynatree-title:text('Favorites'))")
      .click();
    await page
      .locator("a.dynatree-title:text('Group Selection for Power BI')")
      .click();
    await page.locator("button.btn-primary:text('Export List')").click();

    await page.waitForTimeout(3000); // Wait 3 sec before hitting cancel

    await page
      .locator("button.btn-default[ng-click=\"$dismiss('Cancelled')\"]")
      .click();
    console.log("--- Group selection and modal interaction finished. ---");
  } catch (error) {
    const errorName = error.name || "";
    const errorMessage = error.message || "";
    if (errorName === "TimeoutError") {
      console.error(
        `Timeout error while locating elements during group selection: ${errorMessage}`
      );
      throw new Error(`Timeout error during group selection: ${errorMessage}`);
    } else {
      console.error(`Error during group selection: ${errorMessage}`, error);
      throw new Error(`Error during group selection: ${errorMessage}`);
    }
  }
}

// helper function to select and download files from message center
async function waitForInboxIncrease(page, increaseBy = 2, timeoutMs = 300000) {
  const start = Date.now();
  const inbox = page.locator("#dh-header-inbox");

  // Read initial value
  let initial = parseInt((await inbox.getAttribute("data-count")) || "0", 10);
  const target = initial + increaseBy;
  console.log(`📩 Inbox started at ${initial}, waiting until ${target}...`);

  while (Date.now() - start < timeoutMs) {
    // Reload page and read new count
    await page.reload({ waitUntil: "networkidle" });
    await page.waitForTimeout(2000);

    const current = parseInt(
      (await inbox.getAttribute("data-count")) || "0",
      10
    );
    console.log(`🔄 Current inbox count: ${current}`);

    if (current >= target) {
      console.log(`✅ Inbox increased by ${increaseBy} (now ${current})`);
      return true;
    }

    // Wait 10 seconds before checking again
    await page.waitForTimeout(10000);
  }

  throw new Error(`❌ Inbox did not increase by ${increaseBy} within timeout.`);
}

/**
 * Translates select_file_for_download function from provided snippet.
 * Navigates to message center, reloads, handles iframe, and clicks message links for downloads.
 * @param {import('playwright').Page} page - The Playwright page object.
 */
async function selectFileForDownload(page) {
  console.log(
    "--- Executing selectFileForDownload (Message Center Download) ---"
  );
  mainWindow.webContents.send(
    "update-status",
    "Checking message center for downloads..."
  );

  try {
    await page.waitForTimeout(10000); // Wait for 10 seconds to ensure files are ready

    // Refresh the page to ensure new messages appear
    await page.reload({ waitUntil: "networkidle" });

    // Wait up to 5 minutes (300000 ms) for data-count=2, refresh every 30s
    await waitForInboxIncrease(page, 2, 600000); // wait for 10 minutes max

    // Once ready, click the message center icon
    await page.locator("a:has(span#dh-header-inbox)").click();

    // Wait for the page to load completely after clicking message center
    await page.waitForLoadState("networkidle", { timeout: 30000 });
    await page.waitForTimeout(5000); // wait 5 sec after reload, to ensure messages are loaded

    // Wait for the iframe with file to download to appear
    const iframeElement = await page.waitForSelector("iframe#messages-frame", {
      timeout: 10000,
    });
    const iframe = await iframeElement.contentFrame();
    if (!iframe) {
      throw new Error("Could not get content frame of messages iframe.");
    }
    await iframe.waitForSelector("div.message-link", { timeout: 10000 });

    // Click the first message link
    const messageLinks = iframe.locator("div.message-link");
    const count = await messageLinks.count(); // Await count()

    if (count > 0) {
      await messageLinks.nth(0).click();
      console.log("✅ Clicked first message link.");
      await downloadFile(page); // Call the download function to handle the first file

      await page.waitForTimeout(2000); // Wait for 2 seconds before next click

      if (count > 1) {
        // Only attempt to click the second if it exists
        await messageLinks.nth(1).click();
        console.log("✅ Clicked second message link.");
        await downloadFile(page); // Call the download function to handle the second file
      } else {
        console.log("Only one message link found for download.");
      }
    } else {
      console.log(
        "❌ No message links found in the message center for download."
      );
    }

    // Go back to the main page
    await page.locator("a#url_myworkspace").click();
    console.log("--- Returned to main workspace after download attempts. ---");
  } catch (error) {
    const errorName = error.name || "";
    const errorMessage = error.message || "";
    if (errorName === "TimeoutError") {
      console.error(
        `Timeout error during message center download process: ${errorMessage}`
      );
      throw new Error(
        `Timeout error during message center download: ${errorMessage}`
      );
    } else {
      console.error(
        `Error during message center download process: ${errorMessage}`,
        error
      );
      throw new Error(`Error during message center download: ${errorMessage}`);
    }
  }
}

/**
 * Translates exportFile.py's export_file function.
 * @param {import('playwright').Page} page - The Playwright page object.
 * @returns {Promise<{success: boolean, message: string, filePath?: string}>}
 */
async function exportFile(page) {
  try {
    mainWindow.webContents.send("update-status", "Starting export process...");
    console.log("=== Starting export process... ===");

    console.log("=== Export Custom Attributes... ===");
    mainWindow.webContents.send(
      "update-status",
      "Navigating to custom attributes export..."
    );
    await page.locator("#url_customattributes").click();
    await page.locator("text=Export / Import").click();
    await page.locator("#import_export_actions").click();
    await page.locator("a:has(span:text('Export custom attributes'))").click();

    await selectGroupSelectionForPowerBi(page); // Perform selection

    console.log("=== Export process completed for Custom Attributes. ===");
    mainWindow.webContents.send(
      "update-status",
      "Custom attributes export complete. Starting standard attributes export..."
    );

    console.log("=== Export Standard Attributes... ===");
    await page.locator("#import_export_actions").click();
    await page
      .locator("a:has(span:text('Export standard attributes'))")
      .click();

    await selectGroupSelectionForPowerBi(page); // Perform selection

    // After export, initiate the download from the message center
    await selectFileForDownload(page); // Call the new function for downloading from message center

    console.log("=== All export processes completed successfully. ===");
    return {
      success: true,
      message:
        "Export and download completed." /* filePath will come from downloadFile's internal logs now */,
    };
  } catch (error) {
    const msg = `Export process failed: ${error.message}`;
    console.error(msg, error);
    return { success: false, message: msg };
  }
}

// ======== Import and Sales Download Functions ========

// Add this class definition
class ImportResult {
  constructor(success = false, hasError = false, message = "") {
    this.success = success;
    this.hasError = hasError;
    this.message = message;
  }
}

/**
 * Translates resubmit_report from importFile.py.
 * Resubmits the report after import.
 * @param {import('playwright').Page} page - The Playwright page object.
 */
// async function resubmitReport(page) {
//   console.log("=== Starting resubmit process... ===");
//   mainWindow.webContents.send("update-status", "Resubmitting report...");
//   try {
//     await page.locator(".btn.dropdown-toggle.undraggable").first().click();
//     await page.locator(".context-menu-text").first().click();
//     await page.waitForTimeout(2000);
//     await page.locator("#btnSubmit").first().click();
//     console.log("✅ Resubmit process initiated.");
//     mainWindow.webContents.send("update-status", "Report resubmitted.");
//   } catch (error) {
//     console.error(`Error during report resubmission: ${error.message}`, error);
//     throw new Error(`Error during report resubmission: ${error.message}`);
//   }
// }

async function resubmitReport(page) {
  console.log("=== Starting resubmit process... ===");
  mainWindow.webContents.send("update-status", "Resubmitting report...");
  try {
    // Find the row containing "Store Level Report"
    const storeReportRow = page
      .locator('table tbody tr:has-text("Store Level Report")')
      .first();

    // Verify the row exists
    const rowCount = await storeReportRow.count();
    if (rowCount === 0) {
      throw new Error("Store Level Report row not found");
    }

    console.log("✅ Found Store Level Report row");

    // Click the actions dropdown within that specific row
    await storeReportRow.locator(".btn.dropdown-toggle.undraggable").click();
    await page.waitForTimeout(500);

    // Click the resubmit option from context menu
    await page.locator(".context-menu-text").first().click();
    await page.waitForTimeout(2000);

    // Click submit button
    await page.locator("#btnSubmit").first().click();

    console.log("✅ Resubmit process initiated.");
    mainWindow.webContents.send("update-status", "Report resubmitted.");
  } catch (error) {
    console.error(`Error during report resubmission: ${error.message}`, error);
    throw new Error(`Error during report resubmission: ${error.message}`);
  }
}

/**
 * Translates expand_if_needed from importFile.py.
 * Expands the dynatree folders if needed.
 * @param {import('playwright').Page} page - The Playwright page object.
 */
async function expandIfNeeded(page) {
  console.log("--- Executing expandIfNeeded ---");
  mainWindow.webContents.send("update-status", "Expanding report folders...");
  try {
    await page.locator("#url_reports>>span:text('Reports')").click();
    await page.waitForLoadState("domcontentloaded"); // Ensure DOM is ready after click

    const labelsToExpand = [
      "Shared",
      "FNZ Nestle New Zealand Limited",
      "9. Store Level",
    ];

    for (const label of labelsToExpand) {
      const node = page.locator(
        `span.dynatree-node:has(span.dynatree-expander):has(a.dynatree-title:has-text('${label}'))`
      );
      const classAttr = (await node.getAttribute("class")) || "";

      if (!classAttr.includes("dynatree-expanded")) {
        console.log(`📂 Expanding: ${label}`);
        const expander = node.locator("span.dynatree-expander");
        await expander.click();
        await page.waitForTimeout(500);
      } else {
        console.log(`✅ Already expanded: ${label}`);
      }
    }

    // Finally, click F&B
    try {
      const fnbLink = page.getByRole("link", { name: "F&B", exact: true });
      if (await fnbLink.isVisible()) {
        await fnbLink.click();
        console.log("🎯 Clicked F&B");
        mainWindow.webContents.send("update-status", "F&B report selected.");
      } else {
        const node = page.locator(
          `span.dynatree-node:has(span.dynatree-expander):has(a.dynatree-title:has-text("9. Store Level"))`
        );
        const expander = node.locator("span.dynatree-expander");
        await expander.click();
        await page.waitForTimeout(500);
        const fnbLink = page.getByRole("link", { name: "F&B", exact: true });
        if (await fnbLink.isVisible()) {
          await fnbLink.click();
          console.log("🎯 Clicked F&B");
          mainWindow.webContents.send("update-status", "F&B report selected.");
        }
      }
    } catch (innerError) {
      throw new Error("❌ F&B link is not visible after expanding.");
    }
  } catch (error) {
    const errorName = error.name || "";
    const errorMessage = error.message || "";

    if (errorName === "TimeoutError") {
      console.error(`Timeout error during folder expansion: ${errorMessage}`);
      throw new Error(`Timeout error during folder expansion: ${errorMessage}`);
    } else {
      console.error(`Error during folder expansion: ${errorMessage}`, error);
      throw new Error(`Error during folder expansion: ${errorMessage}`);
    }
  }
}

/**
 * Translates accept_file from importFile.py.
 * Accepts the file after import and checks the status.
 * @param {import('playwright').Page} page - The Playwright page object.
 * @returns {Promise<ImportResult>}
 */
async function acceptFile(page) {
  mainWindow.webContents.send("update-status", "Checking import status...");
  try {
    // Wait for at least one row to appear in the table
    await page.waitForSelector("table tbody tr", { timeout: 30000 });

    const statusRow = page.locator("table tbody tr").nth(0); // select the first row
    const status = await statusRow.locator("td").nth(2).textContent(); // Get status from third cell
    const statusError = await statusRow.locator("td").nth(5).textContent(); // Check errors from sixth cell

    console.log(`Import Status: ${status}, Errors: ${statusError}`);
    mainWindow.webContents.send(
      "update-status",
      `Import status: ${status}, Errors: ${statusError}`
    );

    if (status === "PENDING") {
      if (statusError === "0") {
        console.log("✅ No errors found in the import process. Accepting...");
        const acceptBtn = page.locator("span[uib-tooltip='Accept']");
        await acceptBtn.waitFor({ state: "visible", timeout: 10000 });
        await acceptBtn.click();
        await page.waitForTimeout(2000);

        const yesBtn = page.locator('button[ng-click="yes()"]');
        await yesBtn.waitFor({ state: "visible", timeout: 10000 });
        await yesBtn.click();
        await page.waitForTimeout(2000);

        const closeXBtn = page.locator('button[aria-label="Close"]');
        if (await closeXBtn.isVisible()) {
          await closeXBtn.click();
          console.log("✅ Closed rejection modal using 'x' button.");
        } else {
          // Fallback to the 'Ok' button if 'x' is not found
          console.log(
            "❌ 'x' button not found, trying 'Ok' button in rejection modal."
          );
          const okBtn = page.locator(
            "button.btn-primary:has(span:has-text('Ok'))"
          );
          await okBtn.waitFor({ state: "visible", timeout: 10000 });
          await okBtn.click();
          console.log("✅ Closed rejection modal using 'Ok' button.");
        }
        return new ImportResult(true, false, "Import accepted successfully");
      } else {
        console.log("⚠️ Errors found in the import process. Rejecting...");
        const rejectBtn = page.locator("span[uib-tooltip='Reject']");
        await rejectBtn.waitFor({ state: "visible", timeout: 10000 });
        await rejectBtn.click();
        await page.waitForTimeout(2000);

        const yesBtn = page.locator('button[ng-click="yes()"]');
        await yesBtn.waitFor({ state: "visible", timeout: 10000 });
        await yesBtn.click();
        await page.waitForTimeout(2000);

        const closeXBtn = page.locator('button[aria-label="Close"]');
        if (await closeXBtn.isVisible()) {
          await closeXBtn.click();
          console.log("✅ Closed rejection modal using 'x' button.");
        } else {
          // Fallback to the 'Ok' button if 'x' is not found
          console.log(
            "❌ 'x' button not found, trying 'Ok' button in rejection modal."
          );
          const okBtnAfterReject = page.locator(
            "button.btn-primary:has(span:has-text('Ok'))"
          );

          await okBtnAfterReject.waitFor({ state: "visible", timeout: 10000 });
          await okBtnAfterReject.click();
          console.log("✅ Closed rejection modal using 'Ok' button.");
        }

        console.log("❌ Import rejected due to errors.");
        return new ImportResult(
          false,
          true,
          `Import rejected due to ${statusError} errors.`
        );
      }
    } else if (status === "REJECTED") {
      console.log(
        "❌ Import process was rejected. Please check the logs for errors."
      );
      return new ImportResult(false, true, "Import rejected by the system.");
    } else {
      console.log("Page is getting refreshed, please wait...");
      mainWindow.webContents.send(
        "update-status",
        "Import status not final, refreshing page..."
      );
      await page.waitForTimeout(3000);
      await page.reload({ waitUntil: "networkidle" });
      // Recursively call acceptFile after reload to recheck status
      return await acceptFile(page);
    }
  } catch (error) {
    const errorName = error.name || "";
    const errorMessage = error.message || "";

    if (errorName === "TimeoutError") {
      console.error(
        `Timeout error while checking import status: ${errorMessage}`
      );
      return new ImportResult(
        false,
        true,
        `Timeout error checking import status: ${errorMessage}`
      );
    } else {
      console.error(
        `An unexpected error occurred in acceptFile: ${errorMessage}`,
        error
      );
      return new ImportResult(
        false,
        true,
        `An unexpected error occurred while checking import status: ${errorMessage}`
      );
    }
  }
}

/**
 * Translates import_file from importFile.py.
 * Imports the specified file to the page.
 * @param {import('playwright').Page} page - The Playwright page object.
 * @param {string} filePath - The path to the file to import.
 * @returns {Promise<ImportResult>}
 */
async function importFile(page, filePath) {
  try {
    console.log("=== Starting import process... ===");
    mainWindow.webContents.send(
      "update-status",
      "Navigating to import page..."
    );
    console.log("=== Import File... === ");

    // Navigate to the import page
    await page.locator("#url_customattributes").click();
    await page.locator("text=Export / Import").click();
    await page.locator("#import_export_actions").click();
    await page.locator("a:has(span:text('Import custom attributes'))").click();
    await page.waitForLoadState("domcontentloaded"); // Ensure page is ready

    // Inject the file and click 'Import Data'
    console.log(`Uploading file: ${filePath}`);
    mainWindow.webContents.send(
      "update-status",
      `Uploading ${path.basename(filePath)}...`
    );
    await page.setInputFiles('input[name="file"]', filePath);
    await page.click('button:has-text("Import Data")');

    // Wait for import to process
    await page.waitForLoadState("networkidle", { timeout: 30000 });
    mainWindow.webContents.send(
      "update-status",
      "Import data sent, checking status..."
    );

    // Check import result
    const result = await acceptFile(page);

    if (result.success) {
      await page.reload(); // Reload the page after successful import
      console.log("=== ✅ Import process completed successfully. ===");
      mainWindow.webContents.send(
        "update-status",
        "Import completed. Resubmitting report..."
      );
      await expandIfNeeded(page);
      await resubmitReport(page);

      // small buffer before checking status
      await page.waitForTimeout(15000);
      await page.reload({ waitUntil: "networkidle" });

      // // === begin resubmit check loop === // not needed as client requested removal
      // let running = false;
      // const timeoutMs = 300000; // 5 min timeout
      // const start = Date.now();

      // while (!running && Date.now() - start < timeoutMs) {
      //   // reload page to get updated status
      //   await page.waitForTimeout(15000); // small buffer
      //   await page.reload({ waitUntil: "networkidle" });

      //   const row = page.locator("table tbody tr").nth(0);
      //   const status_icon = row.locator('span[ng-show="!!objstatus"].icon');
      //   const status_text =
      //     (await status_icon.getAttribute("uib-tooltip")) || "";

      //   console.log(`📊 Current status: ${status_text}`);

      //   if (status_text.toUpperCase().startsWith("RUNNING")) {
      //     running = true;
      //   } else {
      //     console.log("⚠️ Status not RUNNING, resubmitting...");
      //     await resubmitReport(page);
      //   }
      // }

      // if (!running) {
      //   throw new Error(
      //     "❌ Report never reached RUNNING status after resubmit."
      //   );
      // }
      // // === end resubmit check loop ===

      mainWindow.webContents.send(
        "update-status",
        "Report resubmitted after import."
      );
      return new ImportResult(
        true,
        false,
        "Import and resubmit completed successfully."
      );
    } else {
      console.log(
        "=== ❌ Import process failed. Please check the logs for errors. ==="
      );
      return new ImportResult(false, true, result.message);
    }
  } catch (error) {
    const errorName = error.name || "";
    const errorMessage = error.message || "";
    if (errorName === "TimeoutError") {
      console.error(`Timeout error during import process: ${errorMessage}`);
      return new ImportResult(
        false,
        true,
        `Timeout error during import: ${errorMessage}`
      );
    } else {
      console.error(
        `An unexpected error occurred during import process: ${errorMessage}`,
        error
      );
      return new ImportResult(
        false,
        true,
        `An unexpected error occurred during import: ${errorMessage}`
      );
    }
  }
}

/**
 * Checks for file availability and handles download.
 * Translated from Python's checkDownload.py -> file_available.
 * @param {import('playwright').Page} page - The Playwright page object.
 * @returns {Promise<string|null>} - Path to the downloaded file or null if not found/error.
 */

async function fileAvailable(page) {
  const row = page.locator("table tbody tr").nth(0);
  const cell = row.locator("td").nth(1);
  const name = "job" + (await cell.innerText()).trim();
  const status_icon = row.locator('span[ng-show="!!objstatus"].icon');
  const status_text = (await status_icon.getAttribute("uib-tooltip")) || "";
  console.log("Status tooltip:", status_text);

  if (status_text.toUpperCase().startsWith("COMPLETE")) {
    const file_link_locator = row
      .locator("td")
      .nth(4)
      .locator("a:has-text('Store Level Report')");

    const MAX_RETRY_ATTEMPTS = 5;
    let attemptCount = 0;
    let popupPage = null;
    let download = null;

    while (attemptCount < MAX_RETRY_ATTEMPTS) {
      attemptCount++;
      console.log(`📥 Download attempt ${attemptCount}/${MAX_RETRY_ATTEMPTS}`);

      try {
        // 1. Click the link and wait for the popup page to open
        [popupPage] = await Promise.all([
          page.waitForEvent("popup", { timeout: 60000 }),
          file_link_locator.click(),
        ]);

        console.log("✅ New tab opened for download.");

        // 2. Wait for download listener on the popup page
        download = await popupPage.waitForEvent("download", { timeout: 60000 });

        // 3. Close the popup page after getting the download
        if (popupPage && !popupPage.isClosed()) {
          await popupPage.close();
          console.log("✅ Download tab closed.");
        }

        // 4. Save the downloaded file
        let newName = name;
        let ext = path.extname(download.suggestedFilename());
        if (!ext) {
          ext = ".zip";
        }
        const downloadsFolder = store.get("downloadFolder");
        const finalNewPath = path.join(downloadsFolder, newName + ext);
        await download.saveAs(finalNewPath);

        console.log(`✅ Downloaded file: ${finalNewPath}`);
        return finalNewPath;
      } catch (e) {
        console.error(
          `❌ Download attempt ${attemptCount} failed: ${e.message}`
        );

        // Clean up popup after failed attempt
        try {
          if (popupPage && !popupPage.isClosed()) {
            await popupPage.close();
            console.log(
              `✅ Popup page closed after attempt ${attemptCount} error.`
            );
          }
        } catch (closeError) {
          console.error("Error closing popup:", closeError.message);
        }

        // Reset popupPage for next attempt
        popupPage = null;

        // Log specific error types
        const errorName = e.name || "";
        if (errorName === "TimeoutError") {
          console.log(`❌ Timeout occurred on attempt ${attemptCount}.`);
        }

        // If this was the last attempt, return null
        if (attemptCount >= MAX_RETRY_ATTEMPTS) {
          console.error(
            `❌ All ${MAX_RETRY_ATTEMPTS} download attempts failed.`
          );
          return null;
        }

        // Wait a bit before retrying (5 seconds)
        console.log(
          `⏳ Waiting 5 seconds before retry attempt ${attemptCount + 1}...`
        );
        await new Promise((resolve) => setTimeout(resolve, 5000));
      }
    }

    // If loop completes without success
    console.error(
      `❌ Failed to download after ${MAX_RETRY_ATTEMPTS} attempts.`
    );
    return null;
  } else {
    console.log(`Status is not COMPLETE: ${status_text}. Returning null.`);
    return null;
  }
}

/**
 * Executes the full download check process: login, expand, and attempt file download.
 * @param {import('playwright').BrowserContext} context - The Playwright browser context.
 * @returns {Promise<string|null>} - Path to the downloaded file or null if process fails.
 */
async function runCheck(context) {
  // login with context
  const creds = {
    username: await keytar.getPassword(SERVICE_NAME, "username"),
    password: await keytar.getPassword(SERVICE_NAME, "password"),
  };
  const loginUrl = LOGIN_URL;

  if (!creds.username || !creds.password || !loginUrl) {
    const msg =
      "Environment variables EMAIL, PASSWORD, or LOGIN_URL are not set or are empty.";
    console.error(msg);
    mainWindow.webContents.send("automation-error", msg, "Configuration Error");
    return null;
  }

  let page = null;
  try {
    // Create a new page from the provided context
    page = await context.pages()[0];
    console.log(`Navigating to login URL: ${loginUrl}`);
    mainWindow.webContents.send("update-status", "Navigating to login page...");

    await page.goto(loginUrl, { waitUntil: "load", timeout: 30000 });

    await page.waitForSelector("#userNameInput", { timeout: 10000 });
    console.log("Login page detected. Proceeding with login...");
    mainWindow.webContents.send("update-status", "Entering credentials...");

    await page.fill("#userNameInput", "");
    await page.fill("#userNameInput", creds.username, { timeout: 5000 });
    await page.fill("#passwordInput", creds.password, { timeout: 5000 });
    await page.click("#submitButton");

    await page.waitForSelector("text=Reports", { timeout: 20000 });
    console.log("Login successful. Reached landing page.");
    mainWindow.webContents.send("update-status", "Login successful!");
  } catch (error) {
    console.error("❌ Login failed during download process:", error.message);
    if (page) {
      await page.close();
    }
    return null;
  }

  await expandIfNeeded(page);
  await page.waitForLoadState("networkidle"); // Wait for page to settle after expansion
  let downloadAttempts = 0;
  while (downloadAttempts < MAX_DOWNLOAD_ATTEMPTS) {
    try {
      console.log("Attempt FileAvailable: Checking for file availability...");
      const downloadPath = await fileAvailable(page);

      if (downloadPath) {
        console.log("✅ File downloaded successfully on attempt");
        return downloadPath; // Return path on success
      }
      // If fileAvailable returns null (meaning status is not COMPLETE or download timed out in there)
      console.log(
        "Download not yet ready or timed out on attempt. Retrying in 60 minutes..."
      );
      await new Promise((resolve) => setTimeout(resolve, CHECK_INTERVAL_MS)); // Wait for 1 hour
    } catch (error) {
      console.error(`❌ Error during download attempt: ${error.message}`);

      // Use string matching instead of instanceof to avoid import issues
      const errorName = error.name || "";
      const errorMessage = error.message || "";

      if (
        errorName === "TimeoutError" ||
        errorMessage.includes("locator") ||
        errorMessage.includes("navigation")
      ) {
        console.log(
          "Retrying due to Playwright error. Waiting for 60 minutes..."
        );
        await new Promise((resolve) => setTimeout(resolve, CHECK_INTERVAL_MS)); // Wait before next attempt
      } else {
        console.error(
          "Unhandled critical error in runCheck, stopping attempts."
        );
        mainWindow.webContents.send(
          "automation-error",
          error.message,
          "Sales Download Error"
        );
        throw error; // Re-throw critical errors
      }
    }
    downloadAttempts++;
  }

  console.log(
    `❌ Max download attempts (${MAX_DOWNLOAD_ATTEMPTS}) reached. File not downloaded.`
  );
  mainWindow.webContents.send(
    "automation-error",
    "Max download attempts reached",
    "Sales Download Error"
  );
  return null; // Return null if max attempts reached without success
}

// Run the sales download process
async function runSalesDownload() {
  let finalFilePath = null;
  let message = "";
  let calendarMessage = "";
  let success = false;
  const downloadsFolder = store.get("downloadFolder");
  try {
    // Create browser context using system Chrome only
    if (!browserContext) {
      const executablePath = findChromeExecutable();

      if (!executablePath) {
        const errorMsg = `Chrome or Chromium browser not found on this system.
        Please install one of the following browsers:
        • Google Chrome: https://www.google.com/chrome/
        • Chromium: https://www.chromium.org/getting-involved/download-chromium/

        After installation, restart this application.`;

        mainWindow.webContents.send(
          "automation-error",
          errorMsg,
          "Browser Not Found"
        );
        throw new Error("Chrome or Chromium browser not found.");
      }

      console.log(
        `🚀 Using system browser for sales download: ${executablePath}`
      );
      browserContext = await chromium.launchPersistentContext(user_data_dir, {
        headless: false,
        executablePath: executablePath,
        acceptDownloads: true,
        downloadsPath: downloadsFolder,
        args: [
          "--start-maximized",
          "--disable-web-security",
          "--disable-features=IsolateOrigins,site-per-process",
        ],
      });
    }

    // Pass the global 'browserContext' to your runCheck function
    const downloadedPath = await runCheck(browserContext); // 'browserContext' is now properly defined

    if (downloadedPath) {
      console.log(`✅ Sales file downloaded to: ${downloadedPath}`);

      // Now unzip the downloaded file
      const extractTo = path.join(downloadsFolder, "temp_extracted");
      console.log(`📁 Starting unzip process...`);

      const unzipResult = await unzipFile(downloadedPath, extractTo);

      if (unzipResult) {
        console.log(`✅ File successfully unzipped to: ${unzipResult}`);
        try {
          // Clean up and rename the extracted file
          finalFilePath = await cleanupExtractedFiles(
            extractTo,
            downloadsFolder
          );

          const destinationFolder = store.get("destinationFolder") || null;
          let destinationFilePath = null;
          destinationFilePath = await moveFileToDestination(
            finalFilePath,
            destinationFolder
          );

          if (destinationFilePath) {
            console.log(
              `✅ File cleaned up and renamed to: ${destinationFilePath}`
            );

            // Delete the original zip file
            if (fs.existsSync(downloadedPath)) {
              fs.unlinkSync(downloadedPath);
              console.log(`🗑️ Deleted original zip file: ${downloadedPath}`);
            }

            // Send success message with final file path
            mainWindow.webContents.send(
              "download-finished",
              true,
              "Sales file downloaded, extracted, and cleaned up, and moved to destination successfully.",
              destinationFilePath
            );
            console.log(
              "✅ Sales file download, cleanup, and GUI signaling complete."
            );
          } else {
            throw new Error(
              "Failed to move file to destination folder.\n file is at: " +
                finalFilePath
            );
          }
        } catch (cleanupError) {
          console.error("❌ Error during file cleanup:", cleanupError.message);

          mainWindow.webContents.send(
            "download-finished",
            false,
            `File extracted but cleanup failed: ${cleanupError.message}`,
            unzipResult
          );
        }
      } else {
        console.log("❌ File downloaded but unzip failed.");

        // Send partial success message - downloaded but not unzipped
        mainWindow.webContents.send(
          "download-finished",
          false,
          "Sales file downloaded but failed to unzip.",
          null
        );
      }

      // 📅 Call FSCalendar.xlsx update function

      // get path to FSCalendar.xlsx from electron-store
      const excelPath = store.get("fcCalendar") || "";
      console.log(`📅 Starting FSCalendar.xlsx update process...`);

      try {
        if (!excelPath || !fs.existsSync(excelPath)) {
          throw new Error(`Excel file not found at path: ${excelPath}`);
        } else {
          console.log(`📅 Found Excel file at path: ${excelPath}`);
          const result = await runPythonExcelUpdate(excelPath);
          calendarMessage = `FSCalendar.xlsx updated successfully. ${result}`;
        }
      } catch (err) {
        calendarMessage = `FSCalendar update failed: ${err}`;
      }
      success = true;
      message = `Sales file processed successfully. ${calendarMessage}`;
      console.log("✅ Sales file download and calendar update complete.");

      // End Calendar update
    } else {
      console.log("❌ Sales file download failed or file not found.");
      message = "Failed to download sales file.";
    }
  } catch (error) {
    console.error(
      "❌ Unhandled error during sales file download/unzip process:",
      error
    );
    message = `An unexpected error occurred: ${error.message}`;
  } finally {
    // Close browser ONLY after both download and unzip are complete (or failed)
    console.log("🔄 Attempting to close browser context...");
    mainWindow.webContents.send(
      "download-finished",
      success,
      message,
      finalFilePath
    );
    try {
      if (browserContext) {
        // For Playwright contexts, we don't need to check isClosed(), just try to close
        await browserContext.close();
        console.log("🔒 Playwright browser context closed successfully.");
      } else {
        console.log("ℹ️ Browser context was null.");
      }
    } catch (closeError) {
      console.error("❌ Error closing browser context:", closeError.message);
    } finally {
      browserContext = null; // Always nullify the global variable
    }
  }
}

// --- IPC Main Handlers ---

ipcMain.on("select-file-dialog", async (event) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openFile"],
    filters: [
      { name: "Excel Files", extensions: ["xlsx", "xls"] },
      { name: "All Files", extensions: ["*"] },
    ],
  });

  if (!result.canceled && result.filePaths.length > 0) {
    event.sender.send("selected-file", result.filePaths[0]);
  } else {
    event.sender.send("selected-file", null);
  }
});

ipcMain.on("login-request", async (event) => {
  mainWindow.webContents.send("update-status", "Attempting login...");
  const page = await loginHumby();
  if (page) {
    event.sender.send("login-success");
  } else {
    event.sender.send("login-failure");
  }
});

ipcMain.on("start-export", async (event) => {
  if (!currentPage) {
    const msg =
      "Browser page not available. Login may have failed or timed out.";
    mainWindow.webContents.send("automation-error", msg, "Automation Error");
    event.sender.send("export-finished", false, msg);
    return;
  }
  const result = await exportFile(currentPage);
  event.sender.send("export-finished", result.success, result.message);

  if (browserContext) {
    await browserContext.close();
    browserContext = null; // Clear the reference
    currentPage = null; // Clear the reference
    console.log("Playwright browser context closed after export.");
  }
});

ipcMain.on("login-request-for-import", async (event, filePath) => {
  mainWindow.webContents.send(
    "update-status",
    "Attempting login for import..."
  );
  const page = await loginHumby();
  if (page) {
    event.sender.send("login-success-for-import", filePath);
  } else {
    event.sender.send("login-failure");
  }
});

// Locate this existing block in main.js
ipcMain.on("start-import", async (event, filePath) => {
  if (!currentPage) {
    const msg =
      "Browser page not available for import. Login may have failed or timed out.";
    mainWindow.webContents.send("automation-error", msg, "Automation Error");
    event.sender.send("import-finished", false, msg);
    return;
  }

  mainWindow.webContents.send(
    "update-status",
    `Starting import of ${path.basename(filePath)}...`
  );
  let result;
  try {
    result = await importFile(currentPage, filePath);
    event.sender.send("import-finished", result.success, result.message);

    if (result.success) {
      // close browser context after import
      if (browserContext) {
        try {
          await browserContext.close();
        } catch (closeErr) {
          console.error("Error closing browser context:", closeErr);
        }
        browserContext = null;
        currentPage = null;
      }
      // ✅ Schedule sales download after 1 hour and 30 minutes
      const oneHourThirtyMinutesMs = 1 * 90 * 60 * 1000;
      mainWindow.webContents.send(
        "update-status",
        "⏳ Import successful, scheduling sales download in 3 hours..."
      );
      setTimeout(() => {
        mainWindow.webContents.send(
          "update-status",
          "🚀 Triggering scheduled sales download..."
        );
        runSalesDownload(); // run sales download after 1 hour and 30 minutes
      }, oneHourThirtyMinutesMs);
    }
  } catch (err) {
    console.error("Import process crashed:", err);
    event.sender.send("import-finished", false, err.message || "Unknown error");
  }
});

// IPC handler sales download file
// IPC Handlers for Sales Download Flow (login, then start sales download)

ipcMain.on("start-download-sales-file", async (event, args) => {
  await runSalesDownload();
});

ipcMain.handle("save-credentials", async (event, { username, password }) => {
  try {
    await keytar.setPassword(SERVICE_NAME, "username", username);
    await keytar.setPassword(SERVICE_NAME, "password", password);
    return { success: true, message: "Credentials saved securely." };
  } catch (err) {
    return { success: false, message: err.message };
  }
});

ipcMain.handle("get-credentials", async () => {
  try {
    const username = await keytar.getPassword(SERVICE_NAME, "username");
    const password = await keytar.getPassword(SERVICE_NAME, "password");
    return { username, password };
  } catch (err) {
    return { username: null, password: null, error: err.message };
  }
});

// listen for folder selection
ipcMain.handle("select-folder", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openDirectory"],
  });

  if (result.canceled) return null;
  return result.filePaths[0];
});

// listen for excel file selection
ipcMain.handle("select-excel-file", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openFile"],
    filters: [{ name: "Excel Files", extensions: ["xlsx", "xls"] }],
  });

  if (result.canceled) return null;
  return result.filePaths[0];
});

// Save Excel and Destination folder paths
ipcMain.handle(
  "save-paths",
  async (event, { fcCalendar, downloadFolder, destinationFolder, url }) => {
    try {
      store.set("fcCalendar", fcCalendar);
      store.set("destinationFolder", destinationFolder);
      store.set("downloadFolder", downloadFolder);
      store.set("url", url);
      return { success: true, message: "Paths saved successfully." };
    } catch (err) {
      return { success: false, message: err.message };
    }
  }
);

// Get Excel and Destination folder paths
ipcMain.handle("get-paths", async () => {
  return {
    fcCalendar: store.get("fcCalendar") || "",
    destinationFolder: store.get("destinationFolder") || "",
    downloadFolder: store.get("downloadFolder") || "",
    url: store.get("url") || "",
  };
});
