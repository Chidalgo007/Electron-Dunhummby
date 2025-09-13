// renderer.js
const { ipcRenderer } = require("electron");

document.addEventListener("DOMContentLoaded", () => {
  // Get references to our HTML elements
  const downloadExportButton = document.getElementById("downloadExportButton");
  const uploadBrowseButton = document.getElementById("uploadBrowseButton");
  const downloadSalesButton = document.getElementById("downloadSalesButton");
  const filePathLabel = document.getElementById("filePathLabel");
  const statusLabel = document.getElementById("statusLabel");
  const errorOverlay = document.getElementById("errorOverlay");
  const errorMessage = document.getElementById("errorMessage");
  const errorCloseButton = document.getElementById("errorCloseButton");
  const credsModal = document.getElementById("credentialsModal");
  const closeCredsModal = document.getElementById("closeCredsModal");
  const saveCredsButton = document.getElementById("saveCredsButton");
  const usernameInput = document.getElementById("usernameInput");
  const passwordInput = document.getElementById("passwordInput");
  const fcCalendar = document.getElementById("FSCalendar");
  const destinationFolder = document.getElementById("destinationFolder");

  let selectedFile = null; // To store the path of the selected file

  // Function to update the status message
  function updateStatus(message, type = "info") {
    if (statusLabel) {
      statusLabel.textContent = `Status: ${message}`;
      switch (type) {
        case "error":
          statusLabel.style.color = "red";
          break;
        case "success":
          statusLabel.style.color = "green";
          break;
        default:
          statusLabel.style.color = "rgb(131, 107, 208)"; // Default color
      }
    } else {
      console.log(`Status update: ${message}`);
    }
  }

  // Function to show the error overlay
  function showImportError(msg, type = "Unknown Error") {
    errorMessage.textContent = `${msg}\n\nError Type: ${type}`;
    errorOverlay.classList.remove("hidden"); // Show the overlay
  }

  // Function to hide the error overlay
  function hideErrorOverlay() {
    errorOverlay.classList.add("hidden"); // Hide the overlay
  }

  // Function to enable/disable all main interaction buttons
  function setButtonsEnabled(enabled) {
    downloadExportButton.disabled = !enabled;
    uploadBrowseButton.disabled = !enabled;
    downloadSalesButton.disabled = !enabled;
    // errorCloseButton.disabled = !enabled; // Keep close button enabled even if main buttons are disabled during error
  }

  // Event listeners for buttons

  // Download Export Files Button (Login + Export)
  if (downloadExportButton) {
    downloadExportButton.addEventListener("click", () => {
      updateStatus("Starting download process: Attempting login...");
      setButtonsEnabled(false); // Disable buttons during automation
      ipcRenderer.send("login-request"); // Send message to main process to initiate login
    });
  }

  // Browse/Import Button (dynamic behavior like PyQt)
  if (uploadBrowseButton) {
    uploadBrowseButton.addEventListener("click", () => {
      if (uploadBrowseButton.textContent === "Browse") {
        ipcRenderer.send("select-file-dialog"); // Request file selection from the main process
      } else {
        // It's "Import"
        if (selectedFile) {
          updateStatus("Starting import process: Attempting login...");
          setButtonsEnabled(false); // Disable buttons during automation
          // Send login request for import as well, then trigger import
          ipcRenderer.send("login-request-for-import", selectedFile);
        } else {
          updateStatus("Please select a file to import.");
        }
      }
    });
  }

  // Download Sales File Button
  if (downloadSalesButton) {
    downloadSalesButton.addEventListener("click", () => {
      updateStatus("Starting sales file download: Attempting login...");
      setButtonsEnabled(false); // Disable buttons during automation
      ipcRenderer.send("start-download-sales-file"); // Send message to main process
    });
  }

  // Error Overlay Close Button
  if (errorOverlay) {
    errorCloseButton.addEventListener("click", hideErrorOverlay);
  }

  // --- IPC Renderer Listeners ---
  // Listen for file path response from the main process
  ipcRenderer.on("selected-file", (event, path) => {
    if (path) {
      selectedFile = path;
      filePathLabel.textContent = `Upload file: ${path.split("/").pop()}`; // Display just the filename
      uploadBrowseButton.textContent = "Import";
      updateStatus("File selected. Ready to import.");
    } else {
      filePathLabel.textContent = "Upload file: No file selected";
      uploadBrowseButton.textContent = "Browse";
      selectedFile = null;
      updateStatus("File selection cancelled.");
    }
    setButtonsEnabled(true); // Re-enable buttons after file dialog
  });

  // Listen for status updates from the main process
  ipcRenderer.on("update-status", (event, msg) => {
    updateStatus(msg);
  });

  // Listen for login success for general download/export flow
  ipcRenderer.on("login-success", () => {
    updateStatus("Login successful. Starting export process...");
    // Now that we're logged in, send a request to start the export
    ipcRenderer.send("start-export");
  });

  // Listen for login success for import flow
  ipcRenderer.on("login-success-for-import", (event, filePath) => {
    updateStatus("Login successful. Starting import...");
    // Now that we're logged in, send a request to start the import
    ipcRenderer.send("start-import", filePath);
  });

  // Listen for login failure (for any flow)
  ipcRenderer.on("login-failure", () => {
    showImportError(
      "Failed to log in. Please check your credentials and network connection.",
      "Login Error"
    );
    setButtonsEnabled(true); // Re-enable buttons on failure
  });

  // Listen for errors from the main process (e.g., during export/import)
  ipcRenderer.on("automation-error", (event, msg, type) => {
    showImportError(msg, type);
    setButtonsEnabled(true); // Re-enable buttons on error
  });

  // Listen for export process completion (success or failure)
  ipcRenderer.on("export-finished", (event, success, message) => {
    if (success) {
      updateStatus(`Export completed successfully: ${message}`);
    } else {
      updateStatus(`Export failed: ${message}`);
      showImportError(message, "Export Error");
    }
    setButtonsEnabled(true); // Re-enable buttons after completion
  });

  // Listen for import process completion (success or failure)
  ipcRenderer.on("import-finished", (event, success, message) => {
    if (success) {
      updateStatus(`Import completed successfully: ${message}`);
    } else {
      updateStatus(`Import failed: ${message}`);
      showImportError(message, "Import Error");
    }
    setButtonsEnabled(true); // Re-enable buttons after completion
  });

  // Listen for sales download process completion (success or failure)
  ipcRenderer.on("download-finished", (event, success, message, filePath) => {
    if (success) {
      updateStatus(`✅ Sales file download completed: ${message}`);
      console.log(`Downloaded file path: ${filePath}`); // Path is now included
    } else {
      updateStatus(`❌ Sales file download failed: ${message}`, "error"); // Pass 'error' type
      showImportError(message, "Sales Download Error");
    }
    setButtonsEnabled(true); // Re-enable buttons after completion
  });

  // ======= Credentials creation and retrieval ===========
  // open modal from a menu/settings button
  const openCredsButton = document.getElementById("openCredsButton");
  openCredsButton.addEventListener("click", () => {
    credsModal.classList.remove("hidden");
  });

  // Close modal
  closeCredsModal.addEventListener("click", () => {
    credsModal.classList.add("hidden");
  });

  // Load stored creds on startup
  ipcRenderer.invoke("get-credentials").then((creds) => {
    if (creds.username) usernameInput.value = creds.username;
    if (creds.password) passwordInput.value = creds.password;
  });

  // Save creds
  saveCredsButton.addEventListener("click", async () => {
    const username = usernameInput.value;
    const password = passwordInput.value;
    const result = await ipcRenderer.invoke("save-credentials", {
      username,
      password,
    });
    const storeResult = await ipcRenderer.invoke("save-paths", {
      fcCalendar: fcCalendar.value,
      destinationFolder: destinationFolder.value,
    });
    if (result.success && storeResult.success) {
      updateStatus("✅ Credentials and paths saved.");
      credsModal.classList.add("hidden");
    } else {
      updateStatus(`❌ Failed to save: ${result.message}`, "error");
    }
  });

  // Load stored folder and calendar on startup
  document
    .getElementById("browseFolderButton")
    .addEventListener("click", async () => {
      const folderPath = await ipcRenderer.invoke("select-folder");
      if (folderPath) {
        destinationFolder.value = folderPath;
      } else {
        destinationFolder.value = "No folder selected";
      }
    });

  // Select path for FSCalendar
  document
    .getElementById("browseExcelButton")
    .addEventListener("click", async () => {
      const excelPath = await ipcRenderer.invoke("select-excel-file");
      if (excelPath) {
        fcCalendar.value = excelPath;
      } else {
        fcCalendar.value = "No folder selected";
      }
    });

  // Load stored paths on startup
  ipcRenderer.invoke("get-paths").then((paths) => {
    if (paths.fcCalendar) fcCalendar.value = paths.fcCalendar;
    if (paths.destinationFolder)
      destinationFolder.value = paths.destinationFolder;
  });
});
