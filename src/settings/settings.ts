/* global Office */

// Use MutationObserver to attach event listener when element is added to DOM
const observer = new MutationObserver((mutations) => {
  mutations.forEach((mutation) => {
    if (mutation.type === "childList" && mutation.addedNodes.length > 0) {
      attachSaveSettingsListener();
      observer.disconnect(); // Disconnect observer after listener is attached
    }
  });
});

// Start observing changes in the document body
observer.observe(document.body, { childList: true, subtree: true });

// Call loadSettings to populate the input fields with saved settings
document.addEventListener("DOMContentLoaded", loadSettings);

function attachSaveSettingsListener() {
  const saveSettingsButton = document.getElementById("save-settings");
  if (saveSettingsButton) {
    saveSettingsButton.addEventListener("click", saveSettings);
    console.log("settings.ts: Event listener attached to save-settings button");
  } else {
    console.error("settings.ts: Element with id 'save-settings' not found");
  }
}

async function saveSettings() {
  console.log("settings.ts: saveSettings called");
  const syncroUrl = (document.getElementById("syncro-url") as HTMLInputElement).value;
  const syncroApiKey = (document.getElementById("syncro-api-key") as HTMLInputElement).value;
  const errorMessageDiv = document.getElementById("error-message");

  if (!syncroUrl || !syncroApiKey) {
    showStatus("Please enter both Syncro URL and API Key.", "error");
    return;
  }

  try {
    await saveSyncroSettings(syncroUrl, syncroApiKey);
    console.log("settings.ts: Settings saved successfully");
    showStatus("Settings saved successfully!", "success");
    // Redirect back to the main taskpane
    window.close();
  } catch (error) {
    console.error("settings.ts: Error saving settings:", error);
    showStatus("Error saving settings: " + error.message, "error");
  }
}

function loadSettings() {
  console.log("settings.ts: loadSettings called");
  const settings = getSyncroSettings();
  const syncroUrl = settings.syncroUrl;
  const syncroApiKey = settings.syncroApiKey;

  if (syncroUrl) {
    (document.getElementById("syncro-url") as HTMLInputElement).value = syncroUrl;
  }
  if (syncroApiKey) {
    (document.getElementById("syncro-api-key") as HTMLInputElement).value = syncroApiKey;
  }
}

// Export functions to be used in other files
export function getSyncroSettings(): { syncroUrl: string; syncroApiKey: string } {
  console.log("settings.ts: getSyncroSettings called");
  let syncroUrl = "";
  let syncroApiKey = "";

  if (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) {
    syncroUrl = (Office.context.roamingSettings.get("syncroUrl") as string) || "";
    syncroApiKey = (Office.context.roamingSettings.get("syncroApiKey") as string) || "";
  } else {
    syncroUrl = localStorage.getItem("syncroUrl") || "";
    syncroApiKey = localStorage.getItem("syncroApiKey") || "";
  }

  console.log("settings.ts: Retrieved settings:", { syncroUrl, syncroApiKey });
  return { syncroUrl, syncroApiKey };
}

export function saveSyncroSettings(syncroUrl: string, syncroApiKey: string): Promise<void> {
  console.log("settings.ts: saveSyncroSettings called", { syncroUrl, syncroApiKey });
  return new Promise((resolve, reject) => {
    if (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) {
      try {
        Office.context.roamingSettings.set("syncroUrl", syncroUrl);
        Office.context.roamingSettings.set("syncroApiKey", syncroApiKey);
        Office.context.roamingSettings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error("Error saving settings in Office environment: " + JSON.stringify(result.error)));
          }
        });
      } catch (error) {
        reject(new Error("Error saving settings in Office environment: " + error.message));
      }
    } else {
      try {
        localStorage.setItem("syncroUrl", syncroUrl);
        localStorage.setItem("syncroApiKey", syncroApiKey);
        resolve();
      } catch (error) {
        reject(new Error("Error saving settings in local storage: " + error.message));
      }
    }
  });
}

function showStatus(message: string, type: "info" | "error" | "success" = "info") {
  const statusElement = document.getElementById("status-message");
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.className = `status-message ${type}`;
    statusElement.style.display = "block";
  }
}
