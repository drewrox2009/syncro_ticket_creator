/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// Add event listener for DOMContentLoaded
document.addEventListener("DOMContentLoaded", () => {
  console.log("settings.ts: DOMContentLoaded event fired");
  if (typeof Office !== "undefined" && Office.context && Office.context.host === Office.HostType.Outlook) {
    loadSettings();
    attachSaveSettingsListener();
  } else {
    attachSaveSettingsListener();
  }
});

function attachSaveSettingsListener() {
  const saveSettingsButton = document.getElementById("save-settings");
  if (saveSettingsButton) {
    saveSettingsButton.addEventListener("click", saveSettings);
    console.log("settings.ts: Event listener attached to save-settings button");
  } else {
    console.error("settings.ts: Element with id 'save-settings' not found");
  }
}

function saveSettings() {
  console.log("settings.ts: saveSettings called");
  const syncroUrl = (document.getElementById("syncro-url") as HTMLInputElement).value;
  const syncroApiKey = (document.getElementById("syncro-api-key") as HTMLInputElement).value;

  saveSyncroSettings(syncroUrl, syncroApiKey)
    .then(() => {
      console.log("settings.ts: Settings saved successfully");
      // TODO: Show success message to user
    })
    .catch((error) => {
      console.error("settings.ts: Error saving settings:", error);
      // TODO: Show error message to user
    });
}

function loadSettings() {
  console.log("settings.ts: loadSettings called");
  const syncroUrl = getSyncroSettings().syncroUrl;
  const syncroApiKey = getSyncroSettings().syncroApiKey;

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
      Office.context.roamingSettings.set("syncroUrl", syncroUrl);
      Office.context.roamingSettings.set("syncroApiKey", syncroApiKey);
      Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error);
        }
      });
    } else {
      localStorage.setItem("syncroUrl", syncroUrl);
      localStorage.setItem("syncroApiKey", syncroApiKey);
      resolve();
    }
  });
}
