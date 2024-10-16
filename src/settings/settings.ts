/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, document, console */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("save-settings")!.onclick = saveSettings;
    loadSettings();
  }
});

function saveSettings() {
  const syncroUrl = (document.getElementById("syncro-url") as HTMLInputElement).value;
  const syncroApiKey = (document.getElementById("syncro-api-key") as HTMLInputElement).value;

  Office.context.roamingSettings.set("syncroUrl", syncroUrl);
  Office.context.roamingSettings.set("syncroApiKey", syncroApiKey);
  Office.context.roamingSettings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Settings saved successfully");
      // TODO: Show success message to user
    } else {
      console.error("Error saving settings:", result.error);
      // TODO: Show error message to user
    }
  });
}

function loadSettings() {
  const syncroUrl = Office.context.roamingSettings.get("syncroUrl");
  const syncroApiKey = Office.context.roamingSettings.get("syncroApiKey");

  if (syncroUrl) {
    (document.getElementById("syncro-url") as HTMLInputElement).value = syncroUrl;
  }
  if (syncroApiKey) {
    (document.getElementById("syncro-api-key") as HTMLInputElement).value = syncroApiKey;
  }
}

// Export functions to be used in other files
export function getSyncroSettings(): { syncroUrl: string; syncroApiKey: string } {
  const syncroUrl = Office.context.roamingSettings.get("syncroUrl") as string;
  const syncroApiKey = Office.context.roamingSettings.get("syncroApiKey") as string;
  return { syncroUrl, syncroApiKey };
}
