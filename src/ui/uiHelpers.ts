/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

export function showStatus(
  message: string,
  type: "info" | "error" | "success" = "info",
  elementId: string = "status-message"
) {
  console.log(`UI Helpers: showStatus called - ${type}: ${message}`);
  const statusElement = document.getElementById(elementId);
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.className = `status-message ${type}`;
    statusElement.style.display = "block";
  }
}

export function hideStatus(elementId: string = "status-message") {
  console.log("UI Helpers: hideStatus called");
  const statusElement = document.getElementById(elementId);
  if (statusElement) {
    statusElement.style.display = "none";
  }
}

export function sanitizeHtml(input: string): string {
  const div = document.createElement("div");
  div.textContent = input;
  return div.innerHTML;
}
