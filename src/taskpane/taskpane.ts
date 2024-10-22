/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import { getSyncroSettings, saveSyncroSettings } from "../settings/settings";
import * as syncroApi from "../api/syncroApi";
import * as uiHelpers from "../ui/uiHelpers";

interface EmailInfo {
  subject: string;
  content: string;
  senderEmail: string;
  senderName: string;
}

// Declare syncroApiKey and syncroUrl at the top
let syncroApiKey: string;
let syncroUrl: string;

// Function to initialize the app
function initializeApp() {
  console.log("Syncro Ticket Creator: initializeApp called");
  loadSyncroSettings();
  document.getElementById("create-ticket")!.onclick = createTicket;
}

if (typeof Office !== "undefined") {
  console.log("Syncro Ticket Creator: Office environment detected");
  Office.onReady((info: { host: Office.HostType; platform: Office.PlatformType }) => {
    console.log("Syncro Ticket Creator: Office.onReady called", info);
    if (info.host === Office.HostType.Outlook) {
      document.addEventListener("DOMContentLoaded", initializeApp); // Ensure DOM is fully loaded
    }
  });
} else {
  console.log("Syncro Ticket Creator: Non-Office environment detected");
  document.addEventListener("DOMContentLoaded", () => {
    console.log("Syncro Ticket Creator: DOMContentLoaded event fired");
    initializeApp();
  });
}

async function loadSyncroSettings() {
  console.log("Syncro Ticket Creator: loadSyncroSettings called");
  const settings = getSyncroSettings();
  console.log("Retrieved settings:", settings);
  syncroApiKey = settings.syncroApiKey;
  syncroUrl = settings.syncroUrl;

  showSettingsUI();
}

function showSettingsUI() {
  console.log("Syncro Ticket Creator: showSettingsUI called");
  const appBody = document.getElementById("app-body");
  if (appBody) {
    console.log("Syncro Ticket Creator: app-body element found");
    appBody.innerHTML = `
      <h2>Syncro API Settings</h2>
      <div class="ms-TextField">
        <label class="ms-Label" for="syncro-url">Syncro URL</label>
        <input type="text" id="syncro-url" class="ms-TextField-field" value="${syncroUrl || ""}" required />
      </div>
      <div class="ms-TextField">
        <label class="ms-Label" for="syncro-api-key">Syncro API Key</label>
        <input type="password" id="syncro-api-key" class="ms-TextField-field" value="${syncroApiKey || ""}" required />
      </div>
      <div class="ms-TextField">
        <button id="save-settings" class="ms-Button ms-Button--primary">
          <span class="ms-Button-label">Save Settings</span>
        </button>
        <button id="test-api-settings" class="ms-Button ms-Button--secondary">
          <span class="ms-Button-label">Test API Settings</span>
        </button>
      </div>
    `;
    document.getElementById("save-settings")!.onclick = saveSettings;
    document.getElementById("test-api-settings")!.onclick = testApiSettings;
    console.log("Syncro Ticket Creator: Settings UI rendered");
  } else {
    console.error("Syncro Ticket Creator: Element with id 'app-body' not found");
  }
}

async function saveSettings() {
  console.log("Syncro Ticket Creator: saveSettings called");
  const newSyncroUrl = (document.getElementById("syncro-url") as HTMLInputElement).value;
  const newSyncroApiKey = (document.getElementById("syncro-api-key") as HTMLInputElement).value;

  if (!newSyncroUrl || !newSyncroApiKey) {
    uiHelpers.showStatus("Please enter both Syncro URL and API Key.", "error");
    return;
  }

  try {
    await saveSyncroSettings(newSyncroUrl, newSyncroApiKey);
    syncroUrl = newSyncroUrl;
    syncroApiKey = newSyncroApiKey;
    uiHelpers.showStatus("Settings saved successfully. Initializing app...", "success");
    await verifyApiSettings();
    await populateCustomers();
    uiHelpers.hideStatus();
  } catch (error) {
    console.error("Error saving settings:", error);
    uiHelpers.showStatus("Failed to save settings. Please try again.", "error");
  }
}

// New function to test API settings
async function testApiSettings() {
  console.log("Syncro Ticket Creator: testApiSettings called");
  const testSyncroUrl = (document.getElementById("syncro-url") as HTMLInputElement).value;
  const testSyncroApiKey = (document.getElementById("syncro-api-key") as HTMLInputElement).value;

  if (!testSyncroUrl || !testSyncroApiKey) {
    uiHelpers.showStatus("Please enter both Syncro URL and API Key.", "error");
    return;
  }

  try {
    const response = await syncroApi.fetchSyncroCustomers(testSyncroUrl, testSyncroApiKey);
    console.log("Syncro Ticket Creator: API settings verified successfully", response);
    uiHelpers.showStatus("API settings verified successfully!", "success");
  } catch (error) {
    console.error("Error verifying API settings:", error);
    uiHelpers.showStatus("Failed to verify API settings. Please try again.", "error");
  }
}

async function verifyApiSettings() {
  console.log("Syncro Ticket Creator: verifyApiSettings called");
  try {
    await syncroApi.fetchSyncroCustomers(syncroUrl, syncroApiKey);
  } catch (error) {
    console.error("Error verifying API settings:", error);
    throw new Error("Invalid API settings. Please check your Syncro URL and API Key.");
  }
}

async function populateCustomers() {
  console.log("Syncro Ticket Creator: populateCustomers called");
  try {
    uiHelpers.showStatus("Loading customers...", "info");
    const customers = await syncroApi.fetchSyncroCustomers(syncroUrl, syncroApiKey);
    const customerSelect = document.getElementById("customer-select") as HTMLSelectElement;
    if (!customerSelect) {
      console.error("Syncro Ticket Creator: Element with id 'customer-select' not found");
      return;
    }
    customerSelect.innerHTML = '<option value="">Select a customer</option>';
    customers.forEach((customer) => {
      const option = document.createElement("option");
      option.value = customer.id.toString();
      option.textContent = uiHelpers.sanitizeHtml(customer.name);
      customerSelect.appendChild(option);
    });
    customerSelect.onchange = populateContacts;
    uiHelpers.hideStatus();
  } catch (error) {
    console.error("Error populating customers:", error);
    uiHelpers.showStatus("Failed to load customers. Please check your Syncro settings.", "error");
    throw error;
  }
}

async function populateContacts() {
  console.log("Syncro Ticket Creator: populateContacts called");
  try {
    uiHelpers.showStatus("Loading contacts...", "info");
    const customerSelect = document.getElementById("customer-select") as HTMLSelectElement;
    const customerId = customerSelect.value;
    const contacts = await syncroApi.fetchSyncroContacts(syncroUrl, syncroApiKey, parseInt(customerId));
    const contactSelect = document.getElementById("contact-select") as HTMLSelectElement;
    if (!contactSelect) {
      console.error("Syncro Ticket Creator: Element with id 'contact-select' not found");
      return;
    }
    contactSelect.innerHTML = '<option value="">Select a contact</option>';
    contacts.forEach((contact) => {
      const option = document.createElement("option");
      option.value = contact.id.toString();
      option.textContent = uiHelpers.sanitizeHtml(contact.name);
      contactSelect.appendChild(option);
    });
    uiHelpers.hideStatus();
    populateAssets();
  } catch (error) {
    console.error("Error populating contacts:", error);
    uiHelpers.showStatus("Failed to load contacts. Please try again.", "error");
  }
}

async function populateAssets() {
  console.log("Syncro Ticket Creator: populateAssets called");
  try {
    uiHelpers.showStatus("Loading assets...", "info");
    const customerSelect = document.getElementById("customer-select") as HTMLSelectElement;
    const customerId = customerSelect.value;
    const assets = await syncroApi.fetchSyncroAssets(syncroUrl, syncroApiKey, parseInt(customerId));
    const assetSelect = document.getElementById("asset-select") as HTMLSelectElement;
    if (!assetSelect) {
      console.error("Syncro Ticket Creator: Element with id 'asset-select' not found");
      return;
    }
    assetSelect.innerHTML = '<option value="">Select an asset</option>';
    assets.forEach((asset) => {
      const option = document.createElement("option");
      option.value = asset.id.toString();
      option.textContent = uiHelpers.sanitizeHtml(asset.name);
      assetSelect.appendChild(option);
    });
    uiHelpers.hideStatus();
  } catch (error) {
    console.error("Error populating assets:", error);
    uiHelpers.showStatus("Failed to load assets. Please try again.", "error");
  }
}

async function createTicket(): Promise<void> {
  console.log("Syncro Ticket Creator: createTicket called");
  try {
    uiHelpers.showStatus("Creating ticket...", "info");
    const emailInfo = await getEmailInfo();
    const customerId = (document.getElementById("customer-select") as HTMLSelectElement).value;
    const contactId = (document.getElementById("contact-select") as HTMLSelectElement).value;
    const assetId = (document.getElementById("asset-select") as HTMLSelectElement).value;
    const ticketTitle = (document.getElementById("ticket-title") as HTMLInputElement).value || emailInfo.subject;
    const ticketMessage = (document.getElementById("ticket-message") as HTMLTextAreaElement).value || emailInfo.content;

    if (!customerId || !contactId || !ticketTitle || !ticketMessage) {
      uiHelpers.showStatus("Please fill in all required fields.", "error");
      return;
    }

    const ticketData = {
      customer_id: parseInt(customerId),
      contact_id: parseInt(contactId),
      asset_id: assetId ? parseInt(assetId) : null,
      subject: ticketTitle,
      problem_type: "Other",
      status: "New",
      comment: ticketMessage,
    };

    const createdTicket = await syncroApi.createSyncroTicket(syncroUrl, syncroApiKey, ticketData);
    console.log("Ticket created:", createdTicket);
    uiHelpers.showStatus("Ticket created successfully!", "success");
  } catch (error) {
    console.error("Error creating ticket:", error);
    uiHelpers.showStatus("Failed to create ticket. Please try again.", "error");
  } finally {
    uiHelpers.hideStatus();
  }
}

async function getEmailInfo(): Promise<EmailInfo> {
  console.log("Syncro Ticket Creator: getEmailInfo called");
  return new Promise((resolve) => {
    if (typeof Office !== "undefined" && Office.context && Office.context.mailbox) {
      const item = Office.context.mailbox.item;
      const emailInfo: EmailInfo = {
        subject: item?.subject || "",
        content: "",
        senderEmail: item?.from?.emailAddress || "",
        senderName: item?.from?.displayName || "",
      };

      item?.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          emailInfo.content = result.value;
        }
        resolve(emailInfo);
      });
    } else {
      resolve({
        subject: "",
        content: "",
        senderEmail: "",
        senderName: "",
      });
    }
  });
}
