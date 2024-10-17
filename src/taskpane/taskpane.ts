/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// Declare fetch as a global for TypeScript
declare const fetch: (input: RequestInfo, init?: RequestInit) => Promise<Response>;

import { getSyncroSettings, saveSyncroSettings } from "../settings/settings";

// Define environment-based configuration
const BASE_URL = "https://drewrox2009.github.io/syncro_ticket_creator";

interface EmailInfo {
  subject: string;
  content: string;
  senderEmail: string;
  senderName: string;
}

interface SyncroCustomer {
  id: number;
  name: string;
}

interface SyncroContact {
  id: number;
  name: string;
}

interface SyncroAsset {
  id: number;
  name: string;
}

let syncroApiKey: string;
let syncroUrl: string;

Office.onReady((info: { host: Office.HostType; platform: Office.PlatformType }) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("create-ticket")!.onclick = createTicket;
    loadSyncroSettings();
  }
});

async function loadSyncroSettings() {
  const settings = getSyncroSettings();
  console.log("Retrieved settings:", settings); // Add this line for debugging
  syncroApiKey = settings.syncroApiKey;
  syncroUrl = settings.syncroUrl;

  if (!syncroApiKey || !syncroUrl) {
    showSettingsUI();
  } else {
    try {
      showStatus("Verifying API settings...");
      await verifyApiSettings();
      showStatus("Loading customers...");
      await populateCustomers();
      hideStatus();
    } catch (error) {
      console.error("Error initializing app:", error);
      showStatus("Failed to initialize app. Please check your Syncro settings.", "error");
      showSettingsUI();
    }
  }
}

function showSettingsUI() {
  document.getElementById("app-body")!.innerHTML = `
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
    </div>
  `;
  document.getElementById("save-settings")!.onclick = saveSettings;
}

async function saveSettings() {
  const newSyncroUrl = (document.getElementById("syncro-url") as HTMLInputElement).value;
  const newSyncroApiKey = (document.getElementById("syncro-api-key") as HTMLInputElement).value;

  if (!newSyncroUrl || !newSyncroApiKey) {
    showStatus("Please enter both Syncro URL and API Key.", "error");
    return;
  }

  try {
    await saveSyncroSettings(newSyncroUrl, newSyncroApiKey);
    syncroUrl = newSyncroUrl;
    syncroApiKey = newSyncroApiKey;
    showStatus("Settings saved successfully. Initializing app...", "success");
    await verifyApiSettings();
    await populateCustomers();
    hideStatus();
  } catch (error) {
    console.error("Error saving settings:", error);
    showStatus("Failed to save settings. Please try again.", "error");
  }
}

async function verifyApiSettings() {
  try {
    await fetchSyncroCustomers();
  } catch (error) {
    console.error("Error verifying API settings:", error);
    throw new Error("Invalid API settings. Please check your Syncro URL and API Key.");
  }
}

function openSettings() {
  Office.context.ui.displayDialogAsync(
    `${BASE_URL}/settings/settings.html`,
    { height: 60, width: 30, displayInIframe: true },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to open settings dialog:", result.error);
        showStatus("Failed to open settings. Please try again.", "error");
      }
    }
  );
}

async function populateCustomers() {
  try {
    const customers = await fetchSyncroCustomers();
    const customerSelect = document.getElementById("customer-select") as HTMLSelectElement;
    customerSelect.innerHTML = '<option value="">Select a customer</option>';
    customers.forEach((customer) => {
      const option = document.createElement("option");
      option.value = customer.id.toString();
      option.textContent = sanitizeHtml(customer.name);
      customerSelect.appendChild(option);
    });
    customerSelect.onchange = populateContacts;
  } catch (error) {
    console.error("Error populating customers:", error);
    showStatus("Failed to load customers. Please check your Syncro settings.", "error");
    throw error; // Rethrow the error to be caught by the caller
  }
}

async function populateContacts() {
  try {
    showStatus("Loading contacts...");
    const customerId = (document.getElementById("customer-select") as HTMLSelectElement).value;
    const contacts = await fetchSyncroContacts(parseInt(customerId));
    const contactSelect = document.getElementById("contact-select") as HTMLSelectElement;
    contactSelect.innerHTML = '<option value="">Select a contact</option>';
    contacts.forEach((contact) => {
      const option = document.createElement("option");
      option.value = contact.id.toString();
      option.textContent = sanitizeHtml(contact.name);
      contactSelect.appendChild(option);
    });
    hideStatus();
    populateAssets();
  } catch (error) {
    console.error("Error populating contacts:", error);
    showStatus("Failed to load contacts. Please try again.", "error");
  }
}

async function populateAssets() {
  try {
    showStatus("Loading assets...");
    const customerId = (document.getElementById("customer-select") as HTMLSelectElement).value;
    const assets = await fetchSyncroAssets(parseInt(customerId));
    const assetSelect = document.getElementById("asset-select") as HTMLSelectElement;
    assetSelect.innerHTML = '<option value="">Select an asset</option>';
    assets.forEach((asset) => {
      const option = document.createElement("option");
      option.value = asset.id.toString();
      option.textContent = sanitizeHtml(asset.name);
      assetSelect.appendChild(option);
    });
    hideStatus();
  } catch (error) {
    console.error("Error populating assets:", error);
    showStatus("Failed to load assets. Please try again.", "error");
  }
}

async function createTicket(): Promise<void> {
  try {
    showStatus("Creating ticket...");
    const emailInfo = await getEmailInfo();
    const customerId = (document.getElementById("customer-select") as HTMLSelectElement).value;
    const contactId = (document.getElementById("contact-select") as HTMLSelectElement).value;
    const assetId = (document.getElementById("asset-select") as HTMLSelectElement).value;
    const ticketTitle = (document.getElementById("ticket-title") as HTMLInputElement).value || emailInfo.subject;
    const ticketMessage = (document.getElementById("ticket-message") as HTMLTextAreaElement).value || emailInfo.content;

    const ticketData = {
      customer_id: parseInt(customerId),
      contact_id: parseInt(contactId),
      asset_id: assetId ? parseInt(assetId) : null,
      subject: ticketTitle,
      problem_type: "Other",
      status: "New",
      comment: ticketMessage,
    };

    const createdTicket = await createSyncroTicket(ticketData);
    console.log("Ticket created:", createdTicket);
    showStatus("Ticket created successfully!", "success");
  } catch (error) {
    console.error("Error creating ticket:", error);
    showStatus("Failed to create ticket. Please try again.", "error");
  }
}

async function getEmailInfo(): Promise<EmailInfo> {
  return new Promise((resolve) => {
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
  });
}

// Syncro API functions
async function fetchSyncroCustomers(): Promise<SyncroCustomer[]> {
  const response = await fetch(`${syncroUrl}/api/v1/customers`, {
    headers: {
      Authorization: `Bearer ${syncroApiKey}`,
    },
  });
  if (!response.ok) {
    throw new Error("Failed to fetch customers");
  }
  return response.json();
}

async function fetchSyncroContacts(customerId: number): Promise<SyncroContact[]> {
  const response = await fetch(`${syncroUrl}/api/v1/customers/${customerId}/contacts`, {
    headers: {
      Authorization: `Bearer ${syncroApiKey}`,
    },
  });
  if (!response.ok) {
    throw new Error("Failed to fetch contacts");
  }
  return response.json();
}

async function fetchSyncroAssets(customerId: number): Promise<SyncroAsset[]> {
  const response = await fetch(`${syncroUrl}/api/v1/customers/${customerId}/assets`, {
    headers: {
      Authorization: `Bearer ${syncroApiKey}`,
    },
  });
  if (!response.ok) {
    throw new Error("Failed to fetch assets");
  }
  return response.json();
}

async function createSyncroTicket(ticketData: any): Promise<any> {
  const response = await fetch(`${syncroUrl}/api/v1/tickets`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${syncroApiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(ticketData),
  });
  if (!response.ok) {
    throw new Error("Failed to create ticket");
  }
  return response.json();
}

function showStatus(message: string, type: "info" | "error" | "success" = "info") {
  const statusElement = document.getElementById("status-message");
  if (!statusElement) {
    const newStatusElement = document.createElement("div");
    newStatusElement.id = "status-message";
    document.body.insertBefore(newStatusElement, document.body.firstChild);
  }
  const element = statusElement || document.getElementById("status-message")!;
  element.textContent = sanitizeHtml(message);
  element.className = `status-message ${type}`;
  element.style.display = "block";
}

function hideStatus() {
  const statusElement = document.getElementById("status-message");
  if (statusElement) {
    statusElement.style.display = "none";
  }
}

// Simple HTML sanitization function
function sanitizeHtml(input: string): string {
  const div = document.createElement("div");
  div.textContent = input;
  return div.innerHTML;
}
