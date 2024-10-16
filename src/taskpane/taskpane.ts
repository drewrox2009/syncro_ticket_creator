/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, document, console, window */

// Declare fetch as a global for TypeScript
declare const fetch: (input: RequestInfo, init?: RequestInit) => Promise<Response>;

import { getSyncroSettings } from "../settings/settings";

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
  syncroApiKey = settings.syncroApiKey;
  syncroUrl = settings.syncroUrl;

  if (!syncroApiKey || !syncroUrl) {
    // Prompt user to set up settings
    document.getElementById("app-body")!.innerHTML = `
      <p>Please set up your Syncro settings before using this add-in.</p>
      <button id="open-settings" class="ms-Button ms-Button--primary">
        <span class="ms-Button-label">Open Settings</span>
      </button>
    `;
    document.getElementById("open-settings")!.onclick = openSettings;
  } else {
    // Settings are available, initialize the app
    showStatus("Loading customers...");
    await populateCustomers();
    hideStatus();
  }
}

function openSettings() {
  Office.context.ui.displayDialogAsync(
    `${window.location.origin}/settings/settings.html`,
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
      option.textContent = customer.name;
      customerSelect.appendChild(option);
    });
    customerSelect.onchange = populateContacts;
  } catch (error) {
    console.error("Error populating customers:", error);
    showStatus("Failed to load customers. Please check your Syncro settings.", "error");
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
      option.textContent = contact.name;
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
      option.textContent = asset.name;
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
  element.textContent = message;
  element.className = `status-message ${type}`;
  element.style.display = "block";
}

function hideStatus() {
  const statusElement = document.getElementById("status-message");
  if (statusElement) {
    statusElement.style.display = "none";
  }
}
