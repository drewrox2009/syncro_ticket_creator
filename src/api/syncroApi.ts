/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Declare fetch as a global for TypeScript
declare const fetch: (input: RequestInfo, init?: RequestInit) => Promise<Response>;

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

export async function fetchSyncroCustomers(syncroUrl: string, syncroApiKey: string): Promise<SyncroCustomer[]> {
  console.log("Syncro API: fetchSyncroCustomers called");
  const response = await fetch(`${syncroUrl}/api/v1/customers`, {
    headers: {
      Authorization: `Bearer ${syncroApiKey}`,
    },
  });
  if (!response.ok) {
    throw new Error(`Failed to fetch customers: ${response.status} ${response.statusText}`);
  }
  return response.json();
}

export async function fetchSyncroContacts(
  syncroUrl: string,
  syncroApiKey: string,
  customerId: number
): Promise<SyncroContact[]> {
  console.log("Syncro API: fetchSyncroContacts called");
  const response = await fetch(`${syncroUrl}/api/v1/customers/${customerId}/contacts`, {
    headers: {
      Authorization: `Bearer ${syncroApiKey}`,
    },
  });
  if (!response.ok) {
    throw new Error(`Failed to fetch contacts: ${response.status} ${response.statusText}`);
  }
  return response.json();
}

export async function fetchSyncroAssets(
  syncroUrl: string,
  syncroApiKey: string,
  customerId: number
): Promise<SyncroAsset[]> {
  console.log("Syncro API: fetchSyncroAssets called");
  const response = await fetch(`${syncroUrl}/api/v1/customers/${customerId}/assets`, {
    headers: {
      Authorization: `Bearer ${syncroApiKey}`,
    },
  });
  if (!response.ok) {
    throw new Error(`Failed to fetch assets: ${response.status} ${response.statusText}`);
  }
  return response.json();
}

export async function createSyncroTicket(syncroUrl: string, syncroApiKey: string, ticketData: any): Promise<any> {
  console.log("Syncro API: createSyncroTicket called");
  const response = await fetch(`${syncroUrl}/api/v1/tickets`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${syncroApiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(ticketData),
  });
  if (!response.ok) {
    throw new Error(`Failed to create ticket: ${response.status} ${response.statusText}`);
  }
  return response.json();
}
