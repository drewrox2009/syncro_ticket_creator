/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/settings/settings.ts":
/*!**********************************!*\
  !*** ./src/settings/settings.ts ***!
  \**********************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getSyncroSettings: function() { return /* binding */ getSyncroSettings; },
/* harmony export */   saveSyncroSettings: function() { return /* binding */ saveSyncroSettings; }
/* harmony export */ });
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("save-settings").onclick = saveSettings;
    loadSettings();
  }
});
function saveSettings() {
  const syncroUrl = document.getElementById("syncro-url").value;
  const syncroApiKey = document.getElementById("syncro-api-key").value;
  saveSyncroSettings(syncroUrl, syncroApiKey).then(() => {
    console.log("Settings saved successfully");
    // TODO: Show success message to user
  }).catch(error => {
    console.error("Error saving settings:", error);
    // TODO: Show error message to user
  });
}
function loadSettings() {
  const syncroUrl = Office.context.roamingSettings.get("syncroUrl");
  const syncroApiKey = Office.context.roamingSettings.get("syncroApiKey");
  if (syncroUrl) {
    document.getElementById("syncro-url").value = syncroUrl;
  }
  if (syncroApiKey) {
    document.getElementById("syncro-api-key").value = syncroApiKey;
  }
}

// Export functions to be used in other files
function getSyncroSettings() {
  const syncroUrl = Office.context.roamingSettings.get("syncroUrl");
  const syncroApiKey = Office.context.roamingSettings.get("syncroApiKey");
  return {
    syncroUrl,
    syncroApiKey
  };
}
function saveSyncroSettings(syncroUrl, syncroApiKey) {
  return new Promise((resolve, reject) => {
    Office.context.roamingSettings.set("syncroUrl", syncroUrl);
    Office.context.roamingSettings.set("syncroApiKey", syncroApiKey);
    Office.context.roamingSettings.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error);
      }
    });
  });
}

/***/ }),

/***/ "./node_modules/html-loader/dist/runtime/getUrl.js":
/*!*********************************************************!*\
  !*** ./node_modules/html-loader/dist/runtime/getUrl.js ***!
  \*********************************************************/
/***/ (function(module) {



module.exports = function (url, options) {
  if (!options) {
    // eslint-disable-next-line no-param-reassign
    options = {};
  }

  if (!url) {
    return url;
  } // eslint-disable-next-line no-underscore-dangle, no-param-reassign


  url = String(url.__esModule ? url.default : url);

  if (options.hash) {
    // eslint-disable-next-line no-param-reassign
    url += options.hash;
  }

  if (options.maybeNeedQuotes && /[\t\n\f\r "'=<>`]/.test(url)) {
    return "\"".concat(url, "\"");
  }

  return url;
};

/***/ }),

/***/ "./assets/logo-filled.png":
/*!********************************!*\
  !*** ./assets/logo-filled.png ***!
  \********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "assets/logo-filled.png";

/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "fd6ebc4970a52eece689.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	!function() {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = function(module) {
/******/ 			var getter = module && module.__esModule ?
/******/ 				function() { return module['default']; } :
/******/ 				function() { return module; };
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript && document.currentScript.tagName.toUpperCase() === 'SCRIPT')
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = document.baseURI || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other entry modules.
!function() {
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _settings_settings__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../settings/settings */ "./src/settings/settings.ts");
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// Declare fetch as a global for TypeScript



// Define environment-based configuration
const BASE_URL = "https://drewrox2009.github.io/syncro_ticket_creator";
let syncroApiKey;
let syncroUrl;
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("create-ticket").onclick = createTicket;
    loadSyncroSettings();
  }
});
async function loadSyncroSettings() {
  const settings = (0,_settings_settings__WEBPACK_IMPORTED_MODULE_0__.getSyncroSettings)();
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
  document.getElementById("app-body").innerHTML = `
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
  document.getElementById("save-settings").onclick = saveSettings;
}
async function saveSettings() {
  const newSyncroUrl = document.getElementById("syncro-url").value;
  const newSyncroApiKey = document.getElementById("syncro-api-key").value;
  if (!newSyncroUrl || !newSyncroApiKey) {
    showStatus("Please enter both Syncro URL and API Key.", "error");
    return;
  }
  try {
    await (0,_settings_settings__WEBPACK_IMPORTED_MODULE_0__.saveSyncroSettings)(newSyncroUrl, newSyncroApiKey);
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
  Office.context.ui.displayDialogAsync(`${BASE_URL}/settings/settings.html`, {
    height: 60,
    width: 30,
    displayInIframe: true
  }, result => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to open settings dialog:", result.error);
      showStatus("Failed to open settings. Please try again.", "error");
    }
  });
}
async function populateCustomers() {
  try {
    const customers = await fetchSyncroCustomers();
    const customerSelect = document.getElementById("customer-select");
    customerSelect.innerHTML = '<option value="">Select a customer</option>';
    customers.forEach(customer => {
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
    const customerId = document.getElementById("customer-select").value;
    const contacts = await fetchSyncroContacts(parseInt(customerId));
    const contactSelect = document.getElementById("contact-select");
    contactSelect.innerHTML = '<option value="">Select a contact</option>';
    contacts.forEach(contact => {
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
    const customerId = document.getElementById("customer-select").value;
    const assets = await fetchSyncroAssets(parseInt(customerId));
    const assetSelect = document.getElementById("asset-select");
    assetSelect.innerHTML = '<option value="">Select an asset</option>';
    assets.forEach(asset => {
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
async function createTicket() {
  try {
    showStatus("Creating ticket...");
    const emailInfo = await getEmailInfo();
    const customerId = document.getElementById("customer-select").value;
    const contactId = document.getElementById("contact-select").value;
    const assetId = document.getElementById("asset-select").value;
    const ticketTitle = document.getElementById("ticket-title").value || emailInfo.subject;
    const ticketMessage = document.getElementById("ticket-message").value || emailInfo.content;
    const ticketData = {
      customer_id: parseInt(customerId),
      contact_id: parseInt(contactId),
      asset_id: assetId ? parseInt(assetId) : null,
      subject: ticketTitle,
      problem_type: "Other",
      status: "New",
      comment: ticketMessage
    };
    const createdTicket = await createSyncroTicket(ticketData);
    console.log("Ticket created:", createdTicket);
    showStatus("Ticket created successfully!", "success");
  } catch (error) {
    console.error("Error creating ticket:", error);
    showStatus("Failed to create ticket. Please try again.", "error");
  }
}
async function getEmailInfo() {
  return new Promise(resolve => {
    const item = Office.context.mailbox.item;
    const emailInfo = {
      subject: item?.subject || "",
      content: "",
      senderEmail: item?.from?.emailAddress || "",
      senderName: item?.from?.displayName || ""
    };
    item?.body.getAsync(Office.CoercionType.Text, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        emailInfo.content = result.value;
      }
      resolve(emailInfo);
    });
  });
}

// Syncro API functions
async function fetchSyncroCustomers() {
  const response = await fetch(`${syncroUrl}/api/v1/customers`, {
    headers: {
      Authorization: `Bearer ${syncroApiKey}`
    }
  });
  if (!response.ok) {
    throw new Error("Failed to fetch customers");
  }
  return response.json();
}
async function fetchSyncroContacts(customerId) {
  const response = await fetch(`${syncroUrl}/api/v1/customers/${customerId}/contacts`, {
    headers: {
      Authorization: `Bearer ${syncroApiKey}`
    }
  });
  if (!response.ok) {
    throw new Error("Failed to fetch contacts");
  }
  return response.json();
}
async function fetchSyncroAssets(customerId) {
  const response = await fetch(`${syncroUrl}/api/v1/customers/${customerId}/assets`, {
    headers: {
      Authorization: `Bearer ${syncroApiKey}`
    }
  });
  if (!response.ok) {
    throw new Error("Failed to fetch assets");
  }
  return response.json();
}
async function createSyncroTicket(ticketData) {
  const response = await fetch(`${syncroUrl}/api/v1/tickets`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${syncroApiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(ticketData)
  });
  if (!response.ok) {
    throw new Error("Failed to create ticket");
  }
  return response.json();
}
function showStatus(message, type = "info") {
  const statusElement = document.getElementById("status-message");
  if (!statusElement) {
    const newStatusElement = document.createElement("div");
    newStatusElement.id = "status-message";
    document.body.insertBefore(newStatusElement, document.body.firstChild);
  }
  const element = statusElement || document.getElementById("status-message");
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
function sanitizeHtml(input) {
  const div = document.createElement("div");
  div.textContent = input;
  return div.innerHTML;
}
}();
// This entry need to be wrapped in an IIFE because it need to be isolated against other entry modules.
!function() {
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../node_modules/html-loader/dist/runtime/getUrl.js */ "./node_modules/html-loader/dist/runtime/getUrl.js");
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__);
// Imports

var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_1___ = new URL(/* asset import */ __webpack_require__(/*! ../../assets/logo-filled.png */ "./assets/logo-filled.png"), __webpack_require__.b);
// Module
var ___HTML_LOADER_REPLACEMENT_0___ = _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default()(___HTML_LOADER_IMPORT_0___);
var ___HTML_LOADER_REPLACEMENT_1___ = _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default()(___HTML_LOADER_IMPORT_1___);
var code = "<!DOCTYPE html>\n<html>\n  <head>\n    <meta charset=\"UTF-8\" />\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />\n    <title>Syncro Ticket Creator</title>\n\n    <!-- Office JavaScript API -->\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js\"><" + "/script>\n\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\n    <link\n      rel=\"stylesheet\"\n      href=\"https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css\"\n    />\n\n    <!-- Template styles -->\n    <link href=\"" + ___HTML_LOADER_REPLACEMENT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\n  </head>\n\n  <body class=\"ms-font-m ms-Fabric\">\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">\n      <img width=\"90\" height=\"90\" src=\"" + ___HTML_LOADER_REPLACEMENT_1___ + "\" alt=\"Syncro\" title=\"Syncro\" />\n      <h1 class=\"ms-font-su\">Syncro Ticket Creator</h1>\n    </header>\n    <main id=\"app-body\" class=\"ms-welcome__main\">\n      <form id=\"ticket-form\">\n        <div class=\"ms-TextField\">\n          <label class=\"ms-Label\" for=\"customer-select\">Syncro Customer</label>\n          <select id=\"customer-select\" class=\"ms-Dropdown\" required aria-label=\"Select Syncro Customer\">\n            <option value=\"\">Select a customer</option>\n          </select>\n        </div>\n        <div class=\"ms-TextField\">\n          <label class=\"ms-Label\" for=\"contact-select\">Contact</label>\n          <select id=\"contact-select\" class=\"ms-Dropdown\" required aria-label=\"Select Contact\">\n            <option value=\"\">Select a contact</option>\n          </select>\n        </div>\n        <div class=\"ms-TextField\">\n          <label class=\"ms-Label\" for=\"asset-select\">Asset (Optional)</label>\n          <select id=\"asset-select\" class=\"ms-Dropdown\" aria-label=\"Select Asset\">\n            <option value=\"\">Select an asset</option>\n          </select>\n        </div>\n        <div class=\"ms-TextField\">\n          <label class=\"ms-Label\" for=\"ticket-title\">Ticket Title</label>\n          <input type=\"text\" id=\"ticket-title\" class=\"ms-TextField-field\" required />\n        </div>\n        <div class=\"ms-TextField\">\n          <label class=\"ms-Label\" for=\"ticket-message\">Ticket Message</label>\n          <textarea id=\"ticket-message\" class=\"ms-TextField-field\" rows=\"5\" required></textarea>\n        </div>\n        <div class=\"ms-TextField\">\n          <button id=\"create-ticket\" class=\"ms-Button ms-Button--primary\">\n            <span class=\"ms-Button-label\">Create Ticket</span>\n          </button>\n        </div>\n      </form>\n    </main>\n  </body>\n</html>\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map