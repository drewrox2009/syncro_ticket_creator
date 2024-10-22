!function(){"use strict";var e,t,n,o,s={14385:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},98362:function(e,t,n){e.exports=n.p+"assets/logo-filled.png"},58394:function(e,t,n){e.exports=n.p+"fd6ebc4970a52eece689.css"}},c={};function r(e){var t=c[e];if(void 0!==t)return t.exports;var n=c[e]={exports:{}};return s[e](n,n.exports,r),n.exports}r.m=s,r.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return r.d(t,{a:t}),t},r.d=function(e,t){for(var n in t)r.o(t,n)&&!r.o(e,n)&&Object.defineProperty(e,n,{enumerable:!0,get:t[n]})},r.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),r.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;r.g.importScripts&&(e=r.g.location+"");var t=r.g.document;if(!e&&t&&(t.currentScript&&"SCRIPT"===t.currentScript.tagName.toUpperCase()&&(e=t.currentScript.src),!e)){var n=t.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!e||!/^http(s?):/.test(e));)e=n[o--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),r.p=e}(),r.b=document.baseURI||self.location.href,function(){const e=new MutationObserver((n=>{n.forEach((n=>{"childList"===n.type&&n.addedNodes.length>0&&(function(){const e=document.getElementById("save-settings");e?(e.addEventListener("click",t),console.log("settings.ts: Event listener attached to save-settings button")):console.error("settings.ts: Element with id 'save-settings' not found")}(),e.disconnect())}))}));async function t(){console.log("settings.ts: saveSettings called");const e=document.getElementById("syncro-url").value,t=document.getElementById("syncro-api-key").value;if(document.getElementById("error-message"),e&&t)try{await o(e,t),console.log("settings.ts: Settings saved successfully"),s("Settings saved successfully!","success"),window.close()}catch(e){console.error("settings.ts: Error saving settings:",e),s("Error saving settings: "+e.message,"error")}else s("Please enter both Syncro URL and API Key.","error")}function n(){console.log("settings.ts: getSyncroSettings called");let e="",t="";return"undefined"!=typeof Office&&Office.context&&Office.context.roamingSettings?(e=Office.context.roamingSettings.get("syncroUrl")||"",t=Office.context.roamingSettings.get("syncroApiKey")||""):(e=localStorage.getItem("syncroUrl")||"",t=localStorage.getItem("syncroApiKey")||""),console.log("settings.ts: Retrieved settings:",{syncroUrl:e,syncroApiKey:t}),{syncroUrl:e,syncroApiKey:t}}function o(e,t){return console.log("settings.ts: saveSyncroSettings called",{syncroUrl:e,syncroApiKey:t}),new Promise(((n,o)=>{if("undefined"!=typeof Office&&Office.context&&Office.context.roamingSettings)try{Office.context.roamingSettings.set("syncroUrl",e),Office.context.roamingSettings.set("syncroApiKey",t),Office.context.roamingSettings.saveAsync((e=>{e.status===Office.AsyncResultStatus.Succeeded?n():o(new Error("Error saving settings in Office environment: "+JSON.stringify(e.error)))}))}catch(e){o(new Error("Error saving settings in Office environment: "+e.message))}else try{localStorage.setItem("syncroUrl",e),localStorage.setItem("syncroApiKey",t),n()}catch(e){o(new Error("Error saving settings in local storage: "+e.message))}}))}function s(e,t="info"){const n=document.getElementById("status-message");n&&(n.textContent=e,n.className=`status-message ${t}`,n.style.display="block")}async function c(e,t){console.log("Syncro API: fetchSyncroCustomers called");const n=await fetch(`${e}/api/v1/customers`,{headers:{Authorization:`Bearer ${t}`}});if(!n.ok)throw new Error(`Failed to fetch customers: ${n.status} ${n.statusText}`);return n.json()}function r(e,t="info",n="status-message"){console.log(`UI Helpers: showStatus called - ${t}: ${e}`);const o=document.getElementById(n);o&&(o.textContent=e,o.className=`status-message ${t}`,o.style.display="block")}function i(e="status-message"){console.log("UI Helpers: hideStatus called");const t=document.getElementById(e);t&&(t.style.display="none")}function a(e){const t=document.createElement("div");return t.textContent=e,t.innerHTML}let l,d;function u(){console.log("Syncro Ticket Creator: initializeApp called"),async function(){console.log("Syncro Ticket Creator: loadSyncroSettings called");const e=n();console.log("Retrieved settings:",e),l=e.syncroApiKey,d=e.syncroUrl,function(){console.log("Syncro Ticket Creator: showSettingsUI called");const e=document.getElementById("app-body");e?(console.log("Syncro Ticket Creator: app-body element found"),e.innerHTML=`\n      <h2>Syncro API Settings</h2>\n      <div class="ms-TextField">\n        <label class="ms-Label" for="syncro-url">Syncro URL</label>\n        <input type="text" id="syncro-url" class="ms-TextField-field" value="${d||""}" required />\n      </div>\n      <div class="ms-TextField">\n        <label class="ms-Label" for="syncro-api-key">Syncro API Key</label>\n        <input type="password" id="syncro-api-key" class="ms-TextField-field" value="${l||""}" required />\n      </div>\n      <div class="ms-TextField">\n        <button id="save-settings" class="ms-Button ms-Button--primary">\n          <span class="ms-Button-label">Save Settings</span>\n        </button>\n        <button id="test-api-settings" class="ms-Button ms-Button--secondary">\n          <span class="ms-Button-label">Test API Settings</span>\n        </button>\n      </div>\n    `,document.getElementById("save-settings").onclick=y,document.getElementById("test-api-settings").onclick=g,console.log("Syncro Ticket Creator: Settings UI rendered")):console.error("Syncro Ticket Creator: Element with id 'app-body' not found")}()}(),document.getElementById("create-ticket").onclick=m}async function y(){console.log("Syncro Ticket Creator: saveSettings called");const e=document.getElementById("syncro-url").value,t=document.getElementById("syncro-api-key").value;if(e&&t)try{await o(e,t),d=e,l=t,r("Settings saved successfully. Initializing app...","success"),await async function(){console.log("Syncro Ticket Creator: verifyApiSettings called");try{await c(d,l)}catch(e){throw console.error("Error verifying API settings:",e),new Error("Invalid API settings. Please check your Syncro URL and API Key.")}}(),await async function(){console.log("Syncro Ticket Creator: populateCustomers called");try{r("Loading customers...","info");const e=await c(d,l),t=document.getElementById("customer-select");if(!t)return void console.error("Syncro Ticket Creator: Element with id 'customer-select' not found");t.innerHTML='<option value="">Select a customer</option>',e.forEach((e=>{const n=document.createElement("option");n.value=e.id.toString(),n.textContent=a(e.name),t.appendChild(n)})),t.onchange=f,i()}catch(e){throw console.error("Error populating customers:",e),r("Failed to load customers. Please check your Syncro settings.","error"),e}}(),i()}catch(e){console.error("Error saving settings:",e),r("Failed to save settings. Please try again.","error")}else r("Please enter both Syncro URL and API Key.","error")}async function g(){console.log("Syncro Ticket Creator: testApiSettings called");const e=document.getElementById("syncro-url").value,t=document.getElementById("syncro-api-key").value;if(e&&t)try{const n=await c(e,t);console.log("Syncro Ticket Creator: API settings verified successfully",n),r("API settings verified successfully!","success")}catch(e){console.error("Error verifying API settings:",e),r("Failed to verify API settings. Please try again.","error")}else r("Please enter both Syncro URL and API Key.","error")}async function f(){console.log("Syncro Ticket Creator: populateContacts called");try{r("Loading contacts...","info");const e=document.getElementById("customer-select").value,t=await async function(e,t,n){console.log("Syncro API: fetchSyncroContacts called");const o=await fetch(`${e}/api/v1/customers/${n}/contacts`,{headers:{Authorization:`Bearer ${t}`}});if(!o.ok)throw new Error(`Failed to fetch contacts: ${o.status} ${o.statusText}`);return o.json()}(d,l,parseInt(e)),n=document.getElementById("contact-select");if(!n)return void console.error("Syncro Ticket Creator: Element with id 'contact-select' not found");n.innerHTML='<option value="">Select a contact</option>',t.forEach((e=>{const t=document.createElement("option");t.value=e.id.toString(),t.textContent=a(e.name),n.appendChild(t)})),i(),async function(){console.log("Syncro Ticket Creator: populateAssets called");try{r("Loading assets...","info");const e=document.getElementById("customer-select").value,t=await async function(e,t,n){console.log("Syncro API: fetchSyncroAssets called");const o=await fetch(`${e}/api/v1/customers/${n}/assets`,{headers:{Authorization:`Bearer ${t}`}});if(!o.ok)throw new Error(`Failed to fetch assets: ${o.status} ${o.statusText}`);return o.json()}(d,l,parseInt(e)),n=document.getElementById("asset-select");if(!n)return void console.error("Syncro Ticket Creator: Element with id 'asset-select' not found");n.innerHTML='<option value="">Select an asset</option>',t.forEach((e=>{const t=document.createElement("option");t.value=e.id.toString(),t.textContent=a(e.name),n.appendChild(t)})),i()}catch(e){console.error("Error populating assets:",e),r("Failed to load assets. Please try again.","error")}}()}catch(e){console.error("Error populating contacts:",e),r("Failed to load contacts. Please try again.","error")}}async function m(){console.log("Syncro Ticket Creator: createTicket called");try{r("Creating ticket...","info");const e=await async function(){return console.log("Syncro Ticket Creator: getEmailInfo called"),new Promise((e=>{if("undefined"!=typeof Office&&Office.context&&Office.context.mailbox){const t=Office.context.mailbox.item,n={subject:t?.subject||"",content:"",senderEmail:t?.from?.emailAddress||"",senderName:t?.from?.displayName||""};t?.body.getAsync(Office.CoercionType.Text,(t=>{t.status===Office.AsyncResultStatus.Succeeded&&(n.content=t.value),e(n)}))}else e({subject:"",content:"",senderEmail:"",senderName:""})}))}(),t=document.getElementById("customer-select").value,n=document.getElementById("contact-select").value,o=document.getElementById("asset-select").value,s=document.getElementById("ticket-title").value||e.subject,c=document.getElementById("ticket-message").value||e.content;if(!(t&&n&&s&&c))return void r("Please fill in all required fields.","error");const i={customer_id:parseInt(t),contact_id:parseInt(n),asset_id:o?parseInt(o):null,subject:s,problem_type:"Other",status:"New",comment:c},a=await async function(e,t,n){console.log("Syncro API: createSyncroTicket called");const o=await fetch(`${e}/api/v1/tickets`,{method:"POST",headers:{Authorization:`Bearer ${t}`,"Content-Type":"application/json"},body:JSON.stringify(n)});if(!o.ok)throw new Error(`Failed to create ticket: ${o.status} ${o.statusText}`);return o.json()}(d,l,i);console.log("Ticket created:",a),r("Ticket created successfully!","success")}catch(e){console.error("Error creating ticket:",e),r("Failed to create ticket. Please try again.","error")}finally{i()}}e.observe(document.body,{childList:!0,subtree:!0}),document.addEventListener("DOMContentLoaded",(function(){console.log("settings.ts: loadSettings called");const e=n(),t=e.syncroUrl,o=e.syncroApiKey;t&&(document.getElementById("syncro-url").value=t),o&&(document.getElementById("syncro-api-key").value=o)})),"undefined"!=typeof Office?(console.log("Syncro Ticket Creator: Office environment detected"),Office.onReady((e=>{console.log("Syncro Ticket Creator: Office.onReady called",e),e.host===Office.HostType.Outlook&&document.addEventListener("DOMContentLoaded",u)}))):(console.log("Syncro Ticket Creator: Non-Office environment detected"),document.addEventListener("DOMContentLoaded",(()=>{console.log("Syncro Ticket Creator: DOMContentLoaded event fired"),u()})))}(),e=r(14385),t=r.n(e),n=new URL(r(58394),r.b),o=new URL(r(98362),r.b),t()(n),t()(o)}();
//# sourceMappingURL=taskpane.js.map