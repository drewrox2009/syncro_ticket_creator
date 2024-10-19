!function(){"use strict";var e,t,n,o,c={14385:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},98362:function(e,t,n){e.exports=n.p+"assets/logo-filled.png"},58394:function(e,t,n){e.exports=n.p+"fd6ebc4970a52eece689.css"}},r={};function s(e){var t=r[e];if(void 0!==t)return t.exports;var n=r[e]={exports:{}};return c[e](n,n.exports,s),n.exports}s.m=c,s.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return s.d(t,{a:t}),t},s.d=function(e,t){for(var n in t)s.o(t,n)&&!s.o(e,n)&&Object.defineProperty(e,n,{enumerable:!0,get:t[n]})},s.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),s.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;s.g.importScripts&&(e=s.g.location+"");var t=s.g.document;if(!e&&t&&(t.currentScript&&"SCRIPT"===t.currentScript.tagName.toUpperCase()&&(e=t.currentScript.src),!e)){var n=t.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!e||!/^http(s?):/.test(e));)e=n[o--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),s.p=e}(),s.b=document.baseURI||self.location.href,function(){function e(){t(document.getElementById("syncro-url").value,document.getElementById("syncro-api-key").value).then((()=>{console.log("Settings saved successfully")})).catch((e=>{console.error("Error saving settings:",e)}))}function t(e,t){return new Promise(((n,o)=>{Office.context.roamingSettings.set("syncroUrl",e),Office.context.roamingSettings.set("syncroApiKey",t),Office.context.roamingSettings.saveAsync((e=>{e.status===Office.AsyncResultStatus.Succeeded?n():o(e.error)}))}))}let n,o;function c(){console.log("Syncro Ticket Creator: initializeApp called"),document.getElementById("create-ticket").onclick=a,async function(){console.log("Syncro Ticket Creator: loadSyncroSettings called");const e={syncroUrl:Office.context.roamingSettings.get("syncroUrl"),syncroApiKey:Office.context.roamingSettings.get("syncroApiKey")};console.log("Retrieved settings:",e),n=e.syncroApiKey,o=e.syncroUrl,function(){console.log("Syncro Ticket Creator: showSettingsUI called");const e=document.getElementById("app-body");e?(console.log("Syncro Ticket Creator: app-body element found"),e.innerHTML=`\n      <h2>Syncro API Settings</h2>\n      <div class="ms-TextField">\n        <label class="ms-Label" for="syncro-url">Syncro URL</label>\n        <input type="text" id="syncro-url" class="ms-TextField-field" value="${o||""}" required />\n      </div>\n      <div class="ms-TextField">\n        <label class="ms-Label" for="syncro-api-key">Syncro API Key</label>\n        <input type="password" id="syncro-api-key" class="ms-TextField-field" value="${n||""}" required />\n      </div>\n      <div class="ms-TextField">\n        <button id="save-settings" class="ms-Button ms-Button--primary">\n          <span class="ms-Button-label">Save Settings</span>\n        </button>\n      </div>\n    `,document.getElementById("save-settings").onclick=r,console.log("Syncro Ticket Creator: Settings UI rendered")):console.error("Syncro Ticket Creator: Element with id 'app-body' not found")}()}()}async function r(){console.log("Syncro Ticket Creator: saveSettings called");const e=document.getElementById("syncro-url").value,c=document.getElementById("syncro-api-key").value;if(e&&c)try{await t(e,c),o=e,n=c,l("Settings saved successfully. Initializing app...","success"),await async function(){console.log("Syncro Ticket Creator: verifyApiSettings called");try{await i()}catch(e){throw console.error("Error verifying API settings:",e),new Error("Invalid API settings. Please check your Syncro URL and API Key.")}}(),await async function(){console.log("Syncro Ticket Creator: populateCustomers called");try{const e=await i(),t=document.getElementById("customer-select");t.innerHTML='<option value="">Select a customer</option>',e.forEach((e=>{const n=document.createElement("option");n.value=e.id.toString(),n.textContent=d(e.name),t.appendChild(n)})),t.onchange=s}catch(e){throw console.error("Error populating customers:",e),l("Failed to load customers. Please check your Syncro settings.","error"),e}}(),u()}catch(e){console.error("Error saving settings:",e),l("Failed to save settings. Please try again.","error")}else l("Please enter both Syncro URL and API Key.","error")}async function s(){console.log("Syncro Ticket Creator: populateContacts called");try{l("Loading contacts...");const e=document.getElementById("customer-select").value,t=await async function(e){console.log("Syncro Ticket Creator: fetchSyncroContacts called");const t=await fetch(`${o}/api/v1/customers/${e}/contacts`,{headers:{Authorization:`Bearer ${n}`}});if(!t.ok)throw new Error("Failed to fetch contacts");return t.json()}(parseInt(e)),c=document.getElementById("contact-select");c.innerHTML='<option value="">Select a contact</option>',t.forEach((e=>{const t=document.createElement("option");t.value=e.id.toString(),t.textContent=d(e.name),c.appendChild(t)})),u(),async function(){console.log("Syncro Ticket Creator: populateAssets called");try{l("Loading assets...");const e=document.getElementById("customer-select").value,t=await async function(e){console.log("Syncro Ticket Creator: fetchSyncroAssets called");const t=await fetch(`${o}/api/v1/customers/${e}/assets`,{headers:{Authorization:`Bearer ${n}`}});if(!t.ok)throw new Error("Failed to fetch assets");return t.json()}(parseInt(e)),c=document.getElementById("asset-select");c.innerHTML='<option value="">Select an asset</option>',t.forEach((e=>{const t=document.createElement("option");t.value=e.id.toString(),t.textContent=d(e.name),c.appendChild(t)})),u()}catch(e){console.error("Error populating assets:",e),l("Failed to load assets. Please try again.","error")}}()}catch(e){console.error("Error populating contacts:",e),l("Failed to load contacts. Please try again.","error")}}async function a(){console.log("Syncro Ticket Creator: createTicket called");try{l("Creating ticket...");const e=await async function(){return console.log("Syncro Ticket Creator: getEmailInfo called"),new Promise((e=>{if("undefined"!=typeof Office&&Office.context&&Office.context.mailbox){const t=Office.context.mailbox.item,n={subject:t?.subject||"",content:"",senderEmail:t?.from?.emailAddress||"",senderName:t?.from?.displayName||""};t?.body.getAsync(Office.CoercionType.Text,(t=>{t.status===Office.AsyncResultStatus.Succeeded&&(n.content=t.value),e(n)}))}else e({subject:"",content:"",senderEmail:"",senderName:""})}))}(),t=document.getElementById("customer-select").value,c=document.getElementById("contact-select").value,r=document.getElementById("asset-select").value,s=document.getElementById("ticket-title").value||e.subject,a=document.getElementById("ticket-message").value||e.content,i={customer_id:parseInt(t),contact_id:parseInt(c),asset_id:r?parseInt(r):null,subject:s,problem_type:"Other",status:"New",comment:a},u=await async function(e){console.log("Syncro Ticket Creator: createSyncroTicket called");const t=await fetch(`${o}/api/v1/tickets`,{method:"POST",headers:{Authorization:`Bearer ${n}`,"Content-Type":"application/json"},body:JSON.stringify(e)});if(!t.ok)throw new Error("Failed to create ticket");return t.json()}(i);console.log("Ticket created:",u),l("Ticket created successfully!","success")}catch(e){console.error("Error creating ticket:",e),l("Failed to create ticket. Please try again.","error")}}async function i(){console.log("Syncro Ticket Creator: fetchSyncroCustomers called");const e=await fetch(`${o}/api/v1/customers`,{headers:{Authorization:`Bearer ${n}`}});if(!e.ok)throw new Error("Failed to fetch customers");return e.json()}function l(e,t="info"){console.log(`Syncro Ticket Creator: showStatus called - ${t}: ${e}`);const n=document.getElementById("status-message");if(!n){const e=document.createElement("div");e.id="status-message",document.body.insertBefore(e,document.body.firstChild)}const o=n||document.getElementById("status-message");o.textContent=d(e),o.className=`status-message ${t}`,o.style.display="block"}function u(){console.log("Syncro Ticket Creator: hideStatus called");const e=document.getElementById("status-message");e&&(e.style.display="none")}function d(e){const t=document.createElement("div");return t.textContent=e,t.innerHTML}Office.onReady((t=>{t.host===Office.HostType.Outlook&&(document.getElementById("save-settings").onclick=e,function(){const e=Office.context.roamingSettings.get("syncroUrl"),t=Office.context.roamingSettings.get("syncroApiKey");e&&(document.getElementById("syncro-url").value=e),t&&(document.getElementById("syncro-api-key").value=t)}())})),"undefined"!=typeof Office?(console.log("Syncro Ticket Creator: Office environment detected"),Office.onReady((e=>{console.log("Syncro Ticket Creator: Office.onReady called",e),e.host===Office.HostType.Outlook&&c()}))):(console.log("Syncro Ticket Creator: Non-Office environment detected"),document.addEventListener("DOMContentLoaded",(()=>{console.log("Syncro Ticket Creator: DOMContentLoaded event fired"),c()})))}(),e=s(14385),t=s.n(e),n=new URL(s(58394),s.b),o=new URL(s(98362),s.b),t()(n),t()(o)}();
//# sourceMappingURL=taskpane.js.map