// 1AV Project Health Monitor - settings
// Client ID and tenant ID are not secrets; they are safe to keep in this file.
// Never put a client secret in this repo.
window.MONITOR_CONFIG = {
  // Entra ID app registration (Single-page application platform)
  clientId: "PASTE-CLIENT-ID-HERE",
  tenantId: "fab3949d-3cb0-4097-802f-053ad7075219",

  // HTTP trigger URL of the "API" flow (When an HTTP request is received,
  // Who can trigger the flow = Any user in my tenant).
  // Leave empty to run in demo mode with data/sample.json.
  apiUrl: "",

  // Token audience required by Power Automate HTTP triggers
  scopes: ["https://service.flow.microsoft.com//.default"],

  refreshMinutes: 15,   // auto refresh interval
  staleMinutes: 35      // warn when the collector has not run for this long
};
