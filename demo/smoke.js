import * as esm from '../dist/index.browser.js';
import { broadcastResponseToMainFrame } from '../dist/redirect-bridge.browser.js';
const output = document.getElementById('output');
const originalFetch = window.fetch;
const lines = [];
try {
  if (typeof broadcastResponseToMainFrame !== 'function') throw new Error('Callback export missing');
  for (const [name, api] of [['CDN IIFE', window.VibeToolkitMicrosoft], ['CDN ESM', esm]]) {
    let initialized = false;
    const account = { homeAccountId: 'test-account' };
    const msalInstance = {
      initialize: async () => { initialized = true; },
      getActiveAccount: () => account,
      acquireTokenSilent: async ({ scopes }) => {
        if (!initialized || scopes[0] !== 'https://contoso.sharepoint.com/.default') throw new Error('Incorrect authentication lifecycle/scopes');
        return { accessToken: 'browser-test-token' };
      },
    };
    window.fetch = async (url, init) => {
      if (new URL(url).origin !== 'https://contoso.sharepoint.com') throw new Error('Incorrect target');
      if (new Headers(init.headers).get('Authorization') !== 'Bearer browser-test-token') throw new Error('Missing authorization header');
      return new Response(JSON.stringify({ value: [{ Id: 1, Title: 'Browser test' }] }), { headers: { 'Content-Type': 'application/json' } });
    };
    const client = api.createVibeMicrosoftClient({
      auth: { clientId: 'test', tenantId: 'test', redirectUri: new URL('./msal-callback.html', location.href).href },
      sharePoint: { siteUrl: 'https://contoso.sharepoint.com/sites/test', lists: { requests: { name: 'Requests' } } },
      copilot: { environmentId: 'test-environment', schemaName: 'test-agent' },
    }, { msalInstance });
    const items = await client.getListItems('requests');
    if (items[0]?.Id !== 1) throw new Error('Response parsing failed');
    if (client.getCopilot().scopes[0] !== 'https://api.powerplatform.com/.default') throw new Error('Copilot scope mismatch');
    lines.push(`${name} v${api.VIBE_TOOLKIT_VERSION}: OK — loading, simulated authentication, real PnPjs, and Copilot configuration.`);
  }
  lines.push('Callback ESM: OK. No requests were sent to a real tenant.');
  output.dataset.status = 'passed';
} catch (error) { lines.push(`FAILED: ${error.message}`); output.dataset.status = 'failed'; }
finally { window.fetch = originalFetch; output.textContent = lines.join('\n'); }
