# @vibe-toolkit/vibe-toolkit-microsoft

[![npm version](https://img.shields.io/npm/v/@vibe-toolkit/vibe-toolkit-microsoft?color=cb3837&logo=npm)](https://www.npmjs.com/package/@vibe-toolkit/vibe-toolkit-microsoft)
[![npm downloads](https://img.shields.io/npm/dm/@vibe-toolkit/vibe-toolkit-microsoft?color=blue&logo=npm)](https://www.npmjs.com/package/@vibe-toolkit/vibe-toolkit-microsoft)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![TypeScript](https://img.shields.io/badge/TypeScript-ready-blue?logo=typescript)](https://www.typescriptlang.org/)

Framework-independent browser toolkit for Microsoft Entra ID, SharePoint/PnPjs and Copilot Studio. Distributed as npm ESM, a standalone browser ESM bundle, and a classic script bundle.

## Requirements and supported scope

- Modern browsers with ES2022, Web Crypto, fetch, streams and BroadcastChannel; HTTPS (HTTP localhost for development).
- Node >=22.12 for package installation/build tools. Authentication runs in the browser, not during SSR. Importing the npm package during SSR is supported.
- Microsoft public cloud and delegated user authentication. App-only secrets, sovereign clouds, cross-origin iframe authentication and automatic SPFx/Teams SSO are not implemented by this wrapper. Use the host's supported authentication adapter through `TokenProvider` where appropriate.
- An Entra SPA app registration with the exact callback URL and the required delegated permissions/consent. A successful build does not validate tenant policy, CORS, permissions or a published agent.

## npm

```sh
npm install @vibe-toolkit/vibe-toolkit-microsoft
```

```ts
import { createVibeMicrosoftClient } from '@vibe-toolkit/vibe-toolkit-microsoft';

const client = createVibeMicrosoftClient({
  auth: {
    clientId: 'YOUR_APPLICATION_ID',
    tenantId: 'YOUR_TENANT_ID',
    redirectUri: new URL('/auth/callback.html', location.origin).href,
  },
  sharePoint: {
    siteUrl: 'https://contoso.sharepoint.com/sites/operations',
    lists: { requests: { name: 'Service requests' } },
  },
  copilot: {
    environmentId: 'YOUR_ENVIRONMENT_ID',
    schemaName: 'YOUR_PUBLISHED_AGENT_SCHEMA_NAME',
  },
});

await client.initialize(); // During startup, before enabling sign-in buttons.
// In a user click handler:
await client.login();
```

Only configure the integrations you need. IDs are public configuration; never put client secrets or permanent access tokens in frontend code.

## CDN

Use the exact version you published; these examples target 0.1.2. Local changes do not update npm or CDN until you publish a new, unused version.

```html
<script src="https://cdn.jsdelivr.net/npm/@vibe-toolkit/vibe-toolkit-microsoft@0.1.2/dist/index.global.js"
        crossorigin="anonymous"></script>
<script>
  const { createVibeMicrosoftClient } = VibeToolkitMicrosoft;
  // Use the same configuration and API as the npm example.
</script>
```

For native browser modules:

```html
<script type="module">
  import { createVibeMicrosoftClient } from
    'https://cdn.jsdelivr.net/npm/@vibe-toolkit/vibe-toolkit-microsoft@0.1.2/dist/index.browser.js';
</script>
```

UNPKG can serve the same versioned paths. `dist/index.js` is the npm entry and requires a bundler; use `index.browser.js` for direct browser imports. Both browser bundles include their dependencies. Do not load both variants on one page.

Each build generates `dist/integrity.json` with SHA-384 values. When publishing, add the matching `integrity` value to classic script tags and use `crossorigin="anonymous"`. Serve versioned assets with immutable caching. CSP must permit the selected script host and the Microsoft endpoints used by your application; avoid wildcard policies.

## Required authentication callback (MSAL 5)

The callback page **and its scripts must be served from your application's own origin**, even when the main toolkit comes from CDN. Copy `dist/redirect-bridge.global.js` to your application's public assets, together with `THIRD-PARTY-NOTICES.txt`. Create `/auth/callback.html`:

```html
<!doctype html>
<html lang="en">
<head><meta charset="utf-8"><title>Signing in</title></head>
<body>
  <p id="status">Completing sign-in…</p>
  <script src="/auth/redirect-bridge.global.js"></script>
  <script src="/auth/callback.js"></script>
</body>
</html>
```

Contents of the same-origin `/auth/callback.js`:

```js
VibeToolkitRedirectBridge.broadcastResponseToMainFrame().catch(() => {
  document.getElementById('status').textContent = 'Sign-in could not be completed. Close this window and try again.';
});
```

With a bundler, alternatively import `broadcastResponseToMainFrame` from `@vibe-toolkit/vibe-toolkit-microsoft/redirect-bridge` in a dedicated callback entry. Do not load the application router or initialize another MSAL client on this page.

Register its exact URL as a **SPA redirect URI** in Entra. Serve the callback with `Cache-Control: no-store`, without `Cross-Origin-Opener-Policy`. Do not log or display its query/hash. The toolkit uses this callback for login and popup logout. [Microsoft redirect bridge guidance](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/redirect-bridge.md).

## Tokens and consent

`getAccessToken(scopes)` is silent. If consent, MFA or another interaction is required, it propagates the MSAL error; it never opens a background popup. Show a button whose click calls `acquireTokenInteractive(scopes)`, then retry the operation. You can import `InteractionRequiredAuthError` from the toolkit.

Request each resource separately:

| Resource | Scopes | Entra setup |
| --- | --- | --- |
| Graph | `https://graph.microsoft.com/User.Read` (default login) | Microsoft Graph delegated permissions |
| SharePoint REST | `https://contoso.sharepoint.com/.default` (default SharePoint behavior) | SharePoint delegated permissions, e.g. `AllSites.Read` for reads |
| Copilot | SDK-derived `https://api.powerplatform.com/.default` | Power Platform delegated `CopilotStudio.Copilot.Invoke`, published agent and user access |

`.default` uses permissions configured and consented for that resource; it does not grant access on its own. Follow least privilege. Consent requirements depend on tenant policy. Graph `Sites.Read.All` is not a SharePoint REST permission.

```ts
// Each interactive call belongs to a separate user action.
await client.acquireTokenInteractive(['https://contoso.sharepoint.com/.default']);
const items = await client.getListItems('requests');

const token = await client.getAccessToken(['https://graph.microsoft.com/User.Read']);
const response = await fetch('https://graph.microsoft.com/v1.0/me', {
  headers: { Authorization: `Bearer ${token}` },
});
if (!response.ok) throw new Error(`Graph failed: ${response.status}`);
const me = await response.json();
```

Tokens are obtained before every SharePoint/Copilot operation; MSAL handles cache and renewal. The toolkit doesn't retry writes or Copilot turns after authentication failures. PnPjs supplies its native transport retry behavior. Do not log tokens.

A restored single account can be used without a new login. With multiple accounts and no active account, call `getAccounts()` and `selectAccount(homeAccountId)` explicitly. `logout()` signs out the selected account. To reuse your application's MSAL 5 instance, pass `{ msalInstance }` as the second argument to the factory, with matching authentication/callback settings. Do not mix separate bundled copies of MSAL on the same page.

## SharePoint

`getListItems<T>(key)` returns **one server page**, not all items in a large list. `getSharePoint()` exposes native PnPjs with web/list/item extensions:

```ts
const query = client.getSharePoint().web.lists.getByTitle('Service requests')
  .items.select('Id', 'Title').top(100);
for await (const page of query) {
  // Process each page; avoid collecting an unlimited list in memory.
}
```

Additional PnPjs extensions can be imported by npm consumers using the same installed PnP version. CDN consumers need a custom build to include extensions beyond web/list/item. Requests are restricted to the configured origin; redirects are rejected to avoid forwarding tokens elsewhere. Validate SharePoint REST CORS from your deployment origin. If browser access is blocked, use a same-origin backend or a suitable Graph endpoint rather than disabling browser protections.

## Copilot Studio

```ts
const copilot = client.getCopilot();
// In a consent button click, if required:
await client.acquireTokenInteractive([...copilot.scopes]);
const conversation = await copilot.startConversation();
const answer = await copilot.sendActivity(
  { type: 'message', text: 'Hello' }, conversation.conversationId,
);
for await (const activity of copilot.sendActivityStreaming(
  { type: 'message', text: 'Tell me more' }, conversation.conversationId,
)) {
  // Render activities using your UI; treat agent content as untrusted content.
}
```

Keep the conversation ID with its signed-in user and discard it when switching accounts. Start/send calls fetch a fresh token from the provider. A long-running stream can still fail on network/token expiry; the application controls recovery. `agentId` is a deprecated alias for `schemaName`, not an agent GUID. Existing configurations must now include `environmentId`.

## Modular imports

```ts
import { VibeAuthClient } from '@vibe-toolkit/vibe-toolkit-microsoft/auth';
import { createSharePointClient } from '@vibe-toolkit/vibe-toolkit-microsoft/sharepoint';
import { VibeCopilotClient } from '@vibe-toolkit/vibe-toolkit-microsoft/copilot';
```

SharePoint and Copilot accept a `TokenProvider` with `getAccessToken(scopes): Promise<string>`, so a host application can supply its own supported authentication. Importing only `/auth` excludes PnPjs and Copilot from the application bundle. The main CDN bundle intentionally includes all integrations. CommonJS `require()` is not an advertised entry.

## Development and release

```sh
npm ci
npm run check
npm run demo
```

Open `http://localhost:8000/demo/msal-login.html` for the tenant test or `/demo/smoke.html` for a credential-free browser smoke test. The demo server serves only demo/build files on loopback and applies callback headers. It is not a production server.

`check` validates source/test types, unit and real-PnPjs HTTP tests, browser builds, tarball contents, package exports, a TypeScript consumer, npm browser bundling, CDN IIFE/ESM execution, subpath isolation and integrity hashes. HTTP/authentication boundaries are simulated in automated tests; live Entra, SharePoint CORS and Copilot streaming require your own tenant test before release.

`npm pack` and `npm publish` run these checks through `prepack`. Before publishing, choose an unused version, update CDN examples, run checks with `npm ci` on supported Node versions, and verify the packed artifact. Commit the lockfile for reproducible releases. Published files are allowlisted; third-party license notices and integrity hashes are included in `dist`.
