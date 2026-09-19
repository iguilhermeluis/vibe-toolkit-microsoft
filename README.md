# @vibe-toolkit/vibe-toolkit-microsoft

[![npm version](https://img.shields.io/npm/v/@vibe-toolkit/vibe-toolkit-microsoft?color=cb3837&logo=npm)](https://www.npmjs.com/package/@vibe-toolkit/vibe-toolkit-microsoft)
[![npm downloads](https://img.shields.io/npm/dm/@vibe-toolkit/vibe-toolkit-microsoft?color=blue&logo=npm)](https://www.npmjs.com/package/@vibe-toolkit/vibe-toolkit-microsoft)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![TypeScript](https://img.shields.io/badge/TypeScript-ready-blue?logo=typescript)](https://www.typescriptlang.org/)

> A TypeScript package for Microsoft integrations. The public API is currently under active development.

## Overview

`@vibe-toolkit/vibe-toolkit-microsoft` will provide a consistent foundation for building modern web applications that integrate with Microsoft services.

## Installation

```bash
npm install @vibe-toolkit/vibe-toolkit-microsoft
```

## Configure and use

```ts
import { createVibeMicrosoftClient } from "@vibe-toolkit/vibe-toolkit-microsoft";

const client = createVibeMicrosoftClient({
	auth: {
		clientId: import.meta.env.VITE_ENTRA_CLIENT_ID,
		tenantId: import.meta.env.VITE_ENTRA_TENANT_ID,
		redirectUri: window.location.origin,
	},
	sharePoint: {
		siteUrl: "https://contoso.sharepoint.com/sites/operations",
		lists: {
			requests: { name: "Service requests" },
		},
	},
	copilot: {
		agentId: import.meta.env.VITE_COPILOT_AGENT_ID,
	},
});

const requestsList = client.getList("requests");

console.log(requestsList.name); // "Service requests"
```

Use environment variables for public browser configuration such as the Entra application ID, tenant ID, and Copilot agent ID. Do not include client secrets in frontend applications.

`createVibeMicrosoftClient` validates and exposes configuration. Microsoft Entra authentication, Copilot conversations, and SharePoint list reads and writes will be introduced as separate APIs after the authenticated session is implemented.

## API status

Available now:

- `createVibeMicrosoftClient`, which validates the Entra ID, SharePoint, and Copilot configuration.
- `client.getList(key)`, which resolves a named SharePoint list configuration.

Not implemented yet:

- Microsoft Entra sign-in and token acquisition.
- Copilot Studio conversations.
- SharePoint list reads and writes.

## Planned SharePoint access

After authentication is implemented, the client will expose the native PnPjs `SPFI` interface for advanced scenarios. This API is planned and is not available in the current release:

```ts
await client.initialize();

const sp = client.getSharePoint(); // SPFI
const requestsList = client.getList("requests");

const items = await sp.web.lists
	.getByTitle(requestsList.name)
	.items();
```

The package will also provide higher-level list methods for common operations. The native `SPFI` escape hatch will remain available for PnPjs APIs not covered by those helpers.

## Package design

The toolkit is designed to provide a framework-agnostic core that can be used with React, Vue, Next.js, SPFx, and other web applications.

Planned package extensions include:

```text
@vibe-toolkit/core
@vibe-toolkit/react
@vibe-toolkit/next
@vibe-toolkit/cli
```

## Development

Clone the repository and install the dependencies:

```bash
git clone https://github.com/iguilhermeluis/vibe-toolkit-microsoft.git
cd vibe-toolkit-microsoft
npm install
```

Available scripts:

```bash
npm run build
npm test
```

## Project status

This project is currently under active development. APIs may change before the first stable release.

## Links

- 📦 [npm package](https://www.npmjs.com/package/@vibe-toolkit/vibe-toolkit-microsoft)
- 💻 [GitHub repository](https://github.com/iguilhermeluis/vibe-toolkit-microsoft)
- 🐛 [Report an issue](https://github.com/iguilhermeluis/vibe-toolkit-microsoft/issues)

## Author

Created and maintained by [Guilherme Luis Faustino](https://github.com/iguilhermeluis).

## License

MIT © [Guilherme Luis Faustino](https://github.com/iguilhermeluis)
