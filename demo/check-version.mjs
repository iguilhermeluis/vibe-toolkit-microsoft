import {
  VIBE_TOOLKIT_VERSION,
  createVibeMicrosoftClient,
} from "../dist/index.js";

console.log("VIBE_TOOLKIT_VERSION:", VIBE_TOOLKIT_VERSION);

const client = createVibeMicrosoftClient({
  auth: {
    clientId: "application-client-id",
    tenantId: "directory-tenant-id",
    redirectUri: "https://app.contoso.com/auth/callback",
  },
  sharePoint: {
    siteUrl: "https://contoso.sharepoint.com/sites/operations",
    lists: {
      requests: { name: "Service requests" },
    },
  },
  copilot: {
    environmentId: "environment-id",
    schemaName: "agent-schema-name",
  },
});

console.log("list name:", client.getList("requests").name);
console.log("demo ok");
