import { describe, expect, it } from "vitest";

import {
  createVibeMicrosoftClient,
  VIBE_TOOLKIT_VERSION,
} from "../src/index.js";

describe("VIBE_TOOLKIT_VERSION", () => {
  it("exposes the current package version", () => {
    expect(VIBE_TOOLKIT_VERSION).toBe("0.1.0");
  });
});

describe("createVibeMicrosoftClient", () => {
  it("configures Entra, Copilot, and named SharePoint lists", () => {
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
        agentId: "copilot-agent-id",
      },
    });

    expect(client.configuration.copilot?.agentId).toBe("copilot-agent-id");
    expect(client.getList("requests")).toEqual({ name: "Service requests" });
  });

  it("rejects a missing Entra client ID", () => {
    expect(() =>
      createVibeMicrosoftClient({
        auth: {
          clientId: "",
          tenantId: "directory-tenant-id",
          redirectUri: "https://app.contoso.com/auth/callback",
        },
      }),
    ).toThrow('Configuration value "auth.clientId" is required.');
  });
});