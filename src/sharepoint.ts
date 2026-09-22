import { spfi, SPBrowser, type SPFI } from "@pnp/sp";
import "@pnp/sp/webs/index.js";
import "@pnp/sp/lists/index.js";
import "@pnp/sp/items/index.js";
import type { TokenProvider } from "./auth";
import { nonEmpty, secureUrl, validateScopes } from "./validation";

export interface SharePointListConfiguration { name: string; }
export interface SharePointConfiguration {
  siteUrl: string;
  /** Defaults to the SharePoint origin + /.default; configure delegated SharePoint permissions in Entra. */
  scopes?: string[];
  lists?: Record<string, SharePointListConfiguration>;
}

export function validateSharePoint(configuration: SharePointConfiguration): void {
  const url = secureUrl(configuration.siteUrl, "sharePoint.siteUrl");
  const scopes = configuration.scopes ?? [`${url.origin}/.default`];
  validateScopes(scopes);
  if (scopes.some(scope => !scope.startsWith(`${url.origin}/`))) throw new Error("SharePoint scopes must target the configured SharePoint origin.");
  for (const [key, list] of Object.entries(configuration.lists ?? {})) {
    nonEmpty(key, "sharePoint list key");
    nonEmpty(list?.name, `sharePoint.lists.${key}.name`);
  }
}

/** Native PnPjs instance with web, list and item extensions installed. */
export function createSharePointClient(configuration: SharePointConfiguration, tokens: TokenProvider): SPFI {
  validateSharePoint(configuration);
  const origin = new URL(configuration.siteUrl).origin;
  const scopes = [...(configuration.scopes ?? [`${origin}/.default`])];
  return spfi(configuration.siteUrl).using(SPBrowser(), instance => {
    instance.on.auth.replace(async (url: URL, init: RequestInit) => {
      if (new URL(url).origin !== origin) throw new Error("Refusing to send a SharePoint token to a different origin.");
      const token = await tokens.getAccessToken(scopes);
      const headers = new Headers(init.headers);
      headers.set("Authorization", `Bearer ${token}`);
      return [url, { ...init, headers: Object.fromEntries(headers.entries()), redirect: "error" }];
    });
    return instance;
  });
}
