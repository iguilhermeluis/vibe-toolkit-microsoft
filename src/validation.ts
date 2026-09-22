export function nonEmpty(value: unknown, key: string): asserts value is string {
  if (typeof value !== "string" || !value.trim()) throw new Error(`Configuration value "${key}" is required.`);
}

export function secureUrl(value: string, key: string, allowLocalhost = false): URL {
  let url: URL;
  try { url = new URL(value); } catch { throw new Error(`Configuration value "${key}" must be a valid URL.`); }
  const local = allowLocalhost && url.protocol === "http:" && ["localhost", "127.0.0.1", "[::1]"].includes(url.hostname);
  if ((!local && url.protocol !== "https:") || url.username || url.password || url.hash || url.search) {
    throw new Error(`Configuration value "${key}" must use HTTPS (HTTP localhost is allowed for callbacks), without credentials, query or fragment.`);
  }
  return url;
}

export function validateScopes(scopes: string[]): void {
  if (!Array.isArray(scopes) || !scopes.length || scopes.some(scope => typeof scope !== "string" || !scope.trim() || /\s/.test(scope))) {
    throw new Error("At least one non-empty scope is required.");
  }
  const apiScopes = scopes.filter(s => !["openid", "profile", "offline_access"].includes(s));
  const resources = new Set(apiScopes.map(scope => {
    const slash = scope.lastIndexOf("/");
    return slash < 0 ? "https://graph.microsoft.com" : scope.slice(0, slash).toLowerCase();
  }));
  if (resources.size > 1) throw new Error("Request scopes for one resource at a time (Graph, SharePoint or Copilot).");
  if (scopes.some(s => s.endsWith("/.default")) && apiScopes.length > 1) throw new Error("Do not mix .default with individual API scopes.");
}
