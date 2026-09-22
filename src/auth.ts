import { PublicClientApplication, type IPublicClientApplication, type AccountInfo, type AuthenticationResult } from "@azure/msal-browser";
import { nonEmpty, secureUrl, validateScopes } from "./validation";

export { InteractionRequiredAuthError } from "@azure/msal-browser";
export interface TokenProvider { getAccessToken(scopes: string[]): Promise<string>; }
export interface EntraConfiguration {
  clientId: string;
  tenantId: string;
  /** Same-origin page implementing the MSAL 5 redirect bridge. */
  redirectUri: string;
}
export interface AuthClientOptions {
  /** Reuse the host application's MSAL 5 instance with matching auth and callback settings. */
  msalInstance?: IPublicClientApplication;
}
const graphScopes = ["https://graph.microsoft.com/User.Read"];
const initializations = new WeakMap<IPublicClientApplication, Promise<void>>();
const interactions = new WeakSet<IPublicClientApplication>();

export class VibeAuthClient implements TokenProvider {
  private instance?: IPublicClientApplication;
  private loginPromise?: Promise<AuthenticationResult>;
  private loginScopes?: string;

  constructor(readonly authConfiguration: EntraConfiguration, options: AuthClientOptions = {}) {
    nonEmpty(authConfiguration?.clientId, "auth.clientId");
    nonEmpty(authConfiguration?.tenantId, "auth.tenantId");
    if (!/^[a-zA-Z0-9.-]+$/.test(authConfiguration.tenantId)) throw new Error("auth.tenantId must be a tenant ID or domain.");
    secureUrl(authConfiguration.redirectUri, "auth.redirectUri", true);
    this.instance = options.msalInstance;
  }

  private getInstance(): IPublicClientApplication {
    if (!this.instance) {
      if (typeof window === "undefined" || !window.crypto?.subtle) throw new Error("Authentication requires a browser in a secure context (HTTPS or localhost).");
      if (new URL(this.authConfiguration.redirectUri).origin !== window.location.origin) throw new Error("auth.redirectUri must have the same origin as the application.");
      this.instance = new PublicClientApplication({
        auth: {
          clientId: this.authConfiguration.clientId,
          authority: `https://login.microsoftonline.com/${this.authConfiguration.tenantId}`,
          redirectUri: this.authConfiguration.redirectUri,
          postLogoutRedirectUri: this.authConfiguration.redirectUri,
        },
        cache: { cacheLocation: "sessionStorage" },
      });
    }
    return this.instance;
  }

  /** Call on application startup, before enabling sign-in buttons. Safe to call repeatedly. */
  async initialize(): Promise<void> {
    const instance = this.getInstance();
    let promise = initializations.get(instance);
    if (!promise) {
      promise = instance.initialize().catch(error => { initializations.delete(instance); throw error; });
      initializations.set(instance, promise);
    }
    await promise;
  }

  private async ready(): Promise<IPublicClientApplication> {
    await this.initialize();
    return this.getInstance();
  }

  private account(instance: IPublicClientApplication): AccountInfo | null {
    const active = instance.getActiveAccount();
    if (active) return active;
    const accounts = instance.getAllAccounts();
    if (accounts.length > 1) throw new Error("Multiple accounts found. Call selectAccount(homeAccountId) before continuing.");
    return accounts[0] ?? null;
  }

  async selectAccount(homeAccountId: string): Promise<void> {
    const instance = await this.ready();
    const account = instance.getAllAccounts().find(a => a.homeAccountId === homeAccountId);
    if (!account) throw new Error("The selected account is not signed in.");
    instance.setActiveAccount(account);
  }

  async getAccounts(): Promise<AccountInfo[]> { return (await this.ready()).getAllAccounts(); }

  private async interactive<T>(instance: IPublicClientApplication, action: () => Promise<T>): Promise<T> {
    if (interactions.has(instance)) throw new Error("An authentication interaction is already in progress.");
    interactions.add(instance);
    try { return await action(); } finally { interactions.delete(instance); }
  }

  async login(scopes: string[] = graphScopes): Promise<AuthenticationResult> {
    validateScopes(scopes);
    const key = [...scopes].sort().join(" ");
    if (this.loginPromise) {
      if (this.loginScopes !== key) throw new Error("A login with different scopes is already in progress.");
      return this.loginPromise;
    }
    this.loginScopes = key;
    this.loginPromise = this.loginInternal([...scopes]);
    try { return await this.loginPromise; } finally { this.loginPromise = undefined; this.loginScopes = undefined; }
  }

  private async loginInternal(scopes: string[]): Promise<AuthenticationResult> {
    const instance = await this.ready();
    return this.interactive(instance, async () => {
      const result = await instance.loginPopup({ scopes });
      instance.setActiveAccount(result.account);
      return result;
    });
  }

  /** Silent only: handle InteractionRequiredAuthError in the UI, then call acquireTokenInteractive from a user gesture. */
  async getAccessToken(scopes: string[] = graphScopes): Promise<string> {
    validateScopes(scopes);
    const instance = await this.ready();
    const account = this.account(instance);
    if (!account) throw new Error("No signed-in account found. Call login() before requesting a token.");
    return (await instance.acquireTokenSilent({ account, scopes: [...scopes] })).accessToken;
  }

  async acquireTokenInteractive(scopes: string[]): Promise<string> {
    validateScopes(scopes);
    const instance = await this.ready();
    return this.interactive(instance, async () => {
      const account = this.account(instance);
      const result = await instance.acquireTokenPopup({ scopes: [...scopes], ...(account ? { account } : {}) });
      instance.setActiveAccount(result.account);
      return result.accessToken;
    });
  }

  async logout(): Promise<void> {
    const instance = await this.ready();
    const account = this.account(instance);
    if (!account) return;
    await this.interactive(instance, () => instance.logoutPopup({ account, postLogoutRedirectUri: this.authConfiguration.redirectUri }));
    instance.setActiveAccount(null);
  }
}
