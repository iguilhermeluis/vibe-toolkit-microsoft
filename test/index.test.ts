import { afterEach, describe, expect, it, vi } from 'vitest';
import type { IPublicClientApplication, AccountInfo } from '@azure/msal-browser';
import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client';
import { createVibeMicrosoftClient, VIBE_TOOLKIT_VERSION, VibeAuthClient, VibeCopilotClient, createSharePointClient, InteractionRequiredAuthError } from '../src/index';
import packageJson from '../package.json';
import { Web } from '@pnp/sp/webs/index.js';
const auth = { clientId: 'app', tenantId: 'tenant', redirectUri: 'https://example.com/callback' };
const sharePoint = { siteUrl: 'https://contoso.sharepoint.com/sites/test', lists: { requests: { name: 'Requests' } } };
function fakeMsal() {
  const account = { homeAccountId: 'one', username: 'user@contoso.com' } as AccountInfo;
  let active: AccountInfo | null = null;
  const instance = {
    initialize: vi.fn().mockResolvedValue(undefined),
    getAllAccounts: vi.fn(() => [account]),
    getActiveAccount: vi.fn(() => active),
    setActiveAccount: vi.fn((value: AccountInfo | null) => { active = value; }),
    loginPopup: vi.fn().mockResolvedValue({ account, accessToken: 'login-token' }),
    logoutPopup: vi.fn().mockResolvedValue(undefined),
    acquireTokenSilent: vi.fn().mockResolvedValue({ account, accessToken: 'silent-token' }),
    acquireTokenPopup: vi.fn().mockResolvedValue({ account, accessToken: 'interactive-token' }),
  };
  return { instance, account, options: { msalInstance: instance as unknown as IPublicClientApplication } };
}
afterEach(() => { vi.restoreAllMocks(); vi.unstubAllGlobals(); });

describe('configuration', () => {
  it('exports version and configured lists', () => {
    expect(VIBE_TOOLKIT_VERSION).toBe(packageJson.version);
    const client = createVibeMicrosoftClient({ auth, sharePoint });
    expect(client.getList('requests').name).toBe('Requests');
    expect(() => client.getList('toString')).toThrow('not configured');
  });
  it('validates factory and direct construction', () => {
    expect(() => new VibeAuthClient({ ...auth, clientId: '' })).toThrow('auth.clientId');
    expect(() => createVibeMicrosoftClient({ auth: { ...auth, redirectUri: 'javascript:alert(1)' } })).toThrow('HTTPS');
    expect(() => new VibeAuthClient({ ...auth, tenantId: 'tenant/path' })).toThrow('tenant ID');
  });
  it('allows import and configuration on SSR but rejects authentication there', async () => {
    await expect(new VibeAuthClient(auth).initialize()).rejects.toThrow('browser');
  });
  it('does not accept Graph scopes for SharePoint', () => {
    expect(() => createSharePointClient({ ...sharePoint, scopes: ['User.Read'] }, { getAccessToken: vi.fn() })).toThrow('SharePoint origin');
  });
});

describe('authentication lifecycle', () => {
  it('initializes once across concurrent operations and shared clients', async () => {
    const { instance, options } = fakeMsal();
    const a = new VibeAuthClient(auth, options), b = new VibeAuthClient(auth, options);
    await Promise.all([a.getAccessToken(), b.getAccounts(), a.initialize()]);
    expect(instance.initialize).toHaveBeenCalledTimes(1);
    expect(instance.acquireTokenSilent).toHaveBeenCalledWith(expect.objectContaining({ scopes: ['https://graph.microsoft.com/User.Read'] }));
  });
  it('retries failed initialization', async () => {
    const { instance, options } = fakeMsal();
    instance.initialize.mockRejectedValueOnce(new Error('initialization failed'));
    const client = new VibeAuthClient(auth, options);
    await expect(client.initialize()).rejects.toThrow('initialization failed');
    await client.initialize();
    expect(instance.initialize).toHaveBeenCalledTimes(2);
  });
  it('deduplicates login and selects the returned account', async () => {
    const { instance, options, account } = fakeMsal();
    const client = new VibeAuthClient(auth, options);
    await Promise.all([client.login(), client.login()]);
    expect(instance.loginPopup).toHaveBeenCalledTimes(1);
    expect(instance.setActiveAccount).toHaveBeenCalledWith(account);
  });
  it('releases the interaction lock after a cancelled login', async () => {
    const { instance, options } = fakeMsal();
    instance.loginPopup.mockRejectedValueOnce(new Error('cancelled'));
    const client = new VibeAuthClient(auth, options);
    await expect(client.login()).rejects.toThrow('cancelled');
    await client.login();
    expect(instance.loginPopup).toHaveBeenCalledTimes(2);
  });
  it('requires account selection instead of choosing the first of many', async () => {
    const { instance, options, account } = fakeMsal();
    instance.getAllAccounts.mockReturnValue([account, { ...account, homeAccountId: 'two' }]);
    const client = new VibeAuthClient(auth, options);
    await expect(client.getAccessToken()).rejects.toThrow('Multiple accounts');
    await client.selectAccount('two');
    await client.getAccessToken();
    expect(instance.acquireTokenSilent).toHaveBeenCalledWith(expect.objectContaining({ account: expect.objectContaining({ homeAccountId: 'two' }) }));
  });
  it('does not open a popup during background token acquisition', async () => {
    const { instance, options } = fakeMsal();
    const error = new InteractionRequiredAuthError('interaction_required', 'User interaction required');
    instance.acquireTokenSilent.mockRejectedValue(error);
    const client = new VibeAuthClient(auth, options);
    await expect(client.getAccessToken()).rejects.toBe(error);
    expect(instance.acquireTokenPopup).not.toHaveBeenCalled();
    expect(await client.acquireTokenInteractive(['User.Read'])).toBe('interactive-token');
  });
  it('rejects mixed resources and empty scopes', async () => {
    const { options } = fakeMsal();
    const client = new VibeAuthClient(auth, options);
    await expect(client.getAccessToken(['User.Read', 'https://contoso.sharepoint.com/.default'])).rejects.toThrow('one resource');
    await expect(client.getAccessToken([])).rejects.toThrow('scope');
  });
  it('initializes before logout and configures the bridge return URL', async () => {
    const { instance, options, account } = fakeMsal();
    await new VibeAuthClient(auth, options).logout();
    expect(instance.initialize).toHaveBeenCalledTimes(1);
    expect(instance.logoutPopup).toHaveBeenCalledWith({ account, postLogoutRedirectUri: auth.redirectUri });
    expect(instance.setActiveAccount).toHaveBeenLastCalledWith(null);
  });
  it('reports a missing session without opening a popup', async () => {
    const { instance, options } = fakeMsal();
    instance.getAllAccounts.mockReturnValue([]);
    const client = new VibeAuthClient(auth, options);
    await expect(client.getAccessToken()).rejects.toThrow('No signed-in account');
    await client.logout();
    expect(instance.logoutPopup).not.toHaveBeenCalled();
  });
});

describe('real PnPjs integration (only HTTP is stubbed)', () => {
  it('reads lists with the SharePoint token, native imports and response parsing', async () => {
    const { instance, options } = fakeMsal();
    const fetch = vi.fn().mockResolvedValue(new Response(JSON.stringify({ value: [{ Id: 1, Title: 'Alpha' }] }), { headers: { 'Content-Type': 'application/json' } }));
    vi.stubGlobal('fetch', fetch);
    const client = createVibeMicrosoftClient({ auth, sharePoint }, options);
    expect(await client.getListItems('requests')).toEqual([{ Id: 1, Title: 'Alpha' }]);
    expect(instance.acquireTokenSilent).toHaveBeenCalledWith(expect.objectContaining({ scopes: ['https://contoso.sharepoint.com/.default'] }));
    const [url, init] = fetch.mock.calls[0]!;
    expect(String(url)).toContain("/_api/web/lists/getByTitle('Requests')/items");
    expect(new Headers(init.headers).get('Authorization')).toBe('Bearer silent-token');
    expect(init.redirect).toBe('error');
    expect(client.getSharePoint()).toBe(client.getSharePoint());
  });
  it('renews tokens for each request', async () => {
    const tokens = { getAccessToken: vi.fn().mockResolvedValueOnce('one').mockResolvedValueOnce('two') };
    const fetch = vi.fn().mockImplementation(async () => new Response(JSON.stringify({ value: [] })));
    vi.stubGlobal('fetch', fetch);
    const sp = createSharePointClient(sharePoint, tokens);
    await sp.web.lists(); await sp.web.lists();
    expect(fetch.mock.calls.map(([, init]) => new Headers(init.headers).get('Authorization'))).toEqual(['Bearer one', 'Bearer two']);
  });
  it('blocks authentication to a different origin before requesting a token', async () => {
    const tokens = { getAccessToken: vi.fn().mockResolvedValue('token') };
    const fetch = vi.fn(); vi.stubGlobal('fetch', fetch);
    const sp = createSharePointClient(sharePoint, tokens);
    const other = Web([sp.web, 'https://other.example/_api/web']);
    await expect(other()).rejects.toThrow('different origin');
    expect(tokens.getAccessToken).not.toHaveBeenCalled();
    expect(fetch).not.toHaveBeenCalled();
  });
  it('propagates HTTP failures rather than returning an empty list', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('Forbidden', { status: 403 })));
    const sp = createSharePointClient(sharePoint, { getAccessToken: vi.fn().mockResolvedValue('token') });
    await expect(sp.web.lists()).rejects.toThrow('403');
  });
});

describe('Copilot Studio', () => {
  it('requires environment and schema configuration', () => {
    expect(() => new VibeCopilotClient({ environmentId: '', schemaName: 'agent' }, { getAccessToken: vi.fn() })).toThrow('environmentId');
    expect(() => new VibeCopilotClient({ environmentId: 'environment' }, { getAccessToken: vi.fn() })).toThrow('schemaName');
  });
  it('renews tokens per operation and preserves explicit conversation IDs', async () => {
    const start = vi.spyOn(CopilotStudioClient.prototype, 'startConversationWithResponse').mockResolvedValue({ activities: [], conversationId: 'conversation', isNewConversation: true });
    const send = vi.spyOn(CopilotStudioClient.prototype, 'executeWithResponse').mockResolvedValue({ activities: [], conversationId: 'conversation', activityCount: 0 });
    const tokens = { getAccessToken: vi.fn().mockResolvedValue('token') };
    const client = new VibeCopilotClient({ environmentId: 'environment', schemaName: 'agent' }, tokens);
    const result = await client.startConversation();
    const activity = { type: 'message', text: 'Hello' };
    await client.sendActivity(activity, result.conversationId);
    expect(start).toHaveBeenCalledTimes(1);
    expect(send).toHaveBeenCalledWith(expect.objectContaining(activity), 'conversation');
    expect(tokens.getAccessToken).toHaveBeenCalledTimes(2);
    expect(tokens.getAccessToken).toHaveBeenCalledWith(['https://api.powerplatform.com/.default']);
    await expect(client.sendActivity(activity, '')).rejects.toThrow('conversationId');
  });
});
