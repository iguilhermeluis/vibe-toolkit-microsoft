const { createVibeMicrosoftClient } = VibeToolkitMicrosoft;
const output = document.getElementById('output');
const fields = ['clientId', 'tenantId', 'redirectUri', 'sharePointSiteUrl', 'sharePointListName'];
const buttons = ['loginButton', 'logoutButton', 'consentButton', 'readButton'];
const storageKey = 'vibe-toolkit-msal-demo-config';
let client;
const show = value => { output.textContent = typeof value === 'string' ? value : JSON.stringify(value, null, 2); };
const setEnabled = enabled => buttons.forEach(id => { document.getElementById(id).disabled = !enabled || ((id === 'consentButton' || id === 'readButton') && !client?.configuration.sharePoint); });
document.getElementById('redirectUri').value = new URL('./msal-callback.html', location.href).href;
try {
  const saved = JSON.parse(localStorage.getItem(storageKey) || '{}');
  fields.forEach(id => { if (saved[id]) document.getElementById(id).value = saved[id]; });
} catch { /* Configuration persistence is optional. */ }
fields.forEach(id => document.getElementById(id).addEventListener('input', () => { client = undefined; setEnabled(false); }));
setEnabled(false);
document.getElementById('configureButton').onclick = async () => {
  client = undefined; setEnabled(false);
  try {
    const values = Object.fromEntries(fields.map(id => [id, document.getElementById(id).value.trim()]));
    const candidate = createVibeMicrosoftClient({
      auth: { clientId: values.clientId, tenantId: values.tenantId, redirectUri: values.redirectUri },
      ...(values.sharePointSiteUrl ? { sharePoint: { siteUrl: values.sharePointSiteUrl, lists: { requests: { name: values.sharePointListName } } } } : {}),
    });
    await candidate.initialize();
    client = candidate;
    try { localStorage.setItem(storageKey, JSON.stringify(values)); } catch {}
    setEnabled(true); show('Configuration ready. Sign in with Microsoft, then authorize SharePoint separately if required.');
  } catch (error) { show(error.message); }
};
async function run(action) {
  if (!client) return;
  setEnabled(false);
  try { await action(); } catch (error) { show(error.message); } finally { setEnabled(Boolean(client)); }
}
document.getElementById('loginButton').onclick = () => run(async () => {
  const result = await client.login();
  show({ status: 'login-success', username: result.account?.username });
});
document.getElementById('logoutButton').onclick = () => run(async () => { await client.logout(); show('Sign-out completed.'); });
document.getElementById('consentButton').onclick = () => run(async () => {
  const origin = new URL(client.configuration.sharePoint.siteUrl).origin;
  await client.acquireTokenInteractive([`${origin}/.default`]); show('SharePoint authorized.');
});
document.getElementById('readButton').onclick = () => run(async () => { show(await client.getListItems('requests')); });
