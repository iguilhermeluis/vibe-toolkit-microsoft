VibeToolkitRedirectBridge.broadcastResponseToMainFrame().catch(() => {
  document.getElementById('status').textContent = 'Authentication could not be completed. Close this window and try again.';
});
