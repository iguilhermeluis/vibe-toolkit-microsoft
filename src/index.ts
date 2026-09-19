export const VIBE_TOOLKIT_VERSION = "0.1.0";

export interface EntraConfiguration {
	clientId: string;
	tenantId: string;
	redirectUri: string;
}

export interface SharePointListConfiguration {
	name: string;
}

export interface SharePointConfiguration {
	siteUrl: string;
	lists?: Record<string, SharePointListConfiguration>;
}

export interface CopilotConfiguration {
	agentId: string;
}

export interface VibeMicrosoftClientConfiguration {
	auth: EntraConfiguration;
	sharePoint?: SharePointConfiguration;
	copilot?: CopilotConfiguration;
}

export class VibeMicrosoftClient {
	constructor(readonly configuration: VibeMicrosoftClientConfiguration) {}

	getList(listKey: string): SharePointListConfiguration {
		const list = this.configuration.sharePoint?.lists?.[listKey];

		if (!list) {
			throw new Error(`SharePoint list "${listKey}" is not configured.`);
		}

		return list;
	}
}

export function createVibeMicrosoftClient(
	configuration: VibeMicrosoftClientConfiguration,
): VibeMicrosoftClient {
	assertNonEmptyValue(configuration.auth.clientId, "auth.clientId");
	assertNonEmptyValue(configuration.auth.tenantId, "auth.tenantId");
	assertUrl(configuration.auth.redirectUri, "auth.redirectUri");

	if (configuration.sharePoint) {
		assertUrl(configuration.sharePoint.siteUrl, "sharePoint.siteUrl");

		for (const [listKey, list] of Object.entries(configuration.sharePoint.lists ?? {})) {
			assertNonEmptyValue(listKey, "sharePoint list key");
			assertNonEmptyValue(list.name, `sharePoint.lists.${listKey}.name`);
		}
	}

	if (configuration.copilot) {
		assertNonEmptyValue(configuration.copilot.agentId, "copilot.agentId");
	}

	return new VibeMicrosoftClient(configuration);
}

function assertNonEmptyValue(value: string, configurationKey: string): void {
	if (!value.trim()) {
		throw new Error(`Configuration value "${configurationKey}" is required.`);
	}
}

function assertUrl(value: string, configurationKey: string): void {
	try {
		new URL(value);
	} catch {
		throw new Error(`Configuration value "${configurationKey}" must be a valid URL.`);
	}
}