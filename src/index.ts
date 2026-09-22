import packageJson from "../package.json";
import { VibeAuthClient, type EntraConfiguration, type AuthClientOptions } from "./auth";
import { createSharePointClient, validateSharePoint, type SharePointConfiguration } from "./sharepoint";
import { VibeCopilotClient, type CopilotConfiguration } from "./copilot";
import type { SPFI } from "@pnp/sp";

export * from "./auth";
export * from "./sharepoint";
export * from "./copilot";
export const VIBE_TOOLKIT_VERSION = packageJson.version;

export interface VibeMicrosoftClientConfiguration {
  auth: EntraConfiguration;
  sharePoint?: SharePointConfiguration;
  copilot?: CopilotConfiguration;
}

export class VibeMicrosoftClient extends VibeAuthClient {
  private sharePointClient?: SPFI;
  private copilotClient?: VibeCopilotClient;

  constructor(readonly configuration: VibeMicrosoftClientConfiguration, options?: AuthClientOptions) {
    super(configuration.auth, options);
    if (configuration.sharePoint) validateSharePoint(configuration.sharePoint);
    if (configuration.copilot) this.copilotClient = new VibeCopilotClient(configuration.copilot, this);
  }

  getSharePoint(): SPFI {
    const configuration = this.configuration.sharePoint;
    if (!configuration) throw new Error("SharePoint is not configured.");
    return this.sharePointClient ??= createSharePointClient(configuration, this);
  }

  getList(listKey: string) {
    const lists = this.configuration.sharePoint?.lists;
    if (!lists || !Object.hasOwn(lists, listKey)) throw new Error(`SharePoint list "${listKey}" is not configured.`);
    return lists[listKey]!;
  }

  /** Returns one server page. Use getSharePoint() for filters and async pagination. */
  async getListItems<T = Record<string, unknown>>(listKey: string): Promise<T[]> {
    return this.getSharePoint().web.lists.getByTitle(this.getList(listKey).name).items<T[]>();
  }

  getCopilot(): VibeCopilotClient {
    if (!this.copilotClient) throw new Error("Copilot Studio is not configured.");
    return this.copilotClient;
  }
}

export function createVibeMicrosoftClient(configuration: VibeMicrosoftClientConfiguration, options?: AuthClientOptions) {
  return new VibeMicrosoftClient(configuration, options);
}
