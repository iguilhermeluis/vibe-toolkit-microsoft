import { Activity } from "@microsoft/agents-activity";
import { ConnectionSettings, CopilotStudioClient, ScopeHelper } from "@microsoft/agents-copilotstudio-client";
import type { TokenProvider } from "./auth";
import { nonEmpty } from "./validation";

export interface CopilotConfiguration {
  environmentId: string;
  schemaName?: string;
  /** @deprecated Use schemaName (the published agent's schema name, not its GUID). */
  agentId?: string;
}
export type CopilotActivity = Partial<Activity> & Pick<Activity, "type">;
export type CopilotStartRequest = Parameters<CopilotStudioClient["startConversationWithResponse"]>[0];

/** Public-cloud, delegated-user integration. Tokens are renewed before each operation. */
export class VibeCopilotClient {
  private readonly settings: ConnectionSettings;
  readonly scopes: readonly string[];

  constructor(configuration: CopilotConfiguration, private readonly tokens: TokenProvider) {
    nonEmpty(configuration.environmentId, "copilot.environmentId");
    nonEmpty(configuration.schemaName ?? configuration.agentId, "copilot.schemaName");
    this.settings = new ConnectionSettings({ environmentId: configuration.environmentId, schemaName: configuration.schemaName ?? configuration.agentId });
    this.scopes = Object.freeze([ScopeHelper.getScopeFromSettings(this.settings)]);
  }

  private async client(): Promise<CopilotStudioClient> {
    return new CopilotStudioClient(this.settings, await this.tokens.getAccessToken([...this.scopes]));
  }

  async startConversation(request?: CopilotStartRequest) {
    return (await this.client()).startConversationWithResponse(request);
  }

  async sendActivity(activity: CopilotActivity, conversationId: string) {
    nonEmpty(conversationId, "conversationId");
    return (await this.client()).executeWithResponse(Activity.fromObject(activity), conversationId);
  }

  async *sendActivityStreaming(activity: CopilotActivity, conversationId: string) {
    nonEmpty(conversationId, "conversationId");
    yield* (await this.client()).executeStreaming(Activity.fromObject(activity), conversationId);
  }
}
