// Shared types for the Brillienta Copilot agent builder. Everything in core/ is pure (no SharePoint calls) and unit
// tested with node:test against lib-commonjs.

export type Json = Record<string, unknown>;

export type AppKey = 'helpdesk' | 'assets' | 'directory';

/** A Brillienta app found on a site: where its lists are and the values the agent needs from its settings. */
export interface IAppTarget {
  key: AppKey;
  siteUrl: string;
  /** Graph site id: "<host>,<site collection guid>,<web guid>". */
  graphSiteId: string;
  /** List title prefix the app was installed with (HD_, AM_, BrillED_...). */
  prefix: string;
  /** List ids by the module's list keys. */
  listIds: Record<string, string>;
  /** App-specific values read from its settings (e.g. the helpdesk ticket prefix). */
  options: Record<string, string>;
}

export interface IListRef {
  key: string;
  /** List title = prefix + suffix. */
  suffix: string;
}

export interface IFunctionSpec {
  name: string;
  description: string;
  capabilities?: Json;
}

/** One Brillienta app's part of the agent: its lists, instructions, Graph calls and plugin functions. */
export interface IAppModule {
  key: AppKey;
  label: string;
  /** Default list prefix the app installs with. */
  defaultPrefix: string;
  /** Lists the agent uses; all must exist for the app to be included. */
  lists: IListRef[];
  /** Plugin namespace (letters only). */
  namespace: string;
  /** The app's section of the agent instructions. Kept short: all apps share the 8,000-character limit. */
  instructions(t: IAppTarget): string;
  /** OpenAPI paths (Graph, relative to https://graph.microsoft.com/v1.0). */
  paths(t: IAppTarget): Json;
  functions: IFunctionSpec[];
  pluginDescription(orgName: string): { human: string; model: string };
  starters: Array<{ title: string; text: string }>;
  /** Plain-language reasons the app cannot be included as found (e.g. a required setting is off). */
  check?(t: IAppTarget): string[];
}

export interface IAgentInput {
  agentName: string;
  orgName: string;
  /** Stable GUID of the Teams/Copilot app package. */
  appId: string;
  appVersion: string;
  oauthRegistrationId: string;
  tenantId?: string;
  developerName: string;
  websiteUrl: string;
  privacyUrl: string;
  termsUrl: string;
  brandColor: string;
  apps: IAppTarget[];
}

export type AgentFiles = Record<string, string>;
