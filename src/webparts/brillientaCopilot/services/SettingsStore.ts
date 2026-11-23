// Brillienta Copilot's own settings, in the list BrillCP_Settings on the site where the setup page is (Title = key,
// SettingValue = JSON). The list is created the first time an admin saves.

import { SPHttpClient } from '@microsoft/sp-http';

export const SETTINGS_LIST = 'BrillCP_Settings';
const KEY = 'Agent';

export interface ICopilotSettings {
  agentName: string;
  orgName: string;
  /** Stable id of the generated Teams / Copilot app: keep it so re-uploads update the same app. */
  appId: string;
  oauthRegistrationId: string;
  brandColor: string;
  websiteUrl: string;
  privacyUrl: string;
  termsUrl: string;
  /** Sites that hold Brillienta apps. */
  sites: string[];
  version: number;
}

export class SettingsStore {
  public constructor(private readonly _sp: SPHttpClient, private readonly _webUrl: string) {}

  private _list(): string {
    return `${this._webUrl}/_api/web/lists/getbytitle('${SETTINGS_LIST}')`;
  }

  public async load(): Promise<ICopilotSettings | undefined> {
    const res = await this._sp.get(`${this._list()}/items?$select=Id,Title,SettingValue&$filter=Title eq '${KEY}'&$top=1`, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata' }
    });
    if (res.status === 404) return undefined;
    if (!res.ok) throw new Error(`Could not read ${SETTINGS_LIST} (${res.status}).`);
    const rows = ((await res.json()) as { value: Array<{ SettingValue: string | null }> }).value;
    if (!rows[0] || !rows[0].SettingValue) return undefined;
    try {
      return JSON.parse(rows[0].SettingValue) as ICopilotSettings;
    } catch {
      return undefined;
    }
  }

  private async _ensureList(): Promise<void> {
    const check = await this._sp.get(`${this._list()}?$select=Id`, SPHttpClient.configurations.v1);
    if (check.ok) return;
    const created = await this._sp.post(`${this._webUrl}/_api/web/lists`, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
      body: JSON.stringify({ Title: SETTINGS_LIST, BaseTemplate: 100, Description: 'Brillienta Copilot settings (agent name, app id, sites).', Hidden: true })
    });
    if (!created.ok) throw new Error(`Could not create ${SETTINGS_LIST} (${created.status}). Site owners can create it.`);
    const field = await this._sp.post(`${this._list()}/fields`, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
      body: JSON.stringify({ '@odata.type': '#SP.Field', Title: 'SettingValue', FieldTypeKind: 3 })
    });
    if (!field.ok) throw new Error(`Could not add the SettingValue column (${field.status}).`);
  }

  public async save(s: ICopilotSettings): Promise<void> {
    await this._ensureList();
    const json = { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' };
    const existing = await this._sp.get(`${this._list()}/items?$select=Id&$filter=Title eq '${KEY}'&$top=1`, SPHttpClient.configurations.v1, { headers: json });
    const rows = existing.ok ? ((await existing.json()) as { value: Array<{ Id: number }> }).value : [];
    const value = JSON.stringify(s);
    const res = rows[0]
      ? await this._sp.post(`${this._list()}/items(${rows[0].Id})`, SPHttpClient.configurations.v1, {
          headers: { ...json, 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
          body: JSON.stringify({ SettingValue: value })
        })
      : await this._sp.post(`${this._list()}/items`, SPHttpClient.configurations.v1, { headers: json, body: JSON.stringify({ Title: KEY, SettingValue: value }) });
    if (!res.ok) throw new Error(`Could not save settings (${res.status}).`);
  }
}
