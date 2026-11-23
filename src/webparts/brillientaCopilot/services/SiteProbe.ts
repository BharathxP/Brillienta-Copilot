// Reads a SharePoint site with the signed-in admin's permissions: its ids, its lists and, when Helpdesk is there,
// the Helpdesk settings the agent needs. Nothing is written.

import { SPHttpClient } from '@microsoft/sp-http';
import { IAppTarget } from '../core/types';
import { findApps, helpdeskOptions, ISiteInfo, ISiteList } from '../core/discovery';

export class ProbeError extends Error {
  public readonly status: number;
  public constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}

export class SiteProbe {
  public constructor(private readonly _sp: SPHttpClient) {}

  private async _get<T>(url: string): Promise<T> {
    const res = await this._sp.get(url, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
    if (!res.ok) {
      const reason = res.status === 403 || res.status === 401 ? 'You do not have access to this site.' : res.status === 404 ? 'The site does not exist.' : `SharePoint answered ${res.status}.`;
      throw new ProbeError(res.status, reason);
    }
    return (await res.json()) as T;
  }

  public async site(url: string): Promise<ISiteInfo> {
    const [site, web, lists] = await Promise.all([
      this._get<{ value: string }>(`${url}/_api/site/id`),
      this._get<{ value: string }>(`${url}/_api/web/id`),
      this._get<{ value: ISiteList[] }>(`${url}/_api/web/lists?$select=Id,Title&$filter=Hidden eq false&$top=1000`)
    ]);
    return { url, host: new URL(url).hostname, siteId: site.value, webId: web.value, lists: lists.value };
  }

  /** Apps found on the site, with their options filled in. */
  public async apps(url: string): Promise<IAppTarget[]> {
    const info = await this.site(url);
    const apps = findApps(info);
    for (const a of apps) {
      if (a.key === 'helpdesk') {
        const settingsTitle = `${a.prefix}Settings`.replace(/'/g, "''");
        const rows = await this._get<{ value: Array<{ SettingValue: string | null }> }>(
          `${url}/_api/web/lists/getbytitle('${settingsTitle}')/items?$select=SettingValue&$filter=Title eq 'Config'&$top=1`
        ).catch(() => ({ value: [] as Array<{ SettingValue: string | null }> }));
        a.options = helpdeskOptions(rows.value[0] ? rows.value[0].SettingValue || undefined : undefined);
      }
    }
    return apps;
  }
}
