// Finds Brillienta apps among a site's lists. Pure: the web part fetches the list titles and ids, this decides.
//
// An app is found when every list it needs exists with one common prefix. That works for the default prefixes
// (HD_, AM_, BrillED_) and for custom ones, and it never mistakes another app's lists for it (Helpdesk needs
// <prefix>CopilotActions, Asset Management needs <prefix>AssetTypes...).

import { AppKey, IAppTarget } from './types';
import { APP_ORDER, MODULES } from './compose';

export interface ISiteList {
  Id: string;
  Title: string;
}

export interface ISiteInfo {
  url: string;
  /** Host name, e.g. contoso.sharepoint.com. */
  host: string;
  siteId: string;
  webId: string;
  lists: ISiteList[];
}

export function graphSiteId(host: string, siteId: string, webId: string): string {
  return `${host},${siteId.replace(/[{}]/g, '')},${webId.replace(/[{}]/g, '')}`;
}

/** Every app (and prefix) whose lists are all present on the site. */
export function findApps(site: ISiteInfo): IAppTarget[] {
  const byTitle: Record<string, string> = {};
  site.lists.forEach((l) => {
    byTitle[l.Title.toLowerCase()] = l.Id;
  });
  const found: IAppTarget[] = [];
  APP_ORDER.forEach((key: AppKey) => {
    const m = MODULES[key];
    const anchor = m.lists[0].suffix.toLowerCase();
    const prefixes = site.lists
      .map((l) => l.Title)
      .filter((t) => t.toLowerCase().endsWith(anchor) && t.length > anchor.length)
      .map((t) => t.substring(0, t.length - anchor.length));
    for (const prefix of prefixes) {
      const ids: Record<string, string> = {};
      const complete = m.lists.every((l) => {
        const id = byTitle[(prefix + l.suffix).toLowerCase()];
        if (id) ids[l.key] = id;
        return !!id || !!l.optional;
      });
      if (complete) {
        found.push({ key, siteUrl: site.url, graphSiteId: graphSiteId(site.host, site.siteId, site.webId), prefix, listIds: ids, options: {} });
        break;
      }
    }
  });
  return found;
}

/** Reads what the agent needs from Helpdesk's Config setting (JSON in <prefix>Settings, item "Config"). */
export function helpdeskOptions(configJson: string | undefined): Record<string, string> {
  let cfg: { ticketPrefix?: string; copilot?: { enabled?: boolean } } = {};
  try {
    cfg = configJson ? JSON.parse(configJson) : {};
  } catch {
    cfg = {};
  }
  return {
    ticketPrefix: typeof cfg.ticketPrefix === 'string' && cfg.ticketPrefix ? cfg.ticketPrefix : 'HD-',
    copilotEnabled: cfg.copilot && cfg.copilot.enabled === true ? 'true' : 'false'
  };
}

/** Normalises what an admin types as a site address. Returns undefined when it is not a SharePoint site URL. */
export function normaliseSiteUrl(input: string): string | undefined {
  const raw = (input || '').trim();
  let u: URL;
  try {
    u = new URL(raw);
  } catch {
    return undefined;
  }
  if (u.protocol !== 'https:' || !/\.sharepoint(-df)?\.com$/i.test(u.hostname)) return undefined;
  const pathname = u.pathname.replace(/\/(SitePages|Shared%20Documents|Lists|_layouts)\/.*$/i, '').replace(/\/+$/, '');
  return `${u.origin}${pathname}`;
}

/** Reads what the agent needs from Asset Management's Config setting (JSON in <prefix>Settings, item "Config"). */
export function assetOptions(configJson: string | undefined): Record<string, string> {
  let cfg: { features?: { copilot?: boolean; acknowledgements?: boolean } } = {};
  try {
    cfg = configJson ? JSON.parse(configJson) : {};
  } catch {
    cfg = {};
  }
  const f = cfg.features || {};
  return {
    copilotEnabled: f.copilot === true ? 'true' : 'false',
    // Receipt confirmations are on by default in Asset Management.
    acknowledgementsEnabled: f.acknowledgements === false ? 'false' : 'true'
  };
}
