// The agent's sign-in registered once by Brillienta for every organisation (same model and same Graph permissions as
// Brillienta-Helpdesk, so one registration can serve both agents). Until both values are filled in, each organisation
// pastes its own OAuth client registration ID (docs/setup.md).

import { GRAPH_SCOPES } from './openapi';

export interface IPublisherSignIn {
  publisher: string;
  /** Application (client) id of Brillienta's multi-tenant Entra app. */
  appId: string;
  /** Teams Developer Portal OAuth client registration id ("any Microsoft 365 organization"). */
  oauthRegistrationId: string;
}

export const PUBLISHER_SIGN_IN: IPublisherSignIn = { publisher: 'Brillienta', appId: '', oauthRegistrationId: '' };

export const OAUTH_REDIRECT = 'https://teams.microsoft.com/api/platform/v1.0/oAuthRedirect';

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function publisherSignIn(s: IPublisherSignIn = PUBLISHER_SIGN_IN): IPublisherSignIn | undefined {
  return GUID.test(s.appId.trim()) && s.oauthRegistrationId.trim().length >= 8 ? s : undefined;
}

export function adminConsentUrl(appId: string, tenantId?: string): string {
  const q = ['client_id=' + encodeURIComponent(appId), 'scope=' + encodeURIComponent(GRAPH_SCOPES.join(' ')), 'redirect_uri=' + encodeURIComponent(OAUTH_REDIRECT)].join('&');
  return `https://login.microsoftonline.com/${tenantId || 'organizations'}/v2.0/adminconsent?${q}`;
}

export function newGuid(random: () => number = Math.random): string {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = Math.floor(random() * 16);
    return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16);
  });
}
