import * as React from 'react';
import {
  DefaultButton, Icon, Link, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, Stack, Text, TextField
} from '@fluentui/react';
import { IAgentInput, IAppTarget } from '../core/types';
import { buildAgent, MODULES, validateAgentInput } from '../core/compose';
import { normaliseSiteUrl } from '../core/discovery';
import { buildPackageZip } from '../core/package';
import { base64ToBytes, COLOR_PNG_BASE64, OUTLINE_PNG_BASE64 } from '../core/icons';
import { adminConsentUrl, newGuid, publisherSignIn } from '../core/publisher';
import { SiteProbe } from '../services/SiteProbe';
import { ICopilotSettings, SettingsStore } from '../services/SettingsStore';

export interface ISetupPageProps {
  probe: SiteProbe;
  store: SettingsStore;
  currentSiteUrl: string;
  siteTitle: string;
  tenantId: string;
  appVersion: string;
}

interface ISiteResult {
  url: string;
  apps: IAppTarget[];
  error?: string;
}

const card: React.CSSProperties = { background: 'var(--white, #fff)', border: '1px solid #e8eaed', borderRadius: 16, padding: 24, marginBottom: 16 };
const stepTitle = (n: number, title: string): JSX.Element => (
  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} style={{ marginBottom: 12 }}>
    <span style={{ width: 28, height: 28, borderRadius: 14, background: '#008d62', color: '#fff', display: 'grid', placeItems: 'center', fontWeight: 600 }}>{n}</span>
    <Text variant="xLarge" style={{ fontWeight: 600 }}>{title}</Text>
  </Stack>
);

function defaults(p: ISetupPageProps): ICopilotSettings {
  return {
    agentName: 'Brillienta',
    orgName: p.siteTitle || 'Our company',
    appId: newGuid(),
    oauthRegistrationId: '',
    brandColor: '#008D62',
    websiteUrl: '',
    privacyUrl: '',
    termsUrl: '',
    sites: [p.currentSiteUrl],
    version: 1
  };
}

export const SetupPage: React.FC<ISetupPageProps> = (props) => {
  const [settings, setSettings] = React.useState<ICopilotSettings>(() => defaults(props));
  const [loading, setLoading] = React.useState(true);
  const [scanning, setScanning] = React.useState(false);
  const [results, setResults] = React.useState<ISiteResult[]>([]);
  const [newSite, setNewSite] = React.useState('');
  const [message, setMessage] = React.useState<{ type: MessageBarType; text: string } | undefined>();
  const publisher = publisherSignIn();

  const scan = React.useCallback(async (sites: string[]) => {
    setScanning(true);
    const out: ISiteResult[] = [];
    for (const url of sites) {
      try {
        out.push({ url, apps: await props.probe.apps(url) });
      } catch (e) {
        out.push({ url, apps: [], error: (e as Error).message });
      }
    }
    setResults(out);
    setScanning(false);
  }, [props.probe]);

  React.useEffect(() => {
    props.store
      .load()
      .then((saved) => {
        const s = saved ? { ...defaults(props), ...saved } : defaults(props);
        setSettings(s);
        return scan(s.sites);
      })
      .catch((e) => setMessage({ type: MessageBarType.error, text: (e as Error).message }))
      .finally(() => setLoading(false));
  }, []);

  const update = (patch: Partial<ICopilotSettings>): void => setSettings((s) => ({ ...s, ...patch }));

  // One target per app: the first site where it was found.
  const apps: IAppTarget[] = [];
  results.forEach((r) => r.apps.forEach((a) => {
    if (!apps.some((x) => x.key === a.key)) apps.push(a);
  }));

  const input: IAgentInput = {
    agentName: settings.agentName,
    orgName: settings.orgName,
    appId: settings.appId,
    appVersion: props.appVersion,
    oauthRegistrationId: publisher ? publisher.oauthRegistrationId : settings.oauthRegistrationId.trim(),
    tenantId: props.tenantId,
    developerName: settings.orgName,
    websiteUrl: settings.websiteUrl,
    privacyUrl: settings.privacyUrl,
    termsUrl: settings.termsUrl,
    brandColor: settings.brandColor,
    apps
  };
  const problems = validateAgentInput(input);

  const addSite = (): void => {
    const url = normaliseSiteUrl(newSite);
    if (!url) {
      setMessage({ type: MessageBarType.warning, text: 'Enter a SharePoint site address, e.g. https://contoso.sharepoint.com/sites/hr' });
      return;
    }
    if (settings.sites.indexOf(url) >= 0) return;
    const sites = settings.sites.concat([url]);
    update({ sites });
    setNewSite('');
    void scan(sites);
  };

  const removeSite = (url: string): void => {
    const sites = settings.sites.filter((s) => s !== url);
    update({ sites });
    setResults((r) => r.filter((x) => x.url !== url));
  };

  const save = async (): Promise<void> => {
    try {
      await props.store.save(settings);
      setMessage({ type: MessageBarType.success, text: 'Saved.' });
    } catch (e) {
      setMessage({ type: MessageBarType.error, text: (e as Error).message });
    }
  };

  const download = async (): Promise<void> => {
    const zip = buildPackageZip(buildAgent(input), { color: base64ToBytes(COLOR_PNG_BASE64), outline: base64ToBytes(OUTLINE_PNG_BASE64) });
    const blob = new Blob([zip], { type: 'application/zip' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${settings.agentName.replace(/[^A-Za-z0-9-]+/g, '-')}-copilot-agent.zip`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
    await save();
  };

  if (loading) return <Spinner size={SpinnerSize.large} label="Looking for your Brillienta apps…" style={{ padding: 48 }} />;

  return (
    <div style={{ maxWidth: 900, margin: '0 auto', padding: '24px 16px', fontFamily: "'Segoe UI', system-ui, sans-serif" }}>
      <Text variant="xxLarge" block style={{ fontWeight: 700 }}>Brillienta Copilot</Text>
      <Text variant="large" block style={{ color: '#5f6368', margin: '8px 0 24px' }}>
        One Microsoft 365 Copilot agent for your Brillienta apps. It reads and writes your own SharePoint lists as each signed-in person, so SharePoint permissions still decide what everyone sees.
      </Text>
      {message && <MessageBar messageBarType={message.type} onDismiss={() => setMessage(undefined)} style={{ marginBottom: 16 }}>{message.text}</MessageBar>}

      <section style={card} aria-labelledby="s1">
        {stepTitle(1, 'Find your apps')}
        <Text block style={{ marginBottom: 12 }}>Add every site that holds a Brillienta app. Lists are found by name, including custom prefixes.</Text>
        {results.map((r) => (
          <div key={r.url} style={{ borderTop: '1px solid #e8eaed', padding: '12px 0' }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text style={{ fontWeight: 600, overflowWrap: 'anywhere' }}>{r.url}</Text>
              {r.url !== props.currentSiteUrl && <DefaultButton text="Remove" onClick={() => removeSite(r.url)} />}
            </Stack>
            {r.error && <MessageBar messageBarType={MessageBarType.error}>{r.error}</MessageBar>}
            {!r.error && !r.apps.length && <Text style={{ color: '#5f6368' }}>No Brillienta app found on this site.</Text>}
            {r.apps.map((a) => {
              const m = MODULES[a.key];
              const issues = m.check ? m.check(a) : [];
              const notes = m.notes ? m.notes(a) : [];
              return (
                <Stack key={a.key} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginTop: 6 }}>
                  <Icon iconName={issues.length ? 'Warning' : 'CompletedSolid'} style={{ color: issues.length ? '#9a6a00' : '#008d62' }} />
                  <Text>
                    <b>{m.label}</b> (lists {a.prefix}…){a.key === 'directory' || notes.length ? ' · read-only' : ' · read and write'}
                    {issues.map((x) => <span key={x} style={{ display: 'block', color: '#9a6a00' }}>{x}</span>)}
                    {notes.map((x) => <span key={x} style={{ display: 'block', color: '#5f6368' }}>{x}</span>)}
                  </Text>
                </Stack>
              );
            })}
          </div>
        ))}
        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 12 }} wrap>
          <TextField placeholder="https://contoso.sharepoint.com/sites/hr" value={newSite} onChange={(_, v) => setNewSite(v || '')} styles={{ root: { flex: 1, minWidth: 240 } }} ariaLabel="Site address" />
          <DefaultButton text="Add site" onClick={addSite} />
          <DefaultButton text="Scan again" iconProps={{ iconName: 'Refresh' }} disabled={scanning} onClick={() => void scan(settings.sites)} />
        </Stack>
        {scanning && <Spinner label="Scanning…" style={{ marginTop: 12 }} />}
      </section>

      <section style={card} aria-labelledby="s2">
        {stepTitle(2, 'Name your agent')}
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField label="Agent name (up to 30 characters)" value={settings.agentName} maxLength={30} onChange={(_, v) => update({ agentName: v || '' })} />
          <TextField label="Organisation name" value={settings.orgName} onChange={(_, v) => update({ orgName: v || '' })} />
          <TextField label="Brand colour" value={settings.brandColor} onChange={(_, v) => update({ brandColor: v || '' })} />
          <TextField label="Website (optional)" value={settings.websiteUrl} onChange={(_, v) => update({ websiteUrl: v || '' })} />
        </Stack>
      </section>

      <section style={card} aria-labelledby="s3">
        {stepTitle(3, 'Sign-in')}
        {publisher ? (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text>A Global Administrator approves the agent&apos;s Microsoft Graph permissions for everyone, once.</Text>
            <Link href={adminConsentUrl(publisher.appId, props.tenantId)} target="_blank">Grant admin consent</Link>
          </Stack>
        ) : (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text>
              Paste the OAuth client registration ID from the Teams Developer Portal (Tools › OAuth client registration). The agent signs in with
              Sites.ReadWrite.All, User.Read and offline_access, the same as the Helpdesk agent, so an existing Helpdesk registration works here too.
            </Text>
            <TextField label="OAuth client registration ID" value={settings.oauthRegistrationId} onChange={(_, v) => update({ oauthRegistrationId: v || '' })} />
          </Stack>
        )}
      </section>

      <section style={card} aria-labelledby="s4">
        {stepTitle(4, 'Download and upload')}
        {problems.length ? (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            <b>Before you can download:</b>
            <ul style={{ margin: '4px 0 0', paddingLeft: 18 }}>{problems.map((p) => <li key={p}>{p}</li>)}</ul>
          </MessageBar>
        ) : (
          <Text block>Ready: {apps.map((a) => MODULES[a.key].label).join(', ')}.</Text>
        )}
        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ margin: '12px 0' }}>
          <PrimaryButton text="Download agent package" iconProps={{ iconName: 'Download' }} disabled={problems.length > 0} onClick={() => void download()} />
          <DefaultButton text="Save settings" onClick={() => void save()} />
        </Stack>
        <Text block style={{ color: '#5f6368' }}>
          Then upload the zip in the Microsoft 365 admin center › Settings › Integrated apps › Upload custom apps, and deploy it to the people who should use it.
          Upload a new package whenever you add an app or rename the agent; it keeps the same app id ({settings.appId}) so it updates in place.
        </Text>
      </section>
    </div>
  );
};
