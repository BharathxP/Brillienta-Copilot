import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import * as strings from 'BrillientaCopilotWebPartStrings';
import { SetupPage } from './components/SetupPage';
import { SiteProbe } from './services/SiteProbe';
import { SettingsStore } from './services/SettingsStore';

export default class BrillientaCopilotWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const ctx = this.context;
    const element = React.createElement(SetupPage, {
      probe: new SiteProbe(ctx.spHttpClient),
      store: new SettingsStore(ctx.spHttpClient, ctx.pageContext.web.absoluteUrl),
      currentSiteUrl: ctx.pageContext.web.absoluteUrl,
      siteTitle: ctx.pageContext.web.title,
      tenantId: ctx.pageContext.aadInfo ? ctx.pageContext.aadInfo.tenantId.toString() : '',
      appVersion: this.manifest.version || '1.0.0'
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: strings.PropertyPaneDescription }, groups: [] }] };
  }
}
