import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'SparqOnlinePageAlertsApplicationCustomizerStrings';

import { sp } from "@pnp/sp/presets/all";
import PageAlerts, { IPageAlertProps } from './components/PageAlerts';

export const QUALIFIED_NAME = 'Extension.ApplicationCustomizer.SparqOnlinePageAlerts';

const LOG_SOURCE: string = 'SparqOnlinePageAlertsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISparqOnlinePageAlertsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  siteUrl: string;
  listName: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SparqOnlinePageAlertsApplicationCustomizer
  extends BaseApplicationCustomizer<ISparqOnlinePageAlertsApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    await super.onInit();

    sp.setup(this.context);
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    if (!this.properties.siteUrl || !this.properties.listName)
     {
      const e: Error = new Error('Missing required configuration parameters');
      Log.error(QUALIFIED_NAME, e);
      return Promise.reject(e);
     }

     const header = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

        if (!header) {
            const error = new Error('Could not find placeholder Top');
            Log.error(QUALIFIED_NAME, error);
            return Promise.reject(error);
        }

        let site = this.context.pageContext.site;
        let tenantUrl = site.absoluteUrl.replace(site.serverRelativeUrl, "");

        const elem: React.ReactElement<IPageAlertProps> = React.createElement(PageAlerts, { 
            siteUrl: `${tenantUrl}${this.properties.siteUrl}`, 
            listName: this.properties.listName,
            culture: this.context.pageContext.cultureInfo.currentUICultureName
         });

        ReactDOM.render(elem, header.domElement);

  }
}
