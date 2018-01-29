import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import {AppInsights} from "applicationinsights-js";
import * as strings from 'GoogleAnalyticsApplicationCustomizerStrings';
import ApplicationCustomizerContext from '@microsoft/sp-application-base/lib/extensibility/ApplicationCustomizerContext';

const LOG_SOURCE: string = 'GoogleAnalyticsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGoogleAnalyticsApplicationCustomizerProperties {  
  trackingID: string;  
}

/*
http://sharepoint.handsontek.net/2017/12/21/how-to-add-google-analytics-to-the-modern-sharepoint/
Connect-PnPOnline -UseWebLogin -Url https://arupwestadvisory.sharepoint.com
Add-PnPCustomAction -ClientSideComponentId "00be02e5-93bc-449e-87c9-fa127ca564a3" -Name "Analytics" -Title "Analytics" 
-Location ClientSideExtension.ApplicationCustomizer -ClientSideComponentProperties: '{"trackingID":"UA-112734844-1"}'
*/
/** A Custom Action which can be run during execution of a Client Side Application */
export default class GoogleAnalyticsApplicationCustomizer
  extends BaseApplicationCustomizer<IGoogleAnalyticsApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    let trackingID: string = this.properties.trackingID;
    
    if (!trackingID) {
      Log.info(LOG_SOURCE, `${strings.MissingID}`);
    }else{
      var gtagScript = document.createElement("script");
      gtagScript.type = "text/javascript";
      gtagScript.src = `https://www.googletagmanager.com/gtag/js?id=${trackingID}`;    
      gtagScript.async = true;
      document.head.appendChild(gtagScript);  

      eval(`
        window.dataLayer = window.dataLayer || [];
        function gtag(){dataLayer.push(arguments);}
        gtag('js', new Date());    
        gtag('config',  '${trackingID}');
      `);
    }
    
     
    return Promise.resolve();
  }
}
