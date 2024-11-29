import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import * as strings from 'InjectCssApplicationCustomizerStrings';

const LOG_SOURCE: string = 'InjectCssApplicationCustomizer';

export interface IInjectCssApplicationCustomizerProperties {
  cssurl: string;
}

export default class InjectCssApplicationCustomizer
  extends BaseApplicationCustomizer<IInjectCssApplicationCustomizerProperties> {
  
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const windowURL = window.location.href;
    const siteURL = windowURL.split('/').splice(0,5).join('/');

    const cssUrl: string =  siteURL + this.properties.cssurl;
    if (cssUrl) {
        console.log("CSSURL siteURL", siteURL);
        // inject the style sheet
        const head: HTMLElement = document.getElementsByTagName("head")[0] || document.documentElement;
        const customStyle: HTMLLinkElement = document.createElement("link");
        customStyle.href = cssUrl;
        customStyle.rel = "stylesheet";
        customStyle.type = "text/css";
        head.insertAdjacentElement("beforeend", customStyle);
    }

    return Promise.resolve();
  }
}
