import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'SpfxTeamFullWidthApplicationCustomizerStrings';

const LOG_SOURCE: string = 'SpfxTeamFullWidthApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpfxTeamFullWidthApplicationCustomizerProperties {
  // This is an example; replace with your own property
  cssurl: string;
  enableFullWidth:boolean;
  
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxTeamFullWidthApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxTeamFullWidthApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const cssUrl: string = this.properties.cssurl;
    if(cssUrl) {
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      const customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.href = cssUrl;
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";
      head.insertAdjacentElement("beforeEnd", customStyle);
    }

    if(this.properties.enableFullWidth) {
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      const fullWidthStyle: HTMLStyleElement = document.createElement("style");
      fullWidthStyle.innerText = `
        .CanvasComponent .fui-FluentProvider:first-of-type .CanvasZoneSectionContainer:first-of-type {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 100%;
            margin: unset;
            align-items: center;
            justify-content: center;
            display: flex;
        }
      `;
      fullWidthStyle.type = "text/css";
      head.insertAdjacentElement("beforeEnd", fullWidthStyle);
    }

    
    return Promise.resolve();
  }
}
