import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import CopyPageComponent, { ICopyPageComponentProps } from './components/CopyPageComponent';

const LOG_SOURCE: string = 'CopyPageApplicationCustomizer';

// export interface ICopyPageApplicationCustomizerProperties {
//   testMessage: string;
// }

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CopyPageApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {
  private _topPlaceholder?: HTMLElement; // Reference to the placeholder DOM element

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${LOG_SOURCE}`);

    // Wait for page to load, then inject button into the page canvas
    this._waitAndInjectButton();

    return Promise.resolve();
  }

  private _waitAndInjectButton(): void {
    // Wait for the Republish button to be available
    const checkInterval = setInterval(() => {
      // Find the Republish button - it contains the text "Republish"
      const buttons = document.querySelectorAll('[role="menuitem"], button');
      let republishButton: Element | undefined;
      
      buttons.forEach((button: Element) => {
        if (button.textContent?.includes('Republish')) {
          republishButton = button;
        }
      });
      
      if (republishButton && republishButton.parentElement) {
        clearInterval(checkInterval);
        
        // Create a container div for the button
        const container = document.createElement('div');
        container.id = 'copy-page-button-container';
        container.style.display = 'inline-block';
        container.style.marginLeft = '8px';
        container.style.verticalAlign = 'middle';
        
        // Insert container right after the Republish button
        republishButton.parentElement.insertBefore(container, republishButton.nextSibling);
        
        this._topPlaceholder = container;

        // Retrieve site URL and page details
        const siteUrl = this.context.pageContext.web.absoluteUrl;
        const serverRequestPath = this.context.pageContext.site.serverRequestPath;
        const pageName = serverRequestPath.substring(serverRequestPath.lastIndexOf('/') + 1);

        // Create the React element for the CopyPageComponent
        const elem: React.ReactElement<ICopyPageComponentProps> = React.createElement(CopyPageComponent, {
          context: this.context,
          siteUrl: siteUrl,
          pageName: pageName,
          pageUrl: serverRequestPath,
        });

        // Render the React element into the container
        ReactDOM.render(elem, container);
      }
    }, 100);

    // Timeout after 5 seconds to prevent infinite checking
    setTimeout(() => clearInterval(checkInterval), 5000);
  }

  public onDispose(): void {
    Log.info(LOG_SOURCE, `Disposed ${LOG_SOURCE}`);

    // Unmount the React component from the placeholder DOM element
    if (this._topPlaceholder) {
      ReactDOM.unmountComponentAtNode(this._topPlaceholder);
    }
  }
}