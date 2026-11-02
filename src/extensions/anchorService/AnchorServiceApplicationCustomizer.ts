import * as React from 'react';
import * as ReactDom from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

import * as strings from 'AnchorServiceApplicationCustomizerStrings';
import {
  EXCLUDE_ATTR,
  IAnchorEngineConfig,
  anchorEngine,
  anchorStrings,
  normalizeEngineConfig,
  pageNavigation
} from '../../anchors';
import { AnchorEngineHost } from './components/AnchorEngineHost';

const LOG_SOURCE: string = 'AnchorServiceApplicationCustomizer';
const CONTAINER_ID: string = 'spfx-anchor-service';

/**
 * Properties (ClientSideComponentProperties), e.g. in elements.xml:
 * { "headingLevels": [1,2,3,4,5], "showAnchorIcons": true, "convertInlineAnchors": true,
 *   "scrollToHashOnLoad": true, "contentSelector": "" }
 */
export interface IAnchorServiceApplicationCustomizerProperties {
  /** Anchored levels, array or text "2,3,4". Default: 1–5 */
  headingLevels?: number[] | string;
  showAnchorIcons?: boolean;
  /** Convert {anchor-#id} markers in the content into anchor targets. */
  convertInlineAnchors?: boolean;
  scrollToHashOnLoad?: boolean;
  /** Optional CSS selector of the content area. */
  contentSelector?: string;
}

/**
 * Anchor service as an application customizer: runs on every page of the site, assigns anchors
 * to the headings, converts {anchor-#id} markers and scrolls to the #hash of the URL.
 * The table of contents web part uses the anchors of this extension.
 */
export default class AnchorServiceApplicationCustomizer
  extends BaseApplicationCustomizer<IAnchorServiceApplicationCustomizerProperties> {

  private _container: HTMLDivElement | undefined;
  private _detachNavigation: (() => void) | undefined;

  @override
  public onInit(): Promise<void> {
    anchorStrings.permalinkFormat = strings.PermalinkFormat;

    const config: IAnchorEngineConfig = normalizeEngineConfig(this.properties);
    Log.info(LOG_SOURCE, `Initialized: ${JSON.stringify(config)}`);

    // Provide the configuration to the table of contents web part
    anchorEngine.register(config);

    // Page changes: navigatedEvent is officially available to application customizers
    this._detachNavigation = pageNavigation.attach(this.context);

    // Own invisible container – the engine renders nothing visible
    let container: HTMLDivElement | null = document.getElementById(CONTAINER_ID) as HTMLDivElement | null;
    if (!container) {
      container = document.createElement('div');
      container.id = CONTAINER_ID;
      container.setAttribute(EXCLUDE_ATTR, 'true');
      container.setAttribute('aria-hidden', 'true');
      container.style.display = 'none';
      document.body.appendChild(container);
    }
    this._container = container;

    ReactDom.render(
      React.createElement(AnchorEngineHost, { instanceId: `extension-${this.componentId}`, config: config }),
      container);

    return Promise.resolve();
  }

  @override
  public onDispose(): void {
    if (this._container) {
      ReactDom.unmountComponentAtNode(this._container);
      if (this._container.parentNode) {
        this._container.parentNode.removeChild(this._container);
      }
      this._container = undefined;
    }
    if (this._detachNavigation) {
      this._detachNavigation();
      this._detachNavigation = undefined;
    }
    anchorEngine.unregister();
  }
}
