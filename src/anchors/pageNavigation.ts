import { Log } from '@microsoft/sp-core-library';
import { anchorServiceManager } from './AnchorServiceManager';
import { LOG_SOURCE } from './anchorUtils';
import { getGlobal } from './globalRegistry';

/**
 * Page changes during client-side navigation (SharePoint does not reload the page,
 * our bundle and all module singletons stay in memory).
 *
 * Sources of a page change:
 *  1. context.application.navigatedEvent – if the context provides it
 *     (officially documented for application customizers, not guaranteed in the web part context)
 *  2. popstate (browser back/forward)
 *  3. checkForPageChange() – on every DOM change the engine checks whether the URL
 *     (without hash) has changed. This detects a change even without an event.
 *
 * Pure hash changes (jumping to an anchor) are NOT a page change.
 */

export type PageChangePhase = 'beforeChange' | 'changed';
export type PageChangeListener = (phase: PageChangePhase) => void;

interface ISPEvent {
  add(observer: object, handler: (args?: any) => void): void;
  remove(observer: object, handler: (args?: any) => void): void;
}

/**
 * Key of the current page: path + query without hash and without "Mode=Edit".
 * (SharePoint appends ?Mode=Edit when editing – that is not a page change.)
 */
function currentPageKey(): string {
  const query: string = window.location.search
    .replace(/^\?/, '')
    .split('&')
    .filter((part: string) => !!part && !/^mode=/i.test(part))
    .join('&');
  return (window.location.pathname + (query ? `?${query}` : '')).toLowerCase();
}

class PageNavigationHub {
  private _listeners: PageChangeListener[] = [];
  private _pageKey: string = currentPageKey();
  private _attachCount: number = 0;
  private _zoneOwner: string | undefined;

  /**
   * Call in onInit of a web part (or application customizer).
   * @returns function to detach (call it in onDispose)
   */
  public attach(context: any): () => void {
    const navigatedEvent: ISPEvent | undefined =
      context && context.application && context.application.navigatedEvent;
    const observer: object = {};
    const handler: () => void = () => this.checkForPageChange('navigatedEvent');

    if (navigatedEvent && typeof navigatedEvent.add === 'function') {
      navigatedEvent.add(observer, handler);
    }
    if (this._attachCount === 0) {
      window.addEventListener('popstate', this._onPopState);
    }
    this._attachCount++;

    let attached: boolean = true;
    return () => {
      if (!attached) {
        return;
      }
      attached = false;
      if (navigatedEvent && typeof navigatedEvent.remove === 'function') {
        navigatedEvent.remove(observer, handler);
      }
      this._attachCount--;
      if (this._attachCount === 0) {
        window.removeEventListener('popstate', this._onPopState);
      }
    };
  }

  /** Subscribe: 'beforeChange' (clean up the old page), then 'changed' (build the new page). */
  public subscribe(listener: PageChangeListener): () => void {
    this._listeners.push(listener);
    return () => {
      const i: number = this._listeners.indexOf(listener);
      if (i >= 0) {
        this._listeners.splice(i, 1);
      }
    };
  }

  /**
   * Checks whether the page (path + query, without hash) has changed and then triggers the change.
   * @returns true if a page change was detected
   */
  public checkForPageChange(source: string = 'check'): boolean {
    const key: string = currentPageKey();
    if (key === this._pageKey) {
      return false;
    }
    this._pageKey = key;
    Log.verbose(LOG_SOURCE, `Page changed (${source}): ${key}`);

    this._notify('beforeChange');
    anchorServiceManager.reset();
    this._zoneOwner = undefined;
    this._notify('changed');
    return true;
  }

  // ---- Zone ownership: only one component per page anchors the headings ----

  public claimZone(ownerId: string): boolean {
    if (!this._zoneOwner || this._zoneOwner === ownerId) {
      this._zoneOwner = ownerId;
      return true;
    }
    return false;
  }

  public releaseZone(ownerId: string): void {
    if (this._zoneOwner === ownerId) {
      this._zoneOwner = undefined;
    }
  }

  public get zoneOwner(): string | undefined {
    return this._zoneOwner;
  }

  private _onPopState = (): void => {
    this.checkForPageChange('popstate');
  }

  private _notify(phase: PageChangePhase): void {
    this._listeners.slice().forEach((listener: PageChangeListener) => {
      try {
        listener(phase);
      } catch (e) {
        Log.error(LOG_SOURCE, e instanceof Error ? e : new Error(String(e)));
      }
    });
  }
}

export const pageNavigation: PageNavigationHub = getGlobal('pageNavigation', () => new PageNavigationHub());
