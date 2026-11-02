import { Log, Validate } from '@microsoft/sp-core-library';
import { AnchorKind, IAnchorData, IAnchorEvent, ISortedAnchor, IScrollTask, AnchorsUpdatedCallback } from './IAnchor';
import { assignAnchorIds, sanitizeAnchorId } from './anchorIds';
import { LOG_SOURCE, getHeadingLevel } from './anchorUtils';
import { getGlobal } from './globalRegistry';

// ID allocation lives in anchorIds.ts (shared with the prediction used in edit mode)
export { INVALID_ANCHOR_ID_CHARS, allocateAnchorId, handleAnchorCollision } from './anchorIds';

// ---------------------------------------------------------------------------
// AnchorService: registration, ID allocation, scrolling
// ---------------------------------------------------------------------------

export class AnchorService {
  private _componentIdToAnchorDataMap: Map<string, IAnchorData[]> = new Map<string, IAnchorData[]>();
  private _anchorIdToAnchorDataMap: Map<string, IAnchorData> = new Map<string, IAnchorData>();
  private _onAnchorsUpdated: AnchorsUpdatedCallback;

  constructor(onAnchorsUpdated: AnchorsUpdatedCallback) {
    this._onAnchorsUpdated = onAnchorsUpdated;
  }

  /** Text -> valid, URL-safe ID (max. 128 characters). */
  public static sanitizeAnchorId(id: string): string {
    return sanitizeAnchorId(id);
  }

  private static _compareAnchorLocation(a: IAnchorData, b: IAnchorData): number {
    // tslint:disable-next-line:no-bitwise
    return (a.anchorTargetElement.compareDocumentPosition(b.anchorTargetElement) & Node.DOCUMENT_POSITION_FOLLOWING)
      ? -1
      : 1;
  }

  public getElementByAnchorId(anchorId: string): HTMLElement | undefined {
    const data: IAnchorData | undefined = this._anchorIdToAnchorDataMap.get(anchorId);
    return data ? data.anchorTargetElement : undefined;
  }

  public scrollTo(anchorId: string): boolean {
    const data: IAnchorData | undefined = this._anchorIdToAnchorDataMap.get(anchorId);
    if (data) {
      data.anchorTargetElement.scrollIntoView();
      if (data.onScrollToAnchorId) {
        data.onScrollToAnchorId(data.anchorTargetElement, anchorId);
      }
      return true;
    }
    // Fallback: any element with this ID
    const element: HTMLElement | null = typeof document !== 'undefined' ? document.getElementById(anchorId) : null;
    if (element) {
      element.scrollIntoView();
      return true;
    }
    return false;
  }

  public register(componentId: string, anchorList: IAnchorData[]): void {
    Validate.isNonemptyString(componentId, 'uniqueId');
    Validate.isNotNullOrUndefined(anchorList, 'anchorList');
    this._componentIdToAnchorDataMap.set(componentId, anchorList.map((a: IAnchorData) => ({ ...a, anchorId: undefined })));
    Log.verbose(LOG_SOURCE, `Component '${componentId}' registered with ${anchorList.length} anchors`);
  }

  public unregister(componentId: string): void {
    const list: IAnchorData[] | undefined = this._componentIdToAnchorDataMap.get(componentId);
    if (!list) {
      // Expected after a page change (reset) – therefore no longer an error
      Log.verbose(LOG_SOURCE, `Trying to unregister a non-exists component '${componentId}'`);
      return;
    }
    list.forEach((a: IAnchorData) => {
      if (a.anchorId) {
        if (a.onDisposeAnchorId) {
          a.onDisposeAnchorId(a.anchorTargetElement, a.anchorId);
        }
        this._anchorIdToAnchorDataMap.delete(a.anchorId);
      }
    });
    this._componentIdToAnchorDataMap.delete(componentId);
    Log.verbose(LOG_SOURCE, `Successfully unregistered component '${componentId}'`);
  }

  public finishRegistration(): void {
    Log.verbose(LOG_SOURCE, 'finishRegistration invoked');

    let all: IAnchorData[] = [];
    this._componentIdToAnchorDataMap.forEach((list: IAnchorData[]) => { all = all.concat(list); });
    all.sort(AnchorService._compareAnchorLocation);

    this._anchorIdToAnchorDataMap.clear();
    // Fixed inline IDs first, then headings – identical to the prediction (anchorIds.ts)
    const ids: string[] = assignAnchorIds(all);

    const sorted: ISortedAnchor[] = all.map((data: IAnchorData, index: number) => {
      const id: string = ids[index];
      data.anchorId = id;
      this._anchorIdToAnchorDataMap.set(id, data);
      if (data.onReceiveAnchorId) {
        data.onReceiveAnchorId(data.anchorTargetElement, id);
      }
      Log.verbose(LOG_SOURCE, `Anchor ID '${id}' allocated`);
      const kind: AnchorKind = data.kind || 'heading';
      return {
        anchorText: data.suggestedAnchorId,
        anchoredElementTagName: data.anchorTargetElement.tagName,
        anchorId: id,
        level: data.level !== undefined ? data.level : (kind === 'inline' ? 0 : getHeadingLevel(data.anchorTargetElement)),
        kind: kind
      };
    });

    this._onAnchorsUpdated(sorted);
  }

}

// ---------------------------------------------------------------------------
// Queue: waits until all reserved components have registered
// ---------------------------------------------------------------------------

class AnchorServiceQueue {
  private _pendingRegistrations: Map<string, boolean> = new Map<string, boolean>();
  private _anchorService: AnchorService;
  private _resolveScroll: (() => void) | undefined;

  constructor(onAnchorsUpdated: AnchorsUpdatedCallback) {
    this._anchorService = new AnchorService(onAnchorsUpdated);
  }

  public initialize(events: IAnchorEvent[]): void {
    events.forEach((e: IAnchorEvent) => this.handleAnchorEvent(e));
  }

  public scrollTo(anchorId: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (e: Error) => void) => {
      const doScroll: () => void = () => {
        if (this._anchorService.scrollTo(anchorId)) {
          resolve();
        } else {
          reject(new Error('Failed to find anchor to scroll'));
        }
      };
      if (this._pendingRegistrations.size === 0) {
        doScroll();
      } else {
        this._resolveScroll = doScroll;
      }
    });
  }

  public handleAnchorEvent(e: IAnchorEvent): void {
    switch (e.action) {
      case 'ReserveRegistration':
        this._pendingRegistrations.set(e.uniqueId, true);
        break;
      case 'Register':
      case 'Refresh':
        this._anchorService.register(e.uniqueId, e.anchorList || []);
        this._pendingRegistrations.delete(e.uniqueId);
        this._checkFinishRegistration();
        break;
      case 'Unregister':
        this._anchorService.unregister(e.uniqueId);
        this._pendingRegistrations.delete(e.uniqueId);
        this._checkFinishRegistration();
        break;
    }
  }

  private _checkFinishRegistration(): void {
    if (this._pendingRegistrations.size === 0) {
      this._anchorService.finishRegistration();
      if (this._resolveScroll) {
        const fn: () => void = this._resolveScroll;
        this._resolveScroll = undefined;
        fn();
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Manager: public singleton API
// ---------------------------------------------------------------------------

export class AnchorServiceManager {
  private _anchorIdToScroll: string | undefined;
  private _lastAnchors: ISortedAnchor[] | undefined;
  private _listeners: AnchorsUpdatedCallback[] = [];
  private _queue: AnchorServiceQueue;
  private _pendingCancels: (() => void)[] = [];

  constructor() {
    this._queue = new AnchorServiceQueue(this._handleAnchorsUpdated);
  }

  /**
   * Resets everything for a new page (SPA navigation between pages).
   * Registrations, pending scrolls and the cached anchor list are discarded.
   * Subscribers clear their own list via pageNavigation ('beforeChange'), see useSortedAnchors.
   */
  public reset(): void {
    Log.verbose(LOG_SOURCE, 'reset (page changed)');
    this._pendingCancels.splice(0).forEach((cancel: () => void) => cancel());
    this._anchorIdToScroll = undefined;
    this._queue = new AnchorServiceQueue(this._handleAnchorsUpdated);
    this._lastAnchors = undefined;
  }

  private _handleAnchorsUpdated = (anchors: ISortedAnchor[]): void => {
    this._lastAnchors = anchors;
    this._listeners.slice().forEach((cb: AnchorsUpdatedCallback) => cb(anchors));
  }

  /** Last known sorted anchor list (or empty). */
  public get anchors(): ISortedAnchor[] {
    return this._lastAnchors || [];
  }

  /**
   * Replacement for SharePoint's sticky event "anchorEvent".
   * Immediately delivers the last state, if any. Returns an unsubscribe function.
   */
  public onAnchorsUpdated(callback: AnchorsUpdatedCallback): () => void {
    this._listeners.push(callback);
    if (this._lastAnchors) {
      callback(this._lastAnchors);
    }
    return () => {
      const i: number = this._listeners.indexOf(callback);
      if (i >= 0) {
        this._listeners.splice(i, 1);
      }
    };
  }

  /** Announces a later registration; scrolling waits until it has happened. */
  public reserve(uniqueId: string): void {
    this._queue.handleAnchorEvent({ action: 'ReserveRegistration', uniqueId: uniqueId });
  }

  public raiseAnchorEvent(e: IAnchorEvent): void {
    this._queue.handleAnchorEvent(e);
  }

  /**
   * Schedules scrolling to an anchor. Scrolls on execute() or once the timeout (ms) has elapsed.
   * timeout = -1: only via execute().
   */
  public scrollTo(anchorId: string, timeout: number = 5000): IScrollTask | undefined {
    if (!anchorId && timeout < -1) {
      return undefined;
    }
    this._anchorIdToScroll = anchorId;

    let execute: () => void = () => { /* assigned below */ };
    let cancel: () => void = () => { /* assigned below */ };

    const racers: Promise<string>[] = [];
    if (timeout !== -1) {
      racers.push(new Promise<string>((resolve: (v: string) => void) => {
        setTimeout(() => resolve('timeout'), timeout);
      }));
    }
    racers.push(new Promise<string>((resolve: (v: string) => void, reject: (v: string) => void) => {
      execute = () => resolve('execute');
      cancel = () => reject('cancel');
    }));
    const cancelRef: () => void = cancel;
    this._pendingCancels.push(cancelRef);
    const forget: () => void = () => {
      const i: number = this._pendingCancels.indexOf(cancelRef);
      if (i >= 0) {
        this._pendingCancels.splice(i, 1);
      }
    };

    const scrollPromise: Promise<void> = Promise.race(racers)
      .then((trigger: string) => { forget(); return trigger; }, (reason: any) => { forget(); throw reason; })
      .then((trigger: string) => this._scroll().then(
        () => {
          Log.verbose(LOG_SOURCE, trigger === 'timeout' ? 'Scroll triggered by timeout' : 'Scroll succeeded');
        },
        (err: any) => {
          Log.error(LOG_SOURCE, err instanceof Error ? err : new Error(`FailedToScroll (${trigger}): ${err}`));
        }
      ))
      .catch((reason: any) => {
        Log.verbose(LOG_SOURCE, `CancelScroll: ${reason}`);
      });

    return { execute: execute, cancel: cancel, scrollPromise: scrollPromise };
  }

  private _scroll(): Promise<void> {
    if (!this._anchorIdToScroll) {
      return Promise.reject('AnchorId is undefined, multiple scroll Task planned') as Promise<void>;
    }
    const p: Promise<void> = this._queue.scrollTo(this._anchorIdToScroll);
    this._anchorIdToScroll = undefined;
    return p;
  }
}

/** A single manager per page – even across bundles (extension + web part). */
export const anchorServiceManager: AnchorServiceManager =
  getGlobal('anchorServiceManager', () => new AnchorServiceManager());
