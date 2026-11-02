import * as React from 'react';
import { ISortedAnchor } from './IAnchor';
import { anchorServiceManager } from './AnchorServiceManager';
import { PageChangePhase, pageNavigation } from './pageNavigation';
import { DEFAULT_ENGINE_CONFIG, IAnchorEngineConfig, anchorEngine } from './anchorEngine';
import { getAvailableAnchors } from './anchorCatalog';
import { defaultTocFilter } from './elementFilters';

export interface IAnchorsState {
  anchors: ISortedAnchor[];
  /** Counts the updates since the last page change (0 = none yet). */
  version: number;
}

/** Current sorted anchor list. Cleared automatically on page changes. */
export function useSortedAnchors(): IAnchorsState {
  const [state, setState] = React.useState<IAnchorsState>({ anchors: anchorServiceManager.anchors, version: 0 });

  React.useEffect(() => {
    const unsubscribeAnchors: () => void = anchorServiceManager.onAnchorsUpdated((anchors: ISortedAnchor[]) => {
      setState((prev: IAnchorsState) => ({ anchors: anchors, version: prev.version + 1 }));
    });
    // Clear when leaving the page – NOT on 'changed', because by then the engine
    // may already have reported the anchors of the new page (order of subscribers).
    const unsubscribePage: () => void = pageNavigation.subscribe((phase: PageChangePhase) => {
      if (phase === 'beforeChange') {
        setState({ anchors: [], version: 0 });
      }
    });
    return () => {
      unsubscribeAnchors();
      unsubscribePage();
    };
  }, []);

  return state;
}

/** Counter that increases on every page change (e.g. as a useEffect dependency). */
export function usePageVersion(): number {
  const [version, setVersion] = React.useState(0);
  React.useEffect(() => pageNavigation.subscribe((phase: PageChangePhase) => {
    if (phase === 'changed') {
      setVersion((v: number) => v + 1);
    }
  }), []);
  return version;
}

/**
 * Calls the callback (debounced) on DOM changes in the document while enabled.
 * Also checks whether a page change happened without an event.
 */
export function useDomChanges(enabled: boolean, callback: () => void, delay: number = 300): void {
  const callbackRef: React.MutableRefObject<() => void> = React.useRef(callback);
  callbackRef.current = callback;

  React.useEffect(() => {
    if (!enabled || typeof MutationObserver === 'undefined') {
      return undefined;
    }
    let timer: number | undefined;
    const observer: MutationObserver = new MutationObserver(() => {
      if (timer !== undefined) {
        window.clearTimeout(timer);
      }
      timer = window.setTimeout(() => {
        timer = undefined;
        pageNavigation.checkForPageChange('mutation');
        callbackRef.current();
      }, delay);
    });
    observer.observe(document.body, { childList: true, subtree: true, characterData: true });
    return () => {
      observer.disconnect();
      if (timer !== undefined) {
        window.clearTimeout(timer);
      }
    };
  }, [enabled, delay]);
}

// ---------------------------------------------------------------------------
// Prediction in edit mode (the engine is inactive there)
// ---------------------------------------------------------------------------

export interface IPredictedAnchorsState {
  anchors: ISortedAnchor[];
  /** true once the page has been evaluated at least once. */
  loaded: boolean;
}

const sameAnchors = (a: ISortedAnchor[], b: ISortedAnchor[]): boolean =>
  a.length === b.length && a.every((x: ISortedAnchor, i: number) =>
    x.anchorId === b[i].anchorId && x.anchorText === b[i].anchorText && x.level === b[i].level && x.kind === b[i].kind);

/**
 * Anchors that will exist in view mode – using the anchor service configuration
 * (levels, content area, inline anchors) and the same filters. Updates on changes in the editor.
 * Used for the table of contents preview and the anchor dropdown in the property pane.
 */
export function usePredictedAnchors(enabled: boolean = true): IPredictedAnchorsState {
  const [state, setState] = React.useState<IPredictedAnchorsState>({ anchors: [], loaded: false });

  const refresh = React.useCallback((): void => {
    const config: IAnchorEngineConfig = anchorEngine.config || DEFAULT_ENGINE_CONFIG;
    const next: ISortedAnchor[] = getAvailableAnchors({
      headingSelector: anchorEngine.headingSelector,
      filter: defaultTocFilter,
      contentSelector: config.contentSelector,
      includeInline: config.convertInlineAnchors
    });
    setState((prev: IPredictedAnchorsState) =>
      prev.loaded && sameAnchors(prev.anchors, next) ? prev : { anchors: next, loaded: true });
  }, []);

  // first run (slightly delayed, SharePoint renders content later) + configuration changes
  React.useEffect(() => {
    if (!enabled) {
      return undefined;
    }
    const timeoutId: number = window.setTimeout(refresh, 0);
    const unsubscribe: () => void = anchorEngine.subscribe(refresh);
    return () => {
      window.clearTimeout(timeoutId);
      unsubscribe();
    };
  }, [enabled, refresh]);

  useDomChanges(enabled, refresh);
  return state;
}
