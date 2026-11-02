import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { ITheme } from '@fluentui/react/lib/Styling';

import { ISortedAnchor, IScrollTask } from './IAnchor';
import { anchorServiceManager } from './AnchorServiceManager';
import { AnchorZone } from './AnchorZone';
import { cleanupAnchorIconsInTree } from './cleanupAnchorIconsInTree';
import { convertInlineAnchors, revertInlineAnchors } from './inlineAnchors';
import { PageChangePhase, pageNavigation } from './pageNavigation';
import { ANCHOR_TARGET_ATTR, EXCLUDE_ATTR, cls, isInDocument } from './anchorUtils';
import { isExcludedElement, resolveContentRoot } from './contentRoot';
import { IAnchorCandidate, collectAnchorCandidates } from './anchorCollector';
import { isInContentEditable } from './anchorEngine';

export { DEFAULT_CONTENT_SELECTORS } from './contentRoot';


export interface IPageAnchorsProps {
  /** Instance ID of the hosting component. */
  instanceId: string;
  displayMode: DisplayMode;
  /** CSS selector of the headings (e.g. from getQuerySelector). Empty = do not anchor headings. */
  headingSelector: string;
  /** Optional selector of the content area. Empty = automatic. */
  contentSelector?: string;
  convertInlineAnchors: boolean;
  showAnchorIcons: boolean;
  scrollToHashOnLoad: boolean;
  /** Additional filter for headings (true = anchor). */
  additionalFilter?: (element: HTMLElement) => boolean;
  /** Theme for the anchor icons (default: page theme). */
  theme?: ITheme;
}

interface IPageAnchorsState {
  root: HTMLElement | null;
  ownsZone: boolean;
  scanVersion: number;
}

/**
 * Anchor engine without any visible output. Assigns anchors to the headings of the page,
 * converts {anchor-#id} markers and scrolls to the URL hash.
 *
 * Robust against:
 *  - page changes via SPA navigation (navigatedEvent / popstate / URL check):
 *    clean up the old page, reset the service, build the new page, scroll to the new hash
 *  - switching view <-> edit: in edit mode everything is removed and the
 *    markers are restored; switching back rebuilds everything
 *  - a re-rendered canvas (e.g. after publishing): the content area is resolved again
 *  - headings that are loaded later or re-rendered (text web parts)
 */
export class PageAnchors extends React.Component<IPageAnchorsProps, IPageAnchorsState> {
  public state: IPageAnchorsState = { root: null, ownsZone: false, scanVersion: 0 };

  private _observer: MutationObserver | undefined;
  private _scanTimer: number | undefined;
  private _lastSignature: string = '';
  private _hashTask: IScrollTask | undefined;
  private _hashAnchorId: string | undefined;
  private _isMounted: boolean = false;
  private _unsubscribePage: (() => void) | undefined;
  private _unsubscribeAnchors: (() => void) | undefined;

  public componentDidMount(): void {
    this._isMounted = true;
    this._unsubscribePage = pageNavigation.subscribe(this._onPageChange);
    this._unsubscribeAnchors = anchorServiceManager.onAnchorsUpdated(this._onAnchorsUpdated);

    if (this._isDisplay(this.props)) {
      this._planHashScroll();
      this._activate();
    }
  }

  public componentDidUpdate(prevProps: IPageAnchorsProps, prevState: IPageAnchorsState): void {
    // After a re-registration (markers set) remember the signature again
    if (prevState.scanVersion !== this.state.scanVersion && this.state.root) {
      this._lastSignature = this._computeSignature(this.state.root);
    }

    const wasDisplay: boolean = this._isDisplay(prevProps);
    const isDisplay: boolean = this._isDisplay(this.props);
    const p: IPageAnchorsProps = this.props;

    if (wasDisplay && !isDisplay) {
      // View -> edit: remove icons/wrappers, restore markers
      this._cancelHashScroll();
      this._deactivate(true);
    } else if (!wasDisplay && isDisplay) {
      // Edit -> view: the canvas may have been re-rendered -> rebuild completely
      this._activate();
    } else if (isDisplay && (
      prevProps.headingSelector !== p.headingSelector ||
      prevProps.contentSelector !== p.contentSelector ||
      prevProps.convertInlineAnchors !== p.convertInlineAnchors ||
      prevProps.showAnchorIcons !== p.showAnchorIcons ||
      prevProps.theme !== p.theme)) {
      this._reactivate();
    }
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    this._cancelHashScroll();
    this._deactivate(false);
    if (this._unsubscribePage) {
      this._unsubscribePage();
    }
    if (this._unsubscribeAnchors) {
      this._unsubscribeAnchors();
    }
  }

  public render(): JSX.Element {
    const { headingSelector, showAnchorIcons, convertInlineAnchors: includeInline, instanceId, theme } = this.props;
    const { root, ownsZone, scanVersion } = this.state;
    const attrs: { [key: string]: string } = { [EXCLUDE_ATTR]: 'true' };

    return (
      <div className={cls(theme).hidden} aria-hidden='true' {...attrs}>
        {this._isDisplay(this.props) && ownsZone && root && (headingSelector || includeInline)
          ? <AnchorZone
            key={showAnchorIcons ? 'icons' : 'noicons'}
            componentId={instanceId}
            theme={theme}
            anchorElementSelector={headingSelector}
            elementForSearch={root}
            filterElement={this._filterElement}
            hideAnchorIcon={!showAnchorIcons}
            includeInline={includeInline}
            refreshKey={scanVersion} />
          : null}
      </div>
    );
  }

  // -------------------------------------------------------------------------
  // Activate / deactivate
  // -------------------------------------------------------------------------

  private _isDisplay(props: IPageAnchorsProps): boolean {
    return props.displayMode === DisplayMode.Read;
  }

  private _activate(): void {
    const root: HTMLElement = this._resolveRoot();
    const ownsZone: boolean = pageNavigation.claimZone(this.props.instanceId);

    if (ownsZone) {
      this._convertInline(root);
    }
    this._lastSignature = ownsZone ? this._computeSignature(root) : '';
    this._startObserver();

    this.setState((prev: IPageAnchorsState) => ({
      root: root,
      ownsZone: ownsZone,
      scanVersion: prev.scanVersion + 1
    }));

    // In case the hash points to an inline anchor (or no headings are anchored)
    this._tryHashScroll(anchorServiceManager.anchors);
  }

  /**
   * @param updateState false on unmount (setState is no longer allowed)
   */
  private _deactivate(updateState: boolean): void {
    this._stopObserver();
    const { root, ownsZone } = this.state;

    // Only clean up if the old content is still in the document (it is gone after a page change)
    if (root && ownsZone && isInDocument(root)) {
      cleanupAnchorIconsInTree(root);
      revertInlineAnchors(root);
    }
    pageNavigation.releaseZone(this.props.instanceId);

    if (updateState) {
      this.setState({ root: null, ownsZone: false });
    }
  }

  private _reactivate(): void {
    this._deactivate(true);
    this._activate();
  }

  private _onPageChange = (phase: PageChangePhase): void => {
    if (!this._isMounted) {
      return;
    }
    if (phase === 'beforeChange') {
      this._cancelHashScroll();
      this._deactivate(true);
    } else if (this._isDisplay(this.props)) {
      this._planHashScroll();
      this._activate();
    }
  }

  private _resolveRoot(): HTMLElement {
    return resolveContentRoot(this.props.contentSelector);
  }

  private _filterElement = (element: HTMLElement): boolean => {
    // Never modify our own areas and never anything inside an editable editor –
    // protection in case edit mode (e.g. in the extension) is detected too late.
    if (isExcludedElement(element) || isInContentEditable(element)) {
      return false;
    }
    return this.props.additionalFilter ? this.props.additionalFilter(element) : true;
  }

  private _convertInline(root: HTMLElement): void {
    if (this.props.convertInlineAnchors) {
      // in the whole content area (paragraphs, lists, tables, headings …)
      convertInlineAnchors(root, undefined, this._filterElement);
    }
  }

  // -------------------------------------------------------------------------
  // Observe the DOM (the whole body, because the canvas can be replaced entirely)
  // -------------------------------------------------------------------------

  private _startObserver(): void {
    if (this._observer || typeof MutationObserver === 'undefined') {
      return;
    }
    this._observer = new MutationObserver(this._onMutation);
    this._observer.observe(document.body, { childList: true, subtree: true, characterData: true });
  }

  private _stopObserver(): void {
    if (this._observer) {
      this._observer.disconnect();
      this._observer = undefined;
    }
    if (this._scanTimer !== undefined) {
      window.clearTimeout(this._scanTimer);
      this._scanTimer = undefined;
    }
  }

  private _onMutation = (): void => {
    if (this._scanTimer !== undefined) {
      window.clearTimeout(this._scanTimer);
    }
    this._scanTimer = window.setTimeout(this._scan, 300);
  }

  /** Public for testing; otherwise called (debounced) by the MutationObserver. */
  public _scan = (): void => {
    this._scanTimer = undefined;
    if (!this._isMounted || !this._isDisplay(this.props)) {
      return;
    }

    // 1. Page change without an event? -> handled via _onPageChange
    if (pageNavigation.checkForPageChange('mutation')) {
      return;
    }

    // 2. Canvas replaced (e.g. after editing/publishing) -> rebuild
    const root: HTMLElement | null = this.state.root;
    if (!root || !isInDocument(root) || this._resolveRoot() !== root) {
      this._reactivate();
      return;
    }

    // 3. Zone owner is gone -> take over
    if (!this.state.ownsZone) {
      if (!pageNavigation.zoneOwner && pageNavigation.claimZone(this.props.instanceId)) {
        this._reactivate();
      }
      return;
    }

    // 4. New markers / headings
    this._convertInline(root);
    this._tryHashScroll(anchorServiceManager.anchors);

    const signature: string = this._computeSignature(root);
    if (signature !== this._lastSignature) {
      this._lastSignature = signature;
      this.setState((prev: IPageAnchorsState) => ({ scanVersion: prev.scanVersion + 1 }));
    }
  }

  /**
   * Signature of all anchor candidates (headings + inline anchors): kind, text/ID and whether already anchored.
   * Re-rendered headings or new markers are not marked yet -> detected.
   * Our own DOM changes (wrappers, icons) do not change the signature.
   */
  private _computeSignature(root: HTMLElement): string {
    return collectAnchorCandidates(root, {
      headingSelector: this.props.headingSelector,
      filter: this._filterElement,
      includeInline: this.props.convertInlineAnchors
    })
      .map((c: IAnchorCandidate) =>
        `${c.kind}:${c.suggestedAnchorId}:${c.element && c.element.hasAttribute(ANCHOR_TARGET_ATTR) ? 1 : 0}`)
      .join('\n');
  }

  // -------------------------------------------------------------------------
  // Scroll to the hash on load / after a page change
  // -------------------------------------------------------------------------

  private _onAnchorsUpdated = (anchors: ISortedAnchor[]): void => {
    if (this._isMounted) {
      this._tryHashScroll(anchors);
    }
  }

  private _planHashScroll(): void {
    this._cancelHashScroll();
    if (!this.props.scrollToHashOnLoad || !window.location.hash || window.location.hash.length < 2) {
      return;
    }
    let id: string = window.location.hash.substring(1);
    try {
      id = decodeURIComponent(id);
    } catch (err) {
      // use the raw value
    }
    this._hashAnchorId = id;
    // Scroll after 5 s at the latest, even if the anchor has not been reported by then.
    this._hashTask = anchorServiceManager.scrollTo(id, 5000);
  }

  private _cancelHashScroll(): void {
    if (this._hashTask) {
      this._hashTask.cancel();
    }
    this._hashTask = undefined;
    this._hashAnchorId = undefined;
  }

  private _tryHashScroll(anchors: ISortedAnchor[]): void {
    if (!this._hashTask || !this._hashAnchorId) {
      return;
    }
    const id: string = this._hashAnchorId;
    const found: boolean = anchors.some((a: ISortedAnchor) => a.anchorId === id) || !!document.getElementById(id);
    if (found) {
      const task: IScrollTask = this._hashTask;
      this._hashTask = undefined;
      this._hashAnchorId = undefined;
      task.execute();
    }
  }
}
