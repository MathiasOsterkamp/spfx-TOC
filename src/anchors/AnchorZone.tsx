import * as React from 'react';
import { IAnchorData } from './IAnchor';
import { anchorServiceManager } from './AnchorServiceManager';
import { INLINE_LINK_STATIC_CLASS } from './anchorStyles';
import { ANCHOR_ATTR, ANCHOR_TARGET_ATTR, applyPermalink, cls, createLinkGlyph, isRTL, toArray } from './anchorUtils';
import { IAnchorCandidate, collectAnchorCandidates } from './anchorCollector';

export interface IAnchorZoneProps {
  /** Unique ID of the registering component (e.g. the web part instance ID). */
  componentId: string;
  /** CSS selector of the elements to anchor, e.g. from getQuerySelector(). */
  anchorElementSelector: string;
  /**
   * Element to search in. If set, the zone renders nothing and works
   * directly on this element. Otherwise the children (wrapper) are searched.
   */
  elementForSearch?: HTMLElement;
  /** Optional filter (true = anchor). */
  filterElement?: (element: HTMLElement) => boolean;
  /** Returns the suggested anchor ID; default: textContent. */
  suggestAnchorId?: (element: HTMLElement, index: number) => string;
  /** Do not show the link icon (IDs are still assigned). */
  hideAnchorIcon?: boolean;
  /**
   * Also register converted inline anchors ({anchor-#id} -> <span data-spfx-inline-anchor>).
   * They keep their ID unchanged and appear with kind = 'inline' in the anchor list.
   */
  includeInline?: boolean;
  /** Changing this value triggers a rescan (e.g. when the content has changed). */
  refreshKey?: number | string;
  children?: React.ReactNode;
}

/**
 * Assigns anchor IDs to headings and shows a permalink icon on hover.
 * (Class component, as in the SharePoint original.)
 */
export class AnchorZone extends React.Component<IAnchorZoneProps> {
  private static _anchorIconElementCache: HTMLAnchorElement | undefined;
  private _wrapperElementRef: React.RefObject<HTMLDivElement> = React.createRef<HTMLDivElement>();

  private static _createAnchorIconElement(): HTMLAnchorElement {
    if (!AnchorZone._anchorIconElementCache) {
      const link: HTMLAnchorElement = document.createElement('a');
      link.className = `${cls().anchor} ${cls().inline}`;
      link.appendChild(createLinkGlyph());
      AnchorZone._anchorIconElementCache = link;
    }
    return AnchorZone._anchorIconElementCache.cloneNode(true) as HTMLAnchorElement;
  }

  private static _isElementAnchorIcon(element: Element): boolean {
    return element.hasAttribute(ANCHOR_ATTR) && element.classList.contains(cls().anchor);
  }

  public componentDidMount(): void {
    this._registerAnchors(false);
  }

  public componentDidUpdate(prevProps: IAnchorZoneProps): void {
    // Only re-register on relevant changes (otherwise an endless loop with onAnchorsUpdated)
    if (
      prevProps.anchorElementSelector !== this.props.anchorElementSelector ||
      prevProps.elementForSearch !== this.props.elementForSearch ||
      prevProps.hideAnchorIcon !== this.props.hideAnchorIcon ||
      prevProps.includeInline !== this.props.includeInline ||
      prevProps.refreshKey !== this.props.refreshKey ||
      prevProps.children !== this.props.children
    ) {
      this._registerAnchors(true);
    }
  }

  public componentWillUnmount(): void {
    anchorServiceManager.raiseAnchorEvent({ action: 'Unregister', uniqueId: this.props.componentId });
  }

  public render(): JSX.Element | null {
    if (this.props.elementForSearch) {
      return null;
    }
    return (
      <div className={cls().wrapper} ref={this._wrapperElementRef}>
        {this.props.children}
      </div>
    );
  }

  private _registerAnchors(isRefresh: boolean): void {
    const root: HTMLElement | null = this.props.elementForSearch || this._wrapperElementRef.current;
    if (!root || (!this.props.anchorElementSelector && !this.props.includeInline)) {
      return;
    }

    // Headings + inline anchors in ONE list (same function as the prediction in edit mode)
    const candidates: IAnchorCandidate[] = collectAnchorCandidates(root, {
      headingSelector: this.props.anchorElementSelector,
      filter: this.props.filterElement,
      includeInline: !!this.props.includeInline
    });

    let headingIndex: number = 0;
    const anchorList: IAnchorData[] = [];
    candidates.forEach((c: IAnchorCandidate) => {
      if (!c.element) {
        return;
      }
      if (c.kind === 'inline') {
        anchorList.push({
          suggestedAnchorId: c.suggestedAnchorId,
          anchorTargetElement: c.element,
          kind: 'inline',
          preserveId: true,
          level: 0,
          onReceiveAnchorId: (target: HTMLElement, anchorId: string) =>
            this._addAnchorIcon(target, anchorId, anchorId, true)
        });
        return;
      }
      const suggested: string = this.props.suggestAnchorId
        ? this.props.suggestAnchorId(c.element, headingIndex)
        : c.suggestedAnchorId;
      headingIndex++;
      anchorList.push({
        suggestedAnchorId: suggested,
        anchorTargetElement: c.element,
        kind: 'heading',
        level: c.level,
        onReceiveAnchorId: (target: HTMLElement, anchorId: string) => this._addAnchorIcon(target, anchorId, suggested),
        onDisposeAnchorId: this._removeAnchorIcon
      });
    });

    anchorServiceManager.raiseAnchorEvent({
      action: isRefresh ? 'Refresh' : 'Register',
      uniqueId: this.props.componentId,
      anchorList: anchorList
    });
  }

  private _addAnchorIcon(target: HTMLElement, anchorId: string, label: string, inline: boolean = false): void {
    target.setAttribute(ANCHOR_TARGET_ATTR, anchorId);
    if (this.props.hideAnchorIcon) {
      target.setAttribute('id', anchorId);
      return;
    }
    const link: HTMLElement = inline ? this._insertInlineIcon(target) : this._insertAnchorIcon(target);
    try {
      applyPermalink(target, link, anchorId, label, true);
    } catch (e) {
      // intentionally ignored (as in the original)
    }
  }

  /** Placement for inline targets: the icon sits in the text flow inside the target span. */
  private _insertInlineIcon(target: HTMLElement): HTMLElement {
    const existing: Element | null = target.querySelector(`a[${ANCHOR_ATTR}]`);
    if (existing) {
      return existing as HTMLElement;
    }
    const link: HTMLAnchorElement = AnchorZone._createAnchorIconElement();
    link.className = `${link.className} ${cls().inlineLink} ${INLINE_LINK_STATIC_CLASS}`;
    target.appendChild(link);
    return link;
  }

  private _insertAnchorIcon(target: HTMLElement): HTMLElement {
    const next: Element | null = target.nextElementSibling;
    if (next && AnchorZone._isElementAnchorIcon(next)) {
      return next as HTMLElement;
    }

    const link: HTMLAnchorElement = AnchorZone._createAnchorIconElement();

    // Wrap heading + icon so that :hover covers both
    const wrapper: HTMLDivElement = document.createElement('div');
    if (target.parentElement) {
      target.parentElement.insertBefore(wrapper, target);
    }
    wrapper.appendChild(target);
    wrapper.appendChild(link);
    wrapper.className = `anchor-${target.tagName.toLowerCase()} ${cls().headingWrapper}`;

    const offset: number = (parseInt(getComputedStyle(target).lineHeight || '0', 10) - link.offsetHeight) / 2;
    const placeholder: HTMLSpanElement = document.createElement('span');
    placeholder.className = cls().anchor;
    this._insertPlaceholder(placeholder, target);

    wrapper.onmouseover = () => this._setAnchorPosition(offset, link, placeholder);
    link.onfocus = () => this._setAnchorPosition(offset, link, placeholder);

    return link;
  }

  /** Position the icon right after the end of the heading text. */
  private _setAnchorPosition(offset: number, link: HTMLElement, placeholder: HTMLElement): void {
    const left: number = isRTL()
      ? placeholder.offsetLeft - link.offsetWidth - 5
      : placeholder.offsetLeft;
    link.style.left = `${left}px`;
    link.style.top = `${placeholder.offsetTop + (isNaN(offset) ? 0 : offset)}px`;
  }

  /** Insert the placeholder before the first <br>, otherwise at the end. */
  private _insertPlaceholder(placeholder: HTMLElement, target: HTMLElement): void {
    const children: HTMLElement[] = toArray<HTMLElement>(target.children as any);
    for (let i: number = 0; i < children.length; i++) {
      if (children[i].tagName === 'BR') {
        target.insertBefore(placeholder, children[i]);
        return;
      }
    }
    target.appendChild(placeholder);
  }

  private _removeAnchorIcon = (target: HTMLElement): void => {
    const last: Element | null = target.lastElementChild;
    if (last && AnchorZone._isElementAnchorIcon(last)) {
      target.removeChild(last);
    }
  }
}
