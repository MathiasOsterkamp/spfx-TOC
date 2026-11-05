import * as React from 'react';
import { ITheme } from '@fluentui/react/lib/Styling';
import { IAnchorData } from './IAnchor';
import { anchorServiceManager } from './AnchorServiceManager';
import { ANCHOR_STATIC_CLASS, IAnchorClassNames, INLINE_LINK_STATIC_CLASS } from './anchorStyles';
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
  /** Theme for the icon colors (default: page theme). */
  theme?: ITheme;
  children?: React.ReactNode;
}

/**
 * Assigns anchor IDs to headings and shows a permalink icon on hover.
 * (Class component, as in the SharePoint original.)
 */
export class AnchorZone extends React.Component<IAnchorZoneProps> {
  private _anchorIconElementCache: HTMLAnchorElement | undefined;
  private _wrapperElementRef: React.RefObject<HTMLDivElement> = React.createRef<HTMLDivElement>();

  private static _isElementAnchorIcon(element: Element): boolean {
    return element.hasAttribute(ANCHOR_ATTR) && element.classList.contains(ANCHOR_STATIC_CLASS);
  }

  private _createAnchorIconElement(): HTMLAnchorElement {
    if (!this._anchorIconElementCache) {
      const link: HTMLAnchorElement = document.createElement('a');
      link.className = `${cls(this.props.theme).anchor} ${cls(this.props.theme).inline} ${ANCHOR_STATIC_CLASS}`;
      link.appendChild(createLinkGlyph(this.props.theme));
      this._anchorIconElementCache = link;
    }
    return this._anchorIconElementCache.cloneNode(true) as HTMLAnchorElement;
  }

  public componentDidMount(): void {
    this._registerAnchors(false);
  }

  public componentDidUpdate(prevProps: IAnchorZoneProps): void {
    if (prevProps.theme !== this.props.theme) {
      this._anchorIconElementCache = undefined;
    }
    // Only re-register on relevant changes (otherwise an endless loop with onAnchorsUpdated)
    if (
      prevProps.anchorElementSelector !== this.props.anchorElementSelector ||
      prevProps.elementForSearch !== this.props.elementForSearch ||
      prevProps.hideAnchorIcon !== this.props.hideAnchorIcon ||
      prevProps.includeInline !== this.props.includeInline ||
      prevProps.refreshKey !== this.props.refreshKey ||
      prevProps.theme !== this.props.theme ||
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
      <div className={cls(this.props.theme).wrapper} ref={this._wrapperElementRef}>
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
    const root: HTMLElement | null = this.props.elementForSearch || this._wrapperElementRef.current;
    if (!inline && root && !root.contains(target)) {
      // Page title etc. is rendered and removed by SharePoint (React): only the ID, never move/wrap the element
      target.setAttribute('id', anchorId);
      return;
    }
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
    const link: HTMLAnchorElement = this._createAnchorIconElement();
    // Without the "inline" margin class: the non-breaking space provides the gap, the icon itself
    // takes no space until it is shown on hover/focus (see inlineLink).
    const c: IAnchorClassNames = cls(this.props.theme);
    link.className = `${c.anchor} ${c.inlineLink} ${ANCHOR_STATIC_CLASS} ${INLINE_LINK_STATIC_CLASS}`;
    // Non-breaking space: stays visible and gives the otherwise empty target a hoverable area
    target.appendChild(document.createTextNode('\u00a0'));
    target.appendChild(link);
    return link;
  }

  private _insertAnchorIcon(target: HTMLElement): HTMLElement {
    const next: Element | null = target.nextElementSibling;
    if (next && AnchorZone._isElementAnchorIcon(next)) {
      return next as HTMLElement;
    }

    const link: HTMLAnchorElement = this._createAnchorIconElement();

    // Wrap heading + icon so that :hover covers both
    const wrapper: HTMLDivElement = document.createElement('div');
    if (target.parentElement) {
      target.parentElement.insertBefore(wrapper, target);
    }
    wrapper.appendChild(target);
    wrapper.appendChild(link);
    wrapper.className = `anchor-${target.tagName.toLowerCase()} ${cls(this.props.theme).headingWrapper}`;

    const placeholder: HTMLSpanElement = document.createElement('span');
    placeholder.className = `${cls(this.props.theme).anchor} ${ANCHOR_STATIC_CLASS}`;
    // In text flow (not absolute) so its rect reflects the line box at the end of the text
    placeholder.style.position = 'static';
    this._insertPlaceholder(placeholder, target);

    wrapper.onmouseover = () => this._setAnchorPosition(wrapper, link, placeholder);
    link.onfocus = () => this._setAnchorPosition(wrapper, link, placeholder);

    return link;
  }

  /**
   * Position the icon right after the end of the heading text.
   * Uses bounding rects relative to the wrapper, so it does not matter which element is the offsetParent
   * of the placeholder (headings are often positioned themselves and have margins).
   */
  private _setAnchorPosition(wrapper: HTMLElement, link: HTMLElement, placeholder: HTMLElement): void {
    const wrapperRect: ClientRect = wrapper.getBoundingClientRect();
    const placeholderRect: ClientRect = placeholder.getBoundingClientRect();
    const textLeft: number = placeholderRect.left - wrapperRect.left;
    const left: number = isRTL() ? textLeft - link.offsetWidth - 5 : textLeft;
    const top: number = placeholderRect.top - wrapperRect.top + (placeholderRect.height - link.offsetHeight) / 2;
    link.style.left = `${left}px`;
    link.style.top = `${isNaN(top) ? 0 : top}px`;
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
