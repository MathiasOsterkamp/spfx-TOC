import { Log, Text } from '@microsoft/sp-core-library';
import { ANCHOR_STATIC_CLASS, GLYPH_STATIC_CLASS, IAnchorClassNames, getAnchorClassNames } from './anchorStyles';
import { ITheme } from '@fluentui/react/lib/Styling';
import { getPageTheme } from './pageTheme';

export const LOG_SOURCE: string = 'AnchorService';

/**
 * Attribute on the permalink links. Deliberately different from the native "data-sp-anchor-id"
 * so this solution does not collide with the built-in SharePoint AnchorService.
 */
export const ANCHOR_ATTR: string = 'data-spfx-anchor-id';

/** Marks headings that received an ID from us (detects re-rendered headings). */
export const ANCHOR_TARGET_ATTR: string = 'data-spfx-anchor-target';

/** Areas with this attribute are never anchored (e.g. the web parts themselves). */
export const EXCLUDE_ATTR: string = 'data-spfx-anchor-exclude';

export const HEADING_TAGS: string[] = ['H1', 'H2', 'H3', 'H4', 'H5', 'H6'];

/** CSS classes of the anchors (mergeStyleSets, created once per theme; default: page theme). */
export function cls(theme?: ITheme): IAnchorClassNames {
  return getAnchorClassNames(theme || getPageTheme());
}

export { ANCHOR_STATIC_CLASS };

/** Texts the components can override from their localization files. */
export const anchorStrings: { permalinkFormat: string } = {
  permalinkFormat: 'Permalink für {0}'
};

export function toArray<T>(list: { length: number; [index: number]: T }): T[] {
  return Array.prototype.slice.call(list);
}

export function removeNode(node: Node | null | undefined): void {
  if (node && node.parentNode) {
    node.parentNode.removeChild(node);
  }
}

/** Is the node (still) in the document? (IE11 does not support isConnected) */
export function isInDocument(node: Node | null | undefined): boolean {
  return !!node && (node === document.body || document.body.contains(node));
}

export function isRTL(): boolean {
  return (document.documentElement.getAttribute('dir') || '').toLowerCase() === 'rtl';
}

export function buildPermalink(anchorId: string): string {
  const base: string = window.location.href.split('#')[0];
  return `${base}#${encodeURIComponent(anchorId)}`;
}

/** Level of a heading: H1–H6 or role="heading" with aria-level. Otherwise 6. */
export function getHeadingLevel(element: HTMLElement): number {
  const tag: string = element.tagName;
  if (HEADING_TAGS.indexOf(tag) >= 0) {
    return parseInt(tag.substring(1), 10);
  }
  if (tag === 'TEXTAREA') {
    // page title field in edit mode
    return 1;
  }
  const ariaLevel: number = parseInt(element.getAttribute('aria-level') || '', 10);
  return isNaN(ariaLevel) ? 6 : ariaLevel;
}

/** Sets the ID on the target and the permalink attributes on the link. */
export function applyPermalink(
  target: HTMLElement,
  link: HTMLElement,
  anchorId: string,
  label: string,
  focusable?: boolean
): void {
  target.setAttribute('id', anchorId);
  link.setAttribute('role', 'link');
  link.setAttribute('title', Text.format(anchorStrings.permalinkFormat, label));
  link.setAttribute(ANCHOR_ATTR, anchorId);
  link.setAttribute('href', buildPermalink(anchorId));
  link.setAttribute('target', '_self');
  link.setAttribute('data-interception', 'off');
  if (!focusable) {
    link.setAttribute('tabIndex', '-1');
  }
  link.onclick = () => {
    Log.verbose(LOG_SOURCE, 'AnchorZone.ClickAnchor');
  };
}

const SVG_NS: string = 'http://www.w3.org/2000/svg';

/** Link glyph as inline SVG (no dependency on the icon font). */
export function createLinkGlyph(theme?: ITheme): HTMLElement {
  const glyph: HTMLElement = document.createElement('span');
  glyph.className = `${cls(theme).glyph} ${GLYPH_STATIC_CLASS}`;
  glyph.setAttribute('aria-hidden', 'true');

  const svg: Element = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('width', '16');
  svg.setAttribute('height', '16');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '2');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.setAttribute('focusable', 'false');

  [
    'M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71',
    'M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71'
  ].forEach((d: string) => {
    const path: Element = document.createElementNS(SVG_NS, 'path');
    path.setAttribute('d', d);
    svg.appendChild(path);
  });

  glyph.appendChild(svg);
  return glyph;
}
