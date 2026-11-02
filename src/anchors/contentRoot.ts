import { EXCLUDE_ATTR } from './anchorUtils';

/** Default selectors for the content area of modern pages (SPO and SharePoint 2019). */
export const DEFAULT_CONTENT_SELECTORS: string[] = [
  '[data-automation-id="Canvas"]',
  '.SPCanvas',
  '#spPageCanvasContent',
  '.CanvasZone'
];

/** Content area of the page (first match), otherwise document.body. */
export function resolveContentRoot(contentSelector?: string): HTMLElement {
  const selectors: string[] = contentSelector ? [contentSelector] : DEFAULT_CONTENT_SELECTORS;
  for (let i: number = 0; i < selectors.length; i++) {
    try {
      const el: HTMLElement | null = document.querySelector(selectors[i]) as HTMLElement | null;
      if (el) {
        return el;
      }
    } catch (err) {
      // invalid selector -> try the next one
    }
  }
  return document.body;
}

/** Web part ID of the built-in text web part: its content is page content and is anchored. */
const TEXT_WEB_PART_ID: string = '868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823';

/** Is the element inside another (non-text) web part? Its DOM belongs to that web part and is never touched. */
function isInForeignWebPart(node: HTMLElement): boolean {
  const webPartId: string = (node.getAttribute('data-sp-web-part-id') || '').toLowerCase();
  if (webPartId) {
    return webPartId !== TEXT_WEB_PART_ID;
  }
  if (node.hasAttribute('data-sp-web-part')) {
    return (node.getAttribute('data-sp-feature-tag') || '').toLowerCase().indexOf('text web part') < 0;
  }
  return false;
}

/**
 * Is the element inside an area with data-spfx-anchor-exclude (e.g. inside our web parts)
 * or inside another web part (only text web part content is anchored)?
 */
export function isExcludedElement(element: HTMLElement): boolean {
  let node: HTMLElement | null = element;
  while (node) {
    if (node.hasAttribute && (node.hasAttribute(EXCLUDE_ATTR) || isInForeignWebPart(node))) {
      return true;
    }
    node = node.parentElement;
  }
  return false;
}
