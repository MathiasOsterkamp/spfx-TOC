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

/** Is the element inside an area with data-spfx-anchor-exclude (e.g. inside our web parts)? */
export function isExcludedElement(element: HTMLElement): boolean {
  let node: HTMLElement | null = element;
  while (node) {
    if (node.hasAttribute && node.hasAttribute(EXCLUDE_ATTR)) {
      return true;
    }
    node = node.parentElement;
  }
  return false;
}
