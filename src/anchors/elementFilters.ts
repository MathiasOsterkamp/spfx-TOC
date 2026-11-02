/**
 * Filters for headings that must not be anchored / must not appear in the table of contents.
 * Ported from TableOfContents.tsx – without Element.closest() (not available in IE11).
 * Every filter returns true = keep the element.
 */

export type ElementFilter = (element: HTMLElement) => boolean;

/** Walks up from the element (optionally including it) and tests every element. */
function hasAncestor(element: HTMLElement, includeSelf: boolean, test: (el: HTMLElement) => boolean): boolean {
  let current: HTMLElement | null = includeSelf ? element : element.parentElement;
  while (current) {
    if (test(current)) {
      return true;
    }
    current = current.parentElement;
  }
  return false;
}

/** Filters elements that are or are inside the global SharePoint Site Header Title. */
export const filterSiteHeaderTitle: ElementFilter = (element: HTMLElement): boolean =>
  !hasAncestor(element, true, (el: HTMLElement) => el.getAttribute('data-automationid') === 'SiteHeaderTitle');

/** Filters elements that are inside an <aside> tag and thus not related to a page. */
export const filterAside: ElementFilter = (element: HTMLElement): boolean =>
  !hasAncestor(element, false, (el: HTMLElement) => el.tagName.toLowerCase() === 'aside');

/**
 * Filters out elements if they themselves or one of their
 * ancestors have the attribute 'data-toc-ignore'.
 */
export const filterTocIgnore: ElementFilter = (element: HTMLElement): boolean => {
  if (!element) {
    return false;
  }
  return !hasAncestor(element, true, (el: HTMLElement) => el.hasAttribute('data-toc-ignore'));
};

/**
 * Keeps only elements inside #spPageChromeAppDiv.
 * If this container does not exist on the page (e.g. workbench), nothing is filtered.
 */
export const hasNoChromeAppDivParent: ElementFilter = (element: HTMLElement): boolean => {
  if (!element) {
    return false;
  }
  if (!document.getElementById('spPageChromeAppDiv')) {
    return true;
  }
  return hasAncestor(element, true, (el: HTMLElement) => el.id === 'spPageChromeAppDiv');
};

/** Filters elements whose parents have been set with a style of 'display: none'. */
export const filterStyleDisplayNone: ElementFilter = (element: HTMLElement): boolean =>
  !hasAncestor(element, false, (el: HTMLElement) => (el.style.display || '').toLowerCase() === 'none');

/** All table of contents filters combined. */
export const defaultTocFilter: ElementFilter = (element: HTMLElement): boolean =>
  filterAside(element) &&
  filterTocIgnore(element) &&
  hasNoChromeAppDivParent(element) &&
  filterStyleDisplayNone(element) &&
  filterSiteHeaderTitle(element);
