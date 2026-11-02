import { cls, removeNode, toArray } from './anchorUtils';

/**
 * Inline anchors in the text, e.g. "Text {anchor-#my-id} continues".
 *
 * The markers are converted into invisible targets (<span>) and then registered with the
 * AnchorService like headings (kind = 'inline', fixed ID) – see anchorCollector.ts.
 * The whole content area is searched (paragraphs, lists, tables, headings …),
 * a separate container selector is no longer needed.
 *
 * IE11: no NodeList.forEach, createTreeWalker with all four parameters.
 */

// Matches exactly {anchor-#some-id} - letters, digits, hyphens, underscores
export const MARKER_REGEX: RegExp = /\{(?:[aA]nker|ANKER|[aA]nchor|ANCHOR)-#([a-zA-Z0-9_-]+)\}/gi;

// allowedCharsRegEx: see linkValidation.ts (Unicode-aware, no conflict with headings)

// Maximum allowed length for anchors and link text
export const MAX_SLUG_LENGTH: number = 255;

/** Stores the original marker on the created target element so it can be restored. */
export const INLINE_ANCHOR_ATTR: string = 'data-spfx-inline-anchor';

/** Markers are never searched/replaced inside these elements. */
const SKIP_TAGS: string[] = ['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEXTAREA', 'INPUT', 'SELECT', 'OPTION', 'CODE', 'PRE'];

/**
 * Returns a query selector based on the specified props.
 * Matches native HTML headings, ARIA heading roles,
 * and the specific page title input textarea when in edit mode.
 */
export function getQuerySelector(
  showHeading1: boolean,
  showHeading2: boolean,
  showHeading3: boolean,
  showHeading4: boolean,
  showHeading5: boolean,
  isEditMode: boolean = false
): string {
  const queryParts: string[] = [];
  const add: (level: number) => void = (level: number) => {
    // :not([role="heading"]) prevents duplicates if an element is hN AND has the ARIA role
    queryParts.push(`h${level}:not([role="heading"])`, `[role="heading"][aria-level="${level}"]`);
  };

  if (showHeading1) {
    add(1);
    // Page title field only in edit mode
    if (isEditMode) {
      queryParts.push('textarea[data-automation-id="pageTitleInput"]');
    }
  }
  if (showHeading2) { add(2); }
  if (showHeading3) { add(3); }
  if (showHeading4) { add(4); }
  if (showHeading5) { add(5); }

  return queryParts.join(',');
}

/** All marker IDs of a text in order (including duplicates). */
export function getMarkerIds(text: string): string[] {
  const ids: string[] = [];
  const re: RegExp = new RegExp(MARKER_REGEX.source, 'gi');
  let match: RegExpExecArray | null = re.exec(text || '');
  while (match) {
    ids.push(match[1].substring(0, MAX_SLUG_LENGTH));
    match = re.exec(text);
  }
  return ids;
}

/** Text without markers – as it is displayed after conversion. */
export function stripInlineMarkers(text: string): string {
  return (text || '').replace(new RegExp(MARKER_REGEX.source, 'gi'), '');
}

function isSkippedTextNode(node: Node): boolean {
  let el: HTMLElement | null = node.parentElement || (node.parentNode as HTMLElement | null);
  while (el) {
    if (SKIP_TAGS.indexOf(el.tagName) >= 0 || el.hasAttribute(INLINE_ANCHOR_ATTR)) {
      return true;
    }
    el = el.parentElement;
  }
  return false;
}

/**
 * Get inline anchors, for example "{anchor-#testid}".
 * Returns all text nodes containing at least one marker.
 *
 * @param root Search area (default: document.body)
 * @param containerSelector Optional: only search inside these elements (formerly fixed to 'p'). Empty = whole area.
 * @param filter Optional: exclude elements (tested on the parent element of the text)
 */
export function getInlineAnchorsElements(
  root?: HTMLElement | Document,
  containerSelector?: string,
  filter?: (element: HTMLElement) => boolean
): Text[] {
  const scope: HTMLElement | Document = root || document.body;
  const containers: Node[] = containerSelector
    ? toArray<Node>(scope.querySelectorAll(containerSelector) as any)
    : [scope];
  const textNodes: Text[] = [];

  containers.forEach((container: Node) => {
    // IE11 requires all four parameters, newer TS typings no longer know the fourth one
    const walker: TreeWalker = (document.createTreeWalker as any)(container, NodeFilter.SHOW_TEXT, null, false);
    let node: Node | null = walker.nextNode();
    while (node) {
      const text: string = node.textContent || '';
      if (text.indexOf('{') >= 0 && getMarkerIds(text).length > 0 && textNodes.indexOf(node as Text) < 0) {
        const parent: HTMLElement | null = node.parentElement;
        if (!isSkippedTextNode(node) && (!filter || !parent || filter(parent))) {
          textNodes.push(node as Text);
        }
      }
      node = walker.nextNode();
    }
  });

  return textNodes;
}

/**
 * Extracts the plain anchor IDs from a list of text nodes (without duplicates).
 *
 * @param textNodes Array of text nodes (e.g. the result of getInlineAnchorsElements)
 */
export function extractAnchorIds(textNodes: Text[]): string[] {
  const ids: string[] = [];
  textNodes.forEach((node: Text) => {
    getMarkerIds(node.data).forEach((id: string) => {
      if (ids.indexOf(id) < 0) {
        ids.push(id);
      }
    });
  });
  return ids;
}

/**
 * Converts all markers in the area into invisible targets:
 * "Text {anchor-#target} more"  ->  "Text <span id="target" data-spfx-inline-anchor="{anchor-#target}"></span> more"
 * The final ID (with a suffix for duplicate markers) is then assigned by the AnchorService.
 *
 * @returns the newly created target elements
 */
export function convertInlineAnchors(
  root?: HTMLElement | Document,
  containerSelector?: string,
  filter?: (element: HTMLElement) => boolean
): HTMLElement[] {
  const created: HTMLElement[] = [];

  // Collect first, then modify – never change the DOM while the TreeWalker is running
  getInlineAnchorsElements(root, containerSelector, filter).forEach((textNode: Text) => {
    const parent: Node | null = textNode.parentNode;
    if (!parent) {
      return;
    }
    const text: string = textNode.data;
    const fragment: DocumentFragment = document.createDocumentFragment();
    const re: RegExp = new RegExp(MARKER_REGEX.source, 'gi');
    let last: number = 0;
    let match: RegExpExecArray | null = re.exec(text);

    while (match) {
      if (match.index > last) {
        fragment.appendChild(document.createTextNode(text.substring(last, match.index)));
      }
      const target: HTMLElement = document.createElement('span');
      target.id = match[1].substring(0, MAX_SLUG_LENGTH);
      target.className = cls().inlineAnchor;
      target.setAttribute(INLINE_ANCHOR_ATTR, match[0]);
      fragment.appendChild(target);
      created.push(target);

      last = match.index + match[0].length;
      match = re.exec(text);
    }

    if (last < text.length) {
      fragment.appendChild(document.createTextNode(text.substring(last)));
    }
    parent.replaceChild(fragment, textNode);
  });

  return created;
}

/** ID from a stored marker ("{anchor-#target}" -> "target"). */
export function idFromMarker(marker: string | null): string {
  const ids: string[] = getMarkerIds(marker || '');
  return ids.length > 0 ? ids[0] : '';
}

/**
 * Reverts convertInlineAnchors: the target elements are replaced by the
 * original marker text again (e.g. before switching to edit mode).
 */
export function revertInlineAnchors(root: HTMLElement | Document): void {
  const targets: HTMLElement[] = toArray<HTMLElement>(root.querySelectorAll(`[${INLINE_ANCHOR_ATTR}]`) as any);
  const parents: Node[] = [];

  targets.forEach((target: HTMLElement) => {
    const parent: Node | null = target.parentNode;
    if (!parent) {
      return;
    }
    const marker: string = target.getAttribute(INLINE_ANCHOR_ATTR) || '';
    if (marker) {
      parent.replaceChild(document.createTextNode(marker), target);
    } else {
      removeNode(target);
    }
    if (parents.indexOf(parent) < 0) {
      parents.push(parent);
    }
  });

  // Merge the split text nodes again
  parents.forEach((p: Node) => p.normalize());
}
