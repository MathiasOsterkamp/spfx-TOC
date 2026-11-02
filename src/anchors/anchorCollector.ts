import { AnchorKind } from './IAnchor';
import { getHeadingLevel, toArray } from './anchorUtils';
import { ElementFilter } from './elementFilters';
import {
  INLINE_ANCHOR_ATTR,
  getInlineAnchorsElements,
  getMarkerIds,
  idFromMarker,
  stripInlineMarkers
} from './inlineAnchors';

/**
 * A possible anchor on the page – heading or inline marker.
 * Basis for both registration (AnchorZone) and prediction (anchorCatalog).
 */
export interface IAnchorCandidate {
  /** Node used for document order (element or – for markers not converted yet – the text node). */
  node: Node;
  /** Target element (undefined for markers not converted yet). */
  element?: HTMLElement;
  kind: AnchorKind;
  /** Heading: text without markers. Inline: the ID from the marker. */
  suggestedAnchorId: string;
  level: number;
  /** Inline IDs are kept unchanged. */
  preserveId: boolean;
}

export interface ICollectAnchorsOptions {
  /** Heading selector (getQuerySelector / extension configuration). Empty = no headings. */
  headingSelector: string;
  /** true = keep the element. */
  filter?: ElementFilter;
  /** Include inline anchors. */
  includeInline: boolean;
  /** Also include markers still present as text (edit mode / prediction). */
  includeUnconvertedMarkers?: boolean;
  /** Additional heading-like elements outside the root (e.g. the page title field). */
  extraElements?: HTMLElement[];
}

/** Heading text as shown in view mode (markers are already removed there). */
export function getHeadingText(element: HTMLElement): string {
  // The page title in edit mode is a textarea: the current text is its value, not its textContent
  const raw: string = element.tagName === 'TEXTAREA' ? (element as HTMLTextAreaElement).value : element.textContent || '';
  return stripInlineMarkers(raw).trim();
}

/**
 * Collects all anchor candidates in the area in document order.
 * Headings and inline anchors end up in ONE list – IDs are allocated together.
 */
export function collectAnchorCandidates(root: HTMLElement, options: ICollectAnchorsOptions): IAnchorCandidate[] {
  const keep: ElementFilter = (el: HTMLElement) => !options.filter || options.filter(el);
  const items: IAnchorCandidate[] = [];

  if (options.headingSelector) {
    let headings: HTMLElement[] = [];
    try {
      headings = toArray<HTMLElement>(root.querySelectorAll(options.headingSelector) as any);
    } catch (e) {
      headings = [];
    }
    (options.extraElements || []).forEach((extra: HTMLElement) => {
      if (headings.indexOf(extra) < 0) {
        headings.push(extra);
      }
    });
    headings.filter(keep).forEach((el: HTMLElement) => {
      items.push({
        node: el,
        element: el,
        kind: 'heading',
        suggestedAnchorId: getHeadingText(el),
        level: getHeadingLevel(el),
        preserveId: false
      });
    });
  }

  if (options.includeInline) {
    // markers already converted (view mode)
    toArray<HTMLElement>(root.querySelectorAll(`[${INLINE_ANCHOR_ATTR}]`) as any)
      .filter(keep)
      .forEach((span: HTMLElement) => {
        items.push({
          node: span,
          element: span,
          kind: 'inline',
          suggestedAnchorId: idFromMarker(span.getAttribute(INLINE_ANCHOR_ATTR)) || span.id,
          level: 0,
          preserveId: true
        });
      });

    // markers still present as text (edit mode)
    if (options.includeUnconvertedMarkers) {
      getInlineAnchorsElements(root, undefined, keep).forEach((textNode: Text) => {
        getMarkerIds(textNode.data).forEach((id: string) => {
          items.push({ node: textNode, kind: 'inline', suggestedAnchorId: id, level: 0, preserveId: true });
        });
      });
    }
  }

  return sortByDocumentPosition(items);
}

/** Stable sort by document position (IE11's sort is not stable). */
function sortByDocumentPosition(items: IAnchorCandidate[]): IAnchorCandidate[] {
  const indexed: { item: IAnchorCandidate; index: number }[] =
    items.map((item: IAnchorCandidate, index: number) => ({ item: item, index: index }));
  indexed.sort((a: { item: IAnchorCandidate; index: number }, b: { item: IAnchorCandidate; index: number }) => {
    if (a.item.node === b.item.node) {
      return a.index - b.index;
    }
    const position: number = a.item.node.compareDocumentPosition(b.item.node);
    // tslint:disable-next-line:no-bitwise
    if (position & Node.DOCUMENT_POSITION_FOLLOWING) {
      return -1;
    }
    // tslint:disable-next-line:no-bitwise
    if (position & Node.DOCUMENT_POSITION_PRECEDING) {
      return 1;
    }
    return a.index - b.index;
  });
  return indexed.map((x: { item: IAnchorCandidate; index: number }) => x.item);
}
