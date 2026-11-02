import { ISortedAnchor } from './IAnchor';
import { assignAnchorIds } from './anchorIds';
import { IAnchorCandidate, collectAnchorCandidates } from './anchorCollector';
import { isExcludedElement, resolveContentRoot } from './contentRoot';
import { ElementFilter } from './elementFilters';
import { cleanLinkTitle } from './linkValidation';
import { PAGE_TITLE_TEXTAREA_SELECTOR } from './inlineAnchors';

/** Predicted anchor – same shape as the entries of the AnchorService. */
export type IAvailableAnchor = ISortedAnchor;

export interface IAnchorCatalogOptions {
  /** Heading selector – the same one the engine uses in view mode. */
  headingSelector: string;
  /** The same additional filter as in the engine (e.g. defaultTocFilter). */
  filter?: ElementFilter;
  contentSelector?: string;
  /** Offer inline anchors {anchor-#…} (only if the engine converts them). Default: true */
  includeInline?: boolean;
}

/**
 * Returns all anchors that will exist on the current page in view mode –
 * also in edit mode, where the engine has not assigned any IDs yet.
 *
 * Same collector (collectAnchorCandidates) and same ID allocation (assignAnchorIds)
 * as the AnchorService – so predicted and actual IDs are identical.
 */
export function getAvailableAnchors(options: IAnchorCatalogOptions): ISortedAnchor[] {
  const root: HTMLElement = resolveContentRoot(options.contentSelector);
  // The page title field is outside the canvas; it is part of the heading selector in edit mode only
  const titleField: HTMLElement | null = options.headingSelector.indexOf(PAGE_TITLE_TEXTAREA_SELECTOR) >= 0
    ? document.querySelector(PAGE_TITLE_TEXTAREA_SELECTOR) as HTMLElement | null
    : null;
  const candidates: IAnchorCandidate[] = collectAnchorCandidates(root, {
    headingSelector: options.headingSelector,
    filter: (el: HTMLElement) => !isExcludedElement(el) && (!options.filter || options.filter(el)),
    includeInline: options.includeInline !== false,
    includeUnconvertedMarkers: true,
    extraElements: titleField ? [titleField] : []
  });
  const ids: string[] = assignAnchorIds(candidates);

  return candidates.map((c: IAnchorCandidate, i: number): ISortedAnchor => ({
    anchorId: ids[i],
    // Heading without text (whitespace/image only): use the ID as title so a selection stays valid
    anchorText: c.kind === 'heading' ? (cleanLinkTitle(c.suggestedAnchorId) || ids[i]) : ids[i],
    anchoredElementTagName: c.element ? c.element.tagName : 'SPAN',
    level: c.level,
    kind: c.kind
  }));
}
