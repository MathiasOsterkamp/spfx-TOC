import { MAX_SLUG_LENGTH } from './inlineAnchors';

/**
 * ID allocation – the ONE place where IDs are created.
 * Used by the AnchorService (view mode) and by the prediction (edit mode, dropdown),
 * so that both are guaranteed to produce the same IDs.
 */

/** Characters that sanitizeAnchorId replaces with "-" (as in the SharePoint original). */
export const INVALID_ANCHOR_ID_CHARS: RegExp = /[`"'<>#%{}|\\^~\/\[\];?:@=&\s\0]/g;

/** Text -> valid, URL-safe ID (max. 128 characters). */
export function sanitizeAnchorId(id: string): string {
  return id
    ? id.trim().toLowerCase()
      .replace(INVALID_ANCHOR_ID_CHARS, '-')
      .replace(/-{2,}/g, '-')
      .replace(/^-+|-+$/g, '')
      .substring(0, 128)
    : '';
}

/** Duplicate IDs: "title", "title-1", "title-2", … (as in the SharePoint original) */
export function handleAnchorCollision(id: string, counters: Map<string, number>): string {
  let candidate: string = id;
  let count: number = 0;
  while (counters.has(candidate)) {
    id = candidate;
    count = counters.get(id) || 0;
    candidate = `${id}-${count}`;
  }
  counters.set(id, count + 1);
  counters.set(candidate, 1);
  return candidate;
}

/** ID for a heading: sanitize the text, resolve collisions. */
export function allocateAnchorId(suggestedAnchorId: string, counters: Map<string, number>): string {
  return handleAnchorCollision(sanitizeAnchorId(suggestedAnchorId) || 'anchor', counters);
}

export interface IAnchorIdRequest {
  suggestedAnchorId: string;
  /** true = keep the ID unchanged (inline anchor {anchor-#…}, deliberately chosen by the author). */
  preserveId?: boolean;
}

/**
 * Allocates IDs for a list in document order.
 *
 * 1. Fixed IDs (inline anchors) are reserved first – exactly as written (case is preserved).
 *    This way links like "#Important_1" always work, even if a heading has the same name.
 *    Only if the same marker occurs twice does the second one get a suffix ("Important_1-1").
 * 2. Headings get their IDs afterwards; if they equal a fixed ID they move aside ("introduction-1").
 */
export function assignAnchorIds(items: IAnchorIdRequest[]): string[] {
  const counters: Map<string, number> = new Map<string, number>();
  const ids: string[] = [];
  items.forEach((item: IAnchorIdRequest, i: number) => {
    if (item.preserveId) {
      ids[i] = handleAnchorCollision((item.suggestedAnchorId || '').substring(0, MAX_SLUG_LENGTH) || 'anchor', counters);
    }
  });
  items.forEach((item: IAnchorIdRequest, i: number) => {
    if (!item.preserveId) {
      ids[i] = allocateAnchorId(item.suggestedAnchorId, counters);
    }
  });
  return ids;
}
