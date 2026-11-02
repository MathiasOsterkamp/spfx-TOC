import { MAX_SLUG_LENGTH } from './inlineAnchors';
import { INVALID_ANCHOR_ID_CHARS } from './anchorIds';

/**
 * Validation of custom links (title + URL) in the table of contents.
 *
 * Principle: everything that can exist as an anchor on a page is valid.
 *  - heading IDs: lower case, umlauts/accents, dots, parentheses …   e.g. "veröffentlicht-1", "a.1", "été-(résumé)"
 *  - inline IDs from {anchor-#…}: mixed case, "_" and "-"                      e.g. "Important_1"
 *  - native SharePoint IDs (same algorithm as ours)
 * Only values that can never be an anchor or that are unsafe are rejected.
 */

/** Result: undefined = valid, otherwise an error code (translated via the strings). */
export type LinkValidationError =
  | 'titleRequired'
  | 'titleTooLong'
  | 'titleInvalidChars'
  | 'urlRequired'
  | 'urlTooLong'
  | 'urlInvalidChars'
  | 'urlUnsafeScheme'
  | 'urlEmptyAnchor'
  | 'urlInvalidAnchor';

export const MAX_URL_LENGTH: number = 2048;

/** Longest possible anchor ID: inline IDs up to MAX_SLUG_LENGTH, headings 128 + collision suffix. */
export const MAX_ANCHOR_ID_LENGTH: number = MAX_SLUG_LENGTH;

/** Allowed URL schemes. Everything else (javascript:, data:, vbscript:, file: …) is rejected. */
export const SAFE_URL_SCHEMES: string[] = ['http:', 'https:', 'mailto:', 'tel:'];

const CONTROL_CHARS: RegExp = /[\u0000-\u001F\u007F]/;

/**
 * Allowed characters in the link text.
 * Replaces the former whitelist (a-z, äöüß, éàç …), which rejected e.g. "è", "ê", "œ", „…“ or "–"
 * – i.e. headings that are perfectly normal on French/German pages.
 * Every printable character is allowed now; React escapes the text anyway.
 */
// eslint-disable-next-line no-control-regex
export const allowedCharsRegEx: RegExp = /^[^\u0000-\u001F\u007F]+$/;

/**
 * Characters that can never occur in an anchor ID (after decoding): the ones the ID allocation
 * replaces (INVALID_ANCHOR_ID_CHARS, so validation and allocation cannot drift apart) plus control characters.
 * Non-global copy, a global RegExp keeps state between test() calls.
 */
// eslint-disable-next-line no-control-regex
const INVALID_FRAGMENT_CHARS: RegExp = new RegExp(`${INVALID_ANCHOR_ID_CHARS.source}|[\\u0000-\\u001F\\u007F]`);

/** Take link text from a heading: remove control characters/line breaks, truncate. */
export function cleanLinkTitle(text: string): string {
  return (text || '')
    .replace(/[\u0000-\u001F\u007F]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, MAX_SLUG_LENGTH);
}

export function validateLinkTitle(value: string): LinkValidationError | undefined {
  const text: string = (value || '').trim();
  if (!text) {
    return 'titleRequired';
  }
  if (text.length > MAX_SLUG_LENGTH) {
    return 'titleTooLong';
  }
  if (!allowedCharsRegEx.test(text)) {
    return 'titleInvalidChars';
  }
  return undefined;
}

/** Can this (decoded) ID be an anchor on a page? */
export function isValidAnchorId(id: string): boolean {
  return !!id && id.length <= MAX_ANCHOR_ID_LENGTH && !INVALID_FRAGMENT_CHARS.test(id);
}

export function validateLinkUrl(value: string): LinkValidationError | undefined {
  const url: string = (value || '').trim();
  if (!url) {
    return 'urlRequired';
  }
  if (url.length > MAX_URL_LENGTH) {
    return 'urlTooLong';
  }
  // Spaces in the path are allowed (e.g. "Shared Documents"), the browser encodes them.
  if (CONTROL_CHARS.test(url) || /[<>"`]/.test(url)) {
    return 'urlInvalidChars';
  }

  const scheme: RegExpExecArray | null = /^([a-z][a-z0-9+.\-]*):/i.exec(url);
  if (scheme && SAFE_URL_SCHEMES.indexOf(`${scheme[1].toLowerCase()}:`) < 0) {
    return 'urlUnsafeScheme';
  }

  const hashIndex: number = url.indexOf('#');
  if (hashIndex >= 0) {
    const fragment: string = url.substring(hashIndex + 1);
    // Only anchors of THIS page must follow our ID rules; fragments of other pages/sites
    // (e.g. "#page=2", "#/route?x=1") are arbitrary and are not judged.
    if (!isSamePageUrl(url, hashIndex)) {
      return undefined;
    }
    if (!fragment) {
      return 'urlEmptyAnchor';
    }
    let decoded: string;
    try {
      decoded = decodeURIComponent(fragment);
    } catch (e) {
      return 'urlInvalidAnchor';
    }
    if (!isValidAnchorId(decoded)) {
      return 'urlInvalidAnchor';
    }
  }
  return undefined;
}

/** Does the URL (with a '#' at hashIndex) point to the current page? "#id" always does. */
function isSamePageUrl(url: string, hashIndex: number): boolean {
  if (hashIndex === 0) {
    return true;
  }
  const a: HTMLAnchorElement = document.createElement('a');
  a.href = url;
  const normalize: (p: string) => string = (p: string) => (p.charAt(0) === '/' ? p : `/${p}`).toLowerCase();
  return a.hostname.toLowerCase() === window.location.hostname.toLowerCase() &&
    normalize(a.pathname) === normalize(window.location.pathname);
}

/**
 * Does the URL point to an anchor on the current page? Then its (decoded) ID, otherwise undefined.
 */
export function getSamePageAnchorId(value: string): string | undefined {
  const url: string = (value || '').trim();
  const hashIndex: number = url.indexOf('#');
  if (hashIndex < 0 || hashIndex === url.length - 1 || !isSamePageUrl(url, hashIndex)) {
    return undefined;
  }
  try {
    return decodeURIComponent(url.substring(hashIndex + 1));
  } catch (e) {
    return undefined;
  }
}
