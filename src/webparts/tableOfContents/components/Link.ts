import {
  AnchorService,
  ISortedAnchor,
  MAX_SLUG_LENGTH,
  validateLinkTitle,
  validateLinkUrl
} from '../../../anchors';

export interface Link {
  /** Child nodes of the link. */
  childNodes: Link[];
  /** Parent link. Undefined for the root link. */
  parent: Link | undefined;
  id: string;
  title: string;
  href: string;
  level: number;
  element?: HTMLElement;
}

/** A custom link from PropertyFieldCollectionData. */
export interface ICustomLink {
  uniqueId?: string;
  title: string;
  url: string;
  level: number;
}

export function generateSlug(title: string, index: number): string {
  const slug: string = AnchorService.sanitizeAnchorId(title).substring(0, MAX_SLUG_LENGTH) || 'link';
  return `${slug}-${index}`;
}

/**
 * Builds a tree by level from a flat, sorted list.
 * Like buildLinkTree in the original, additionally:
 *  - skipped levels (H2 -> H4) attach to the nearest higher existing entry
 *  - after a higher entry, deeper "last nodes" are discarded
 */
export function buildLinkTree(items: Link[]): Link[] {
  const rootLinks: Link[] = [];
  const lastActiveNodes: Link[] = [];

  items.forEach((item: Link) => {
    const currentLevel: number = Math.max(1, item.level || 1);
    item.childNodes = [];
    item.parent = undefined;

    let parentLink: Link | undefined;
    for (let i: number = currentLevel - 2; i >= 0; i--) {
      if (lastActiveNodes[i]) {
        parentLink = lastActiveNodes[i];
        break;
      }
    }

    if (parentLink) {
      item.parent = parentLink;
      parentLink.childNodes.push(item);
    } else {
      rootLinks.push(item);
    }

    lastActiveNodes[currentLevel - 1] = item;
    lastActiveNodes.length = currentLevel;
  });

  return rootLinks;
}

/** Links from the anchors reported by the AnchorService. */
export function linksFromAnchors(anchors: ISortedAnchor[]): Link[] {
  return buildLinkTree(anchors.map((a: ISortedAnchor) => ({
    childNodes: [],
    parent: undefined,
    id: a.anchorId,
    title: a.anchorText,
    href: `#${a.anchorId}`,
    level: a.level || 1
  })));
}

/** Only show valid links (e.g. in case the data was changed outside the property pane). */
export function isValidCustomLink(link: ICustomLink): boolean {
  return !validateLinkTitle(link.title) && !validateLinkUrl(link.url);
}

/** Create links from custom elements. */
export function linksFromCustom(customLinks: ICustomLink[]): Link[] {
  return buildLinkTree(customLinks
    .filter(isValidCustomLink)
    .map((c: ICustomLink, index: number) => ({
      childNodes: [],
      parent: undefined,
      id: c.uniqueId || generateSlug(c.title, index),
      title: c.title.trim(),
      href: c.url.trim(),
      level: c.level || 1
    })));
}

/**
 * Reads the value of the web part property "links":
 *  - array from PropertyFieldCollectionData
 *  - or text from the previous version ("Level | Title | URL" per line) – migrated
 */
export function toCustomLinks(value: any): ICustomLink[] {
  if (Array.isArray(value)) {
    return value.map((v: any) => ({
      uniqueId: v && v.uniqueId,
      title: v && typeof v.title === 'string' ? v.title : '',
      url: v && typeof v.url === 'string' ? v.url : '',
      level: Math.min(6, Math.max(1, parseInt(v && v.level, 10) || 1))
    }));
  }
  if (typeof value === 'string') {
    return parseCustomLinks(value);
  }
  return [];
}

/** Text of the previous version: one line per link "Level | Title | URL". */
export function parseCustomLinks(text: string): ICustomLink[] {
  if (!text) {
    return [];
  }
  return text.split(/\r?\n/)
    .map((line: string) => line.trim())
    .filter((line: string) => line.length > 0)
    .map((line: string) => {
      const parts: string[] = line.split('|').map((p: string) => p.trim());
      let level: number = 1;
      if (parts.length > 1 && /^\d+$/.test(parts[0])) {
        level = Math.min(6, Math.max(1, parseInt(parts.shift() as string, 10)));
      }
      return {
        level: level,
        title: parts[0] || '',
        url: parts[1] || '#'
      };
    })
    .filter((c: ICustomLink) => c.title.length > 0);
}
