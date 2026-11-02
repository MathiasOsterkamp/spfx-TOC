import { ANCHOR_ATTR, ANCHOR_TARGET_ATTR, cls, removeNode, toArray } from './anchorUtils';

/**
 * Removes permalink icons, placeholders and wrappers from a DOM tree again
 * (e.g. before switching to edit mode or before copying HTML).
 * The IDs on the headings are kept.
 */
export function cleanupAnchorIconsInTree(root: HTMLElement): void {
  const anchorClass: string = cls().anchor;
  const links: HTMLElement[] = toArray<HTMLElement>(root.querySelectorAll(`a[${ANCHOR_ATTR}]`) as any);
  links.forEach((link: HTMLElement) => {
    const wrapper: HTMLElement | null = link.parentElement;
    removeNode(link);
    if (!wrapper) {
      return;
    }
    const heading: Element | null = wrapper.firstElementChild;
    if (heading) {
      const children: Element[] = toArray<Element>(heading.children as any);
      for (let i: number = children.length - 1; i >= 0; i--) {
        const child: Element = children[i];
        if (child.tagName === 'SPAN' && child.classList.contains(anchorClass)) {
          removeNode(child);
        }
      }
      if (wrapper.parentElement) {
        wrapper.parentElement.insertBefore(heading, wrapper);
      }
    }
    removeNode(wrapper);
  });

  // Remove the markers (IDs stay)
  toArray<HTMLElement>(root.querySelectorAll(`[${ANCHOR_TARGET_ATTR}]`) as any)
    .forEach((el: HTMLElement) => el.removeAttribute(ANCHOR_TARGET_ATTR));
}
