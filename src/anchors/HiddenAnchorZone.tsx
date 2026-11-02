import * as React from 'react';
import { IAnchorData } from './IAnchor';
import { anchorServiceManager } from './AnchorServiceManager';
import { ANCHOR_ATTR, applyPermalink, cls, getHeadingLevel, removeNode, toArray } from './anchorUtils';

export interface IHiddenAnchorZoneProps {
  componentId: string;
  anchorElementSelector: string;
  /** Area containing the (possibly currently hidden) headings. */
  elementForSearch: HTMLElement;
  /** Called before scrolling, e.g. to expand a collapsed section. */
  onShow?: () => void;
  suggestAnchorId?: (element: HTMLElement, index: number) => string;
}

/**
 * Invisible anchors for content that is not always visible (tabs, accordions …).
 * Creates empty mirror elements at the component's position that carry the anchor IDs.
 * (Hooks version, as in the SharePoint original.)
 */
export function HiddenAnchorZone(props: IHiddenAnchorZoneProps): JSX.Element {
  const { anchorElementSelector, componentId, elementForSearch, onShow, suggestAnchorId } = props;

  React.useEffect(() => {
    anchorServiceManager.reserve(componentId);
  }, []);

  // [source, mirror element]
  const pairs: [HTMLElement, HTMLElement][] = React.useMemo(
    () => toArray<HTMLElement>(elementForSearch.querySelectorAll(anchorElementSelector) as any)
      .map((el: HTMLElement): [HTMLElement, HTMLElement] => [el, document.createElement(el.tagName)]),
    [elementForSearch, anchorElementSelector]);

  React.useEffect(() => {
    anchorServiceManager.raiseAnchorEvent({
      action: 'Register',
      uniqueId: componentId,
      anchorList: pairs.map(([source, mirror]: [HTMLElement, HTMLElement], index: number): IAnchorData => {
        const suggested: string = suggestAnchorId ? suggestAnchorId(source, index) : (source.textContent || '').trim();
        return {
          suggestedAnchorId: suggested,
          anchorTargetElement: mirror,
          level: getHeadingLevel(source),
          onReceiveAnchorId: (target: HTMLElement, anchorId: string) => {
            let link: HTMLElement | null = target.querySelector(`[${ANCHOR_ATTR}="${anchorId}"]`) as HTMLElement | null;
            if (!link) {
              link = document.createElement('a');
              applyPermalink(target, link, anchorId, suggested);
              target.appendChild(link);
            }
          },
          onDisposeAnchorId: (target: HTMLElement) => {
            removeNode(target);
          },
          onScrollToAnchorId: (_target: HTMLElement, anchorId: string) => {
            if (onShow) {
              onShow();
            }
            requestAnimationFrame(() => {
              const el: HTMLElement | null = document.getElementById(anchorId);
              if (el) {
                el.scrollIntoView();
              }
            });
          }
        };
      })
    });
    return () => {
      anchorServiceManager.raiseAnchorEvent({ action: 'Unregister', uniqueId: componentId });
    };
  }, [suggestAnchorId, componentId, pairs, onShow]);

  const containerRef: React.MutableRefObject<HTMLDivElement | null> = React.useRef<HTMLDivElement | null>(null);
  const setContainer: (el: HTMLDivElement | null) => void = React.useCallback((el: HTMLDivElement | null) => {
    if (containerRef.current) {
      pairs.forEach(([, mirror]: [HTMLElement, HTMLElement]) => removeNode(mirror));
    }
    if (el) {
      pairs.forEach(([, mirror]: [HTMLElement, HTMLElement]) => el.appendChild(mirror));
    }
    containerRef.current = el;
  }, [pairs]);

  return <div className={cls().hidden} ref={setContainer} aria-hidden='true' />;
}
