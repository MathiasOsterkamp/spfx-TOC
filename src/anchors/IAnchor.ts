/** heading = heading element, inline = anchor from a {anchor-#id} marker in the text */
export type AnchorKind = 'heading' | 'inline';

/** An anchor reported by a zone. */
export interface IAnchorData {
  /** Text the ID is derived from (usually the heading text). */
  suggestedAnchorId: string;
  /** Element that receives the ID and is scrolled to. */
  anchorTargetElement: HTMLElement;
  /** Unique ID assigned by the service. */
  anchorId?: string;
  /** Optional: explicit level (e.g. HiddenAnchorZone, whose mirror elements carry the source level). */
  level?: number;
  /** Kind of anchor (default: heading). */
  kind?: AnchorKind;
  /** true = suggestedAnchorId already is the desired ID (inline anchor) and is not sanitized. */
  preserveId?: boolean;
  onReceiveAnchorId?: (element: HTMLElement, anchorId: string) => void;
  onDisposeAnchorId?: (element: HTMLElement, anchorId: string) => void;
  onScrollToAnchorId?: (element: HTMLElement, anchorId: string) => void;
}

/** An entry of the sorted anchor list (e.g. for a table of contents). */
export interface ISortedAnchor {
  anchorText: string;
  anchoredElementTagName: string;
  anchorId: string;
  /** Heading level (1–6), also for role="heading" with aria-level; 0 for inline anchors. */
  level: number;
  kind: AnchorKind;
}

export interface IAnchorEvent {
  action: 'ReserveRegistration' | 'Register' | 'Unregister' | 'Refresh';
  uniqueId: string;
  anchorList?: IAnchorData[];
}

export interface IScrollTask {
  /** Scroll now (still waits for pending registrations). */
  execute: () => void;
  /** Cancel the scheduled scroll. */
  cancel: () => void;
  scrollPromise: Promise<void>;
}

export type AnchorsUpdatedCallback = (anchors: ISortedAnchor[]) => void;
