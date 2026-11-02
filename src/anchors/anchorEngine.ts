import * as React from 'react';
import { getGlobal } from './globalRegistry';
import { getQuerySelector } from './inlineAnchors';
import { resolveContentRoot } from './contentRoot';

/**
 * Configuration of the anchor engine (comes from the extension's ClientSideComponentProperties).
 */
export interface IAnchorEngineConfig {
  /** Anchored heading levels, e.g. [1, 2, 3, 4, 5]. */
  headingLevels: number[];
  showAnchorIcons: boolean;
  convertInlineAnchors: boolean;
  scrollToHashOnLoad: boolean;
  /** Optional CSS selector of the content area. Empty = automatic. */
  contentSelector: string;
}

export const DEFAULT_ENGINE_CONFIG: IAnchorEngineConfig = {
  headingLevels: [1, 2, 3, 4, 5],
  showAnchorIcons: true,
  convertInlineAnchors: true,
  scrollToHashOnLoad: true,
  contentSelector: ''
};

/** Read properties defensively (JSON from elements.xml / tenant configuration may contain anything). */
export function normalizeEngineConfig(raw: any): IAnchorEngineConfig {
  const r: any = raw || {};
  let levels: number[] = DEFAULT_ENGINE_CONFIG.headingLevels;
  if (Array.isArray(r.headingLevels) || typeof r.headingLevels === 'string' || typeof r.headingLevels === 'number') {
    const list: any[] = Array.isArray(r.headingLevels) ? r.headingLevels : String(r.headingLevels).split(/[,; ]+/);
    const parsed: number[] = list
      .map((v: any) => parseInt(v, 10))
      .filter((n: number) => n >= 1 && n <= 5)
      .filter((n: number, i: number, all: number[]) => all.indexOf(n) === i)
      .sort();
    levels = parsed.length > 0 ? parsed : DEFAULT_ENGINE_CONFIG.headingLevels;
  }
  const bool = (v: any, def: boolean): boolean =>
    v === undefined || v === null || v === '' ? def : (v === true || String(v).toLowerCase() === 'true');

  return {
    headingLevels: levels,
    showAnchorIcons: bool(r.showAnchorIcons, DEFAULT_ENGINE_CONFIG.showAnchorIcons),
    convertInlineAnchors: bool(r.convertInlineAnchors, DEFAULT_ENGINE_CONFIG.convertInlineAnchors),
    scrollToHashOnLoad: bool(r.scrollToHashOnLoad, DEFAULT_ENGINE_CONFIG.scrollToHashOnLoad),
    contentSelector: typeof r.contentSelector === 'string' ? r.contentSelector : ''
  };
}

export function selectorForLevels(levels: number[]): string {
  const has = (n: number): boolean => levels.indexOf(n) >= 0;
  return getQuerySelector(has(1), has(2), has(3), has(4), has(5));
}

type EngineListener = () => void;

/**
 * Page-wide state of the anchor engine:
 *  - is the extension active, and with which configuration?
 *  - do web parts report edit mode? (extensions have no displayMode)
 */
class AnchorEngineState {
  private _config: IAnchorEngineConfig | undefined;
  private _editReports: { [sourceId: string]: boolean } = {};
  private _listeners: EngineListener[] = [];

  /** Called by the extension in onInit. */
  public register(config: IAnchorEngineConfig): void {
    this._config = config;
    this._notify();
  }

  public unregister(): void {
    this._config = undefined;
    this._notify();
  }

  public get isRegistered(): boolean {
    return !!this._config;
  }

  /** Active extension configuration (undefined = extension not active). */
  public get config(): IAnchorEngineConfig | undefined {
    return this._config;
  }

  /** Heading selector used by the engine. */
  public get headingSelector(): string {
    return selectorForLevels((this._config || DEFAULT_ENGINE_CONFIG).headingLevels);
  }

  /** Web parts report their displayMode here (true = edit). */
  public reportEditMode(sourceId: string, isEdit: boolean): void {
    const before: boolean = this.anyReportedEdit;
    if (isEdit) {
      this._editReports[sourceId] = true;
    } else {
      delete this._editReports[sourceId];
    }
    if (before !== this.anyReportedEdit) {
      this._notify();
    }
  }

  public get anyReportedEdit(): boolean {
    return Object.keys(this._editReports).length > 0;
  }

  public subscribe(listener: EngineListener): () => void {
    this._listeners.push(listener);
    return () => {
      const i: number = this._listeners.indexOf(listener);
      if (i >= 0) {
        this._listeners.splice(i, 1);
      }
    };
  }

  private _notify(): void {
    this._listeners.slice().forEach((l: EngineListener) => l());
  }
}

export const anchorEngine: AnchorEngineState = getGlobal('anchorEngine', () => new AnchorEngineState());

// ---------------------------------------------------------------------------
// Detect edit mode (the extension has no displayMode)
// ---------------------------------------------------------------------------

/** Is the element inside an editable area (rich text editor)? Nothing is ever modified there. */
export function isInContentEditable(element: HTMLElement): boolean {
  let node: HTMLElement | null = element;
  while (node) {
    const ce: string | null = node.getAttribute ? node.getAttribute('contenteditable') : null;
    if (ce !== null && ce.toLowerCase() !== 'false') {
      return true;
    }
    node = node.parentElement;
  }
  return false;
}

/**
 * Is the page in edit mode?
 *  1. a web part reports displayMode = Edit
 *  2. the URL contains Mode=Edit (SharePoint sets it when editing)
 *  3. the content area contains an editable editor (contenteditable)
 *  4. typical edit classes of the canvas
 */
export function isPageInEditMode(contentSelector?: string): boolean {
  if (anchorEngine.anyReportedEdit) {
    return true;
  }
  if (/(^|[?&])mode=edit(&|$)/i.test(window.location.search)) {
    return true;
  }
  const root: HTMLElement = resolveContentRoot(contentSelector);
  if (root.querySelector('[contenteditable="true"], [contenteditable=""]')) {
    return true;
  }
  return !!document.querySelector('.CanvasZone--edit, .CanvasSection--edit, .ControlZone--edit');
}

/** Extension configuration (re-renders on change, e.g. when the extension loads later). */
export function useAnchorEngine(): { config: IAnchorEngineConfig | undefined; headingSelector: string } {
  const [, setVersion] = React.useState(0);
  React.useEffect(() => anchorEngine.subscribe(() => setVersion((v: number) => v + 1)), []);
  return { config: anchorEngine.config, headingSelector: anchorEngine.headingSelector };
}
