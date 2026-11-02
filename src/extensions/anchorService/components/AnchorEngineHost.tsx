import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import {
  IAnchorEngineConfig,
  PageAnchors,
  anchorEngine,
  defaultTocFilter,
  isPageInEditMode,
  pageNavigation,
  selectorForLevels,
  useDomChanges
} from '../../../anchors';

export interface IAnchorEngineHostProps {
  instanceId: string;
  config: IAnchorEngineConfig;
}

/**
 * Hosts the anchor engine on the page (rendered by the extension).
 *
 * An extension has no displayMode. Edit mode is therefore detected via:
 * the displayMode reported by web parts, ?Mode=Edit, contenteditable in the canvas and canvas classes.
 * In addition, the engine never modifies elements inside a contenteditable area.
 */
export const AnchorEngineHost: React.FC<IAnchorEngineHostProps> = (props: IAnchorEngineHostProps) => {
  const { instanceId, config } = props;

  const [isEdit, setIsEdit] = React.useState<boolean>(() => isPageInEditMode(config.contentSelector));

  const recheck = React.useCallback((): void => {
    setIsEdit(isPageInEditMode(config.contentSelector));
  }, [config.contentSelector]);

  // Web part reports edit/view
  React.useEffect(() => anchorEngine.subscribe(recheck), [recheck]);
  // Page change
  React.useEffect(() => pageNavigation.subscribe(recheck), [recheck]);
  // Back/forward
  React.useEffect(() => {
    window.addEventListener('popstate', recheck);
    return () => window.removeEventListener('popstate', recheck);
  }, [recheck]);
  // Editor appears/disappears (edit, publish, discard)
  useDomChanges(true, recheck, 150);

  const headingSelector: string = React.useMemo(() => selectorForLevels(config.headingLevels), [config.headingLevels]);

  return (
    <PageAnchors
      instanceId={instanceId}
      displayMode={isEdit ? DisplayMode.Edit : DisplayMode.Read}
      headingSelector={headingSelector}
      contentSelector={config.contentSelector}
      convertInlineAnchors={config.convertInlineAnchors}
      showAnchorIcons={config.showAnchorIcons}
      scrollToHashOnLoad={config.scrollToHashOnLoad}
      additionalFilter={defaultTocFilter} />
  );
};
