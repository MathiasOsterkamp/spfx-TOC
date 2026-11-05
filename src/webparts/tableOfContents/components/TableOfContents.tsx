import * as React from 'react';
import { DisplayMode, Log } from '@microsoft/sp-core-library';
import { Shimmer } from '@fluentui/react/lib/Shimmer';
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import * as strings from 'TableOfContentsWebPartStrings';

import { ITableOfContentsProps } from './ITableOfContentsProps';
import { Link, linksFromAnchors, linksFromCustom } from './Link';
import { getStyles } from './TableOfContents.styles';
import { WebPartTitle, getWebPartTitleStyles } from './WebPartTitle';
import {
  EXCLUDE_ATTR,
  ISortedAnchor,
  IScrollTask,
  anchorServiceManager,
  getPageTheme,
  useAnchorEngine,
  usePageVersion,
  usePredictedAnchors,
  useSortedAnchors
} from '../../../anchors';

const LOG_SOURCE: string = '[TableOfContents]';
/** Maximum wait time in milliseconds (4 seconds) */
const MAX_WAIT_TIME: number = 4000;
const LIST_STYLES: string[] = ['disc', 'circle', 'square', 'decimal', 'none'];

const TableOfContents: React.FC<ITableOfContentsProps> = (props: ITableOfContentsProps) => {
  const {
    hideTitle,
    titleTextValue,
    source,
    customLinks,

    showHeading1,
    showHeading2,
    showHeading3,
    showHeading4,
    showHeading5,
    webpartId,
    hideInMobileView,
    listStyleValue,

    theme,
    displayMode,
    setTitle
  } = props;

  const currentTheme = React.useMemo(() => theme || getPageTheme(), [theme]);
  const isEdit: boolean = displayMode === DisplayMode.Edit;

  const [currentTitle, setCurrentTitle] = React.useState(titleTextValue);
  const [timedOut, setTimedOut] = React.useState(false);

  const pageVersion: number = usePageVersion();
  // Anchor service (extension): is it active, which levels does it anchor?
  const { config: engineConfig } = useAnchorEngine();
  const { anchors, version } = useSortedAnchors();

  const styles = React.useMemo(() => getStyles(currentTheme), [currentTheme]);
  const webTitleStyles = React.useMemo(() => getWebPartTitleStyles(currentTheme), [currentTheme]);

  React.useEffect(() => {
    setCurrentTitle(titleTextValue);
  }, [titleTextValue]);

  // webpart properties and store have to be updated
  const handleTitleChange = (value: string): void => {
    setTitle(value);
    setCurrentTitle(value);
  };

  // Displayed levels = selected in the web part AND anchored by the extension
  const engineLevels: number[] = engineConfig ? engineConfig.headingLevels : [];
  const engineLevelsKey: string = engineLevels.join(',');
  const isLevelEnabled = React.useCallback((level: number): boolean => {
    if (engineLevels.indexOf(level) < 0) {
      return false;
    }
    switch (level) {
      case 1: return showHeading1;
      case 2: return showHeading2;
      case 3: return showHeading3;
      case 4: return showHeading4;
      case 5: return showHeading5;
      default: return false;
    }
  }, [showHeading1, showHeading2, showHeading3, showHeading4, showHeading5, engineLevelsKey]);

  // Loading indicator for at most 4 s – restarted after every page change
  React.useEffect(() => {
    setTimedOut(false);
    const timeoutId: number = window.setTimeout(() => setTimedOut(true), MAX_WAIT_TIME);
    // CLEANUP FUNCTION: Clears the timeout if the component unmounts or the page changes
    return () => window.clearTimeout(timeoutId);
  }, [pageVersion]);

  // ---------------------------------------------------------------------------
  // Anchors: view = list of the anchor service, edit = prediction using the same logic.
  // Both deliver the same IDs, so the preview already shows the real targets.
  // ---------------------------------------------------------------------------

  const predicted = usePredictedAnchors(isEdit && source !== 'custom');
  const pageAnchors: ISortedAnchor[] = isEdit ? predicted.anchors : anchors;

  const anchorTree: Link[] = React.useMemo(() => {
    const headings: ISortedAnchor[] = pageAnchors.filter((a: ISortedAnchor) =>
      (a.kind || 'heading') === 'heading' && isLevelEnabled(a.level));
    Log.verbose(LOG_SOURCE, `Dynamic links loaded (${headings.length})`);
    return linksFromAnchors(headings);
  }, [pageAnchors, isLevelEnabled]);

  const customTree: Link[] = React.useMemo(
    () => linksFromCustom(customLinks || []), [customLinks]);

  let links: Link[];
  let isReady: boolean = true;
  if (source === 'custom') {
    links = customTree;
  } else if (isEdit) {
    links = anchorTree;
    isReady = predicted.loaded;
  } else {
    links = anchorTree;
    isReady = version > 0 || timedOut;
  }

  /**
   * Returns a click handler that scrolls to the header (same page) or lets the browser navigate.
   */
  const scrollToHeader = (href: string) => {
    return (event: React.MouseEvent<HTMLAnchorElement>): void => {
      if (isEdit) {
        event.preventDefault();
        return;
      }
      try {
        // An <a> element resolves relative URLs (works without the URL class in IE11)
        const targetUrl: HTMLAnchorElement = document.createElement('a');
        targetUrl.href = href;

        const normalize = (path: string): string => (path.charAt(0) === '/' ? path : `/${path}`).toLowerCase();
        const isSamePage: boolean =
          targetUrl.protocol === window.location.protocol &&
          targetUrl.hostname.toLowerCase() === window.location.hostname.toLowerCase() &&
          normalize(targetUrl.pathname) === normalize(window.location.pathname);

        // IT IS A DIFFERENT PAGE: let the browser (SharePoint) navigate normally.
        // The page change is detected via navigatedEvent or the URL check.
        if (!isSamePage) {
          return;
        }

        // IT IS THE SAME PAGE: check if it has a hash/anchor
        if (targetUrl.hash && targetUrl.hash.length > 1) {
          let targetId: string = targetUrl.hash.substring(1);
          try {
            targetId = decodeURIComponent(targetId);
          } catch (e) {
            // use the raw value
          }

          // Prevent the default instant jump
          event.preventDefault();

          // Update the browser's location hash without a jump (the back button keeps working)
          const hash: string = `#${encodeURIComponent(targetId)}`;
          if (window.history && window.history.pushState) {
            window.history.pushState(window.history.state, '', hash);
          } else {
            window.location.hash = hash;
          }

          const isRegistered: boolean = anchorServiceManager.anchors.some((a: ISortedAnchor) => a.anchorId === targetId);
          const target: HTMLElement | null = document.getElementById(targetId);
          if (isRegistered || !target) {
            // Registered anchors (including hidden sections with onShow) via the service
            const task: IScrollTask | undefined = anchorServiceManager.scrollTo(targetId, -1);
            if (task) {
              task.execute();
            }
          } else {
            // Find the target element and scroll smoothly to it
            target.scrollIntoView({ behavior: 'smooth', block: 'start', inline: 'nearest' });
          }
        }
        // SAME PAGE BUT NO HASH: let the browser handle it.
      } catch (e) {
        // Fallback: let the browser handle the standard navigation.
        Log.error(LOG_SOURCE, e instanceof Error ? e : new Error(`Error parsing URL in scrollToHeader: ${e}`));
      }
    };
  };

  /**
   * Creates a list of components to display from a list of links.
   */
  const renderLinks = (wpId: string, currentLinks: Link[], listStyle: string, parentId: string): JSX.Element[] => {
    // For each link render a <li> element with a link. If the link has got childNodes, additionaly render <ul> with child links.
    return currentLinks.map((link: Link, index: number) => {
      const id: string = `${index}`;
      const key: string = link.id ? link.id : id;
      const keyField: string = `LinkField_${wpId}_${parentId}_${key}`;
      return (
        <li key={keyField} data-testid={keyField}>
          <a onClick={scrollToHeader(link.href)} href={link.href} data-interception='off'>
            {link.title}
          </a>
          {link.childNodes.length > 0
            ? <ul className={`${styles.subUl} ${listClassName}`} style={listInlineStyle}>{renderLinks(wpId, link.childNodes, listStyle, key)}</ul>
            : ''}
        </li>
      );
    });
  };

  // create components from a list of links
  const listStyle: string = LIST_STYLES.indexOf(listStyleValue) >= 0 ? listStyleValue : '';
  // "decimal" numbers hierarchically (1., 1.1, 1.1.1) via a CSS counter instead of restarting at 1 per level
  const isDecimal: boolean = listStyle === 'decimal';
  const listClassName: string = isDecimal ? styles.decimalList : '';
  const listInlineStyle: React.CSSProperties | undefined = isDecimal ? undefined : { listStyleType: listStyle };
  let toc: JSX.Element;
  if (!isReady) {
    toc = <Shimmer />;
  } else if (links.length === 0) {
    toc = <div className={styles.empty}>{source === 'custom' ? strings.NoCustomLinks : strings.NoHeadings}</div>;
  } else {
    toc = (
      <ul className={`${styles.firstUl} ${listClassName}`} style={listInlineStyle}>
        {renderLinks(webpartId, links, listStyle, 'root')}
      </ul>
    );
  }

  // add CSS class to hide in mobile view if needed
  const hideInMobileViewClass: string = hideInMobileView ? styles.hideInMobileView : '';

  // set title text
  const titleText: string = currentTitle ? currentTitle : strings.TitleDefaultValue;

  const excludeAttr: { [key: string]: string } = { [EXCLUDE_ATTR]: 'true' };

  return (
    <ThemeProvider theme={currentTheme}>
      <section className={styles.tableOfContents} id={webpartId} {...excludeAttr}>
        {/* Anchors are assigned by the anchor service (application customizer). Without it there are no anchors. */}
        {isEdit && !engineConfig &&
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
            {strings.EngineMissingHint}
          </MessageBar>
        }

        <div className={hideInMobileViewClass}>
          <nav data-toc-ignore='true'>
            {hideTitle === false ? (
              <WebPartTitle
                displayMode={displayMode}
                title={isEdit ? (currentTitle || '') : titleText}
                updateProperty={handleTitleChange}
                className={webTitleStyles.root}
                placeholder={strings.TitlePlaceholder}
                themeVariant={currentTheme}
              />
            ) : null}

            {toc}
            {isEdit && <div className={styles.editHint}>{strings.EditModeHint}</div>}
          </nav>
        </div>
      </section>
    </ThemeProvider>
  );
};

export default TableOfContents;
