import { DisplayMode } from '@microsoft/sp-core-library';
import { ITheme } from '@fluentui/react/lib/Styling';
import { ICustomLink } from './Link';

export type SourceType = 'auto' | 'custom';

export interface ITableOfContentsProps {
  webpartId: string;
  displayMode: DisplayMode;
  theme?: ITheme;

  hideTitle: boolean;
  titleTextValue: string;
  /** Updates the web part property when the title is changed in edit mode. */
  setTitle: (value: string) => void;

  /** auto = headings of the page, custom = custom links */
  source: SourceType;
  /** Custom links (already normalized, see toCustomLinks). */
  customLinks: ICustomLink[];

  showHeading1: boolean;
  showHeading2: boolean;
  showHeading3: boolean;
  showHeading4: boolean;
  showHeading5: boolean;

  hideInMobileView: boolean;
  /** default | disc | circle | square | decimal | none */
  listStyleValue: string;
}
