import * as React from 'react';
import { ComboBox, IComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { SelectableOptionMenuItemType } from '@fluentui/react/lib/SelectableOption';
import { Text as FluentText } from '@fluentui/react/lib/Text';
import { ITheme } from '@fluentui/react/lib/Styling';
import { Text } from '@microsoft/sp-core-library';
import * as strings from 'TableOfContentsWebPartStrings';

import {
  ISortedAnchor,
  cleanLinkTitle,
  getPageTheme,
  getSamePageAnchorId,
  usePredictedAnchors
} from '../../../anchors';
import { getTitleErrorMessage } from './linkValidationMessages';

/**
 * Title field for PropertyFieldCollectionData (onCustomRender).
 * Port of the existing LinkDropdown to Fluent UI 8:
 *  - suggests all anchors of the page (headings + inline anchors) – via usePredictedAnchors,
 *    i.e. with the same logic and the same IDs as the anchor service in view mode
 *  - a selection sets title AND URL ("#id"), free text is still possible
 *  - reports validation to the collection via onCustomFieldValidation
 *  - warns (without blocking) if a "#id" of the current page cannot be found here
 */
export interface ILinkDropdownProps {
  field: { id: string; required?: boolean };
  value: string;
  /** Current item (for URL and level). */
  item: any;
  onUpdate: (fieldId: string, value: any) => void;
  onCustomFieldValidation?: (fieldId: string, errorMessage: string) => void;
  /** Theme for the colors (default: page theme). */
  theme?: ITheme;
}

const HEADER_HEADINGS: string = 'header:headings';
const HEADER_INLINE: string = 'header:inline';

/** One line, cut off with an ellipsis (the full text is in the title attribute). */
const ellipsisStyle: React.CSSProperties = {
  display: 'block',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap'
};

/** Hints wrap, break inside long words/IDs and are limited to two lines so they cannot squeeze the URL field. */
const clampStyle: React.CSSProperties = {
  minWidth: 0,
  maxWidth: '100%',
  overflow: 'hidden',
  overflowWrap: 'anywhere',
  wordBreak: 'break-word',
  display: '-webkit-box',
  WebkitLineClamp: 2,
  WebkitBoxOrient: 'vertical'
};

function buildOptions(anchors: ISortedAnchor[], filterText: string): IComboBoxOption[] {
  const needle: string = (filterText || '').toLowerCase();
  const matches = (a: ISortedAnchor): boolean =>
    !needle || a.anchorText.toLowerCase().indexOf(needle) >= 0 || a.anchorId.toLowerCase().indexOf(needle) >= 0;

  const headings: ISortedAnchor[] = anchors.filter((a: ISortedAnchor) => a.kind === 'heading' && matches(a));
  const inline: ISortedAnchor[] = anchors.filter((a: ISortedAnchor) => a.kind === 'inline' && matches(a));
  const options: IComboBoxOption[] = [];

  if (headings.length > 0) {
    options.push({ key: HEADER_HEADINGS, text: strings.DropdownHeadingsHeader, itemType: SelectableOptionMenuItemType.Header });
    // Insert headings
    headings.forEach((a: ISortedAnchor) => options.push({ key: `h:${a.anchorId}`, text: a.anchorText, data: a }));
  }
  if (inline.length > 0) {
    options.push({ key: HEADER_INLINE, text: strings.DropdownInlineHeader, itemType: SelectableOptionMenuItemType.Header });
    // Insert inline text anchors
    inline.forEach((a: ISortedAnchor) => options.push({ key: `i:${a.anchorId}`, text: a.anchorId, data: a }));
  }
  return options;
}

export const LinkDropdown: React.FC<ILinkDropdownProps> = (props: ILinkDropdownProps) => {
  const { field, value, item, onUpdate, onCustomFieldValidation, theme: themeProp } = props;

  // All anchors of the page (the master list) – updates on changes in the editor
  const { anchors, loaded } = usePredictedAnchors(true);
  const isLoading: boolean = !loaded;
  // Filter text; the displayed options are derived from it
  const [filterText, setFilterText] = React.useState<string>('');
  const displayedOptions: IComboBoxOption[] = React.useMemo(
    () => buildOptions(anchors, filterText), [anchors, filterText]);
  const [currentText, setCurrentText] = React.useState<string>(value || '');
  const [touched, setTouched] = React.useState<boolean>(false);

  const theme: ITheme = React.useMemo(() => themeProp || getPageTheme(), [themeProp]);

  React.useEffect(() => {
    setCurrentText(value || '');
  }, [value]);

  // Report validation to PropertyFieldCollectionData ('' = valid)
  const errorMessage: string = getTitleErrorMessage(currentText);
  React.useEffect(() => {
    if (onCustomFieldValidation) {
      onCustomFieldValidation(field.id, errorMessage);
    }
  }, [errorMessage, field.id]);

  // Option selected (Enter/click) or free text committed (blur)
  const handleChange = (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, valueStr?: string): void => {
    setTouched(true);
    if (option && option.data) {
      const anchor: ISortedAnchor = option.data as ISortedAnchor;
      const title: string = cleanLinkTitle(anchor.anchorText) || anchor.anchorId;
      setCurrentText(title);
      onUpdate(field.id, title);
      onUpdate('url', `#${anchor.anchorId}`);
      if (anchor.kind === 'heading' && (!item || !item.level)) {
        onUpdate('level', Math.min(5, Math.max(1, anchor.level)));
      }
      setFilterText(''); // Resets the filter after selection
    } else {
      const text: string = valueStr !== undefined ? valueStr : currentText;
      setCurrentText(text);
      onUpdate(field.id, text);
    }
  };

  // Filter while typing (contains search, also mid-word)
  const handleInputValueChange = (text: string): void => {
    setCurrentText(text);
    setFilterText(text);
  };

  // Hint if the URL points to an anchor on this page that does not exist (yet)
  const samePageId: string | undefined = item ? getSamePageAnchorId(item.url) : undefined;
  const anchorMissing: boolean = !isLoading && !!samePageId &&
    !anchors.some((a: ISortedAnchor) => a.anchorId === samePageId);

  return (
    <div style={{ minWidth: 0, maxWidth: '100%' }}>
      <ComboBox
        options={displayedOptions}
        disabled={isLoading}
        placeholder={isLoading ? strings.DropdownLoading : strings.DropdownPlaceholder}
        allowFreeform={true}
        autoComplete='off'
        text={currentText}
        required={field.required}
        errorMessage={touched && errorMessage ? errorMessage : undefined}
        onChange={handleChange}
        onInputValueChange={handleInputValueChange}
        onMenuDismissed={() => setFilterText('')}
        onBlur={() => setTouched(true)}
        styles={{
          root: { minWidth: 0, maxWidth: '100%' },
          optionsContainerWrapper: { maxHeight: 320 },
          // The menu follows the width of the field; long entries are cut off instead of widening it
          callout: { maxWidth: 'min(480px, 90vw)' }
        }}
        onRenderOption={(option?: IComboBoxOption) => {
          if (!option) {
            return null;
          }
          const anchor: ISortedAnchor | undefined = option.data;
          if (!anchor) {
            return <span style={ellipsisStyle} title={option.text}>{option.text}</span>;
          }
          const suffix: string = anchor.kind === 'heading' ? ` #${anchor.anchorId}` : '';
          return (
            <span
              style={{
                ...ellipsisStyle,
                paddingLeft: anchor.kind === 'heading' ? Math.max(0, anchor.level - 1) * 10 : 0
              }}
              title={`${option.text}${suffix}`}>
              {option.text}
              {suffix && <span style={{ color: theme.semanticColors.bodySubtext, marginLeft: 6 }}>{suffix.trim()}</span>}
            </span>
          );
        }}
      />
      {!isLoading && anchors.length === 0 &&
        <div style={clampStyle}>
          <FluentText variant='small' styles={{ root: { color: theme.semanticColors.bodySubtext } }}>
            {strings.DropdownNoAnchors}
          </FluentText>
        </div>
      }
      {anchorMissing &&
        <div style={clampStyle} title={Text.format(strings.WarningAnchorNotOnPage, samePageId as string)}>
          <FluentText variant='small' styles={{ root: { color: theme.semanticColors.warningText || theme.palette.orangeLighter } }}>
            {Text.format(strings.WarningAnchorNotOnPage, samePageId as string)}
          </FluentText>
        </div>
      }
    </div>
  );
};
