import * as React from 'react';
import { ComboBox, IComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { SelectableOptionMenuItemType } from '@fluentui/react/lib/SelectableOption';
import { Icon } from '@fluentui/react/lib/Icon';
import { ITheme, hiddenContentStyle } from '@fluentui/react/lib/Styling';
import { getId } from '@fluentui/react/lib/Utilities';
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

/**
 * The collection panel is an auto-layout table: every text in a cell widens its column and every
 * extra line makes the row taller. Messages are therefore never rendered as text below the field –
 * they are shown as an icon in a slot of fixed width (tooltip + screen reader text), so title and URL
 * keep their width and all rows keep the same height while validating.
 * (The collection itself shows URL errors only in its error callout, never below the field.)
 */
const HINT_SLOT_WIDTH: number = 24;

/** Read by screen readers only (same as hiddenContentStyle, as inline style). */
const visuallyHiddenStyle: React.CSSProperties = {
  position: 'absolute',
  width: 1,
  height: 1,
  margin: -1,
  padding: 0,
  border: 0,
  overflow: 'hidden',
  whiteSpace: 'nowrap'
};

interface IFieldHint {
  text: string;
  iconName: string;
  color: string;
}

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
  // Controlled selection, reset as soon as the text is edited. Otherwise the ComboBox ignores selecting
  // the same anchor again (no onChange for an unchanged index) and the edited title is not restored.
  const [selectedKey, setSelectedKey] = React.useState<string | null>(null);
  const hintId: string = React.useMemo(() => getId('linkDropdownHint'), []);

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
      setSelectedKey(option.key as string);
      setCurrentText(title);
      onUpdate(field.id, title);
      onUpdate('url', `#${anchor.anchorId}`);
      if (anchor.kind === 'heading' && (!item || !item.level)) {
        onUpdate('level', Math.min(5, Math.max(1, anchor.level)));
      }
      setFilterText(''); // Resets the filter after selection
    } else {
      const text: string = valueStr !== undefined ? valueStr : currentText;
      setSelectedKey(null);
      setCurrentText(text);
      onUpdate(field.id, text);
    }
  };

  // Filter while typing (contains search, also mid-word)
  const handleInputValueChange = (text: string): void => {
    setSelectedKey(null);
    setCurrentText(text);
    setFilterText(text);
  };

  // Hint if the URL points to an anchor on this page that does not exist (yet)
  const samePageId: string | undefined = item ? getSamePageAnchorId(item.url) : undefined;
  const anchorMissing: boolean = !isLoading && !!samePageId &&
    !anchors.some((a: ISortedAnchor) => a.anchorId === samePageId);

  // Errors: only the red border here, the message is shown in the error callout of the collection
  // (reported via onCustomFieldValidation). The icon slot is used for non-blocking hints only.
  const showError: boolean = touched && !!errorMessage;
  let hint: IFieldHint | undefined;
  if (anchorMissing) {
    hint = {
      text: Text.format(strings.WarningAnchorNotOnPage, samePageId as string),
      iconName: 'Warning',
      color: theme.semanticColors.warningText || theme.palette.orangeLighter
    };
  } else if (!isLoading && anchors.length === 0) {
    hint = { text: strings.DropdownNoAnchors, iconName: 'Info', color: theme.semanticColors.bodySubtext };
  }

  return (
    <div style={{ display: 'flex', alignItems: 'center', minWidth: 0, maxWidth: '100%' }}>
      <div style={{ flex: '1 1 auto', minWidth: 0 }}>
        <ComboBox
          options={displayedOptions}
          selectedKey={selectedKey}
          disabled={isLoading}
          placeholder={isLoading ? strings.DropdownLoading : strings.DropdownPlaceholder}
          allowFreeform={true}
          autoComplete='off'
          text={currentText}
          required={field.required}
          // Red border + role="alert" for screen readers; the text itself is visually hidden – it is
          // visible in the field validation (error callout) of PropertyFieldCollectionData
          errorMessage={showError ? errorMessage : undefined}
          ariaDescribedBy={hint ? hintId : undefined}
          onChange={handleChange}
          onInputValueChange={handleInputValueChange}
          onMenuDismissed={() => setFilterText('')}
          onBlur={() => setTouched(true)}
          styles={{
            root: { minWidth: 0, maxWidth: '100%' },
            errorMessage: hiddenContentStyle,
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
      </div>
      {/* Fixed slot: always reserved, so showing/hiding a hint never changes width or height */}
      <span style={{ flex: `0 0 ${HINT_SLOT_WIDTH}px`, width: HINT_SLOT_WIDTH, textAlign: 'center', lineHeight: 1 }}>
        {hint &&
          <Icon
            iconName={hint.iconName}
            title={hint.text}
            styles={{ root: { color: hint.color, fontSize: 16, cursor: 'default', verticalAlign: 'middle' } }}
          />
        }
        {hint &&
          <span id={hintId} style={visuallyHiddenStyle}>{hint.text}</span>
        }
      </span>
    </div>
  );
};
