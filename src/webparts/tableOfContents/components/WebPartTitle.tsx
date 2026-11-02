import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { ITheme, mergeStyleSets } from '@fluentui/react/lib/Styling';
import { getPageTheme } from '../../../anchors';


export interface IWebPartTitleProps {
  displayMode: DisplayMode;
  title: string;
  updateProperty: (value: string) => void;
  className?: string;
  placeholder?: string;
  themeVariant?: ITheme;
}

export const getWebPartTitleStyles = (theme: ITheme) => mergeStyleSets({
  root: {
    marginBottom: 8,
  },
  heading: {
    ...theme.fonts.xLarge,
    fontWeight: 600,
    margin: 0,
    color: theme.semanticColors.bodyText,
  },
  input: {
    ...theme.fonts.xLarge,
    fontWeight: 600,
    display: 'block',
    width: '100%',
    boxSizing: 'border-box',
    padding: 0,
    margin: 0,
    border: 'none',
    borderBottom: '1px solid transparent',
    outline: 'none',
    resize: 'none',
    overflow: 'hidden',
    background: 'transparent',
    color: theme.semanticColors.bodyText,
    selectors: {
      ':focus': { borderBottomColor: theme.palette.themePrimary },
      '::placeholder': { color: theme.semanticColors.inputPlaceholderText },
    },
  },
});

export const WebPartTitle: React.FC<IWebPartTitleProps> = (props: IWebPartTitleProps) => {
  const { displayMode, title, updateProperty, className, placeholder, themeVariant } = props;
  const styles = React.useMemo(() => getWebPartTitleStyles(themeVariant || getPageTheme()), [themeVariant]);

  if (displayMode === DisplayMode.Edit) {
    return (
      <div className={`${styles.root} ${className || ''}`}>
        <textarea
          className={styles.input}
          rows={1}
          placeholder={placeholder}
          aria-label={placeholder}
          value={title}
          onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => updateProperty(e.target.value)} />
      </div>
    );
  }

  return (
    <div className={`${styles.root} ${className || ''}`}>
      <h2 className={styles.heading}>{title}</h2>
    </div>
  );
};
