import { ITheme, mergeStyleSets } from '@fluentui/react/lib/Styling';

export const getStyles = (theme: ITheme) => {
  const dialogStyles = mergeStyleSets({
    tableOfContents: {
      '& a': {
        color: theme.semanticColors.actionLink,
        textDecoration: 'none',
        fontSize: 16,
      },
      '& a:hover': {
        color: theme.semanticColors.actionLink,
        textDecoration: 'underline',
      },
      '& a:visited': {
        color: theme.semanticColors.actionLink,
      },
      '& li': {
        margin: '4px 0',
      },
      '& li::marker': {
        color: theme.semanticColors.primaryButtonBackground,
      },
    },
    firstUl: {
      paddingLeft: 20,
      marginTop: 0,
      marginBottom: 0,
    },
    subUl: {
      paddingLeft: 20,
      marginTop: 4,
      marginBottom: 0,
    },
    // Numbered list with hierarchical numbers: 1., 1.1, 1.1.1 … (CSS counter, one per <ul> level)
    decimalList: {
      listStyleType: 'none',
      counterReset: 'toc-item',
      '& > li': {
        counterIncrement: 'toc-item',
      },
      '& > li::before': {
        content: 'counters(toc-item, ".") ". "',
        color: theme.semanticColors.primaryButtonBackground,
        fontVariantNumeric: 'tabular-nums',
      },
      // Sub levels without the trailing dot: "1.1 "
      '& ul > li::before': {
        content: 'counters(toc-item, ".") " "',
      },
    },
    backItem: {
      borderTop: '1px solid',
      borderBottom: '1px solid',
      borderColor: theme.semanticColors.menuDivider,
      marginBottom: 10,
    },

    // hide on medium screens and smaller
    hideInMobileView: {
      '@media (max-width: 600px)': {
        display: 'none !important',
      },
    },

    title: {
      minHeight: 'auto',
      '& h2': {
        marginBottom: 0,
      },
    },

    hideTitle: {
      display: 'none !important',
    },

    empty: {
      ...theme.fonts.small,
      color: theme.semanticColors.bodySubtext,
      fontStyle: 'italic',
    },

    editHint: {
      ...theme.fonts.small,
      color: theme.semanticColors.bodySubtext,
      fontStyle: 'italic',
      marginTop: 8,
    },
  });
  return dialogStyles;
};
