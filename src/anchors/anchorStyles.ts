import { ITheme, mergeStyleSets } from '@fluentui/react/lib/Styling';
import { getPageTheme } from './pageTheme';
import { getGlobal } from './globalRegistry';

/**
 * Styles of the anchor icons (port of the original styles of the SharePoint AnchorService)
 * using mergeStyleSets instead of SCSS.
 *
 * RTL: merge-styles flips left/right margin/padding automatically, so no [dir] selectors are needed.
 */

/** Static class on the glyph so the focus border can be targeted by a selector. */
export const GLYPH_STATIC_CLASS: string = 'spfxAnchor-glyph';

/** Static class on the inline anchor link so the parent can show it on hover/focus. */
export const INLINE_LINK_STATIC_CLASS: string = 'spfxAnchor-inlineLink';

/** Theme independent marker on icons and placeholders (the generated class names differ per theme). */
export const ANCHOR_STATIC_CLASS: string = 'spfxAnchor-icon';

export interface IAnchorClassNames {
  wrapper: string;
  headingWrapper: string;
  anchor: string;
  inline: string;
  glyph: string;
  hidden: string;
  inlineAnchor: string;
  inlineLink: string;
}

/** Page-wide class sets per theme so cleanup also works across bundles. */
const holder: { cache: { [themeKey: string]: IAnchorClassNames } } =
  getGlobal('anchorClassNames.v4', () => ({ cache: {} }));

/** Class names for the given theme (default: page theme). Created once per set of theme colors used. */
export function getAnchorClassNames(theme?: ITheme): IAnchorClassNames {
  const t: ITheme = theme || getPageTheme();
  const focusBorder: string = t.semanticColors.focusBorder || t.palette.neutralSecondary;
  const key: string = `${t.palette.themePrimary}|${focusBorder}`;
  if (holder.cache[key]) {
    return holder.cache[key];
  }
  const visible: { opacity: number; transition: string } = { opacity: 1, transition: 'opacity .1s linear' };

  const created: IAnchorClassNames = mergeStyleSets({
    wrapper: {
      display: 'contents'
    },
    // Wraps heading + icon: positioning context for the absolute icon and hover area
    headingWrapper: {
      position: 'relative',
      selectors: {
        ':hover > a[data-spfx-anchor-id]': visible
      }
    },
    anchor: {
      opacity: 0,
      outline: 'none',
      position: 'absolute',
      zIndex: 1,
      whiteSpace: 'nowrap',
      cursor: 'pointer',
      textDecoration: 'none',
      color: t.palette.themePrimary,
      selectors: {
        ':focus': visible,
        ':focus > &': visible,
        ':hover > &': visible,
        [`:focus:not(:active) .${GLYPH_STATIC_CLASS}`]: {
          borderColor: focusBorder
        }
      }
    },
    inline: {
      marginLeft: 5
    },
    glyph: {
      display: 'inline-block',
      fontStyle: 'normal',
      fontWeight: 400,
      lineHeight: 1,
      padding: 3,
      verticalAlign: 'middle',
      border: '1px solid transparent',
      selectors: {
        svg: { display: 'block', width: 16, height: 16 }
      }
    },
    hidden: {
      height: 0
    },
    // Target created from an {anchor-#id} marker, contains a link icon that is shown on hover/focus
    inlineAnchor: {
      display: 'inline',
      selectors: {
        [`:hover > & .${INLINE_LINK_STATIC_CLASS}`]: visible,
        [`:focus-within .${INLINE_LINK_STATIC_CLASS}`]: visible
      }
    },
    // Modifier for the shared anchor icon inside an inline target: in text flow instead of absolute
    inlineLink: {
      position: 'static'
    }
  }) as IAnchorClassNames;

  holder.cache[key] = created;
  return created;
}
