import { createTheme, getTheme, IPalette, ITheme } from '@fluentui/react/lib/Styling';

/**
 * Theme of the SharePoint page for Fluent UI 8.
 * SPFx 1.4.1 has no ThemeProvider service yet; however, SharePoint stores the page theme in
 * window.__themeState__.theme (load-themed-styles). A Fluent theme is created from it.
 */
export function getPageTheme(): ITheme {
  try {
    const state: any = (window as any).__themeState__;
    const slots: { [key: string]: string } | undefined = state && state.theme;
    if (slots) {
      const palette: Partial<IPalette> = {};
      Object.keys(slots).forEach((key: string) => {
        if (typeof slots[key] === 'string') {
          (palette as any)[key] = slots[key];
        }
      });
      return createTheme({ palette: palette });
    }
  } catch (e) {
    // fallback below
  }
  return getTheme();
}
