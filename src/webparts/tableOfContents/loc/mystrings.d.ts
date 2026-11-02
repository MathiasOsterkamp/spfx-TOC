declare interface ITableOfContentsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ContentGroupName: string;
  AnchorsGroupName: string;
  EngineActiveInfo: string;
  EngineMissingHint: string;
  SourceFieldLabel: string;
  SourceAuto: string;
  SourceCustom: string;
  HideTitleFieldLabel: string;
  HideInMobileViewFieldLabel: string;
  ShowHeadingLevelsLabel: string;
  ShowHeading1FieldLabel: string;
  ShowHeading2FieldLabel: string;
  ShowHeading3FieldLabel: string;
  ShowHeading4FieldLabel: string;
  ShowHeading5FieldLabel: string;

  LinksFieldLabel: string;
  LinksFieldPanelHeader: string;
  LinksFieldPanelDescription: string;
  LinksFieldButtonLabel: string;
  LinksFieldSaveLabel: string;
  LinksFieldSaveAndAddLabel: string;
  LinksFieldCancelLabel: string;
  LinksColumnTitle: string;
  LinksColumnUrl: string;
  LinksColumnUrlPlaceholder: string;
  LinksColumnLevel: string;

  DropdownPlaceholder: string;
  DropdownLoading: string;
  DropdownHeadingsHeader: string;
  DropdownInlineHeader: string;
  DropdownNoAnchors: string;
  WarningAnchorNotOnPage: string;

  ErrorTitleRequired: string;
  ErrorTitleTooLong: string;
  ErrorTitleInvalidChars: string;
  ErrorUrlRequired: string;
  ErrorUrlTooLong: string;
  ErrorUrlInvalidChars: string;
  ErrorUrlUnsafeScheme: string;
  ErrorUrlNotAbsolute: string;
  ErrorUrlEmptyAnchor: string;
  ErrorUrlInvalidAnchor: string;

  ListStyle: string;
  ListStyleDefault: string;
  ListStyleDisc: string;
  ListStyleCircle: string;
  ListStyleSquare: string;
  ListStyleDecimal: string;
  ListStyleNone: string;
  TitleDefaultValue: string;
  TitlePlaceholder: string;
  NoHeadings: string;
  NoCustomLinks: string;
  EditModeHint: string;
}

declare module 'TableOfContentsWebPartStrings' {
  const strings: ITableOfContentsWebPartStrings;
  export = strings;
}
