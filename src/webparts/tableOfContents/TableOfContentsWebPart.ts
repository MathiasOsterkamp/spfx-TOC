import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import {
  CustomCollectionFieldType,
  ICustomCollectionField,
  PropertyFieldCollectionData
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'TableOfContentsWebPartStrings';
import { anchorEngine, pageNavigation } from '../../anchors';
import TableOfContents from './components/TableOfContents';
import { ITableOfContentsProps, SourceType } from './components/ITableOfContentsProps';
import { ICustomLink, toCustomLinks } from './components/Link';
import { LinkDropdown } from './components/LinkDropdown';
import { getUrlErrorMessage } from './components/linkValidationMessages';

export interface ITableOfContentsWebPartProps {
  titleText: string;
  hideTitle: boolean;
  source: SourceType;
  /** Custom links from PropertyFieldCollectionData (text from the previous version is migrated on load). */
  links: ICustomLink[];
  showHeading1: boolean;
  showHeading2: boolean;
  showHeading3: boolean;
  showHeading4: boolean;
  showHeading5: boolean;
  hideInMobileView: boolean;
  listStyle: string;
}

export default class TableOfContentsWebPart extends BaseClientSideWebPart<ITableOfContentsWebPartProps> {
  private _detachNavigation: (() => void) | undefined;

  protected onInit(): Promise<void> {
    // Page changes (the anchor service attaches as well; attaching multiple times is harmless)
    this._detachNavigation = pageNavigation.attach(this.context);
    return super.onInit();
  }

  public render(): void {
    // In case the event was missed: check for a new page when rendering
    pageNavigation.checkForPageChange('render');

    // Report edit mode to the anchor service (extension) – extensions have no displayMode
    anchorEngine.reportEditMode(this.context.instanceId, this.displayMode === DisplayMode.Edit);

    const p: ITableOfContentsWebPartProps = this.properties;

    const element: React.ReactElement<ITableOfContentsProps> = React.createElement(TableOfContents, {
      webpartId: this.context.instanceId,
      displayMode: this.displayMode,

      hideTitle: !!p.hideTitle,
      titleTextValue: p.titleText,
      setTitle: this._updateTitleOnWebpartProperties,

      source: p.source === 'custom' ? 'custom' : 'auto',
      customLinks: toCustomLinks(p.links),

      showHeading1: !!p.showHeading1,
      showHeading2: !!p.showHeading2,
      showHeading3: !!p.showHeading3,
      showHeading4: !!p.showHeading4,
      showHeading5: !!p.showHeading5,

      hideInMobileView: p.hideInMobileView === true,
      listStyleValue: p.listStyle || 'default'
    });

    ReactDom.render(element, this.domElement);
  }

  /** Switching view <-> edit: re-render, the engine cleans up or rebuilds. */
  protected onDisplayModeChanged(oldDisplayMode: DisplayMode): void {
    this.render();
  }

  protected onDispose(): void {
    anchorEngine.reportEditMode(this.context.instanceId, false);
    ReactDom.unmountComponentAtNode(this.domElement);
    if (this._detachNavigation) {
      this._detachNavigation();
      this._detachNavigation = undefined;
    }
  }


  /** Fields of the custom links: title (dropdown with the page's anchors), URL, level. */
  private _getLinkFields(): ICustomCollectionField[] {
    return [
      {
        id: 'title',
        title: strings.LinksColumnTitle,
        type: CustomCollectionFieldType.custom,
        required: true,
        onCustomRender: (field: ICustomCollectionField, value: any, onUpdate: (fieldId: string, value: any) => void,
          item: any, itemId: string, onCustomFieldValidation: (fieldId: string, errorMessage: string) => void) =>
          React.createElement(LinkDropdown, {
            key: itemId,
            field: field,
            value: value,
            item: item,
            onUpdate: onUpdate,
            onCustomFieldValidation: onCustomFieldValidation
          })
      },
      {
        id: 'url',
        title: strings.LinksColumnUrl,
        type: CustomCollectionFieldType.string,
        required: true,
        placeholder: strings.LinksColumnUrlPlaceholder,
        deferredValidationTime: 300,
        onGetErrorMessage: (value: string) => getUrlErrorMessage(value)
      },
      {
        id: 'level',
        title: strings.LinksColumnLevel,
        type: CustomCollectionFieldType.dropdown,
        required: true,
        defaultValue: 1,
        options: [1, 2, 3, 4, 5].map((n: number) => ({ key: n, text: `${n}` }))
      }
    ];
  }

  private _updateTitleOnWebpartProperties = (title: string): void => {
    this.properties.titleText = title;
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    // Show different fields depending on the source
    if (propertyPath === 'source' && oldValue !== newValue) {
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let group: IPropertyPaneField<any>[];

    if (this.properties.source !== 'custom') {
      group = [
        PropertyPaneLabel('showHeadingLevelsLabel', {
          text: strings.ShowHeadingLevelsLabel
        }),
        PropertyPaneCheckbox('showHeading1', { text: strings.ShowHeading1FieldLabel }),
        PropertyPaneCheckbox('showHeading2', { text: strings.ShowHeading2FieldLabel }),
        PropertyPaneCheckbox('showHeading3', { text: strings.ShowHeading3FieldLabel }),
        PropertyPaneCheckbox('showHeading4', { text: strings.ShowHeading4FieldLabel }),
        PropertyPaneCheckbox('showHeading5', { text: strings.ShowHeading5FieldLabel })
      ];
    } else {
      group = [
        PropertyFieldCollectionData('links', {
          key: 'links',
          label: strings.LinksFieldLabel,
          panelHeader: strings.LinksFieldPanelHeader,
          panelDescription: strings.LinksFieldPanelDescription,
          manageBtnLabel: strings.LinksFieldButtonLabel,
          saveBtnLabel: strings.LinksFieldSaveLabel,
          saveAndAddBtnLabel: strings.LinksFieldSaveAndAddLabel,
          cancelBtnLabel: strings.LinksFieldCancelLabel,
          value: toCustomLinks(this.properties.links),
          enableSorting: true,
          fields: this._getLinkFields()
          // The PnP controls ship the property pane typings of newer SPFx versions (sp-property-pane);
          // at runtime it is the same field object, so only a type cast is needed here.
        }) as unknown as IPropertyPaneField<any>
      ];
    }

    group.push(PropertyPaneDropdown('listStyle', {
      label: strings.ListStyle,
      options: [
        { key: 'default', text: strings.ListStyleDefault },
        { key: 'disc', text: strings.ListStyleDisc },
        { key: 'circle', text: strings.ListStyleCircle },
        { key: 'square', text: strings.ListStyleSquare },
        { key: 'decimal', text: strings.ListStyleDecimal },
        { key: 'none', text: strings.ListStyleNone }
      ],
      selectedKey: 'default'
    }));

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('source', {
                  label: strings.SourceFieldLabel,
                  options: [
                    { key: 'auto', text: strings.SourceAuto },
                    { key: 'custom', text: strings.SourceCustom }
                  ],
                  selectedKey: 'auto'
                }),
                PropertyPaneToggle('hideTitle', { label: strings.HideTitleFieldLabel }),
                PropertyPaneToggle('hideInMobileView', { label: strings.HideInMobileViewFieldLabel })
              ]
            },
            {
              groupName: strings.ContentGroupName,
              groupFields: group
            },
            {
              groupName: strings.AnchorsGroupName,
              groupFields: [
                PropertyPaneLabel('anchorServiceInfo', {
                  text: anchorEngine.isRegistered
                    ? `${strings.EngineActiveInfo} ${(anchorEngine.config ? anchorEngine.config.headingLevels : []).map((n: number) => `H${n}`).join(', ')}`
                    : strings.EngineMissingHint
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
