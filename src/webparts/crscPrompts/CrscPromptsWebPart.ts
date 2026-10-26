import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPFI, spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/behaviors/spfx';

import * as strings from 'CrscPromptsWebPartStrings';
import CrscPrompts from './components/CrscPrompts';
import CrscPromptsErrorBoundary from './components/CrscPromptsErrorBoundary';
import { ICrscPromptsProps } from './components/ICrscPromptsProps';
import { IListOption } from './models/ICrscPrompt';
import { ISharePointService } from './services/ISharePointService';
import SharePointService from './services/SharePointService';

export interface ICrscPromptsWebPartProps {
  promptListId: string;
  itemsPerPage: number;
  modalMinWidth: number;
  clearAllButtonColor: string;
  cardBorderColor: string;
  cardTitleColor: string;
  viewButtonTextColor: string;
  viewButtonBackgroundColor: string;
}

export default class CrscPromptsWebPart extends BaseClientSideWebPart<ICrscPromptsWebPartProps> {
  private _service!: ISharePointService;
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _listLoadError: string | undefined;

  public render(): void {
    const promptElement: React.ReactElement<ICrscPromptsProps> = React.createElement(CrscPrompts, {
        listId: this.properties.promptListId,
        itemsPerPage: Number(this.properties.itemsPerPage) || 50,
        modalMinWidth: this.getPositiveNumber(this.properties.modalMinWidth, 760),
        clearAllButtonColor: this.getHexColor(this.properties.clearAllButtonColor, '#d0d0d0'),
        cardBorderColor: this.getHexColor(this.properties.cardBorderColor, '#0078d4'),
        cardTitleColor: this.getHexColor(this.properties.cardTitleColor, '#323130'),
        viewButtonTextColor: this.getHexColor(this.properties.viewButtonTextColor, '#ffffff'),
        viewButtonBackgroundColor: this.getHexColor(this.properties.viewButtonBackgroundColor, '#004578'),
        service: this._service,
        onConfigure: () => this.context.propertyPane.open()
      });
    const element: React.ReactElement = React.createElement(
      CrscPromptsErrorBoundary,
      {},
      promptElement
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      const sp: SPFI = spfi().using(SPFx(this.context));
      this._service = new SharePointService(sp);
    });
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme || !currentTheme.semanticColors) {
      return;
    }

    this.domElement.style.setProperty('--bodyText', currentTheme.semanticColors.bodyText || null);
    this.domElement.style.setProperty('--link', currentTheme.semanticColors.link || null);
    this.domElement.style.setProperty('--linkHovered', currentTheme.semanticColors.linkHovered || null);
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._listLoadError = undefined;
    this._service.getLists()
      .then(lists => {
        this._listOptions = this.toDropdownOptions(lists);
        this.context.propertyPane.refresh();
      })
      .catch((error: Error) => {
        this._listLoadError = error.message || strings.ListLoadError;
        Log.error('CrscPromptsWebPart', error);
        this.context.propertyPane.refresh();
      });
  }

  private toDropdownOptions(lists: IListOption[]): IPropertyPaneDropdownOption[] {
    return lists.map(list => ({ key: list.id, text: list.title }));
  }

  private getPositiveNumber(value: number, fallback: number): number {
    const parsedValue = Number(value);
    return isFinite(parsedValue) && parsedValue > 0 ? Math.floor(parsedValue) : fallback;
  }

  private getHexColor(value: string, fallback: string): string {
    const trimmed = (value || '').trim();
    if (/^(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(trimmed)) {
      return '#' + trimmed;
    }
    if (/^#(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(trimmed)) {
      return trimmed;
    }
    const css = (window as { CSS?: { supports?: (p: string, v: string) => boolean } }).CSS;
    return trimmed && css && css.supports && css.supports('color', trimmed) ? trimmed : fallback;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: this._listLoadError || strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('promptListId', {
                  label: strings.PromptListLabel,
                  options: this._listOptions,
                  selectedKey: this.properties.promptListId,
                  disabled: this._listOptions.length === 0
                }),
                PropertyPaneDropdown('itemsPerPage', {
                  label: strings.ItemsPerPageLabel,
                  options: [
                    { key: 5, text: '5' },
                    { key: 10, text: '10' },
                    { key: 50, text: '50' },
                    { key: 100, text: '100' }
                  ],
                  selectedKey: Number(this.properties.itemsPerPage) || 50
                }),
                PropertyPaneTextField('modalMinWidth', {
                  label: strings.ModalMinWidthLabel,
                  description: strings.ModalMinWidthDescription
                })
              ]
            },
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                PropertyPaneTextField('clearAllButtonColor', {
                  label: strings.ClearAllButtonColorLabel,
                  description: strings.ClearAllButtonColorDescription
                }),
                PropertyPaneTextField('cardBorderColor', {
                  label: strings.CardBorderColorLabel,
                  description: strings.CardBorderColorDescription
                }),
                PropertyPaneTextField('cardTitleColor', {
                  label: strings.CardTitleColorLabel,
                  description: strings.ColorValueDescription
                }),
                PropertyPaneTextField('viewButtonTextColor', {
                  label: strings.ViewButtonTextColorLabel,
                  description: strings.ColorValueDescription
                }),
                PropertyPaneTextField('viewButtonBackgroundColor', {
                  label: strings.ViewButtonBackgroundColorLabel,
                  description: strings.ColorValueDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
