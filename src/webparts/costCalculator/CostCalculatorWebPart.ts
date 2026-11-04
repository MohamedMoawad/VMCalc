import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CostCalculatorWebPartStrings';
import CostCalculator from './components/CostCalculator';
import { SharePointStore } from './sharepoint';

export interface ICostCalculatorWebPartProps {
  configListTitle: string;
  calculationsListTitle: string;
  currency: string;
  pageName: string;
}

export default class CostCalculatorWebPart extends BaseClientSideWebPart<ICostCalculatorWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _root: Root | undefined;
  private _store: SharePointStore | undefined;

  public render(): void {
    if (!this._store) {
      this._store = this._createStore();
    }

    const element: React.ReactElement = React.createElement(CostCalculator, {
      store: this._store,
      userDisplayName: this.context.pageContext.user.displayName,
      theme: this._isDarkTheme ? 'dark' : 'light'
    });

    if (!this._root) {
      this._root = createRoot(this.domElement);
    }
    this._root.render(element);
  }

  protected onInit(): Promise<void> {
    this._store = this._createStore();
    return super.onInit();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
  }

  protected onDispose(): void {
    this._root?.unmount();
    this._root = undefined;
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void { // eslint-disable-line @typescript-eslint/no-explicit-any
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this._store = this._createStore();
    this.render();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneTextField('configListTitle', {
                  label: strings.ConfigListTitleLabel
                }),
                PropertyPaneTextField('calculationsListTitle', {
                  label: strings.CalculationsListTitleLabel
                }),
                PropertyPaneTextField('currency', {
                  label: strings.CurrencyFieldLabel,
                  description: strings.CurrencyFieldDescription,
                  maxLength: 3
                }),
                PropertyPaneTextField('pageName', {
                  label: strings.PageNameLabel,
                  description: strings.PageNameDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _createStore(): SharePointStore {
    return new SharePointStore(this.context, {
      configListTitle: this.properties.configListTitle,
      calculationsListTitle: this.properties.calculationsListTitle,
      currency: this.properties.currency,
      pageName: this.properties.pageName,
      webPartId: this.manifest.id
    });
  }
}
