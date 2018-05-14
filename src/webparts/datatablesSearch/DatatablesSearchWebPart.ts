import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import * as strings from 'DatatablesSearchWebPartStrings';
import DatatablesSearch from './components/DatatablesSearch';
import { IDatatablesSearchProps } from './components/IDatatablesSearchProps';
import { IDataTablesSearchWebPartProps } from './IDatatablesSearchWebPartProps';
import pnp from "sp-pnp-js";


export default class DatatablesSearchWebPart extends BaseClientSideWebPart<IDatatablesSearchProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

        pnp.setup({
        spfxContext: this.context
        });

    });
  }

  public render(): void {
    const element: React.ReactElement<IDatatablesSearchProps > = React.createElement(
      DatatablesSearch,
      {
        title: this.properties.title,
        query: this.properties.query,
        maxResults: this.properties.maxResults,
        sorting: this.properties.sorting,
        duplicates: this.properties.duplicates,
        privateGroups: this.properties.privateGroups,
        resulttype: this.properties.resulttype,
        context: this.context
    }
    );

    ReactDom.render(element, this.domElement);
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
                    groupName: strings.QueryGroupName,
                    groupFields: [
                        PropertyPaneTextField('query', {
                            label: strings.QueryFieldLabel,
                            description: strings.QueryFieldDescription,
                            multiline: true,
                            onGetErrorMessage: this._queryValidation,
                            deferredValidationTime: 500
                        }),
                        PropertyPaneSlider('maxResults', {
                            label: strings.FieldsMaxResults,
                            min: 1,
                            max: 500
                        }),
                        PropertyPaneDropdown('resulttype', {
                            label: strings.ResultTypeLabel,
                            options: [{
                                key: "project",
                                text: strings.ResultTypeProject
                            },{
                                key: "document",
                                text: strings.ResultTypeDocument
                            }],
                            selectedKey: "document"
                        }),
                        PropertyPaneTextField('sorting', {
                            label: strings.SortingFieldLabel
                        }),
                        PropertyPaneToggle('duplicates', {
                            label: strings.DuplicatesFieldLabel,
                            onText: strings.DuplicatesFieldLabelOn,
                            offText: strings.DuplicatesFieldLabelOff
                        }),
                        PropertyPaneToggle('privateGroups', {
                            label: strings.PrivateGroupsFieldLabel,
                            onText: strings.PrivateGroupsFieldLabelOn,
                            offText: strings.PrivateGroupsFieldLabelOff
                        })
                    ]
                }
            ]
        },
        {
            header: {
                description: strings.PropertyPaneDescription
            },
            groups: [
                {
                    groupName: strings.TemplateGroupName,
                    groupFields: [
                        PropertyPaneTextField('title', {
                            label: strings.TitleFieldLabel
                        })
                    ]
                }
            ]
        }
    ]
    };
  }
  /**
     * Validating the query property
     *
     * @param value
     */
    private _queryValidation(value: string): string {
      // Check if a URL is specified
      if (value.trim() === "") {
          return strings.QuertValidationEmpty;
      }

      return '';
  }

  /**
 * Prevent from changing the query on typing
 */
  protected get disableReactivePropertyChanges(): boolean {
      return true;
  }
}
