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
import { PropertyFieldCustomList, CustomListFieldType } from 'sp-client-custom-fields/lib/PropertyFieldCustomList';
import { PropertyFieldSearchPropertiesPicker } from 'sp-client-custom-fields/lib/PropertyFieldSearchPropertiesPicker';

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
        context: this.context,
        columns: this.properties.columns,
        SeachFields: this.properties.SeachFields
    }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private validateMaxResults(value: string): string {
    if ( value === null || value.length === 0 ) {
      return 'Please set a value';
    }
    if ( isNaN(Number(value)) ) {
      return "The value must be a number";
    }
    if ( Number(value) < 1 ) {
      return "The value cannot be less than 1";
    }

    return "";
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
                        PropertyPaneTextField('title', {
                            label: strings.TitleFieldLabel
                        }),
                        PropertyPaneTextField('query', {
                            label: strings.QueryFieldLabel,
                            description: strings.QueryFieldDescription,
                            multiline: true,
                            onGetErrorMessage: this._queryValidation,
                            deferredValidationTime: 500
                        }),
                        PropertyFieldSearchPropertiesPicker('SeachFields', {
                          label: strings.SearchFieldsLabel,
                          selectedProperties: this.properties.SeachFields,
                          loadingText: 'Loading...',
                          noResultsFoundText: 'No properties found',
                          suggestionsHeaderText: 'Suggested Properties',
                          disabled: false,
                          onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                          render: this.render.bind(this),
                          disableReactivePropertyChanges: this.disableReactivePropertyChanges,
                          properties: this.properties,
                          onGetErrorMessage: null,
                          deferredValidationTime: 0,
                          key: 'searchPropertiesFieldId'
                       }),
                        PropertyFieldCustomList('columns', {
                          label: strings.ColumnsLabel,
                          value: this.properties.columns,
                          headerText: strings.ColumnsHeaderText,
                          fields: [
                            { id: 'Title', title: 'Title', required: true, type: CustomListFieldType.string },
                            { id: 'Enable', title: 'Enable', required: true, type: CustomListFieldType.boolean },
                            { id: 'Type', title: 'Type', required: true, hidden: true, type: CustomListFieldType.string },
                            { id: 'SortedBy', title: 'Sort by', required: true, hidden: false, type: CustomListFieldType.boolean },
                            { id: 'MapTo', title: 'Map to search property', required: false, hidden: false, type: CustomListFieldType.string },
                            { id: 'path', title: 'URL field', required: false, hidden: false, type: CustomListFieldType.string }
                          ],
                          onPropertyChange: this.onPropertyPaneFieldChanged,
                          render: this.render.bind(this),
                          disableReactivePropertyChanges: true,//this.disableReactivePropertyChanges,
                          context: this.context,
                          properties: this.properties,
                          key: 'tilesMenuListField'
                        }),
                        PropertyPaneTextField('maxResults', {
                            label: strings.FieldsMaxResults,
                            onGetErrorMessage: this.validateMaxResults.bind(this),
                            validateOnFocusOut: true
                        }),
                        // PropertyPaneSlider('maxResults', {
                        //     label: strings.FieldsMaxResults,
                        //     min: 1,
                        //     max: 100000
                        // }),
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
