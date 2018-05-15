declare interface IDatatablesSearchWebPartStrings {
  /* Fields */
  PropertyPaneDescription: string;
  QueryGroupName: string;
  TemplateGroupName: string;
  TitleFieldLabel: string;
  QueryFieldLabel: string;
  QueryFieldDescription: string;
  FieldsMaxResults: string;
  SortingFieldLabel: string;
  DebugFieldLabel: string;
  DebugFieldLabelOn: string;
  DebugFieldLabelOff: string;
  ExternalFieldLabel: string;
  ScriptloadingFieldLabel: string;
  ScriptloadingFieldLabelOn: string;
  ScriptloadingFieldLabelOff: string;
  DuplicatesFieldLabel: string;
  DuplicatesFieldLabelOn: string;
  DuplicatesFieldLabelOff: string;
  PrivateGroupsFieldLabel: string;
  PrivateGroupsFieldLabelOn: string;
  PrivateGroupsFieldLabelOff: string;
  ResultTypeLabel: string;
  ColumnsLabel: string;
  ColumnsHeaderText: string;
  SearchFieldsLabel: string;

  /* Validation */
  QuertValidationEmpty: string;
  TemplateValidationEmpty: string;
  TemplateValidationHTML: string;

  /* Dialog */
  ScriptsDialogHeader: string;
  ScriptsDialogSubText: string;

  /* Dropdownoptions */
  ResultTypeProject: string;
  ResultTypeDocument: string;

  /* Datatables Language options */
  lengthMenu: string;
  lengthMenuDocs: string;
  search: string;
  info: string;
  emptyTable: string;
  zeroRecords: string;
  infoFiltered: string;
  first: string;
  last: string;
  next: string;
  previous: string;
  infoEmpty: string;

  /* Datatables column titles */
  titleColumn: string;
  projectIdColumn: string;
  projectStartColumn: string;
  projectOwnerColumn: string;
  modifiedColumn: string;
  modifiedByColumn: string;

  /* Loading/error messages and similar */
  loadingMessage: string;
  errorMessage: string;
  showErrorMessage: string;
  hideErrorMessage: string;
}

declare module 'DatatablesSearchWebPartStrings' {
  const strings: IDatatablesSearchWebPartStrings;
  export = strings;
}
