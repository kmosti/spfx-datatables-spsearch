import * as React from 'react';

import styles from './DatatablesSearch.module.scss';
import {
  IDatatablesSearchProps,
  ISearchVisualizerState,
  IMetadata,
  IProjectSearchResult,
  IDocumentSearchResult
} from './IDatatablesSearchProps';
import TypeofHelper from "../helpers/TypeofHelper";
import SearchTokenHelper from "../helpers/SearchTokenHelper";
import { Spinner, SpinnerSize, MessageBar, MessageBarType, Dialog, DialogType } from 'office-ui-fabric-react';
import * as strings from 'DatatablesSearchWebPartStrings';
import * as uuidv4 from 'uuid/v4';
import * as moment from 'moment';
import * as jQuery from 'jquery';
require('../../../../node_modules/datatables.net');
import '../../../../node_modules/datatables.net-dt/css/jquery.dataTables.css';
import pnp, { SearchQueryBuilder, SearchResults, SearchQuery } from "sp-pnp-js";

const dtjQuery: any = (window as any).jQuery;

export default class DatatablesSearch extends React.Component<IDatatablesSearchProps, ISearchVisualizerState> {

  private _results: any[] = [];
  private _fields: string[] = [];
  private _templateMarkup: string = "";
  private _tmplDoc: Document;
  private _totalResults: number = 0;
  private _pageNr: number = 0;
  private _compId: string = "";
  private _tokenHelper: SearchTokenHelper;
  private _columns: string[];
  private _datatableConfig: any;

  constructor(props: IDatatablesSearchProps, state: ISearchVisualizerState) {
    super(props);

    // Specify a unique ID for the component
    this._compId = 'search-' + uuidv4();

    this._tokenHelper = new SearchTokenHelper(this.props.context);

    // Initialize the current component state
    this.state = {
      loading: true,
      template: "",
      error: "",
      showError: false
    };
  }

  public componentDidMount(): void {
    this._processSearchTasks();
  }

  public componentDidUpdate(prevProps: IDatatablesSearchProps, prevState: ISearchVisualizerState): void {
    // Check if the template needs to be updated
    if (prevProps.title !== this.props.title) {
        this._resetLoadingState();
        // Refresh template and search results
        this._processSearchTasks();
    } else if (prevProps.query !== this.props.query ||
        prevProps.maxResults !== this.props.maxResults ||
        prevProps.sorting !== this.props.sorting ||
        prevProps.duplicates !== this.props.duplicates ||
        JSON.stringify(prevProps.columns) !== JSON.stringify(this.props.columns) ||
        JSON.stringify(prevProps.SeachFields) !== JSON.stringify(this.props.SeachFields) ||
        prevProps.privateGroups !== this.props.privateGroups) {
          this._resetLoadingState();
          // Only refresh the search results
          this._processSearchTasks();
    }
  }

  public componentWillUnmount(): void {
    let table = this.refs[this._compId];
    let datatable = ( jQuery(table) as any ).DataTable();
    datatable.destroy();
  }

  private _resetLoadingState() {
    this._columns = [];
    // Reset state
    this.setState({
        loading: true,
        error: "",
        showError: false
    });
  }

  /**
   * Processing the search web part tasks
   */
  private _processSearchTasks(): void {
    this._processResults();
  }

  /**
   * Processing the search result retrieval process for documents
   */
  private _processResults() {

    let _searchQuerySettings: SearchQuery = {
      TrimDuplicates: this.props.duplicates,
      RowLimit: this.props.maxResults,
      SelectProperties: this.props.SeachFields,
      Properties: [{
        Name: "EnableDynamicGroups",
        Value: {
          BoolVal: this.props.privateGroups,
          QueryPropertyValueTypeIndex: 3
        }
      }]
    };

    if (this.props.sorting && this.props.sorting.split(":").length > 0) {
      let sortprop: string = this.props.sorting.split(":")[0];
      if (this.props.sorting.split(":")[1]) {
        let direction: string = this.props.sorting.split(":")[1];
        let directionEnum: number = null;
        switch( direction.toLowerCase() ) {
          case "ascending":
            directionEnum = 0;
            break;
          case "descending":
            directionEnum = 1;
            break;
        }

        _searchQuerySettings.SortList = [
          {
            Property: sortprop,
            Direction: directionEnum
          }
        ];
        _searchQuerySettings.EnableSorting = true;
      }
    }

    const query = !this._isEmptyString(this.props.query) ? `${this._tokenHelper.replaceTokens(this.props.query)}` : "*";
    const q = SearchQueryBuilder.create(query, _searchQuerySettings).rowLimit(this.props.maxResults);

    pnp.sp.search(q).then( (searchResp: SearchResults)  => {

      let itemsHtml: string = "";

      for ( let doc of searchResp.PrimarySearchResults as Array<IDocumentSearchResult> ) {
        itemsHtml += "<tr>";
        for ( let col of this.props.columns ) {

          if ( col.Type.toLowerCase() == "string" ) {
            if ( col.path.length > 0 ){
              let path: any = doc[col.MapTo] || encodeURI(doc.Path);
              itemsHtml += `
                <td data-search="${doc[col.MapTo]}">
                  <a href="${doc[col.path]}" class="${styles.dtLink}" title="${doc[col.MapTo]}">${doc[col.MapTo]}</a>
                </td>`;
            } else {
              itemsHtml += `<td>${doc[col.MapTo]}</td>`;
            }
          } else if (col.Type.toLowerCase() == "date"){
            itemsHtml += `
            <td data-order="${ moment( doc[col.MapTo] ).format("YYYYMMDDHHmm") }">
                ${moment( doc[col.MapTo] ).format("DD/MM/YY HH:mm")}
            </td>`;
          }
        }
        itemsHtml += "</tr>";
      }

      this.setState({
        loading: false,
        result: itemsHtml
      }, () => this.renderDatatables() );

    }).catch((error: any) => {
        this.setState({
            error: error.toString()
        });
    });

  }

  private renderDatatables() {
    let columnOrder: number = 0;
    this.props.columns.forEach( (col, i )=> {
      if (col.SortedBy == "true" && col.SortedBy !== "") {
        columnOrder = i;
      }
    });
    let table = this.refs[this._compId];
    let datatable = ( jQuery(table) as any ).DataTable({
      order: [[columnOrder, "desc"]],
      language: {
        lengthMenu: strings.lengthMenu,
        search: strings.search,
        info: strings.info,
        emptyTable: strings.infoEmpty,
        infoFiltered: strings.infoFiltered,
        paginate: {
            first:'First',
            last: 'Last',
            next: 'Next',
            previous: 'Previous'
        }
      }
    });
  }


  public render(): React.ReactElement<IDatatablesSearchProps> {

    let view = <Spinner size={SpinnerSize.large} label={strings.loadingMessage} />;
    // table compact hover display nowrap dt-responsive

    if (this.state.error !== "") {
      return (
          <MessageBar className={styles.error} messageBarType={MessageBarType.error}>
              <span>{strings.errorMessage}</span>
              {
                  (() => {
                      if (this.state.showError) {
                          return (
                              <div>
                                  <p>
                                      <a href="javascript:;" onClick={this._toggleError.bind(this)} className="ms-fontColor-neutralPrimary ms-font-m">
                                      <i className={`ms-Icon ms-Icon--ChevronUp ${styles.icon}`} aria-hidden="true"></i> {strings.hideErrorMessage}</a>
                                  </p>
                                  <p className="ms-font-m">{this.state.error}</p>
                              </div>
                          );
                      } else {
                          return (
                              <p>
                                  <a href="javascript:;" onClick={this._toggleError.bind(this)} className="ms-fontColor-neutralPrimary ms-font-m">
                                  <i className={`ms-Icon ms-Icon--ChevronDown ${styles.icon}`} aria-hidden="true"></i> {strings.showErrorMessage}</a>
                              </p>
                          );
                      }
                  })()
              }
          </MessageBar>
      );
    }

    if ( !this.state.loading ) {
      view = <div>
        <span className={styles.title}>{this.props.title}</span>
        <table ref={this._compId} className="compact hover">
          <thead>
              <tr>
                {this.props.columns.map( col => {
                  return <th>{col.Title}</th>;
                })}
              </tr>
          </thead>
          <tbody dangerouslySetInnerHTML={{__html: this.state.result}}></tbody>
        </table>
      </div>;

    }

    return (
      <div id={this._compId + "_container"} className={styles.searchVisualizer} ref={this._compId + "_container"}>
          {view}
          <Dialog isOpen={this.state.showScriptDialog} type={DialogType.normal} onDismiss={this._toggleDialog.bind(this)} title={strings.ScriptsDialogHeader} subText={strings.ScriptsDialogSubText}></Dialog>
      </div>
    );
  }
  /**
   * Toggle the show error message
   */
  private _toggleError() {
    this.setState({
        showError: !this.state.showError
    });
  }

  /**
   * Toggle the script dialog visibility
   */
  private _toggleDialog() {
      this.setState({
          showScriptDialog: !this.state.showScriptDialog
      });
  }

  /**
     * Check if the value is null, undefined or empty
     *
     * @param value
     */
    private _isEmptyString(value: string): boolean {
      return value === null || typeof value === "undefined" || !value.length;
  }
}
