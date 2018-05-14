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
        prevProps.resulttype !== this.props.resulttype ||
        prevProps.duplicates !== this.props.duplicates ||
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
    // Retrieve the next set of results
    if ( this.props.resulttype == "document") {
        this._processResults();
    } else {
        this._processProjectResults();
    }
  }

  /**
   * Processing the search result retrieval process for documents
   */
  private _processResults() {

    let _searchQuerySettings: SearchQuery = {
      TrimDuplicates: this.props.duplicates,
      RowLimit: this.props.maxResults,
      SelectProperties: ["FileExtension","AuthorOWSUSER","ModifiedBy","Filename","CreatedBy","LastModifiedTime","path","ServerRedirectedURL"],
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
      let path: any = doc.ServerRedirectedURL || encodeURI(doc.path);
      let croppedTitle: string = "";
      let modifiedDate: moment.Moment = moment(doc.LastModifiedTime, "YYYY-MM-DDTHH:mm:ss.SSSSSSSZ");

      if ( doc.Filename.length > 45 ) {
        croppedTitle = doc.Filename.substr(0,45) + "...";
      } else {
          croppedTitle = doc.Filename;
      }
      itemsHtml += `
        <tr>
            <td data-search="${doc.Filename}">
                <a href="${path}" class="${styles.dtLink}" title="${doc.Filename}">${croppedTitle}</a>
            </td>
            <td data-order="${modifiedDate.format("YYYYMMDDHHmm")}">
                ${modifiedDate.format("DD/MM/YY HH:mm")}
            </td>
            <td>
                ${doc.ModifiedBy}
            </td>
        </tr>`;
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

  /**
   * Processing the search result retrieval process for projects
   */
  private _processProjectResults() {

    let _searchQuerySettings: SearchQuery = {
      TrimDuplicates: this.props.duplicates,
      RowLimit: this.props.maxResults,
      SelectProperties: ["Title", "SPWebUrl", "AarbakkeProjectnr", "AarbakkeProjectOwner", "AarbakkeProjectStart"],
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

    const query = !this._isEmptyString(this.props.query) ? `${this._tokenHelper.replaceTokens(this.props.query)}` : "'*'";

    const q = SearchQueryBuilder.create(query, _searchQuerySettings).rowLimit(this.props.maxResults);

    pnp.sp.search(q).then( (searchResp: SearchResults)  => {

      let itemsHtml: string = "";

      for ( let doc of searchResp.PrimarySearchResults as Array<IProjectSearchResult> ) {
        let path: any = doc.SPWebUrl || encodeURI(doc.Path);
        let croppedTitle: string = "";
        let prosjektid: string = doc.AarbakkeProjectnr || "No number defined";
        let prosjekteier: string = doc.AarbakkeProjectOwner || "No owner defined";
        let prosjektstart: moment.Moment = null;
        let dateValid: boolean = true;

        if ( doc.Title.length > 45 ) {
          croppedTitle = doc.Title.substr(0,45) + "...";
        } else {
            croppedTitle = doc.Title;
        }
        if ( doc.AarbakkeProjectStart && moment(doc.AarbakkeProjectStart).isValid() ) {
          prosjektstart = moment(doc.AarbakkeProjectStart);
        } else {
          dateValid = false;
        }
        itemsHtml += `
          <tr>
              <td data-search="${doc.Title}">
                  <a href="${path}" class="${styles.dtLink}" title="${doc.Title}">${croppedTitle}</a>
              </td>
              <td data-order="${prosjektstart.format("YYYYMMDDHHmm")}">
                  ${prosjektstart.format("DD/MM/YY")}
              </td>
              <td>
                  ${prosjekteier}
              </td>
              <td>
                  ${prosjektid}
              </td>
          </tr>`;
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
    let table = this.refs[this._compId];
    /* tslint:disable */
    let datatable = ( jQuery(table) as any ).DataTable({
      order: [[1, "desc"]],
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
    /* tslint:enable */
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

      if (this.props.resulttype == "document") {
        view = <div>
          <span className={styles.title}>{this.props.title}</span>
          <table ref={this._compId} className="compact hover">
            <thead>
                <tr>
                    <th>Title</th>
                    <th>Modified</th>
                    <th>Modified by</th>
                </tr>
            </thead>
            <tbody dangerouslySetInnerHTML={{__html: this.state.result}}></tbody>
          </table>
        </div>;
      } else {
        view = <div>
          <span className={styles.title}>{this.props.title}</span>
          <table ref={this._compId} className="compact hover">
            <thead>
                <tr>
                    <th>Title</th>
                    <th>Project start</th>
                    <th>Project owner</th>
                    <th>Project number</th>
                </tr>
            </thead>
            <tbody dangerouslySetInnerHTML={{__html: this.state.result}}></tbody>
          </table>
        </div>;
      }

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
