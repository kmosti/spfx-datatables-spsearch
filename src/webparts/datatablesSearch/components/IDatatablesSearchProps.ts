import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { SearchResult } from 'sp-pnp-js';
import * as moment from 'moment';

export interface IDatatablesSearchProps {
    title: string;
    query: string;
    maxResults: number;
    sorting: string;
    duplicates: boolean;
    privateGroups: boolean;
    //resulttype: ResultType;
    context: IWebPartContext;
    columns?: any[];
    SeachFields: string[];
}

export type ResultType = "project" | "document";

export interface ISearchVisualizerState {
    loading?: boolean;
    template?: string;
    result?: string;
    error?: string;
    showError?: boolean;
    showScriptDialog?: boolean;
}

export interface IMetadata {
    fields: string[];
}

export interface SPUser {
    username?: string;
    displayName?: string;
    email?: string;
}

export interface IProjectSearchResult extends SearchResult {
    SPWebUrl: string;
    AarbakkeProjectnr: string;
    AarbakkeProjectOwner: string;
    AarbakkeProjectStart: Date;
}

export interface IDocumentSearchResult extends SearchResult {
  path: string;
  SPWebUrl: string;
  Filename: string;
  CreatedBy: string;
  ModifiedBy: string;
}
