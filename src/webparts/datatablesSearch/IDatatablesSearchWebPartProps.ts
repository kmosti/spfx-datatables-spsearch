export interface IDataTablesSearchWebPartProps {
  title: string;
  query: string;
  maxResults: number;
  sorting: string;
  debug: boolean;
  scriptloading: boolean;
  duplicates: boolean;
  privateGroups: boolean;
  resulttype: ResultType;
}

export type ResultType = "project" | "document";
