import { SPHttpClient } from '@microsoft/sp-http';

export interface IWpBotoneraDescargablesProps {
  siteUrl: string;
  spHttpClient: SPHttpClient;
 applyFilter?: boolean;
  filterFieldInternalName?: string; 
  filterValue?: string;             
}