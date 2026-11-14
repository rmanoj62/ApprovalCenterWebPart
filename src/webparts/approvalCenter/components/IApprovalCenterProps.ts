import { SPHttpClient } from '@microsoft/sp-http';

export interface IApprovalCenterProps {
  siteUrl: string;
  currentUserEmail: string;
  spHttpClient: SPHttpClient;
}