import { SPHttpClient } from '@microsoft/sp-http';

export interface IMisEventVersionProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  spHttpClient: SPHttpClient; // Add the type for spHttpClient
  siteUrl: string;
}
