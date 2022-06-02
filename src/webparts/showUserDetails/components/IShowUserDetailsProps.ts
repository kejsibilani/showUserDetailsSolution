import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IShowUserDetailsProps {
  level: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  listName: string;
}
