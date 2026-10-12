
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDevFixProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
   context: WebPartContext;
}
