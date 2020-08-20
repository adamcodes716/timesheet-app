import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITimesheetAppProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
