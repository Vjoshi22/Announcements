import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAnnouncementsProps {
  description: string;
  currentContext:WebPartContext;
  listName:string;
  displayStyle:boolean;
  headerFont:number;
  contentFont:number;
  lines:number;
  color:string;
  headerColor:string;
}
