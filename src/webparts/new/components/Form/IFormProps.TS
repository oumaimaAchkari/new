/*****************************REACT COMPONENT PROPS FILE *********************************/
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFormProps {
  closeForm:()=>void;
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
/****************************************************END OF FILE ************************/