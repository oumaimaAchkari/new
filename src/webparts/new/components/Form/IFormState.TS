/*****************************REACT COMPONENT STATES FILE *********************************/

/********DIALOG CONTENT POP UP PROPERTIES **************/
interface IdialogContentProps {
  type: any;
  title: string;
  subText: string;
}
/******************COMPONENT STATE INTERFACE ***********/
export interface IFormState {
  form?: any;
  templateFile?: any;
  SiteName?: any;
  code1?:any;
  projectName?: string;
  dialogContentProps?: IdialogContentProps;
  showDialog?: boolean;
  fileExist?:boolean;
}
/****************************************************END OF FILE ************************/
