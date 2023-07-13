/*****************************REACT COMPONENT STATES FILE *********************************/

/********DIALOG CONTENT POP UP PROPERTIES **************/
interface IdialogContentProps{
  type:any;
  title:string;
  subText:string;
}
/******************COMPONENT STATE INTERFACE ***********/
export interface IUploadState {
    form?: any;
    file?:any;
    SiteName?: any[];
    code1?:any;
    projectName?: string;
    dialogContentProps?:IdialogContentProps;
    showDialog?:boolean;
    fileExist?:boolean;
    files?: any[];
  }
/****************************************************END OF FILE ************************/

