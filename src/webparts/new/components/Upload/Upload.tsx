/******************************************************** AUTHOR: ESSORDI ABDELBASSET ****************************************************************************************/
/********************************************** DRAG AND DROP IMPORTS*********************** */
//@ts-ignore
import { ISpfxReactDropzoneProps } from './ISpfxReactDropzoneProps';
//@ts-ignore
import { ISpfxReactDropzoneState } from './ISpfxReactDropzoneState';
//@ts-ignore
import { autobind } from 'office-ui-fabric-react/lib/Utilities';

//@ts-ignore

// import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";





//@ts-ignore
import { FilePond, registerPlugin } from 'react-filepond';
//@ts-ignore
import 'filepond/dist/filepond.min.css';
//@ts-ignore
import FilePondPluginImageExifOrientation from 'filepond-plugin-image-exif-orientation'
//@ts-ignore
import FilePondPluginImagePreview from 'filepond-plugin-image-preview'
//@ts-ignore
import 'filepond-plugin-image-preview/dist/filepond-plugin-image-preview.css'





/*****************************************************************FILE UPLOAD UI **********************************************************************************************/

/****************************COMPONENT IMPORTS ***************************/
import * as React from "react";
import styles from "./Upload.module.scss";
import { IUploadProps } from "./IUploadProps";
import { IUploadState } from "./IUploadState";
import { SPService } from "../../shared/service/SPService";
//@ts-ignore
import { TextField, MaskedTextField } from "@fluentui/react/lib/TextField";
import { Stack, IStackProps, IStackStyles } from "@fluentui/react/lib/Stack";
//@ts-ignore
import { Form, Formik, Field, FormikProps } from "formik";
import { Label } from "office-ui-fabric-react/lib/Label";
import * as yup from "yup";
import "semantic-ui-css/semantic.min.css";
import { Dropdown } from "semantic-ui-react";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import {
  //@ts-ignore
  DatePicker,
  // Dropdown,
  mergeStyleSets,
  PrimaryButton,
  IIconProps,
} from "office-ui-fabric-react";
import { sp } from "@pnp/sp";
import { DefaultButton, MessageBar, MessageBarType } from "@fluentui/react";
//@ts-ignore
import { extendWith } from "lodash";
/*********************************************************** COMPONENT STYLES ****************************************************************************/
//#region 
const modalPropsStyles = { main: { maxWidth: 450 } };
const modalProps = {
  isBlocking: true,
  styles: modalPropsStyles,
};

//@ts-ignore
const stackTokens = { childrenGap: 50 };
//@ts-ignore
const iconProps = { iconName: "Calendar" };
//@ts-ignore
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
//@ts-ignore
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } },
};
//@ts-ignore
const controlClass = mergeStyleSets({
  control: {
    margin: "0 0 15px 0",
    maxWidth: "300px",
  },
});
//#endregion
/********************************************************** CLASS COMPONENT ********************************************************************************/
export default class Upload extends React.Component<
  IUploadProps,
  IUploadState
> {
  //FORM BUTTON ICONS PROPS
  private cancelIcon: IIconProps = { iconName: "Clear" };
  private saveIcon: IIconProps = { iconName: "Save" };
  private _services: SPService = null;

  //CLASS CONSTRUCTOR
  constructor(props: Readonly<IUploadProps>) {
    super(props);
    //COMPONENT STATE INITIALIZATION
    this.state = {
      SiteName: [],
      code1: []
    };
    //PNP LIBRARY SETUP
    sp.setup({
      spfxContext: this.props.context as any,
    });
    //SERVICE INJECTION
    this._services = new SPService();
  }
  //SET UP POP FORM FIELDS
  private getFieldProps = (formik: FormikProps<any>, field: string) => {
    return {
      ...formik.getFieldProps(field),
      errorMessage: formik.errors[field] as string,
    };
  }
  //REACT LIFECYCLE HOOK 
  public componentDidMount(): void {
    //GETTING SITENAME FROM URL
    let urldecoded = decodeURIComponent(window.location.href.split('?').shift().split('/').pop());
    let siteName = urldecoded.substring(0, urldecoded.lastIndexOf('.'));
    this._services.getSiteByName(siteName).then((val) => {
      console.log('val', val);
      this.setState({
        SiteName: val
      });
    });
    //GET THE DROPDOWN LIST MAPPENSTRUCTUR
    this._services.getListSite("Mappenstructuur").then((result) => {
      // console.log('result',result);
      this.setState({
        code1: result,
      });
    });
  }
  //RENDER METHOD TO DISPLAY THE UI
  public render(): React.ReactElement<IUploadProps> {
    //YUP LIBRARY FOR FORM FIELDS VALIDATION
    const validate = yup.object().shape({
      file: yup.mixed().required("Please provide a file"),
      Site: yup.string().required("Please select Mappenstructuur"),
      Name: yup.string().required("file name is required"),
    });

    return (
      <Formik
        validateOnChange={false}
        validateOnBlur={false}
        initialValues={{
          file: undefined,
          Site: "",
          Name: "",
        }}
        validationSchema={validate}
        onSubmit={(values, helpers) => {
          /* //console.log("submit values", values);
          //console.log('date',new Date().getTime());
          const file = values.file;
          const Site = values.Site;
          const Name = values.Name.concat(" "+this.formatDate(new Date())); */
          /* let body = {
            file:file,
            Site: Site,
            Name: Name
          }; */
          this._services._fileExists(values, this.state.SiteName).then((x) => {

            //console.log("file exist lentgh", x);
            this.setState({ fileExist: x > 0 ? true : false });
          }).then(() => {
            //console.log("file exists", this.state.fileExist);
            if (!this.state.fileExist) {
              this._services.CreateFile(false, values, this.state.SiteName, this.state.files);
              this.props.closeForm();
            } else {
              this.setState({
                form: values,
                showDialog: true,
                dialogContentProps: {
                  type: DialogType.normal,
                  title: "File exists!",
                  subText:
                    'File with name "' +
                    values.Name +
                    "." +
                    values.file.name.split(".").pop() +
                    '" already exists !',
                },
              });
            }
          });

          ////console.log("values", values);

        }}
      >
        {(formik:any) => (
          <div className={styles.reactFormik}>
            <Stack>
              <Label className={styles.lblForm}>Current User</Label>
              <PeoplePicker
                context={this.props.context as any}
                personSelectionLimit={1}
                showtooltip={true}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                ensureUser={true}
                disabled={true}
                defaultSelectedUsers={[
                  this.props.context?.pageContext?.user.email as any,
                ]}
              />

             {/*  <Label className={styles.lblForm}>Upload File</Label>
              <input
                type="file"
                multiple={false}
                name="file"
                onChange={(event) => {
                  formik.setFieldValue("file", event.currentTarget.files[0]);
                }}
              /> */}
              




              <br />
              <br />
              <div /*className={styles.spfxReactDropzone}*/>
                <FilePond name="file" allowMultiple={true} onupdatefiles={fileItems => {
                  this.setState({
                    files: fileItems.map(fileItem => fileItem.file)
                  });
                  formik.setFieldValue("file", fileItems[0].file);
                }} />
                <br />
                
              </div>
              <div>
                {formik.touched.file && formik.errors.file ? (
                  <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                    onDismiss={() => {
                      formik.errors.file = undefined;
                    }}
                  >
                    {formik.errors.file}
                  </MessageBar>
                ) : undefined}
                
              </div>







              <Label className={styles.lblForm}>Mappenstructuur</Label>
              <Dropdown
                placeholder="select Mappenstructuur"
                required
                fluid
                search
                className={styles.myDropDown}
                scrolling

                selection
                options={this.state.code1}
                {...this.getFieldProps(formik, "Site")}
                onChange={(event, option) => {
                  formik.setFieldValue("Site", option.value);
                }}
              />
              <div>
                {formik.touched.Site && formik.errors.Site?.length > 0 ? (
                  <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                    onDismiss={() => {
                      formik.errors.Site = "";
                    }}
                  >
                    {formik.errors.Site}
                  </MessageBar>
                ) : undefined}
              </div>
              <Label className={styles.lblForm}>File Name</Label>
              <TextField
                autoComplete={"off"}
                {...this.getFieldProps(formik, "Name")}
              />
            </Stack>
            <PrimaryButton
              type="submit"
              text="Save"
              iconProps={this.saveIcon}
              className={styles.btnsForm}
              onClick={formik.handleSubmit as any}
            />
            <PrimaryButton
              text="Reset"
              iconProps={this.cancelIcon}
              className={styles.btnsForm}
              onClick={formik.handleReset as any}
            />
            <div>
              <Dialog
                hidden={!this.state.showDialog}
                dialogContentProps={this.state.dialogContentProps}
                modalProps={modalProps}
                onDismiss={() => this.setState({ showDialog: false })}
              >
                <DialogFooter>
                  <PrimaryButton
                    onClick={() => {
                      this._handleReplace(this.state.form);
                    }}
                    text="replace"
                  />
                  <DefaultButton
                    onClick={() => {
                      this._hadnleFileExists(this.state.form);
                    }}
                    text="keep them both"
                  />
                </DialogFooter>
              </Dialog>
            </div>
          </div>
        )}
      </Formik>
    );
  }
  //IF USER WANT TO REPLACE FILE THAT EXISTS
  private async _handleReplace(body?: any) {
    this.setState({
      showDialog: false,
    });
    await this._services.CreateFile(true, body, this.state.SiteName);
    await this.props.closeForm();
  }
  //IF USER WANT TO KEEP BOTH FILE INSTEAD OF REPLACING
  private async _hadnleFileExists(body?: any) {
    this.setState({
      showDialog: false,
    });
    const Name = body.Name.concat(" " + this.formatDate(new Date()));
    let newbody = {
      file: body.file,
      Name: Name,
      Site: body.Site,
    };
    this._services.CreateFile(false, newbody, this.state.SiteName);
    this.props.closeForm();
  }
  //MAKE DATE IN A READABLE FORMAT
  private formatDate = (date:any): string => {
    var date1 = new Date(date);

    var year = date1.getFullYear().toString();
    var month = (1 + date1.getMonth()).toString();
    month = month.length > 1 ? month : "0" + month;
    var day = date1.getDate().toString();
    day = day.length > 1 ? day : "0" + day;
    let hours = date1.getHours().toString();
    hours = hours.length > 1 ? hours : "0" + hours;
    let minutes = date1.getMinutes().toString();
    minutes = minutes.length > 1 ? minutes : "0" + minutes;
    let seconds = date1.getSeconds().toString();
    seconds = seconds.length > 1 ? seconds : "0" + seconds;
    return month + "-" + day + "-" + year + " " + hours + "-" + minutes + "-" + seconds;
  }
}
