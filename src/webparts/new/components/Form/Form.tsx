/******************************************************** AUTHOR: ESSORDI ABDELBASSET ****************************************************************************************/

/*****************************************************************FILE COPY UI **********************************************************************************************/

/****************************COMPONENT IMPORTS ***************************/
import * as React from "react";
import styles from "./Form.module.scss";
import { IFormProps } from "./IFormProps";
import { IFormState } from "./IFormState";
import { SPService } from "../../shared/service/SPService";
import { TextField } from "@fluentui/react/lib/TextField";
import { Stack, IStackProps, IStackStyles } from "@fluentui/react/lib/Stack";
import { Formik, FormikProps } from "formik";
import { Label } from "office-ui-fabric-react/lib/Label";
import * as yup from "yup";
import { Dropdown } from "semantic-ui-react";
import "semantic-ui-css/semantic.min.css";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  mergeStyleSets,
  PrimaryButton,
  IIconProps,
} from "office-ui-fabric-react";
import { DefaultButton, MessageBar, MessageBarType } from "@fluentui/react";
import { sp } from "@pnp/sp";
//@ts-ignore
import { lastIndexOf } from "lodash";
/********************************** COMPONENT STYLES ************************************/
//#region 
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
const modalPropsStyles = { main: { maxWidth: 450 } };
const modalProps = {
  isBlocking: true,
  styles: modalPropsStyles,
};
//#endregion
/********************************************************** CLASS COMPONENT ********************************************************************************/
export default class ReactFormik extends React.Component<
  IFormProps,
  IFormState
> {
  //FORM BUTTON ICONS PROPS
  private cancelIcon: IIconProps = { iconName: "Cancel" };
  private saveIcon: IIconProps = { iconName: "Save" };
  private _services: SPService = null;

  //CLASS CONSTRUCTOR
  constructor(props: Readonly<IFormProps>) {
    super(props);
    //COMPONENT STATES INITIALIZATION
    this.state = {
      templateFile: [],
      code1: [],
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
    //GET THE CURRENT SITE 
    this._services.getSiteByName(siteName).then((val) => {
      this.setState({
        SiteName: val
      });
    });
    //GET THE DROPDOWN LIST TEMPLATE OPTIONS
    this._services.getListSite("Templates").then((result) => {
      this.setState({
        templateFile: result,
      });
    });
    //GET THE DROPDOWN LIST MAPPENSTRUCTUR
    this._services.getListSite("Mappenstructuur").then((result) => {
      //console.log('result',result);
      this.setState({
        code1: result,
      });
    });
  }
  //RENDER METHOD TO DISPLAY THE UI
  public render(): React.ReactElement<IFormProps> {

    //YUP LIBRARY FOR FORM FIELDS VALIDATION
    const validate = yup.object().shape({
      Template: yup.string().required("Please select a template file"),
      Site: yup.string().required("Please select a Mappenstructuur"),
      Name: yup.string().required("file name is required"),
    });

    return (
      <Formik
        enableReinitialize
        validateOnChange={false}
        validateOnBlur={false}
        initialValues={{
          Template: "",
          Site: "",
          Name: "",
        }}
        validationSchema={validate}
        onSubmit={(values, helpers) => {
          this._services
            ._fileExists(values, this.state.SiteName)
            .then((x) => {
              //console.log('x exists', x);
              this.setState({ fileExist: x > 0 ? true : false });
            })
            .then(() => {
              if (!this.state.fileExist) {
                this._services.copyFile(false, values, this.state.SiteName);
                this.props.closeForm();
              } else {
                // this._services.copyFile(true);
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
                      values.Template.split(".").pop() +
                      '" already exists !',
                  },
                });
              }

            });

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

              <Label className={styles.lblForm}>Template File</Label>
              <Dropdown
                fluid
                placeholder="select a file template"
                search
                name="Template"
                scrolling
                className={styles.myDropDown}
                selection
                options={this.state.templateFile}
                {...this.getFieldProps(formik, "Template")}
                onChange={(event, option) => {
                  formik.setFieldValue("Template", option.value);
                }}
              />
              <div>
                {formik.touched.Site && formik.errors.Template?.length > 0 ? (
                  <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                    onDismiss={() => {
                      formik.errors.Template = "";
                    }}
                  >
                    {formik.errors.Template}
                  </MessageBar>
                ) : null}
              </div>
              <Label className={styles.lblForm}>Mappenstructuur</Label>
              <Dropdown
                placeholder="select Mappenstructuur"
                required
                fluid
                className={styles.myDropDown}
                // style={{'position':'relative'}}
                search
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
                ) : null}
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
              text="Cancel"
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
    this._services.copyFile(true, body, this.state.SiteName);
    this.props.closeForm();
  }
  //IF USER WANT TO KEEP BOTH FILES INSTEAD OF REPLACING
  private async _hadnleFileExists(body?: any) {
    this.setState({
      showDialog: false,
    });
    const Name = body.Name.concat(" " + this.formatDate(new Date()));
    let newbody = {
      Template: body.Template,
      Name: Name,
      Site: body.Site,
    };
    this._services.copyFile(false, newbody, this.state.SiteName);
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
/*********************************************************************** END OF FILE **********************************************************************************/