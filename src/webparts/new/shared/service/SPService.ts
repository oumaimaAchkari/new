/*************************************************************************************AUTHOR: ESSORDI ABDELBASSET *************************************************************/

/******************************************************NEW + UPLOAD BUTTON POP UPS FOR UPLOADING OR COPYING FILE FROM DOCUMENT LIBRARY TEMPLATES TO WERVEN***********************************/

/*************SERVICE IMPORTS ***********************************************************************************************************************/
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { IDropdownOption } from "office-ui-fabric-react";
import { Dialog } from "@microsoft/sp-dialog";
/**********************SERVICE CLASS ***************************************************/
export class SPService {
  // GET THE PROFIT CENTER BY SITE'S NAME
  public async getSiteByName(SiteName: string): Promise<any> {
    let Site: any = { Name: '', ID: '' };
    return new Promise<any>(async (resolve, reject) => {
      let item = await sp.web.lists
        .getByTitle("Sites")
        .items
        .filter("Werfnaam eq'" + SiteName.trim() + "'").top(1)
        .get();
      item.map(x => {
        Site = {
          Name: x.Werfnaam, ID: x.ID
        };
      });

      resolve(Site);
    });
  }
  // GET THE DROPDOWN LISTS OPTIONS 
  public async getListSite(list: string): Promise<IDropdownOption[]> {
    let docSites: any[] = [];
    let codeOne: any[] = [{ key: "", value: "", text: "" }];
    //@ts-ignore
    let listSites: any[] = [{ key: "", value: "", text: "" }];
    let Items: any[] = [];
    return new Promise<IDropdownOption[]>(async (resolve, reject) => {
      if (list.toLowerCase().trim() === "templates") {
        Items = await sp.web.lists
          .getByTitle(list)
          .items.select("File/Name, File/ServerRelativeUrl")
          .expand("File")
          .get();

        Items.map((item) => {
          docSites.push({
            key: item.File.ServerRelativeUrl,
            value: item.File.ServerRelativeUrl,
            text: item.File.Name,
          });
        });

        resolve(docSites);
      } else if (list.toLowerCase().trim() === "mappenstructuur") {
        Items = await sp.web.lists.getByTitle('Mappenstructuur').items.select("*").get();
        Items.map((item) => {
          codeOne.push({
            key: item.ID,
            value: item.ID,
            text: item.Title
          });
        });
        resolve(codeOne);
      }
    });
  }
  /***********************************COPY FILE FROM TEMPLATE DOCUMENT LIBRARY TOWARDS THE PROFIT CENTERS DOCUMENT LIBRARY WITH META DATA ********************************/
  public async copyFile(replace: boolean, body?: any, siteName?: any) {
    let destUrl = "/sites/LCV/Werven";
    const _SiteId: number[] = [];
    _SiteId.push(body.Site);
    let _fileCreatedUrl = "";
    let _redirectUrl: string = "";
    let fileName = body.Name + "." + body.Template.split(".").pop();
    await sp.web
      .getFileByServerRelativePath(body.Template)
      .copyByPath(
        destUrl + "/" + body.Name + "." + body.Template.split(".").pop(),
        replace,
        false
      )
      .then(() => {
        sp.web.lists
          .getByTitle("Werven")
          .getItemsByCAMLQuery({
            ViewXml: `<View><Query><Where><Eq><FieldRef Name="FileLeafRef"/><Value Type="File">${fileName}</Value></Eq></Where></Query></View>`,
          })
          .then((x) => {
            sp.web.lists
              .getByTitle("Werven")
              .items.expand("File")
              .select("File/ServerRelativeUrl, EncodedAbsUrl")
              .orderBy("Created", false)
              .filter("Id eq '" + x[0].Id + "'")
              .get()
              .then((y) => {
                y.map((item) => {
                  _fileCreatedUrl += item.File.ServerRelativeUrl;
                  _redirectUrl += item.EncodedAbsUrl;
                });
              })
              .then(() => {
                sp.web
                  .getFileByServerRelativeUrl(_fileCreatedUrl)
                  .getItem()
                  .then((y) => {
                    y.validateUpdateListItem([
                      {
                        FieldName: "Site",
                        FieldValue: "" + siteName.ID,
                      },
                      {
                        FieldName: "Code1",
                        FieldValue: "" + _SiteId[0]
                      },
                    ]);
                  });
              })
              .then(() => {
                Dialog.alert("File Added Succesfully !").then(() => {
                  window.location.href = _redirectUrl + "?web=1";
                });
              });
          });
      })
      .catch((err) => {
      });

  }
  /********************************** UPLOAD FILE TO DOCUMENT LIBRARY PROFIT CENTER WITH META DATA INCLUDED *********************************************/
  public async CreateFile(replace: boolean, body?: any, siteName?: any, files?: any) {



    let destUrl = "/sites/LCV/Werven/";
    let _fileCreatedUrl = "";
    //@ts-ignore
    let _redirectUrl: string = "";

    const _SiteId: number[] = [];
    await _SiteId.push(body.Site);

    //foreach files
if(files != null)
{
    let i:any = 1;

     files.forEach(async element => {
      console.log("doumi file",element.name,"  ", element.name.split(".").pop());
      let fileName = "";

      if(files.length > 1 )
      {
         fileName = body.Name+ i + "." + element.name.split(".").pop();
         i++;
      }
      else
      {
        fileName = body.Name + "." + element.name.split(".").pop();
      }

      await sp.web
      .getFolderByServerRelativeUrl(destUrl)
      .files.add(fileName, element, replace)
      .then((x) => {
        let liste = sp.web.lists.getByTitle("Werven");
        liste
          .getItemsByCAMLQuery({
            ViewXml: `<View><Query><Where><Eq><FieldRef Name="FileLeafRef"/><Value Type="File">${fileName}</Value></Eq></Where></Query></View>`,
          })
          .then((d) => {
            sp.web.lists
              .getByTitle("Werven")
              .items.expand("File")
              .select("File/ServerRelativeUrl, EncodedAbsUrl")
              .orderBy("Created", false)
              .filter("Id eq '" + d[0].Id + "'")
              .get()
              .then((y) => {
                y.map((item) => {
                  _fileCreatedUrl = item.File.ServerRelativeUrl;
                  _redirectUrl = item.EncodedAbsU;
                });
              })
              .then(() => {
                sp.web
                  .getFileByServerRelativeUrl(_fileCreatedUrl)
                  .getItem()
                  .then((z) => {
                    z.validateUpdateListItem([
                      {
                        FieldName: "Site",
                        FieldValue: "" + siteName.ID,
                      },
                      {
                        FieldName: "Code1",
                        FieldValue: "" + _SiteId[0]
                      }
                    ]);
                  });
              });
          });
      })
      .then(() => {
        Dialog.alert("Added Successfully !").then(() => {
          // window.location.href=`https://dewaalpalen.sharepoint.com/sites/LCV/SitePages/Werfleiders.aspx`;
        });
      });

    });
  }
  
  }
  /****************************************CHECK IF FILE EXISTS ****************************************************/
  public _fileExists = async (body: any, siteName?: any): Promise<number> => {
    let fileName = body.file?.name
      ? body.Name + "." + body.file.name.split(".").pop()
      : body.Name + "." + body.Template.split("/").pop().split(".").pop();
    return new Promise<number>(async (resolve, reject) => {
      let liste = await sp.web.lists.getByTitle("Werven");
      liste
        .getItemsByCAMLQuery({
          ViewXml: `<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name="FileLeafRef"/><Value Type="File">${fileName}</Value></Eq></Where></Query></View>`,
        })
        .then((d) => {
          resolve(d.length);
        });
    });
  }
}
/**********************************************************************END OF FILE*********************************************************************************************/
