import { Version } from '@microsoft/sp-core-library';
import { sp, Items, ItemVersion, Web } from "@pnp/sp";

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import {
  Environment,
  EnvironmentType
 } from '@microsoft/sp-core-library';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelpfulResourcesWebPart.module.scss';
import * as strings from 'HelpfulResourcesWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse   
 } from '@microsoft/sp-http';

 export interface ISPLists {
  value: ISPList[];
 }

 export interface ISPList {
  Title: string; // this is the department name in the List
  Id: string;
  AnncURL:string;
  DeptURL:string;
  CalURL:string;
  a85u:string; // this is the LINK URL
 }

  //global vars
  var userDept = "";

export interface IHelpfulResourcesWebPartProps {
  description: string;
}

export default class HelpfulResourcesWebPart extends BaseClientSideWebPart<IHelpfulResourcesWebPartProps> {

  // main promoise method, 1st we get the department, second chain is a REST Call to query the list
// third we get the list data and figure out the document libraries
getuser = new Promise((resolve,reject) => {
  // SharePoint PnP Rest Call to get the User Profile Properties
  return sp.profiles.myProperties.get().then(function(result) {
    var props = result.UserProfileProperties;
    var propValue = "";
    var userDepartment = "";

    props.forEach(function(prop) {
      //this call returns key/value pairs so we need to look for the Dept Key
      if(prop.Key == "Department"){
        // set our global var for the users Dept.
        userDept += prop.Value;
      }
    });
    return result;
  }).then((result) =>{
    this._getListData().then((response) =>{
      this._renderList(response.value);
    });
  });

});

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helpfulResources }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
            <h1>Helpful Resources</h1>
            <h3><div id="HelpfulRes"/></h3>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

    // main REST Call to the list...passing in the deaprtment into the call to 
  //return a single list item
  public _getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(`https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '`+ userDept +`'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }

   private _renderList(items: ISPList[]): void {
    let html: string = '';
    let libHTML: string ='';
  
    var siteURL = "";
    //list name
    var helpfulResources =  "";
    // items in the list
    var hrItems = "";

    items.forEach((item: ISPList) => {
      siteURL = item.DeptURL;
      helpfulResources = item.a85u;
    });
    //1st we need to override the current web to go to the department sites web
    const w = new Web("https://girlscoutsrv.sharepoint.com" + siteURL);
    
    // then use PnP to query the list
    w.lists.getByTitle(helpfulResources).items
    .get()
    .then((data) => {
      console.log(data);

      for (var x = 0; x < data.length; x++){
        //console.log(data[x].URL);
        //this gets the HTTP URL of the hyper link
        console.log(data[x].URL.Url);
        //this gets the text for the Hyperlink
        console.log(data[x].URL.Description);
        //hrItems += data[x].URL + '\r\n';
       // libHTML += `<p>${hrItems.toString()}</p>`;
      }
      //document.getElementById("HelpfulRes").innerText = hrItems;
  }).catch(e => { console.error(e); });

    const listContainer: Element = this.domElement.querySelector('#ListItems');
    listContainer.innerHTML = html;
  }

  // this is required to use the SharePoint PnP shorthand REST CALLS
  public onInit():Promise<void> {
    return super.onInit().then (_=> {
      sp.setup({
        spfxContext:this.context
      });
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
