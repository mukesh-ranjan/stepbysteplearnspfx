import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EventRegistrationWpWebPart.module.scss';
import * as strings from 'EventRegistrationWpWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
require('bootstrap');

let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");

import {SPHttpClient,SPHttpClientResponse,ISPHttpClientOptions} from "@microsoft/sp-http"
import {IEventRegistration} from "./IEventRegistration" 
export interface IEventRegistrationWpWebPartProps {
  description: string;
}

export default class EventRegistrationWpWebPart extends BaseClientSideWebPart<IEventRegistrationWpWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

 

  public render(): void {
    this.domElement.innerHTML = `
    <div id="container" class="container">
    <div class="panel">
        <table border="4" class="table">
            <tr>
                <td>
                    Please Enter Registered User ID
                </td>
                <td>
                    <input type="text" id="txtID" class="form-control">
                </td>
                <td>
                    <input type="submit" id="btnSingleItemRead" value="Read Registered User Info" class="btn btn-primary buttons">
                </td>
            </tr>

           <tr>
               <td>User Name</td>
               <td><input type="text" id="txtUserName" class="form-control"></td>
           </tr> 
           <tr>
            <td>Email</td>
            <td><input type="email" id="txtEmail" class="form-control" ></td>
           </tr> 
           <tr>
              <td>Select Batch</td>
               <td>
                   <select name="Batch" id="ddlBatch" class="form-control">
                    <option value="Batch1">Batch 1</option>
                    <option value="Batch2">Batch 2</option>
                    <option value="Batch3">Batch 3</option>
                   </select>
               </td>
           </tr>
           <tr>
           <td>Select Level Of Knowledge</td>
            <td>
                <select name="LevelOfKnowledge" id="ddlLevelOfKnowledge" class="form-control">
                 <option value="Beginner">Beginner</option>
                 <option value="Intermediate">Intermediate</option>
                 <option value="Expert">Expert</option>
                </select>
            </td>
        </tr>
        <tr>
            <td>
                <input type="submit" value="Create" id="btnCreate" class="btn btn-primary buttons">
                <input type="submit" value="Read" id="btnRead" class="btn btn-primary buttons">
                <input type="submit" value="Update" id="btnUpdate" class="btn btn-primary buttons">
                <input type="submit" value="Delete" id="btnDelete" class="btn btn-primary buttons">
            </td>
        </tr>
        </table>
    </div>
    <div id="divStatus"/>

    <hr>
    <div id="listItems"/>
</div>
   `;

   this._bindAllEvents();
   
  }

  

  private CreateItem():void{

    var userName=document.getElementById("txtUserName")["value"];
    var email=document.getElementById("txtEmail")["value"];
    var batch=document.getElementById("ddlBatch")["value"];
    var levelOfKnowledge=document.getElementById("ddlLevelOfKnowledge")["value"];

    //step3

    const siteurl:string=this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('Even Registration')/items"

    const itemBody:any={
      "Title":userName,
    "Email":email,
    "Batch":batch,
    "LevelOfKnowledge":levelOfKnowledge
    }

    const spHttClentOptions:ISPHttpClientOptions={
      "body":JSON.stringify(itemBody)
    };

    //step4
   this.context.spHttpClient.post(siteurl,SPHttpClient.configurations.v1,spHttClentOptions)
   .then((response:SPHttpClientResponse)=>{
     if(response.status===201){
       let statusmessage:Element=this.domElement.querySelector("#divStatus");
       statusmessage.innerHTML="Item Created Successfully";
       
     }
     else{
       let statusmessage:Element=this.domElement.querySelector("#divStatus");
       statusmessage.innerHTML="An error has occured i.e." + response.status + "-" +response.statusText
     }
   })


  }

  private _getListItems():Promise<IEventRegistration[]>{

    const siteurl:string=this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('Even Registration')/items"

    return this.context.spHttpClient.get(siteurl,SPHttpClient.configurations.v1)
    .then((response)=>{
      return response.json();
    }).then((json)=>{
      return json.value;
    })as Promise<IEventRegistration[]>
  }

  private _getListItemById(Id:string):Promise<IEventRegistration>{

    const siteurl:string=this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('Even Registration')/items?$filter eq " +Id

    return this.context.spHttpClient.get(siteurl,SPHttpClient.configurations.v1)
    .then((response)=>{
      return response.json();
    }).then((json)=>{

      const item:any =json.value[0];
      const listItem:IEventRegistration = item as IEventRegistration
      return listItem;
    })as Promise<IEventRegistration>
  }

  //step 4

  private readItemById():void{
  
    let id:string= document.getElementById("txtID")["value"];

    this._getListItemById(id)
    .then(listItem=>{

      document.getElementById("txtUserName")["value"]=listItem.Title
      document.getElementById("txtEmail")["value"]=listItem.Email
      document.getElementById("ddlBatch")["value"]=listItem.Batch
      document.getElementById("ddlLevelOfKnowledge")["value"]=listItem.LevelOfKnowledge
    })
    .catch(error=>{
      let errMessage:Element = this.domElement.querySelector("#divStatus");

      errMessage.innerHTML="Read Failed"
    })

  }


  private readItems():void{
    console.log("Read Items")
    this._getListItems()
    .then((listItems)=>{
      let html:string='<table border=1 width=100% style="border-collapse:collapse;" class="table">'
      listItems.forEach((listItem)=>{
        html+=`<tr>
        <td>${listItem.Id}</td>
        <td>${listItem.Title}</td>
        <td>${listItem.Email}</td>
        <td>${listItem.Batch}</td>
        <td>${listItem.LevelOfKnowledge}</td>
        </tr>`

      });
      html+='</table>'

      const listContainer:Element =this.domElement.querySelector('#listItems');
      listContainer.innerHTML=html

    })
  }

  private updateListItem():void{

    //step 1

    var userName=document.getElementById("txtUserName")["value"];
    var email=document.getElementById("txtEmail")["value"];
    var batch=document.getElementById("ddlBatch")["value"];
    var levelOfKnowledge=document.getElementById("ddlLevelOfKnowledge")["value"];

    var Id=document.getElementById("txtID")["value"];
    
  //step 3

  const siteurl:string=this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('Even Registration')/items("+Id+")"
  
  const itemBody:any={
    "Title":userName,
  "Email":email,
  "Batch":batch,
  "LevelOfKnowledge":levelOfKnowledge
  }
 
  const headers:any={
    "X-HTTP-Method":"MERGE",
    "IF-MATCH":"*"
  };

  const spHttClentOptions:ISPHttpClientOptions={
    "headers":headers,
    "body":JSON.stringify(itemBody)
  };

    //step4
    this.context.spHttpClient.post(siteurl,SPHttpClient.configurations.v1,spHttClentOptions)
    .then((response:SPHttpClientResponse)=>{
      if(response.status===204){
        let statusmessage:Element=this.domElement.querySelector("#divStatus");
        statusmessage.innerHTML="Item Updated Successfully";
        
      }
      else{
        let statusmessage:Element=this.domElement.querySelector("#divStatus");
        statusmessage.innerHTML="An error has occured i.e." + response.status + "-" +response.statusText
      }
    })
 

  }

  private DeleteListItem(){
    var Id=document.getElementById("txtID")["value"];

    const siteurl:string=this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('Even Registration')/items("+Id+")"

    const headers:any={
      "X-HTTP-Method":"DELETE",
      "IF-MATCH":"*"
    };

    const spHttClentOptions:ISPHttpClientOptions={
      "headers":headers
      
    };

       //step4
       this.context.spHttpClient.post(siteurl,SPHttpClient.configurations.v1,spHttClentOptions)
       .then((response:SPHttpClientResponse)=>{
         if(response.status===204){
           let statusmessage:Element=this.domElement.querySelector("#divStatus");
           statusmessage.innerHTML="Item Deleted Successfully";
           
         }
         else{
           let statusmessage:Element=this.domElement.querySelector("#divStatus");
           statusmessage.innerHTML="An error has occured i.e." + response.status + "-" +response.statusText
         }
       })
  }

//step5

protected _bindAllEvents():void{

  this.domElement.querySelector('#btnCreate').addEventListener('click',()=>{
    this.CreateItem()
  });

  this.domElement.querySelector('#btnRead').addEventListener('click',()=>{
    this.readItems()
  });

  this.domElement.querySelector('#btnSingleItemRead').addEventListener('click',()=>{
    this.readItemById()
  });

  this.domElement.querySelector('#btnUpdate').addEventListener('click',()=>{
    this.updateListItem()
  });

  this.domElement.querySelector('#btnDelete').addEventListener('click',()=>{
    this.DeleteListItem()
  });
}


  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

 
  
  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
