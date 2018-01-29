import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  IWebPartContext 
} from '@microsoft/sp-webpart-base';

import {  
  SPHttpClient
} from '@microsoft/sp-http'; 

import * as jQuery from 'jquery';
import 'jqueryui';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as strings from 'AccordionViewWebPartStrings';
import AccordionView from './components/AccordionView';
import { IAccordionViewProps } from './components/IAccordionViewProps';
import SPHttpClientResponse from '@microsoft/sp-http/lib/spHttpClient/SPHttpClientResponse';
import styles from './components/AccordionView.module.scss';

export interface IAccordionViewWebPartProps {
  listName: string;
}

export interface spListItems{  
  value: spListItem[];  
}  
export interface spListItem{  
  Title: string;  
  Body: string;    
}  

export interface ISPList{  
  Title:string;  
  Id: string;  
}  
export interface ISPLists{  
  value: ISPList[];  
} 

export default class AccordionViewWebPart extends BaseClientSideWebPart<IAccordionViewWebPartProps> {
  
  
  private lists: IPropertyPaneDropdownOption[] = [];
  private listsDropdownDisabled: boolean = true;

  public constructor() {
    super();

    SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
  }

  public render(): void {
    const element: React.ReactElement<IAccordionViewProps > = React.createElement(
      AccordionView,
      {
        listName: this.properties.listName
      }
    );

    ReactDom.render(element, this.domElement);
    
    this.LoadData();    

    const accordionOptions: JQueryUI.AccordionOptions = {
      animate: true,
      collapsible: false,
      icons: {
        header: 'ui-icon-circle-arrow-e',
        activeHeader: 'ui-icon-circle-arrow-s'
      }
    };
    
    jQuery('.accordion', this.domElement).accordion(accordionOptions);
    
  }

  private LoadData(): void{  
    if(this.properties.listName != undefined){  
      let url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('"+this.properties.listName +"')/items?$select=Title,Body/Title";  
  
      
    this.GetListData(url).then((response)=>{  
      // Render the data in the web part  
      this.RenderListData(response.value);  
    }).catch((response)=> {
      this.domElement.querySelector(".accordion").innerHTML = "Error: " + JSON.stringify(response.value); 
    });      
    }  
  
  }  
  
  private GetListData(url: string): Promise<spListItems>{  
    // Retrieves data from SP list  
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse)=>{  
       return response.json();  
    });  
  }  
  private RenderListData(listItems: spListItem[]): void{       
    if(listItems != undefined)
    {
      let itemsHtml: string = ""; 
    // Displays the values in table rows  
    listItems.forEach((listItem: spListItem)=>{ 
      itemsHtml += `<h3>${listItem.Title}</h3>`;  
      itemsHtml += `<div><p>${listItem.Body}</p></div>`;       
    });  
    this.domElement.querySelector(".accordion").innerHTML = itemsHtml;  
    
    
  }

  }  
  

  protected onPropertyPaneConfigurationStart(): void {  
    this.listsDropdownDisabled = !this.lists;

    // Stops execution, if the list values already exists  
   if(this.lists.length > 0) return;  
   
   this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

   // Calls function to append the list names to dropdown  
   this.GetLists();    
  
 }  
   
 private GetLists():void{  
   // REST API to pull the list names  
   let listresturl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$select=Id,Title";  
   
   this.LoadLists(listresturl).then((response)=>{  
     // Render the data in the web part  
     this.LoadDropDownValues(response.value);  
     this.listsDropdownDisabled = false;
     this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
   });  
 }  
 private LoadLists(listresturl:string): Promise<ISPLists>{  
   return this.context.spHttpClient.get(listresturl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
     return response.json();  
   });  
 }  
   
 private LoadDropDownValues(lists: ISPList[]): void{  
  if(lists != undefined) {
   lists.forEach((list:ISPList)=>{  
     // Loads the drop down values  
     this.lists.push({key:list.Title,text:list.Title});  
   });  
  }
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
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
