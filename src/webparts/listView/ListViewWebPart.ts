import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListViewWebPart.module.scss';
import * as strings from 'ListViewWebPartStrings';
import {
  SPHttpClient,
  SPHttpClientResponse,
} from '@microsoft/sp-http';

export interface IListViewWebPartProps {
  DropDownProp: string;
}

export interface spList{  
  Title:string;  
  Id: string;  
  }  

export interface spLists{  
    value: spList[];  
  }  


export default class ListViewWebPart extends BaseClientSideWebPart<IListViewWebPartProps> {
  private dropDownOptions: IPropertyPaneDropdownOption[] = []; 
  private listsDropdownDisabled: boolean = true;

  
  public render(): void {
    this.domElement.innerHTML = `
    <div id="box" class="${ styles.listView }">
        <div id="spListContainer" />
        <div class="${styles.icon}"><i class="ms-Icon ms-Icon--CustomList" aria-hidden="false"></i><br/></div>
        <div class="${styles.ddSelect}">
        Select a list to add to this page.
        </div>
    </div>`; 
    this.LoadViews();  
    console.log('Hello!!')
  }



  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }  


  protected onPropertyPaneConfigurationStart(): void {  
    // Stops execution, if the list values already exists  
   this.listsDropdownDisabled = !this.dropDownOptions;
   
    if(this.dropDownOptions.length>0){ 
     console.log('yes');  
     return;
   }

   this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'DropDownProp');


   // Calls function to append the list names to dropdown  
     this.GetLists();  
 
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        // this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
  
 }  



 private GetLists():void{  
  // REST API to pull the list names  
  let listresturl: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$select=Id,title&$filter=Hidden ne true`;  

  this.LoadLists(listresturl).then((response)=>{  
    // Render the data in the web part  
    this.LoadDropDownValues(response.value);  
  });  
}  

private LoadLists(listresturl:string): Promise<spLists>{  


  return this.context.spHttpClient.get(listresturl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse)=>{  
    return response.json();  
  });  
}  

private LoadDropDownValues(lists: spList[]): void{  
  lists.forEach((list:spList)=>{  
    // Loads the drop down values  
    this.dropDownOptions.push({"key":list.Title,"text":list.Title});
  });  
}


private GetPageUrls():Promise<any>{

  let path = this.context.pageContext.web.absoluteUrl + '/' + this.properties.DropDownProp.replace(/\s/g, '');
  
  console.log(path)
  
  let pageUrl = this.context.pageContext.web.absoluteUrl 
  + `/_api/search/query?querytext=%27path:${path}%20STS_ListItem_DocumentLibrary%20fileType:aspx%27&rowlimit=30&sortlist=%27ViewsLifetime:descending%27&selectproperties=%27Title,Description,Path%27`;

  return this.context.spHttpClient.get(pageUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse)=>{  
    return response.json();  
  }); 

}

private RenderPageUrls(items: any): any {
    
  let html: string = '';

  for(var i=0;i<items.length;i++){  
    
    let url = items[i].Cells[4]["Value"];
    let title = items[i].Cells[2]["Value"];
    let description = items[i].Cells[3]["Value"];

    html += `
              <div id="item" class="${styles.column}" draggable="true">
                  <a class="${styles.title} "href="${url}">${title}</a>
                  <div class="${styles.description}" >${description}</div>
              </div> 
   `
  
  };
  
  const listContainer: Element = this.domElement.querySelector('#spListContainer');
  listContainer.innerHTML = html;

  this.DragDrop();
}


private LoadPageUrls(): void {

  this.GetPageUrls().then((data)=>{

    let listItems = data.PrimaryQueryResult.RelevantResults.Table.Rows
    console.log(listItems)
    this.RenderPageUrls(listItems)

  })
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
              groupName:"",
              groupFields: [  
                PropertyPaneDropdown('DropDownProp',{  
                  label: "Select List to Display on the page",  
                  options: this.dropDownOptions,  
                  disabled: this.listsDropdownDisabled,
                  selectedKey: this.properties.DropDownProp  
                  
                })
            ]
            }
            
          ]
        }
      ]
    };
  }

  
  private GetMostViewed(): Promise<any> {

    // query site pages for ViewsLifetime, sort descending and select properties to filter results

    let absUrl = this.context.pageContext.site.absoluteUrl + '/SitePages'

    let url = this.context.pageContext.web.absoluteUrl + 
    `/_api/search/query?querytext=%27path:${absUrl} ShowInListView:yes%27&rowlimit=30&sortlist=%27ViewsLifetime:descending%27&selectproperties=%27DefaultEncodingUrl,%20Title,%20Description,%20promotedstate,ShowInListView,ProgId%27`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();        
      });
    }


  private RenderMostViewed(items: any): any {
    
  
    let html: string = '';

    for(var i=0;i<items.length;i++){  

// Reminder: ProgID index is 6 on connects

      if (items[i].Cells[6]['Value'] == 'SharePoint.Link' || items[i].Cells[5]['Value'] == 0){

        let url = items[i].Cells[2]["Value"];
        let title = items[i].Cells[3]["Value"];
        let description = items[i].Cells[4]["Value"];

     html += 
     `       
              <div id="item" class="${styles.column}" draggable="true" }>
                  <a class="${styles.title} "href="${url}">${title}</a>
                  <div class="${styles.description}" >${description}</div>
              </div>  
    `;  
      }

    };
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html; 

    this.DragDrop();
    
  }
  

  private DragDrop(){

    var dragSrcEl = null;

    function handleDragStart(e) {
      this.style.opacity = "0.4";
    
      dragSrcEl = this;
    
      e.dataTransfer.effectAllowed = "move";
    }
    
    function handleDragOver(e) {
      if (e.preventDefault) {
        e.preventDefault();
      }
    
      e.dataTransfer.dropEffect = "move";
      return false;
    }
    
    function handleDragEnter(e) {
      this.classList.add("over");
    }
    
    function handleDragLeave(e) {
      this.classList.remove("over");
    }
    
    function handleDrop(e) {
      if (e.stopPropagation) {
        e.stopPropagation();
      }
    
      if (dragSrcEl != this) {
        this.replaceWith(this, dragSrcEl);
      }
    
      return false;
    }
    
    function handleDragEnd(e) {
      this.style.opacity = "1";
    
      elements.forEach(function (element) {
        element.classList.remove("over");
      });
    }
   

    let elements = document.querySelectorAll('#spListContainer #item')
    elements.forEach(function (element) {
      element.addEventListener("dragstart", handleDragStart, false);
      element.addEventListener("dragenter", handleDragEnter, false);
      element.addEventListener("dragover", handleDragOver, false);
      element.addEventListener("dragleave", handleDragLeave, false);
      element.addEventListener("drop", handleDrop, false);
      element.addEventListener("dragend", handleDragEnd, false);
    });
  }

  
  private LoadViews(): void {

    this.GetMostViewed().then((data)=>{
     
      let listItems = data.PrimaryQueryResult.RelevantResults.Table.Rows

      console.log(listItems)

      if(this.properties.DropDownProp == 'Site Pages'){

        this.RenderMostViewed(listItems);  
        this.context.propertyPane.refresh();
      }
      else if(this.properties.DropDownProp.includes('links')){
        this.LoadPageUrls()
      }
    })
  }

}
