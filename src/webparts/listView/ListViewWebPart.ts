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
    console.log('Hello!')
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
    
    // console.log(items[i].Cells[4]["Value"])

   html += `
              <div id="item" class="${styles.column}" draggable="true">
                  <a class="${styles.title} "href="${items[i].Cells[4]["Value"]}">${items[i].Cells[2]["Value"]}</a>
                  <div class="${styles.description}" >${items[i].Cells[3]["Value"]}</div>
              </div> 
   `
  
  };
  
  const listContainer: Element = this.domElement.querySelector('#spListContainer');
  listContainer.innerHTML = html;

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

      if (items[i].Cells[7]['Value'] == 'SharePoint.Link' || items[i].Cells[5]['Value'] == 0){


     html += 
     `       
              <div id="item" class="${styles.column}" draggable="true" >
                  <a class="${styles.title} "href="${items[i].Cells[2]["Value"]}">${items[i].Cells[3]["Value"]}</a>
                  <div class="${styles.description}" >${items[i].Cells[4]["Value"]}</div>
              </div>  
    `;  
      }

    };
    
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
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
