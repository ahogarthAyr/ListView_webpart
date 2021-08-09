declare interface IListViewWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  TitleFieldLabel: string;  
}

declare module 'ListViewWebPartStrings' {
  const strings: IListViewWebPartStrings;
  export = strings;
}
