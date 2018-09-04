declare interface IValidatePropertyWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ListNameFieldLabel: string;
}

declare module 'ValidatePropertyWebPartStrings' {
  const strings: IValidatePropertyWebPartStrings;
  export = strings;
}
