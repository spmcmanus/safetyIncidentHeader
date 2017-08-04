declare interface ISafetyIncidentStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'safetyIncidentStrings' {
  const strings: ISafetyIncidentStrings;
  export = strings;
}
