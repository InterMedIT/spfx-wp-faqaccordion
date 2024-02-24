declare interface IFaqAccordionWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  WebpartTitleFieldLabel
  SiteNameFieldLabel: string;
  ListNameFieldLabel: string;
  CategoryChoiceFieldLabel: string;
  QuestionColumnFieldLabel: string;
  AnswerColumnFieldLabel: string;
  SortColumnFieldLabel: string;
  SortDirectionFieldLabel: string;
  AllowZeroExpandFieldLabel: string;
  AllowMultiExpandFieldLabel: string;
}

declare module 'FaqAccordionWebPartStrings' {
  const strings: IFaqAccordionWebPartStrings;
  export = strings;
}
