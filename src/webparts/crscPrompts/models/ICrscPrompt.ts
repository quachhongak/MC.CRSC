export interface ICrscPrompt {
  Id: number;
  Title: string;
  Input: string;
  AITool: string;
  Category: string;
  Deliverable: string;
  RiskLevel: string;
  SourceInputsRequired: string;
  LimitationsandConstraints: string;
  AdditionalInformation: string;
  Modified: string;
}

export interface IListOption {
  id: string;
  title: string;
}
