export interface IGMRequest {
  Id: number;
  Title: string;
  RequestNo: string;
  Status: string;
  RequestDate: string;
  RequestDetails: string;
  OfficeManagerComment: string;
  HODComment: string;
  GMApprovalDetected: boolean;
  SignedPDFUrl: string;
  Author?: {
    Id: number;
    Title: string;
  };
}

export interface IUserContext {
  id: number;
  title: string;
  loginName: string;
  groups: string[];
}

export type SectionName =
  | 'new'
  | 'my'
  | 'office'
  | 'hod'
  | 'details';