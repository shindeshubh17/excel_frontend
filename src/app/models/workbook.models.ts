export interface Workbook {
  workbookId: string;
  name: string;
  createdBy: string;
  createdDate: string;
}

export interface WorkbookVersion {
  versionId: string;
  workbookId: string;
  versionNumber: number;
  fileName: string;
  createdBy: string;
  createdDate: string;
  comments?: string;
}

export interface SaveVersionRequest {
  workbookId: string;
  workbookName: string;
  createdBy: string;
  comments?: string;
  fileBase64: string;
}

export interface SaveVersionResponse {
  versionId: string;
  versionNumber: number;
  fileName: string;
  createdDate: string;
}
