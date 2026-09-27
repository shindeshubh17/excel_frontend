import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import {
  SaveVersionRequest,
  SaveVersionResponse,
  Workbook,
  WorkbookVersion,
} from '../models/workbook.models';

@Injectable({
  providedIn: 'root'
})
export class WorkbookService {
  private apiUrl = 'http://localhost:5000/api';

  constructor(private http: HttpClient) {}

  listWorkbooks(): Observable<Workbook[]> {
    return this.http.get<Workbook[]>(`${this.apiUrl}/workbooks`);
  }

  getWorkbook(id: string): Observable<Workbook> {
    return this.http.get<Workbook>(`${this.apiUrl}/workbooks/${id}`);
  }

  createWorkbook(name: string, createdBy: string): Observable<Workbook> {
    return this.http.post<Workbook>(`${this.apiUrl}/workbooks`, { name, createdBy });
  }

  updateWorkbook(id: string, name: string): Observable<Workbook> {
    return this.http.put<Workbook>(`${this.apiUrl}/workbooks/${id}`, { name });
  }

  listVersions(workbookId: string): Observable<WorkbookVersion[]> {
    return this.http.get<WorkbookVersion[]>(`${this.apiUrl}/workbooks/${workbookId}/versions`);
  }

  saveVersion(data: SaveVersionRequest): Observable<SaveVersionResponse> {
    return this.http.post<SaveVersionResponse>(
      `${this.apiUrl}/workbooks/${data.workbookId}/versions`,
      data
    );
  }

  getVersionFile(workbookId: string, versionId: string): Observable<{ fileBase64: string; fileName: string }> {
    return this.http.get<{ fileBase64: string; fileName: string }>(
      `${this.apiUrl}/workbooks/${workbookId}/versions/${versionId}/file`
    );
  }

  getDownloadUrl(workbookId: string, versionId: string): string {
    return `${this.apiUrl}/workbooks/${workbookId}/versions/${versionId}/download`;
  }
}
