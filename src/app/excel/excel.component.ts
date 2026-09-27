import {
  AfterViewInit,
  ChangeDetectorRef,
  Component,
  ElementRef,
  OnDestroy,
  ViewChild,
} from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { switchMap } from 'rxjs';

import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import { createUniver, defaultTheme, LocaleType, merge } from '@univerjs/presets';

import { Workbook, WorkbookVersion } from '../models/workbook.models';
import { WorkbookService } from '../services/workbook.service';
import { xlsxToWorkbookData, workbookDataToXlsxArrayBuffer } from './xlsx-univer-bridge';

@Component({
  selector: 'app-excel',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './excel.component.html',
  styleUrl: './excel.component.css',
})
export class ExcelComponent implements AfterViewInit, OnDestroy {
  @ViewChild('univerContainer', { static: true }) univerContainer!: ElementRef<HTMLDivElement>;
  @ViewChild('fileInput', { static: true }) fileInput!: ElementRef<HTMLInputElement>;

  private univer: any;
  private univerAPI: any;

  workbooks: Workbook[] = [];
  currentWorkbook: Workbook | null = null;
  workbookName = 'Untitled Workbook';
  versions: WorkbookVersion[] = [];
  selectedVersionId: string | null = null;

  currentUser = 'demo-user';
  isLoading = false;
  isSaving = false;
  showSaveModal = false;
  versionComment = '';
  statusMessage = '';
  mobileMenuOpen = false;

  constructor(
    private workbookService: WorkbookService,
    private cdr: ChangeDetectorRef
  ) {}

  ngAfterViewInit(): void {
    this.initUniver();

    // Check if we just saved a new workbook and reloaded the page
    const savedId = sessionStorage.getItem('active_workbook_id');
    if (savedId) {
      sessionStorage.removeItem('active_workbook_id');
      this.refreshWorkbooks(() => {
        const savedWb = this.workbooks.find((w) => w.workbookId === savedId);
        if (savedWb) {
          this.onWorkbookSelected(savedWb.workbookId);
          this.statusMessage = `Workbook "${savedWb.name}" saved successfully.`;
        } else {
          this.newWorkbook();
        }
        this.cdr.detectChanges();
      });
      return;
    }

    this.newWorkbook();
    this.refreshWorkbooks();
  }

  ngOnDestroy(): void {
    if (this.univer) {
      this.univer.dispose?.();
    }
  }

  private initUniver(): void {
    const { univer, univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: merge({}, UniverPresetSheetsCoreEnUS),
      },
      theme: defaultTheme,
      presets: [
        UniverSheetsCorePreset({
          container: this.univerContainer.nativeElement,
        }),
      ],
    });

    this.univer = univer;
    this.univerAPI = univerAPI;
  }

  newWorkbook(): void {
    // dispose current active sheet unit to prevent stacking
    const active = this.univerAPI?.getActiveWorkbook?.();
    if (active) {
      this.univerAPI?.disposeUnit?.(active.getId());
    }

    this.univerAPI?.createWorkbook?.({
      id: `wb-${Date.now()}`,
      sheetOrder: ['sheet1'],
      sheets: {
        sheet1: { id: 'sheet1', name: 'Sheet1', cellData: {} }
      },
    });

    this.currentWorkbook = {
      workbookId: '',
      name: 'Untitled Workbook',
      createdBy: this.currentUser,
      createdDate: new Date().toISOString(),
    };
    this.workbookName = 'Untitled Workbook';
    this.versions = [];
    this.selectedVersionId = null;
    this.statusMessage = 'New blank workbook created.';
    this.cdr.detectChanges();
  }

  addSheet(): void {
    const active = this.univerAPI?.getActiveWorkbook?.();
    const count = active?.getSheets()?.length ?? 0;
    const name = `Sheet${count + 1}`;
    active?.create(name, 100, 20);

    this.statusMessage = `${name} added.`;
    this.cdr.detectChanges();
  }

  triggerFileChoose(): void {
    this.fileInput.nativeElement.click();
  }

  onFileChosen(event: Event): void {
    const target = event.target as HTMLInputElement;
    const file = target.files?.[0];
    if (!file) return;

    if (!file.name.match(/\.(xlsx|csv)$/i)) {
      this.statusMessage = 'Only .xlsx and .csv files are supported.';
      this.cdr.detectChanges();
      return;
    }

    this.isLoading = true;
    this.cdr.detectChanges();

    file.arrayBuffer()
      .then((buffer) => {
        const id = `wb-${Date.now()}`;
        const snapshot = xlsxToWorkbookData(buffer, id);

        const active = this.univerAPI?.getActiveWorkbook?.();
        if (active) {
          this.univerAPI?.disposeUnit?.(active.getId());
        }

        this.univerAPI?.createWorkbook?.(snapshot);

        const name = file.name.replace(/\.(xlsx|csv)$/i, '');
        this.currentWorkbook = {
          workbookId: '',
          name,
          createdBy: this.currentUser,
          createdDate: new Date().toISOString(),
        };
        this.workbookName = name;
        this.versions = [];
        this.selectedVersionId = null;
        this.statusMessage = `Loaded "${file.name}".`;
      })
      .catch((err) => {
        console.error('Error loading file:', err);
        this.statusMessage = 'Failed to load workbook.';
      })
      .finally(() => {
        this.isLoading = false;
        target.value = '';
        this.cdr.detectChanges();
      });
  }

  onWorkbookNameChanged(): void {
    const trimmed = (this.workbookName || '').trim();
    if (!trimmed) {
      this.workbookName = this.currentWorkbook?.name || 'Untitled Workbook';
      return;
    }

    this.workbookName = trimmed;

    if (!this.currentWorkbook) {
      this.currentWorkbook = {
        workbookId: '',
        name: trimmed,
        createdBy: this.currentUser,
        createdDate: new Date().toISOString(),
      };
      this.statusMessage = `Workbook renamed to "${trimmed}".`;
      this.cdr.detectChanges();
      return;
    }

    if (this.currentWorkbook.name === trimmed) {
      return;
    }

    this.currentWorkbook.name = trimmed;

    if (this.currentWorkbook.workbookId) {
      this.workbookService
        .updateWorkbook(this.currentWorkbook.workbookId, trimmed)
        .subscribe({
          next: () => {
            const match = this.workbooks.find(
              (w) => w.workbookId === this.currentWorkbook!.workbookId
            );
            if (match) {
              match.name = trimmed;
            }
            this.statusMessage = `Workbook renamed to "${trimmed}".`;
            this.cdr.detectChanges();
          },
          error: (err) => {
            console.error('Rename failed:', err);
            this.statusMessage = 'Failed to rename workbook on server.';
            this.cdr.detectChanges();
          }
        });
    } else {
      this.statusMessage = `Workbook renamed to "${trimmed}".`;
      this.cdr.detectChanges();
    }
  }

  onWorkbookNameInputEnter(event: Event): void {
    const el = event.target as HTMLInputElement;
    el?.blur();
  }

  openSaveModal(): void {
    this.versionComment = '';
    this.showSaveModal = true;
    this.cdr.detectChanges();
  }

  closeSaveModal(): void {
    this.showSaveModal = false;
    this.cdr.detectChanges();
  }

  confirmSave(): void {
    const comment = this.versionComment?.trim() || undefined;
    this.showSaveModal = false;
    this.executeSave(comment);
  }

  save(): void {
    this.openSaveModal();
  }

  executeSave(comment?: string): void {
    this.isSaving = true;
    this.cdr.detectChanges();

    try {
      const active = this.univerAPI?.getActiveWorkbook?.();
      const snapshot = active?.save?.();
      const xlsxBuffer = workbookDataToXlsxArrayBuffer(snapshot);
      const fileBase64 = this.arrayBufferToBase64(xlsxBuffer);

      const isNew = !this.currentWorkbook?.workbookId;
      const title =
        this.workbookName?.trim() ||
        this.currentWorkbook?.name ||
        'Untitled Workbook';

      if (isNew) {
        this.workbookService
          .createWorkbook(title, this.currentUser)
          .pipe(
            switchMap((createdWb) => {
              this.currentWorkbook = createdWb;
              return this.workbookService.saveVersion({
                workbookId: createdWb.workbookId,
                workbookName: title,
                createdBy: this.currentUser,
                comments: comment,
                fileBase64,
              });
            })
          )
          .subscribe({
            next: () => {
              sessionStorage.setItem('active_workbook_id', this.currentWorkbook!.workbookId);
              window.location.reload();
            },
            error: (err) => {
              console.error('Save error:', err);
              this.statusMessage = 'Save failed.';
              this.isSaving = false;
              this.cdr.detectChanges();
            }
          });
      } else {
        this.workbookService
          .saveVersion({
            workbookId: this.currentWorkbook!.workbookId,
            workbookName: title,
            createdBy: this.currentUser,
            comments: comment,
            fileBase64,
          })
          .subscribe({
            next: (res) => {
              this.statusMessage = `Saved as V${res.versionNumber}.`;
              this.selectedVersionId = res.versionId;
              this.refreshWorkbooks();
              this.refreshVersions();
              this.isSaving = false;
              this.cdr.detectChanges();
            },
            error: (err) => {
              console.error('Save error:', err);
              this.statusMessage = 'Save failed.';
              this.isSaving = false;
              this.cdr.detectChanges();
            }
          });
      }
    } catch (err) {
      console.error('Save error:', err);
      this.statusMessage = 'Save failed.';
      this.isSaving = false;
      this.cdr.detectChanges();
    }
  }

  refreshWorkbooks(onComplete?: () => void): void {
    this.workbookService.listWorkbooks().subscribe({
      next: (data) => {
        this.workbooks = data ?? [];
        this.cdr.detectChanges();
        if (onComplete) onComplete();
      },
      error: (err) => {
        console.error('Failed to list workbooks', err);
        this.workbooks = [];
        this.cdr.detectChanges();
        if (onComplete) onComplete();
      }
    });
  }

  refreshVersions(onComplete?: () => void): void {
    if (!this.currentWorkbook?.workbookId) {
      this.versions = [];
      this.cdr.detectChanges();
      if (onComplete) onComplete();
      return;
    }

    this.workbookService
      .listVersions(this.currentWorkbook.workbookId)
      .subscribe({
        next: (data) => {
          this.versions = (data ?? []).sort((a, b) => b.versionNumber - a.versionNumber);
          this.cdr.detectChanges();
          if (onComplete) onComplete();
        },
        error: (err) => {
          console.error('Failed to list versions', err);
          this.versions = [];
          this.cdr.detectChanges();
          if (onComplete) onComplete();
        }
      });
  }

  onWorkbookSelected(workbookId: string): void {
    const wb = this.workbooks.find((w) => w.workbookId === workbookId);
    if (!wb) return;

    this.isLoading = true;
    this.currentWorkbook = wb;
    this.workbookName = wb.name;
    this.cdr.detectChanges();

    this.refreshVersions(() => {
      if (this.versions.length > 0) {
        this.onVersionSelected(this.versions[0].versionId);
      } else {
        this.isLoading = false;
        this.statusMessage = `Loaded "${wb.name}".`;
        this.cdr.detectChanges();
      }
    });
  }

  onVersionSelected(versionId: string): void {
    if (!this.currentWorkbook?.workbookId) return;

    this.isLoading = true;
    this.cdr.detectChanges();

    this.workbookService
      .getVersionFile(this.currentWorkbook.workbookId, versionId)
      .subscribe({
        next: (fileData) => {
          try {
            const buffer = this.base64ToArrayBuffer(fileData.fileBase64);
            const snapshot = xlsxToWorkbookData(buffer, this.currentWorkbook!.workbookId);

            const active = this.univerAPI?.getActiveWorkbook?.();
            if (active) {
              this.univerAPI?.disposeUnit?.(active.getId());
            }

            this.univerAPI?.createWorkbook?.(snapshot);
            this.selectedVersionId = versionId;
            this.workbookName = this.currentWorkbook?.name ?? 'Untitled Workbook';
            this.statusMessage = `Loaded ${fileData.fileName}.`;
          } catch (err) {
            console.error('Load version error:', err);
            this.statusMessage = 'Failed to load version.';
          } finally {
            this.isLoading = false;
            this.cdr.detectChanges();
          }
        },
        error: (err) => {
          console.error('Load version error:', err);
          this.statusMessage = 'Failed to load version.';
          this.isLoading = false;
          this.cdr.detectChanges();
        }
      });
  }

  downloadCurrent(): void {
    if (!this.currentWorkbook?.workbookId || !this.selectedVersionId) {
      this.statusMessage = 'Save the workbook before downloading.';
      this.cdr.detectChanges();
      return;
    }

    const downloadUrl = this.workbookService.getDownloadUrl(
      this.currentWorkbook.workbookId,
      this.selectedVersionId
    );
    window.open(downloadUrl, '_blank');
  }

  dismissStatus(): void {
    this.statusMessage = '';
    this.cdr.detectChanges();
  }

  toggleMobileMenu(): void {
    this.mobileMenuOpen = !this.mobileMenuOpen;
  }

  private arrayBufferToBase64(buffer: ArrayBuffer): string {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }

  private base64ToArrayBuffer(base64: string): ArrayBuffer {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes.buffer;
  }
}

export { ExcelComponent as Excel };
