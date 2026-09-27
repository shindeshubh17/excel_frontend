# Frontend Application (Angular 21 + Univer)

This is the frontend client for the spreadsheet application, built with **Angular 21**, **Univer Sheets Core**, and **SheetJS**.

It provides an in-browser spreadsheet interface with canvas rendering, formula support, cell styling, sheet management, and version history.

---

## Key Features

- **Univer Integration:** Uses `@univerjs/preset-sheets-core` (v1.0.2) to render a canvas-based grid with formula bar, ribbon toolbar, and sheet tabs.
- **Client-Side File Import/Export:** Uses SheetJS (`xlsx`) to convert `.xlsx` and `.csv` files into Univer's internal JSON format when uploading, and converts Univer data back to `.xlsx` binary when saving/downloading.
- **Zoneless Angular:** Configured without `zone.js`. State updates use explicit `ChangeDetectorRef.detectChanges()` calls to keep the bundle small and UI responsive.
- **Reactive HTTP Calls:** All API calls use standard RxJS `Observable` and `.subscribe()`. Chained operations (like creating a workbook then immediately saving its first version) use the `switchMap` operator instead of nested callbacks.
- **Header Toolbar:** Allows creating new workbooks, adding sheets, uploading files, saving versions with comments, and downloading the active version.

---

## Project Structure

```text
excel_frontend/
├── src/
│   ├── app/
│   │   ├── excel/
│   │   │   ├── excel.component.ts         # Main spreadsheet logic and UI events
│   │   │   ├── excel.component.html       # Toolbar, canvas container, and save modal
│   │   │   ├── excel.component.css        # Layout styling
│   │   │   └── xlsx-univer-bridge.ts      # Bridge between SheetJS and Univer JSON
│   │   ├── models/
│   │   │   └── workbook.models.ts         # Interfaces for Workbook and Version
│   │   ├── services/
│   │   │   └── workbook.service.ts        # HttpClient methods for backend API
│   │   ├── app.component.ts               # Root component
│   │   └── app.config.ts                  # App providers and HttpClient setup
│   ├── styles.css                         # Global CSS resets and font setup
│   └── index.html                         # Entry HTML
├── package.json                           # Dependencies and build scripts
└── angular.json                           # Angular CLI configuration
```

---

## How it Works

### 1. Excel Bridge (`xlsx-univer-bridge.ts`)
Univer uses a JSON snapshot structure for its grid data (`cellData: { row: { col: { v: value, f: formula } } }`). The bridge performs bidirectional conversion:
- **`xlsxToWorkbookData(buffer, id)`**: Reads an Excel binary buffer with SheetJS and translates sheets and cell coordinates into Univer's snapshot format.
- **`workbookDataToXlsxArrayBuffer(univerData)`**: Takes the Univer snapshot and generates an Excel binary `ArrayBuffer` ready to send to the backend or download.

### 2. Saving with `switchMap`
When saving a new workbook, the client must first create the workbook record to get an ID, and then call the save-version endpoint. We chain these two calls using `switchMap`:

```typescript
this.workbookService.createWorkbook(title, this.currentUser)
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
      console.error('Save failed', err);
      this.statusMessage = 'Save failed.';
      this.isSaving = false;
      this.cdr.detectChanges();
    }
  });
```

---

## Running the Application

### 1. Install dependencies
```bash
npm install
```

### 2. Start development server
```bash
npm start
```
The app will be available at `http://localhost:4200`.

### 3. Build for production
```bash
npm run build
```
Compiled files will be in `dist/frontend/`.
