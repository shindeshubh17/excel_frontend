import { Component, signal } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { ExcelComponent } from './excel/excel.component';

@Component({
  selector: 'app-root',
  imports: [ExcelComponent],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App {
  protected readonly title = signal('frontend');
}
