import { Component } from '@angular/core';
import { ExcelService } from './excel.service';
import {NewexcelService} from './newexcel.service'
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  
  constructor(private excelService: ExcelService, private newExcelService: NewexcelService) {
  }
  generateExcel() {
    this.excelService.generateExcel();
  }
  generateNewExcel(){
    this.newExcelService.generateExcel();
  }
}
