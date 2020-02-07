import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
import { DatePipe } from '@angular/common';
import * as xlxdata from './data.json'; 

@Injectable({
  providedIn: 'root'
})
export class NewexcelService {
  constructor(private datePipe: DatePipe) { }

  async generateExcel(){
    const workbook = new Workbook();
    //worksheetName
    const worksheetName = this.getWorksheetName(xlxdata.worksheetName);
    const worksheet = workbook.addWorksheet(worksheetName);

    const image = xlxdata.image.data[0];
    const title = xlxdata.title.data[0];
    const tables = xlxdata.tables;
    if(this.isImagePresent(image)){
      this.styleImage(workbook, worksheet, image);
    }

    if(this.isTitlePresent(title)){
      this.cellMergeAndStyle(worksheet, title);
    }

    //tables
    tables.forEach(table=>{
      table.headers.data.forEach(header => {
        this.cellMergeAndStyle(worksheet, header);
      });
      worksheet.addRows(table.rowsData);
    });
    
    //column width
    worksheet.columns.forEach(column => {
      column.width = 20;
    })

    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, 'On-Site-Sanitization.xlsx');
    });

  }
  getWorksheetName(name:string):string{
    return name?name:"Report";
  }
  isImagePresent(data):boolean{
    return data.name? true: false;
  }
  isTitlePresent(data):boolean{
    return data.name?true: false;
  }

  styleImage(workbook, worksheet, image){
    const logo = workbook.addImage({
      base64: image.name,
      extension: 'jpeg',
    },);

    const topLeft = image['top-left'];
    const bottomRight = image['bottom-right'];
    worksheet.addImage(logo, {
      tl: { col: topLeft.col, row: topLeft.row }, 
      br: { col: bottomRight.col, row: bottomRight.row}    
    });
  }

  cellMergeAndStyle(worksheet, data){
    const start = data['mergeCells'].start;
    const end = data['mergeCells'].end;
    worksheet.mergeCells(start, end);
    worksheet.getCell(start).value = data['name'];
    worksheet.getCell(start).alignment = { horizontal : 'center'};
    let cellProperties = {};
    let fontProperties = {};
    
    for(const property in data['style']){

      console.log('property', property, data['style'].property);
       if(property === "bgcolor" || property === "fgcolor"){
           cellProperties['property'] = data['style'].property;
        }

       else{
         if(property === "color") {
           let color = {};
           color['argb'] = data['style'].property;
           fontProperties['property'] = color;
          } 
         else fontProperties['property'] = data['style'].property;
       }
    }

    // console.log('fontProperties', fontProperties);
    // console.log('cellProperties', cellProperties);
    worksheet.getCell(start).font = fontProperties;
    worksheet.getCell(start).fill = cellProperties;
  }


}
