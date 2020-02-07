import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
import * as logoFile from './delllogo.js';
import { DatePipe } from '@angular/common';
@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  constructor(private datePipe: DatePipe) {

  }

  async generateExcel() {


    const worksheet_name = 'Onsite Data Sanitization';
    const title = 'Onsite Data Santization Report'
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet(worksheet_name);
    const isImage = true;
    const isTitle = true;

    if(isImage){
      const logo = workbook.addImage({
        base64: logoFile.delllogobase64,
        extension: 'jpeg',
      },);

      worksheet.addImage(logo, {
        tl: { col: 0.1, row: 0.1 },   //0.1 0.1
        br: { col: 0.9, row: 4}       //0.9 4
      });
    }

    // worksheet.addImage(logo, {
    //   tl: { col: 1.1, row: 1.1 },   //0.1 0.1
    //   br: { col: 1.9, row: 4}       //0.9 4
    // });

    //add Title

    if(isTitle){
      worksheet.mergeCells("A1:H4");   //A1 H4
      worksheet.getCell("A1").value = title; //A1
      worksheet.getCell('A1').alignment = { horizontal:'center'} ;  //A1
     
      worksheet.getCell('A1').font = { //A1
        name: 'Calibri',
        family: 4,
        size: 20,
        bold: true
      };
    }



    worksheet.mergeCells("B5", "C5");
    worksheet.getCell('B5').value = 'Dell Information';
    worksheet.getCell('B5').font = {
      name: 'Arial',
      size: 18,
      bold: true,
      underline: true,
      color: { argb: '027CBB'}
    };
    
    worksheet.mergeCells('D5:E5');
    worksheet.getCell('D5').value = 'Customer Information';
    worksheet.getCell('D5').font = {
      name: 'Arial',
      size: 18,
      bold: true,
      underline: true,
      color: { argb: '027CBB'}
    };

    worksheet.mergeCells('F5:G5');
    worksheet.getCell('F5').value = 'Service Information';
    worksheet.getCell('F5').font = {
      name: 'Arial',
      size: 18,
      bold: true,
      underline: true,
      color: { argb: '027CBB'}
    };



    worksheet.getCell('B6').value = 'Dell Job Reference:';
    worksheet.getCell('B7').value = 'Dell Vendor Name:';
    worksheet.getCell('B9').value = 'Technician Name:';
    worksheet.getCell('B10').value = 'Software Name:';
    worksheet.getCell('B11').value = 'Software Version:';

    worksheet.getCell('D6').value = 'Customer Name:';
    worksheet.getCell('D7').value = 'Project Name:';
    worksheet.getCell('D9').value = 'Site/Location';
    worksheet.getCell('D10').value = 'Country:';
    worksheet.getCell('D11').value = 'Site Contact Name';

    worksheet.getCell('F6').value = 'Date of Service:';
    worksheet.getCell('F7').value = 'Start Time';
    worksheet.getCell('F8').value = 'Finish Time';
    worksheet.getCell('F9').value = 'Systems Processed:';
    worksheet.getCell('F10').value = 'Systems Passed:';
    worksheet.getCell('F11').value = 'Systems Failed:';


worksheet.mergeCells('A12:H12');

// Add Header Row
const header = [' ','Computer make', 'Computer model', 'Computer Service Tag', 'Drive Model', 'Drive Serial Number', 'Purge, Clear, Destroy NIST 800-88Rev 1', 'Result (Pass/Fail)', 'Exceptions Comment'];
const headerRow = worksheet.addRow(header);
headerRow.eachCell((cell, number) => {
  if(number>1){
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "D3D3D3" },
    bgColor: { argb: "000000" }
  };
  cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  };
  cell.font = {
    name: 'Arial'
  }
});


worksheet.columns.forEach(column => {
  column.width = 20;
})
worksheet.addRow([]);


// Generate Excel File with given name
    workbook.xlsx.writeBuffer().then((data: any) => {
  const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  fs.saveAs(blob, 'On-Site-Sanitization.xlsx');
});

  }
}
