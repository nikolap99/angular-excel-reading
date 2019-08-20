import { Component, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-excel-import',
  templateUrl: './excel-import.component.html',
  styleUrls: ['./excel-import.component.scss']
})
export class ExcelImportComponent implements OnInit {
  constructor() {}
  colValues: [];
  ngOnInit() {}

  onFileChange(e) {
    const file = e.target.files[0];
    var reader = new FileReader();
    reader.readAsArrayBuffer(file);
    var colValues = [];
    var colLength = 0;
    reader.onload = function(e) {
      var data = reader.result;
      var workbook = XLSX.read(data, { type: 'array', sheetRows: 5 });

      // Adding column names for mapping
      var first_sheet_name = workbook.SheetNames[0];
      var worksheet = workbook.Sheets[first_sheet_name];
      var cells = Object.keys(worksheet);
      for (var i = 0; i < Object.keys(cells).length; i++) {
        if (cells[i].indexOf('1') > -1) {
          colLength++;
          colValues.push(worksheet[cells[i]].v); //Contails all column names
        }
      }
      console.log(colValues);

      // Getting data from single column(to show the first 5 possible cells)
      var arrOfColumns = [];
      var arrInner = [];
      var range = XLSX.utils.decode_range(worksheet['!ref']);
      for (var C = range.s.c; C <= range.e.c; ++C) {
        arrInner = [];
        for (var R = range.s.r; R <= range.e.r; ++R) {
          var cellref = XLSX.utils.encode_cell({ c: C, r: R });
          if (!worksheet[cellref]) continue; // if cell doesn't exist, move on
          var cell = worksheet[cellref];
          arrInner.push(cell.v);
        }
        arrOfColumns.push(arrInner);
      }
      console.log(arrOfColumns);
    };
  }
}
