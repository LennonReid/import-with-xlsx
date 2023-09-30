import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { NgForage } from 'ngforage';

export interface TableListInterface {
  size: number | string;
  count: string;
  total: string | number;
  time: number | string;
  operateTime?: string | Date;
  operateType?: string;
}

const FAKEDATA = [
  ['1', 'a', 'aa'],
  ['2', 'b', 'bb'],
  ['3', 'c', 'cc']
];

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'import';
  // Acceptable file types
  acceptFile = '.xls,.xlsx';
  // File size
  sizeM = 0;
  sizeKb = 0;
  // Imported data
  importUserList: any;
  // Time spent when uploading before getting the data
  period = 0;
  // The type of time accepted (e.g., s, ms, hour, min, day). It will be 0 if the time is less than 1 second.
  type = 'ms';
  // The start time when clicking upload
  beginTime: any;
  // Header information
  importUserHeader: any;
  fileList: any;
  tableList: TableListInterface[] = [];
  // Configure the counts needed for export
  listNeededArr = [4000000];
  tableHeader = [
    {
      title: 'File Size',
      key: 'size'
    },
    {
      title: 'Header Count',
      key: 'count'
    },
    {
      title: 'Total Count',
      key: 'total'
    },
    {
      title: 'Read Time',
      key: 'time'
    },
    {
      title: 'Operation Time',
      key: 'operateTime'
    },
    {
      title: 'Operation Type',
      key: 'operateType'
    }
  ];

  constructor(
  ) {
    this.fileList = [];
    this.importUserHeader = [];
    // this.listNeededArr.forEach((count: number) => {
    //   this.createData(count);
    // });
  }


  /**
   * Calculate time interval
   * @param faultDate Initial time
   * @param type Time format
   */
  timePeriod(faultDate: any, type?: string): any {
    const completeTime = new Date();
    const stime = new Date(faultDate).getTime();
    const etime = new Date(completeTime).getTime();
    const usedTime = etime - stime;  // Difference in milliseconds between two timestamps

    const days = Math.floor(usedTime / (24 * 3600 * 1000));
    const hours = Math.floor(usedTime / (3600 * 1000));
    const minutes = Math.floor(usedTime / (60 * 1000));
    const seconds = Math.floor(usedTime / (1000));
    const millisecond = usedTime;
    switch (type) {
      default:
        return millisecond;
      case 'day':
        return days;
      case 'hour':
        return hours;
      case 'min':
        return minutes;
      case 's':
        return seconds;
      case 'ms':
        return millisecond;
    }
  }

  createData(count: number): any {
    const fakeDataUsed = [];
    const innerData = [];
    let remain = 0;
    if (count > 16384) {
      remain = Math.ceil(count / 16384);
      const remainCount = count % 16384;
      for (let j = 0; j < remain; j++) {
        fakeDataUsed[j] = [];
        for (let i = 0; i < (j === remain - 1 ? (remainCount < 16384 ? remainCount : 16384) : 16384); i++) {
          fakeDataUsed[j].push(i + 'item');
        }
      }
    } else {
      for (let i = 0; i < count; i++) {
        innerData.push(i + 'item');
      }
      fakeDataUsed.push(innerData);
    }
    if (fakeDataUsed.length > 0) {
      this.exportFile(fakeDataUsed, remain);
    }
  }

  // Calculate file size
  fileSize(file: File, fileSizeType: string): any {
    switch (fileSizeType) {
      case 'm':
        return (file.size / 1024 / 1024).toFixed(2);
      case 'kb':
        return (file.size / 1024).toFixed(2);
      case 'b':
        return file.size.toFixed(2);
    }
  }

  /**
   * File change event
   * @param target Result of the change
   */
  fileChange(target: any): any {
    this.beginTime = new Date();
    if (target && target.files) {
      for (const file of target.files) {
        const fileName = file.name; // Get the file name
        const reader: FileReader = new FileReader();
        // When the reading operation is successfully completed, FileReader.onload is called
        reader.onload = (e: any) => {
          const bstr: string = e.target.result;
          const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
          const wsname: string = wb.SheetNames[0];
          const ws: XLSX.WorkSheet = wb.Sheets[wsname];
          this.importUserList = XLSX.utils.sheet_to_json(ws, { header: 1 });
          this.importUserHeader = this.importUserList[0];
          this.fileList = this.fileList.concat(file);

          this.period = this.timePeriod(this.beginTime, this.type);
          let total = 0;
          this.importUserList.forEach((list: any) => {
            total += list.length;
          });
          this.sizeM = this.fileSize(file, 'm');
          this.sizeKb = this.fileSize(file, 'kb');
          this.tableList.unshift({
            size: this.sizeM + 'M',
            count: this.importUserHeader.length,
            total: total + '\n' + 'items',
            time: this.period + this.type,
            operateTime: new Date().toLocaleDateString() + new Date().toLocaleTimeString(),
            operateType: target.files.length > 1 ? 'Batch Upload' : 'Single File Upload',
          });
        };
        reader.readAsBinaryString(file);
      }
    }
  }

  changeToNeededArray(arr: any): any {
    const dataSource = [];
    const headers = [];
    this.tableHeader.forEach((header: any) => {
      const arr2 = [];
      arr2.push(header.title);
      headers.push(arr2);
    });
    dataSource.push(headers);

    this.tableList.forEach((list: any) => {
      const listKeys = [];
      Object.keys(list).forEach(listKey => {
        listKeys.push(list[listKey]);
      });
      dataSource.push(listKeys);
    });
    this.exportFile(dataSource);
  }

  /**
   * Export data
   */
  exportFile(dataSource: any, remain?: number): any {
    /* Generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(dataSource);

    /* Generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    let total = 0;
    dataSource.forEach((list: any) => {
      total += list.length;
    });
    /* Save to file */
    XLSX.writeFile(wb, remain ? `${dataSource[0].length} Columns ${remain} Rows ${total} Records - ${new Date().toLocaleDateString() + new Date().toLocaleTimeString()}.xlsx` : `${new Date().toLocaleDateString() + new Date().toLocaleTimeString()} Test Results.xlsx`, { type: 'file' });
  }


  clear(): void {
    this.tableList = [];
  }
}
