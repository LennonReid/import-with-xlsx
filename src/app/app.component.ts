import {Component} from '@angular/core';
import * as XLSX from 'xlsx';
import {NgForage} from 'ngforage';

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
  // accpt file type
  acceptFile = '.xls,.xlsx';
  // 文件大小
  sizeM = 0;
  sizeKb = 0;
  // the data of import
  importUserList: any;
  // spent time when upload before getting the data
  period = 0;
  // the type of time accept strong: s.ms.hour,min,day It will be 0 if the time low to 1 second一
  type = 'ms';
  // the begin time when click upload
  beginTime: any;
  // the info of headers
  importUserHeader: any;
  fileList: any;
  tableList: TableListInterface[] = [];
  // config the counts which is needed to export
  // listNeededArr = [20000, 50000, 100000, 200000];
  listNeededArr = [];
  tableHeader = [
    {
      title: '文件大小',
      key: 'size'
    },
    {
      title: '表头条数',
      key: 'count'
    },
    {
      title: '总条数',
      key: 'count'
    },
    {
      title: '读取时间',
      key: 'time'
    },
    {
      title: '操作时间',
      key: 'operateTime'
    },
    {
      title: '操作类型',
      key: 'operateType'
    }
  ];

  constructor(
    private ngForage: NgForage
  ) {
    // this.exportFile(FAKEDATA);
    this.fileList = [];
    this.importUserHeader = [];
    this.initTableList();
    // this.exportList();
    this.listNeededArr.forEach((count: number) => {
      this.createData(count);
    });
  }

  async initTableList(): Promise<void> {

    this.tableList = await this.getFromForage() ? await this.getFromForage() : [];
  }


  /**
   * 间隔时间
   * @param faultDate 初始时间
   * @param type 时间格式
   */
  timePeriod(faultDate: any, type?: string): any {
    const completeTime = new Date();
    // let d1 = new Date(faultDate);
    // let d2 = new Date(completeTime);
    const stime = new Date(faultDate).getTime();
    const etime = new Date(completeTime).getTime();
    const usedTime = etime - stime;  // 两个时间戳相差的毫秒数

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
      const remainCount = (count % 16384);
      for (let j = 0; j < remain; j++) {
        fakeDataUsed[j] = [];
        for (let i = 0; i < (j === remain - 1 ? remainCount < 16384 ? remainCount : 16384 : 16384); i++) {
          fakeDataUsed[j].push(i + '条');
        }

      }
    } else {
      for (let i = 0; i < count; i++) {
        innerData.push(i + '条');
      }
      fakeDataUsed.push(innerData);
    }
    if (fakeDataUsed.length > 0) {
      this.exportFile(fakeDataUsed, remain);
    }


  }

  // 文件大小c
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
   * 文件变化
   * @param target 变化结果
   */
  fileChange(target: any): any {

    this.beginTime = new Date();
    if (target && target.files) {
      for (const file of target.files) {
        const fileName = file.name; // 获取文件名
        const reader: FileReader = new FileReader(); // FileReader 对象允许Web应用程序异步读取存储在用户计算机上的文件
        // 当读取操作成功完成时调用FileReader.onload
        reader.onload = (e: any) => {
          const bstr: string = e.target.result;
          const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});
          const wsname: string = wb.SheetNames[0];
          const ws: XLSX.WorkSheet = wb.Sheets[wsname];
          this.importUserList = XLSX.utils.sheet_to_json(ws, {header: 1}); // 解析出文件数据，可以进行后面操作
          this.importUserHeader = this.importUserList[0]; // 获得表头字段
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
            total: total + '条',
            time: this.period + this.type,
            operateTime: new Date().toLocaleDateString() + new Date().toLocaleTimeString(),
            operateType: target.files.length > 1 ? '批量上传' : '单文件上传',
          });
          this.setToForage();
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
   * 导出
   */
  exportFile(dataSource: any, remain?: number): any {
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(dataSource);
    // const ws2: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    // XLSX.utils.book_append_sheet(wb, ws2, 'Sheet2');

    let total = 0;
    dataSource.forEach((list: any) => {
      total += list.length;
    });
    /* save to file */
    XLSX.writeFile(wb, remain ? `${dataSource[0].length}列${remain}行_${total}条-${new Date().toLocaleDateString() + new Date().toLocaleTimeString()}.xlsx` : `${new Date().toLocaleDateString() + new Date().toLocaleTimeString()}测试结果.xlsx`, {type: 'file'});
    // XLSX.readFile('SheetJS.xlsx',{type: 'file'});
    // console.log(XLSX.readFile('SheetJS.xlsx'));
  }


  setToForage(): void {
    this.ngForage.setItem('tableList', this.tableList).then(() => {

      this.initTableList();
    });
  }

  async getFromForage(): Promise<any> {
    return await this.ngForage.getItem('tableList');
  }

  clear(): void {
    this.ngForage.clear();
    this.initTableList();
  }
}

