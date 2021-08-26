import React from 'react';
import './style.css';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';

export default function App() {
  const getIndex = char => {
    return char?.charCodeAt() - 64;
  };
  const createWorkbook = ({ creator = '', createTime = new Date() }) => {
    const workbook = new Workbook();
    workbook.creator = creator;
    workbook.created = createTime;
    return workbook;
  };
  const createFile = excel => {
    //create workbook
    const workbook = createWorkbook(excel);

    //generate sheets
    excel?.sheets?.forEach(sheet => {
      //create worksheet
      const _sheet = workbook.addWorksheet(sheet?.name ?? '');

      // generate rows
      sheet?.rows?.forEach(row => {
        // preprocess row
        const _row = [];
        row?.forEach((cell, index) => {
          _row[cell?.cellRange ? getIndex(cell?.cellRange[0]) : index] =
            cell?.text;
        });
        wor;
      });
    });

    workbook.xlsx.writeBuffer().then(data => {
      let blob = new Blob([data], {
        type:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      fs.saveAs(blob, `${excel.name}.xlsx`);
    });
  };
  return (
    <div>
      <h1>Hello StackBlitz!</h1>
      <p>Start editing to see some magic happen :)</p>
      <button
        onClick={() => {
          createFile({
            name: 'TestReport',
            creator: 'Uday',
            createTime: new Date(),
            sheets: [
              {
                name: '',
                rows: [
                  [
                    // {
                    //   text:"",
                    //   bold:boolean,
                    //   itatlic:boolean,
                    //   underline:boolean,
                    //   font:string,
                    //   textFomat:enums
                    //   fontSize:number
                    //   textColor:"string"
                    //   bgColor:"string,
                    //   border:enums,
                    //   alignment:enums
                    //   cellRange:String // "A1:A3"
                    //   merge:boolean
                    // }
                  ]
                ]
              }
            ]
          });
        }}
      >
        Export
      </button>
    </div>
  );
}
