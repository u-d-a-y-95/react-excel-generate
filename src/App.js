import React from 'react';
import './style.css';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';

export default function App() {
  const getIndex = char => {
    return char?.toUpperCase()?.charCodeAt() - 64;
  };
  const getBorder = borderInfo => {
    const _borderDefinations = borderInfo.split(':');
    if (_borderDefinations[0].split(' ')[0] === 'all') {
      const _borderData = _borderDefinations[0].split(' ');
      if (_borderData?.length == 2) {
        return ['thin', 'dotted', 'medium', 'double'].includes(_borderData[1])
          ? {
              top: { style: _borderData[1] },
              left: { style: _borderData[1] },
              bottom: { style: _borderData[1] },
              right: { style: _borderData[1] }
            }
          : {
              top: { color: { argb: _borderData[1] }, style: 'thin' },
              left: { color: { argb: _borderData[1] }, style: 'thin' },
              bottom: { color: { argb: _borderData[1] }, style: 'thin' },
              right: { color: { argb: _borderData[1] }, style: 'thin' }
            };
      } else {
        return {
          top: { color: { argb: _borderData[1] }, style: _borderData[2] },
          left: { color: { argb: _borderData[1] }, style: _borderData[2] },
          bottom: { color: { argb: _borderData[1] }, style: _borderData[2] },
          right: { color: { argb: _borderData[1] }, style: _borderData[2] }
        };
      }
    } else {
      const obj = {};
      _borderDefinations.forEach(item => {
        const _singleBorderDef = item.split(' ');
        switch (_singleBorderDef?.length) {
          case 1:
            obj[_singleBorderDef[0]] = { color: 'black', style: 'thin' };
            break;
          case 2:
            obj[_singleBorderDef[0]] = [
              'thin',
              'dotted',
              'medium',
              'double'
            ].includes(_singleBorderDef[1])
              ? { style: _singleBorderDef[1] }
              : { color: { argb: _borderData[1] }, style: 'thin' };
            break;
          case 3:
            obj[_singleBorderDef[0]] = {
              color: { argb: _borderData[1] },
              style: _singleBorderDef[2]
            };
            break;
        }
      });

      return obj;
    }
  };
  const createWorkbook = ({ creator = '', createTime = new Date() }) => {
    const workbook = new Workbook();
    workbook.creator = creator;
    workbook.created = createTime;
    return workbook;
  };
  const getTextFormat = formatName => {
    const format = { number: 0, text: '@' };
    return format[formatName];
  };
  const createFile = excel => {
    //create workbook
    const workbook = createWorkbook(excel);

    //generate sheets
    excel?.sheets?.forEach(sheet => {
      //create worksheet
      const _sheet = workbook.addWorksheet(sheet?.name ?? '');

      // generate rows
      sheet?.rows?.forEach((row, rowIndex) => {
        // preprocess row
        const _row = [];
        // _row[0] = 'uday';
        let lastCellIndex = 0;
        row?.forEach((cell, index) => {
          _row[
            cell?.cellRange ? getIndex(cell?.cellRange[0]) : lastCellIndex + 1
          ] = cell?.text;
          lastCellIndex = cell?.cellRange
            ? getIndex(cell?.cellRange?.split(':')[1][0])
            : lastCellIndex + 1;
        });

        const _addedRow = _sheet.addRow(_row);
        let _cellIndex = 0;
        lastCellIndex = 0;
        _addedRow.eachCell((cell, cellIndex) => {
          if (lastCellIndex < cellIndex) {
            // add fonyt style
            cell.font = {
              // name: row[cellIndex]?.fontFamily,
              size: row[_cellIndex]?.fontSize || 10,
              bold: row[_cellIndex]?.bold || false,
              underline: row[_cellIndex]?.underline || false,
              italic: row[_cellIndex]?.italic || false,
              color: { argb: row[_cellIndex]?.textColor } || 'black'
            };

            //merge cell to the cell
            if (row[_cellIndex]?.merge) {
              _sheet.mergeCells(row[_cellIndex]?.cellRange);
            }

            // add alignment to the cell
            cell.alignment = {
              horizontal: row[_cellIndex]?.alignment.split(':')[0] || 'center',
              vertical: row[_cellIndex]?.alignment.split(':')[1] || 'middle'
            };

            // add border to the cell
            if (row[_cellIndex]?.border) {
              cell.border = getBorder(row[_cellIndex].border);
            }

            // add text format to the cell
            if (row[_cellIndex]?.textFormat) {
              cell.numFmt = getTextFormat(row[_cellIndex]?.textFormat);
            } else {
              cell.numFmt = getTextFormat('text');
            }
            lastCellIndex = row[_cellIndex]?.cellRange
              ? getIndex(row[_cellIndex]?.cellRange?.split(':')[1][0])
              : lastCellIndex + 1;
            _cellIndex++;
          }
        });
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
      <h1>Excel!</h1>
      <button
        onClick={() => {
          createFile({
            name: 'TestReport',
            creator: 'Uday',
            createTime: new Date(),
            sheets: [
              {
                name: 'Sheet-1',
                rows: [
                  [
                    {
                      text: 'Name',
                      bold: true,
                      itatlic: true,
                      underline: true,
                      italic: true,
                      // fontFamily: string,
                      textFormat: 'text',
                      fontSize: 10,
                      textColor: '33adff',
                      // bgColor: 'string',
                      border: 'all 33adff thin',
                      alignment: 'center:middle',
                      cellRange: 'A1:B1', // "A1:A3"
                      merge: true
                    },
                    {
                      text: 'Age',
                      bold: true,
                      // itatlic: true,
                      underline: true,
                      // italic: true,
                      // fontFamily: string,
                      textFormat: 'text',
                      fontSize: 10,
                      textColor: '33adff',
                      // bgColor: 'string',
                      border: 'all 33adff thin',
                      alignment: 'right:middle',
                      // cellRange: 'A1:B4', // "A1:A3"
                      merge: false
                    }
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
