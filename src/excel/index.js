// import { Workbook } from 'exceljs';
import * as fs from 'file-saver';

import { getWorkBook, getIndex } from './utils.js';
import { getfontStyle, getTextFormat, getFill } from './font.js';
import { getAlignment } from './alignment.js';
import { getBorder } from './border.js';

export const createFile = excel => {
  //create workbook
  const workbook = getWorkBook(excel);

  //generate sheets
  excel?.sheets?.forEach(sheet => {
    //create worksheet

    const _sheet = workbook.addWorksheet(sheet?.name ?? '', {
      views: [
        {
          showGridLines: sheet?.hasOwnProperty('gridLine')
            ? sheet.gridLine
            : true
        }
      ]
    });

    // generate rows
    sheet?.rows?.forEach((row, rowIndex) => {
      // preprocess row
      const _row = [];

      let lastCellIndex = 0;
      row?.forEach((cell, index) => {
        if (typeof cell === 'string') {
          if (cell.startsWith('_blank')) {
            for (let i = 0; i < Number(cell.split('_blank*')[1]) - 1; i++) {
              const _addedRow = _sheet.addRow([]);
            }
          } else {
            _row[lastCellIndex + 1] = cell;
            lastCellIndex += 1;
          }
        } else {
          _row[
            cell?.cellRange ? getIndex(cell?.cellRange[0]) : lastCellIndex + 1
          ] = cell?.text;
          lastCellIndex = cell?.cellRange
            ? getIndex(cell?.cellRange?.split(':')[1][0])
            : lastCellIndex + 1;
        }
      });

      const _addedRow = _sheet.addRow(_row);
      _addedRow.border = sheet?.border && getBorder(sheet?.border);
      _addedRow.font = getfontStyle({
        fontSize: sheet?.fontSize,
        bold: sheet?.bold,
        underline: sheet?.underline,
        textColor: sheet?.italic,
        italic: sheet?.textColor
      });
      _addedRow.alignment = getAlignment({
        alignment: sheet?.alignment
      });
      _addedRow.fill =
        sheet?.bgColor &&
        getFill({
          bgColor: sheet?.bgColor
        });
      let _cellIndex = 0;
      lastCellIndex = 0;
      _addedRow.eachCell((cell, cellIndex) => {
        if (lastCellIndex < cellIndex) {
          if (typeof row[_cellIndex] === 'object') {
            // add fonyt style
            cell.font = getfontStyle(row[_cellIndex]);

            //merge cell to the cell
            if (row[_cellIndex]?.merge) {
              const points = row[_cellIndex]?.cellRange?.split(':');
              _sheet.mergeCells(
                `${points[0][0]}${_addedRow.number +
                  (Number(points[0].slice(1)) - 1)}:${
                  points[1][0]
                }${_addedRow.number + (Number(points[1].slice(1)) - 1)}`
              );
            }

            // add alignment to the cell
            cell.alignment = getAlignment(row[_cellIndex]);

            // add fill color
            cell.fill = row[_cellIndex]?.bgColor && getFill(row[_cellIndex]);

            // add border to the cell
            cell.border =
              row[_cellIndex]?.border && getBorder(row[_cellIndex].border);

            // add text format to the cell
            cell.numFmt = getTextFormat(row[_cellIndex]?.textFormat);

            lastCellIndex = row[_cellIndex]?.cellRange
              ? getIndex(row[_cellIndex]?.cellRange?.split(':')[1][0])
              : lastCellIndex + 1;
            _cellIndex++;
          } else {
            lastCellIndex += 1;
            _cellIndex++;
          }
        }
      });
    });
  });
  workbook.xlsx.writeBuffer().then(data => {
    let blob = new Blob([data], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    fs.saveAs(blob, `${excel.name}.xlsx`);
  });
};
