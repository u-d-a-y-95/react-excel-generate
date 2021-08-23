import React from 'react';
import './style.css';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';

export default function App() {
  const createRows = () => {
    return map;
  };
  const createFile = () => {
    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet('Car Data');
    let companyName = worksheet.addRow(['Akij Poly Fibre Industries Ltd.']);
    companyName.font = { size: 20, underline: 'single', bold: true };
    worksheet.mergeCells('A1:L1');
    worksheet.getCell('A1').alignment = { horizontal: 'center' };
    let companyLocation = worksheet.addRow([
      'Akij House, 198 Bir Uttam, Gulshan Link Road, Tejgaon, Dhaka-1208.'
    ]);
    companyLocation.font = { size: 14, underline: 'single', bold: true };
    worksheet.mergeCells('A2:L2');
    worksheet.getCell('A2').alignment = { horizontal: 'center' };
    let applicationHeader = worksheet.addRow([
      `To                                                                                                                                                                  Date: 23 August 2021\nThe Manager\nIslami Bank Bangladesh Ltd.\nHead Office Complex Branch, 40 Dilkusha, C/A, Dhaka-1000\nSubject : Payment Instruction by BEFTN.\n\nDear Sir,\nWe do hereby requesting you to make payment by transferring the amount to the respective Account Holder as shown below in detailed by debiting our CD Account No. \nIBBL: 20502130100096513\nDetailed particulars of each Account Holder:`
    ]);

    worksheet.mergeCells('A3:L14');
    worksheet.getCell('A3').alignment = { horizontal: 'left', wrapText: true };
    worksheet.addRow([]);
    let tableHeader = worksheet.addRow([
      'Account Name',
      'Code NO',
      'Bank Name',
      'Branch Name',
      'A/C Type',
      'Account No',
      'Amount',
      'Payment Info',
      'Comments',
      'Routing No',
      'Instrument No',
      'Sl No'
    ]);
    tableHeader.font = { bold: true };
    tableHeader.eachCell(cell => {
      cell.alignment = { horizontal: 'center' };
    });
    worksheet.addRows([
      [
        'SALEHA METAL INDUSTRIES',
        6587,
        'BANK ASIA LTD',
        'JESSORE',
        'saving',
        '0003333002081',
        332000,
        'SI-APFIL-AUG21-68',
        'BP-APFIL-AUG21-8',
        '070410941',
        'APFIL-118112',
        1
      ],
      [
        'SALEHA METAL INDUSTRIES',
        6587,
        'BANK ASIA LTD',
        'JESSORE',
        'saving',
        '0003333002081',
        332000,
        'SI-APFIL-AUG21-68',
        'BP-APFIL-AUG21-8',
        '070410941',
        'APFIL-118112',
        2
      ]
    ]);
    worksheet.addRow();
    let bottom1 = worksheet.addRow([
      'In Word: Fourteen Lakh Twelve Thousand Three Hundred And Twenty Taka Only'
    ]);
    bottom1.font = { size: 11, underline: 'single', bold: true };
    worksheet.mergeCells(`A${bottom1.number}:L${bottom1.number}`);
    let bottom2 = worksheet.addRow([
      'For Akij Poly Fibre Industries Ltd. Industries Ltd.'
    ]);
    bottom2.font = { size: 11, underline: 'single', bold: true };
    worksheet.mergeCells(`A${bottom2.number}:L${bottom2.number}`);
    worksheet.addRow();
    worksheet.addRow();
    let bottom3 = worksheet.addRow([
      'Authorize Signature',
      '',
      '',
      '',
      'Authorize Signature'
    ]);
    worksheet.mergeCells(
      `A${bottom3.number}:C${bottom3.number},D${bottom3.number}:G${
        bottom3.number
      }`
    );
    // worksheet.mergeCells(`D${bottom3.number}:F${bottom3.number}`);
    workbook.xlsx.writeBuffer().then(data => {
      let blob = new Blob([data], {
        type:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      fs.saveAs(blob, 'CarData.xlsx');
    });
    console.log('test');
  };
  return (
    <div>
      <h1>Hello StackBlitz!</h1>
      <p>Start editing to see some magic happen :)</p>
      <button
        onClick={() => {
          createFile();
        }}
      >
        Export
      </button>
    </div>
  );
}
