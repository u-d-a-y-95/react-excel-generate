import React from 'react';
import './style.css';
import { createFile } from './excel/index';

export default function App() {
  return (
    <div>
      <h1>Excel!</h1>
      <button
        onClick={() => {
          createFile({
            name: 'ACCL-Payment',
            creator: 'Uday',
            createTime: new Date(),
            sheets: [
              {
                name: 'PaymentAdviceReport',
                gridLine: false,
                border: 'all 33adff thin',
                bold: true,
                alignment: 'right:middle',
                fontSize: 15,
                bgColor: 'e6ecff',

                rows: [
                  [
                    {
                      text: 'Akij Cement Company Ltd.',
                      bold: true,
                      // itatlic: true,
                      // underline: true,
                      // fontFamily: string,
                      textFormat: 'text',
                      fontSize: 16,
                      // textColor: '33adff',
                      // bgColor: 'ff0000',
                      // border: 'bottom 33adff thin',
                      alignment: 'center:middle',
                      cellRange: 'A1:M2', // "A1:A3"
                      merge: true
                    }
                  ],
                  [
                    {
                      text:
                        'Akij House, 198 Bir Uttam, Gulshan Link Road, Tejgaon, Dhaka-1208.',
                      bold: true,
                      // itatlic: true,
                      underline: true,
                      // fontFamily: string,
                      textFormat: 'text',
                      fontSize: 12,
                      // textColor: '33adff',
                      // bgColor: 'ff0000',
                      // border: 'bottom 33adff thin',
                      alignment: 'center:middle',
                      cellRange: 'A1:M1', // "A1:A3"
                      merge: true
                    }
                  ],
                  [
                    {
                      text: 'To',
                      bold: true,
                      // itatlic: true,
                      // underline: true,
                      // fontFamily: string,
                      textFormat: 'text',
                      fontSize: 9,
                      // textColor: '33adff',
                      // bgColor: 'ff0000',
                      // border: 'bottom 33adff thin',
                      alignment: 'left:middle',
                      cellRange: 'A1:J1', // "A1:A3"
                      merge: true
                    },
                    {
                      text: 'Date : 23-Jun-2021',
                      bold: true,
                      // itatlic: true,
                      // underline: true,
                      // fontFamily: string,
                      textFormat: 'text',
                      fontSize: 9,
                      // textColor: '33adff',
                      // bgColor: 'ff0000',
                      // border: 'bottom 33adff thin',
                      alignment: 'left:middle',
                      cellRange: 'K1:M1', // "A1:A3"
                      merge: true
                    }
                  ],
                  [
                    {
                      text: 'The Manager',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 9,
                      alignment: 'left:middle',
                      cellRange: 'A1:M1', // "A1:A3"
                      merge: true
                    }
                  ],
                  [
                    {
                      text: 'Standard Chartered Bank',
                      textFormat: 'text',
                      fontSize: 9,
                      alignment: 'left:middle',
                      cellRange: 'A1:M1', // "A1:A3"
                      merge: true
                    }
                  ],
                  [
                    {
                      text:
                        'Gulshan Branch, SCB House, Level-1, 67, Gulshan Avenue, Dhaka-1212',
                      textFormat: 'text',
                      fontSize: 9,
                      alignment: 'left:middle',
                      cellRange: 'A1:M1', // "A1:A3"
                      merge: true
                    }
                  ],
                  [
                    {
                      text: 'Subject : Payment Instruction by BEFTN',
                      textFormat: 'text',
                      bold: true,
                      underline: true,
                      fontSize: 9,
                      alignment: 'left:middle',
                      cellRange: 'A1:M1', // "A1:A3"
                      merge: true
                    }
                  ],
                  [
                    {
                      text: 'Dear Sir,',
                      textFormat: 'text',
                      bold: true,
                      fontSize: 9,
                      alignment: 'left:middle',
                      cellRange: 'A1:M1', // "A1:A3"
                      merge: true
                    }
                  ],
                  [
                    {
                      text:
                        'We do hereby requesting you to make payment by transferring the amount to the respective Account Holder as shown below in detailed by debiting our CD Account No. 01114617101												',
                      textFormat: 'text',
                      bold: true,
                      italic: true,
                      fontSize: 9,
                      alignment: 'left:middle',
                      cellRange: 'A1:M1', // "A1:A3"
                      merge: true
                    }
                  ],
                  [
                    {
                      text: 'Detailed particulars of each Account Holder:',
                      textFormat: 'text',
                      bold: true,
                      fontSize: 9,
                      alignment: 'left:middle',
                      cellRange: 'A1:M1', // "A1:A3"
                      merge: true
                    }
                  ],
                  ['_blank*2'],
                  [
                    'name',
                    'Age',
                    {
                      text: 'Gender',
                      bgColor: 'ff6666',
                      textColor: 'ffffff',
                      border: 'all 000000 medium'
                    },
                    'Nationality'
                  ],
                  ['Uday', '25', 'Male', 'Bangladeshi']
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
