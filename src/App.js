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
            name: 'Demo-1',
            sheets: [
              {
                name: 'HIGHEST POPULATION',
                border: 'all 000000 thin',
                alignment: 'center:middle',
                rows: [
                  [
                    {
                      text: 'Country',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 12,
                      textColor: 'ffffff',
                      bgColor: '666699',
                      cellRange: 'A1:A2',
                      merge: true
                    },
                    {
                      text: 'Population(2000)',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 12,
                      textColor: 'ffffff',
                      bgColor: '666699',
                      cellRange: 'B1:B2',
                      merge: true
                    },
                    {
                      text: 'Population(2021)',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 12,
                      textColor: 'ffffff',
                      bgColor: '666699',
                      cellRange: 'C1:C2',
                      merge: true
                    },
                    {
                      text: 'Expected Population(2050)',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 12,
                      textColor: 'ffffff',
                      bgColor: '666699',
                      cellRange: 'D1:D2',
                      merge: true
                    },
                    {
                      text: 'Pop Growth (2000 - 2021)',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 12,
                      textColor: 'ffffff',
                      bgColor: '666699',
                      cellRange: 'E1:E2',
                      merge: true
                    }
                  ],
                  [
                    'China',
                    1268301605,
                    1444216107,
                    '1,329,570,095',
                    new Date()
                  ],
                  [
                    null,
                    '1,006,300,297',
                    '1,393,409,038',
                    '1,623,588,384',
                    '38.5%'
                  ],
                  [
                    'United States',
                    '282,162,411',
                    '332,129,757',
                    '388,922,201',
                    '17.7%'
                  ],
                  [
                    'Indonesia',
                    '214,090,575',
                    '276,361,783',
                    '318,393,046',
                    '29.1%'
                  ],
                  [
                    'Pakistan',
                    '152,429,036',
                    '225,199,937',
                    '290,847,790',
                    '47.7%'
                  ]
                ]
              }
            ]
          });
        }}
      >
        Demo 1
      </button>
      <button
        onClick={() => {
          createFile({
            name: 'Demo-2',
            creator: 'Uday',
            sheets: [
              {
                name: 'HIGHEST Growth',
                gridLine: false,
                rows: [
                  [
                    {
                      text: 'Akij Cement Company Ltd.',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 16,
                      alignment: 'center:middle',
                      cellRange: 'A1:M2',
                      merge: true
                    }
                  ],
                  [
                    {
                      text:
                        'Akij House, 198 Bir Uttam, Gulshan Link Road, Tejgaon, Dhaka-1208.',
                      bold: true,
                      underline: true,
                      textFormat: 'text',
                      fontSize: 12,
                      alignment: 'center:middle',
                      cellRange: 'A1:M1',
                      merge: true
                    }
                  ],
                  [
                    {
                      text: 'To',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 9,
                      alignment: 'left:middle',
                      cellRange: 'A1:J1', // "A1:A3"
                      merge: true
                    },
                    {
                      text: 'Date : 23-Jun-2021',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 9,
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
                    {
                      text: 'Country',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 12,
                      textColor: 'ffffff',
                      bgColor: '666699',
                      cellRange: 'A1:A2',
                      merge: true
                    },
                    {
                      text: 'Population(2000)',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 12,
                      textColor: 'ffffff',
                      bgColor: '666699',
                      cellRange: 'B1:B2',
                      merge: true
                    },
                    {
                      text: 'Population(2021)',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 12,
                      textColor: 'ffffff',
                      bgColor: '666699',
                      cellRange: 'C1:C2',
                      merge: true
                    },
                    {
                      text: 'Expected Population(2050)',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 12,
                      textColor: 'ffffff',
                      bgColor: '666699',
                      cellRange: 'D1:D2',
                      merge: true
                    },
                    {
                      text: 'Pop Growth (2000 - 2021)',
                      bold: true,
                      textFormat: 'text',
                      fontSize: 12,
                      textColor: 'ffffff',
                      bgColor: '666699',
                      cellRange: 'E1:E2',
                      merge: true
                    }
                  ],
                  [
                    'China',
                    '1,268,301,605',
                    '1,444,216,107',
                    '1,329,570,095',
                    '13.8%'
                  ],
                  [
                    'India',
                    '1,006,300,297',
                    '1,393,409,038',
                    '1,623,588,384',
                    '38.5%'
                  ],
                  [
                    'United States',
                    '282,162,411',
                    '332,129,757',
                    '388,922,201',
                    '17.7%'
                  ],
                  [
                    'Indonesia',
                    '214,090,575',
                    '276,361,783',
                    '318,393,046',
                    '29.1%'
                  ],
                  [
                    'Pakistan',
                    '152,429,036',
                    '225,199,937',
                    '290,847,790',
                    '47.7%'
                  ]
                ]
              }
            ]
          });
        }}
      >
        Demo 2
      </button>
    </div>
  );
}
