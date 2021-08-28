


** This is a excel utility ""
Its just help to cut of your time to generate excel file
It's use exceljs and file-saver

***
    Dependency
      1.excel js
      2.file-saver
***

Try to give more flexiblity for customizing excel 
To work properly One's should follow some rules.

File Structure 
      key     Type                                                      Required      default value
    {
      name: String,
      creator: String                                                 // optional,    "",
      createTime:Date                                                 // optional,    new Date()
      sheets: 
        [
          {
            name: String,
            gridLine: boolean                                         // optional,    true,
            border: String#BorderFormatString                         // optional,    null
            bold: boolean,                                            // optional,    false
            alignment: String#AlignmentFormatString,                  // optional,    "center:middle"
            fontSize: number,                                         // optional,    9
            bgColor: String#ColorHexaCode,                            // optional,    no color 
            rows: 
              [
                [
                  {
                      text: String,
                      bold: boolean,                                  // optional,    false
                      itatlic: boolean,                               // optional,    false
                      underline: boolean,                             // optional,    false
                      textFormat: String ,                            // optional,    Text
                      fontSize: Number ,                              // optional,    9
                      textColor: String#ColorHexaCode,                // Optional,    000000(black)
                      bgColor: String#ColorHexaCode,                  // optional,    no color
                      border: String#BorderFormatString               // optional,    null
                      alignment: String#AlignmentFormatString,        // optional,    "center:middle"
                      cellRange: String#cellRangeFormatString,        // optional,    null
                      merge: boolean                                  // optional,    false
                  }
                ],
                ['_blank*2'],                                       // its special format to gap rows
                [
                  String,
                  Number',
                  {
                    text: String,
                    bold: boolean,                                  // optional,    false
                    itatlic: boolean,                               // optional,    false
                    underline: boolean,                             // optional,    false
                    textFormat: String ,                            // optional,    Text
                    fontSize: Number ,                              // optional,    9
                    textColor: String#ColorHexaCode,                // Optional,    000000(black)
                    bgColor: String#ColorHexaCode,                  // optional,    no color
                    border: String#BorderFormatString               // optional,    null
                    alignment: String#AlignmentFormatString,        // optional,    "center:middle"
                    cellRange: String#cellRangeFormatString,        // optional,    null
                    merge: boolean                                  // optional,    false
                  }
                  'String'
                ],
              ]
          }
        ]
    }


    AlignmentFormatString "

