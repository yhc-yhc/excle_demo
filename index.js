var fs = require('fs')
fs.unlink('./streamed-workbook.xlsx', function () {
})
var Excel = require('exceljs');

var options = {
    filename: './streamed-workbook.xlsx',
    useStyles: true,
    useSharedStrings: true
};
var workbook = new Excel.stream.xlsx.WorkbookWriter(options);

var styleSheet = workbook.addWorksheet('styleSheet', {
    properties: {
        tabColor: {
            argb: 'FFC0000'
        }
    }
});

var pattern = ['none', 'solid', 'darkVertical', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625', 'darkHorizontal',
    'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',
    'lightGrid', 'lightTrellis', 'lightGrid'];

for (var i = 0; i < pattern.length; i++) {
    styleSheet.addRow([pattern[i]])
}

var col1 = styleSheet.getColumn(1);
col1.eachCell(function (cell, rowNum) {
    cell.fill = {
        type: 'pattern',
        pattern: pattern[rowNum - 1],
        fgColor: {argb: '88ff55'}
    }
})

styleSheet.columns = [
    {header: 'null', width: 40}
];

styleSheet.commit();
//========================================>

var mergeSheet = workbook.addWorksheet('mergeSheet', {
    views: [
        {
            state: 'frozen',
            xSplit: 1,
            ySplit: 3,
            topLeftCell: 'B4',
            activeCell: 'B4'
        }
    ]
});

mergeSheet.columns = [
    {header: '入帐日期', width: 15},
    {header: '一卡通', width: 15},
    {header: '一卡通', width: 15},
    {header: '一卡通', width: 15},
    {header: '一卡通', width: 15},
    {header: '一卡通', width: 15},
    {header: '一卡通', width: 15},
    {header: '一卡通', width: 15},
    {header: '一卡通', width: 15}
]
mergeSheet.addRow(['入帐日期', '支付宝', '支付宝', '微信', '微信', '总笔数', '笔数比重', '总收入', '收比重'])
mergeSheet.addRow(['入帐日期', '交易笔数', '交易金额', '交易笔数', '交易金额', '总笔数', '笔数比重', '总收入', '收比重'])

mergeSheet.getRow(1).font = {name: 'Comic Sans MS', family: 4, size: 16, bold: true};
mergeSheet.getRow(2).font = {name: 'Comic Sans MS', family: 4, size: 14, bold: true};
mergeSheet.getRow(3).font = {name: 'Comic Sans MS', family: 4, size: 12, bold: true};

mergeSheet.mergeCells('A1:A3');
mergeSheet.mergeCells('B1:I1');
mergeSheet.mergeCells('B2:C2');
mergeSheet.mergeCells('D2:E2');
mergeSheet.mergeCells('F2:F3');
mergeSheet.mergeCells('G2:G3');
mergeSheet.mergeCells('H2:H3');
mergeSheet.mergeCells('I2:I3');

for (var i = 1; i < 4; i++) {
    mergeSheet.getRow(i).eachCell(function (cell, rowNum) {
        cell.alignment = {vertical: 'middle', horizontal: 'center'};
    })
}

function createRandom() {
    return Math.floor(Math.random() * 1000)
}

for (var i = 1; i < 120; i++) {
    mergeSheet.addRow([new Date(2016, 1, i), createRandom(), createRandom(), createRandom(), createRandom(), createRandom(), createRandom(), createRandom(), createRandom()])
}

var col1 = mergeSheet.getColumn(8);
col1.eachCell(function (cell, rowNum) {
    if (rowNum > 1 && rowNum < 4) {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {argb: '88ff55'}
        }
    } else if (rowNum > 3) {
        cell.fill = {
            type: 'pattern',
            pattern: 'lightGray',
            fgColor: {argb: '88ff55'}
        }
    }
})

mergeSheet.commit();

workbook.commit()
    .then(function () {
        // the stream has been written
        console.log('..........')
    });