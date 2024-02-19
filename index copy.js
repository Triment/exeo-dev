import { Workbook } from 'exceljs';

//创建工作簿
const workBook = new Workbook();

// workbook.creator = 'Me';//创建者
// workbook.lastModifiedBy = 'Her';//最后修改者
// workbook.created = new Date(1985, 8, 30);//创建时间
// workbook.modified = new Date();//修改时间
// workbook.lastPrinted = new Date(2016, 9, 27);//最后打印时间

// // 将工作簿日期设置为 1904 年日期系统
// workbook.properties.date1904 = true;

// workbook.views = [
//     {
//       x: 0, y: 0, width: 10000, height: 20000,
//       firstSheet: 0, activeTab: 1, visibility: 'visible'
//     }
// ]//工作表视图，一个表可以有多个视图，视图中的修改会合并到最后的结果

//创建表格
const sheet = workBook.addWorksheet("新的表格", { 
    properties: { 
        tabColor: { argb: '99FF80CC'}, //表格下表颜色
        outlineLevelRow: 0, //工作表行大纲级别
        outlineLevelCol: 0, // 工作表列大纲级别
        defaultRowHeight: 20, //默认行高
        defaultColWidth: 15, //默认列宽
        dyDescent: 55, // "TBD" 在 Excel 中通常代表 "To Be Determined"，意思是“待定”。这通常用于标记某些信息或数值尚未确定，或者在未来会有进一步确定。
    },
    views: [
        {
            state: 'frozen',//冻结
            showGridLines: false, //表格线框隐藏（无网格）
            xSplit: 1, //第一行和第一列
            ySplit:1
        }
    ],
    pageSetup:{paperSize: 9, orientation:'landscape'},//设置表格方向大小
    headerFooter:{firstHeader: "Hello Exceljs", firstFooter: "Hello World"}//页眉页脚
});

workBook.xlsx.writeBuffer().then(buff=>{
    Bun.write('./test.xlsx', buff)
})

// 遍历所有工作表
// 注意： workbook.worksheets.forEach 仍然是可以正常运行的， 但是以下的方式更好
workBook.eachSheet(function(worksheet, sheetId) {
    // ...
    console.log(sheetId)// 1
});

console.log(
    sheet.rowCount,//文档的总行数。 等于具有值的最后一行的行号。
    sheet.actualRowCount, //具有值的行数的计数。 如果中间文档行为空，则该行将不包括在计数中。
    sheet.columnCount, //文档的总列数。 等于所有行的最大单元数。
    sheet.actualColumnCount //具有值的列数的计数。
)

// 按 name 提取工作表
const worksheet1 = workBook.getWorksheet('My Sheet');

// 按 id 提取工作表
const worksheet2 = workBook.getWorksheet(1)


/**
 * sheet显示状态
 */

// 使工作表可见
sheet.state = 'visible';

// 隐藏工作表
sheet.state = 'hidden';

// 从“隐藏/取消隐藏”对话框中隐藏工作表
sheet.state = 'veryHidden';

//通过表的id删除表
//workBook.removeWorksheet(sheet.id)