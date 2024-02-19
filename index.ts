import { Workbook } from 'exceljs';


const convert = async (files: FileList[]) => {
    const sourceFiles = files[0];
    const templateFiles = files[1];
    let 源订单文件: File, 源扫描表格文件: File;
    for (const file of sourceFiles){
        if(file.name.match(/^(\.\/)?转运单导出数据表.*/)){//生产环境使用 /^转运单导出数据表.*/
            源订单文件 = file;
        }
        if(file.name.match(/^(\.\/)?扫描数据表.*/)){
            源扫描表格文件 = file;
        }
    }
    console.log(源扫描表格文件!);
    const workBook = new Workbook();
    await workBook.xlsx.load(await 源扫描表格文件!.arrayBuffer());
    workBook.getWorksheet(1)?.eachRow((row, rIndex)=>{
        if(rIndex > 1){
            console.log(row.getCell('A').value)
        }
    })
}



(async () => {
    //创建工作簿
    // const workBook = new Workbook();

    // await workBook.xlsx.readFile('./转运单导出数据表 (2).xlsx');
    // workBook.getWorksheet(1)!.eachRow((row)=>{
    //     row.eachCell((cell)=>{
    //         console.log(cell.value)
    //     })
    // })
    const file = Bun.file('./转运单导出数据表 (2).xlsx') as unknown as File;
    const scanFile = Bun.file('./扫描数据表.xlsx') as unknown as File;
    const templateFile = Bun.file('./扫描数据表.xlsx') as unknown as File;
    convert([([file, scanFile] as unknown as FileList), [templateFile] as unknown as FileList]);
})()