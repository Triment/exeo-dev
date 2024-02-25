import { Workbook, type CellValue } from "exceljs"

const convert = async () => {
    //mock data
    const surnames = [
        "王", "李", "张", "刘", "陈", "杨", "黄", "赵", "周", "吴", "徐", "孙", "胡", "朱", "高",
        "林", "何", "郭", "马", "罗", "梁", "宋", "郑", "谢", "韩", "唐", "冯", "于", "董", "萧", "程",
        "曹", "袁", "邓", "许", "傅", "沈", "曾", "彭", "吕", "苏", "卢", "蒋", "蔡", "贾", "丁", "魏",
        "薛", "叶", "阎", "余", "潘", "杜", "戴", "夏", "钟", "汪", "田", "任", "姜", "范", "方", "石",
        "姚", "谭", "廖", "邹", "熊", "金", "陆", "郝", "孔", "白", "崔", "康", "毛", "邱", "秦", "江",
        "史", "顾", "侯", "邵", "孟", "龙", "万", "段", "雷", "钱", "汤", "尹", "黎", "易", "常", "武",
        "乔", "贺", "赖", "龚", "文"
    ];

    const names = [
        "丽", "敏", "军", "伟", "芳", "秀英", "娜", "静", "磊", "洋", "燕", "勇", "艳", "强", "杰", "娟",
        "涛", "霞", "明", "刚", "华", "超", "露", "平", "玲", "辉", "慧", "亮", "红", "桂英", "健", "新",
        "云", "博", "丹", "莉", "亚", "宇", "霞", "刘", "琴", "洁", "翔", "婷", "宁", "宏", "凯", "丽",
        "浩", "桂兰", "凤", "翠", "庆", "莹", "波", "平", "凡", "瑞", "金", "欣", "梅", "飞", "泉", "秋",
        "东", "倩", "青", "亮", "宝", "成", "琳", "春", "峰", "艳", "冰", "敏", "建", "栋", "冰", "菲",
        "龙", "云", "霞", "伟", "敏", "宏", "欢", "瑞", "秋", "鹏", "莹", "辉", "燕", "磊", "梅", "坤"
    ];

    const usedNames = new Set();

    function generateChineseName() {
        let surnameIndex = Math.floor(Math.random() * surnames.length);
        let nameIndex = Math.floor(Math.random() * names.length);
        let fullName = surnames[surnameIndex] + names[nameIndex];

        // 如果生成的名字已经被使用过，则重新生成，直到生成一个未被使用的名字为止
        while (usedNames.has(fullName)) {
            surnameIndex = Math.floor(Math.random() * surnames.length);
            nameIndex = Math.floor(Math.random() * names.length);
            fullName = surnames[surnameIndex] + names[nameIndex];
        }

        // 将生成的名字添加到已使用名字的集合中
        usedNames.add(fullName);

        return fullName;
    }

    const positions = ['软件工程师', '产品经理', '市场专员', '财务主管', '人力资源经理', '运营总监', '客户服务专员', '销售经理', '技术支持工程师', '数据分析师'];

function generatePosition() {
  const randomIndex = Math.floor(Math.random() * positions.length);
  return positions[randomIndex];
}
function generateSalary(min:number, max: number) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }
  


    //   
    const departments = ['人力资源', '市场营销', '技术开发', '财务部', '运营管理'];

    function generateDepartment() {
        return departments[Math.floor(Math.random() * departments.length)];
    }
    const usedIDs = new Set();
    function generateEmployeeID() {
        let id;
        do {
            id = generateRandomID();
        } while (usedIDs.has(id)); // 如果生成的工号已经被使用过，则重新生成

        usedIDs.add(id); // 将生成的工号添加到已使用工号的集合中
        return id;
    }
    function generateRandomID() {
        const min = 100000; // 六位数的最小值
        const max = 999999; // 六位数的最大值
        return Math.floor(Math.random() * (max - min + 1)) + min;
    }

    //mock data end
    const xlsx = Bun.file("./工资表.xlsx")
    const workBook = await new Workbook().xlsx.load(await xlsx.arrayBuffer());

    const outFiles: File[] = [];
    let index = 2;
    const sheet = workBook.getWorksheet("Sheet1");
    while (index<50){
        const row = sheet?.getRow(index);
        row!.getCell('A').value = generateChineseName();
        row!.getCell('B').value = generateEmployeeID();
        row!.getCell('C').value = generateDepartment();
        row!.getCell('D').value = generatePosition();
        const sum = [generateSalary(6000, 25000), generateSalary(2000, 10000), generateSalary(2000, 10000), generateSalary(500, 2000), generateSalary(500, 2000)];
        const [e,f,g,h,i] = sum;
        const j = sum.reduce((p, n)=>p+n)
        const k = j < 5000 ? 0 : j <= 8000?j*0.03:j <= 17000?j*0.1:j <= 30000?j*0.2:j <= 40000? j*0.3:j <= 60000?j*0.3:j <= 85000?j*0.35:j*0.45;
        row!.getCell('E').value = generateSalary(6000, 25000);//基本
        row!.getCell('F').value = generateSalary(2000, 10000);//加班
        row!.getCell('G').value = generateSalary(2000, 10000);//绩效
        row!.getCell('H').value = generateSalary(500, 2000);//津贴
        row!.getCell('I').value = generateSalary(500, 2000);//补贴
        row!.getCell('J').value = j;//税前
        row!.getCell('K').value = k;//个税
        const l = j < 4246 ? 4246*0.14:j<21228?j*0.14:21228*0.14;
        row!.getCell('L').value = l;//社保
        const m = j*0.12;
        row!.getCell('M').value = j*0.12;//公积金
        const n = j - k - l - m;
        row!.getCell('N').value = n;//应发
        row!.getCell('O').value = n;//实发
        row!.getCell('P').value = new Date("2024-03-05T12:23:07.123Z");
        row!.getCell('Q').value = 22;//实发
        row!.getCell('R').value = 0;//实发
        row!.getCell('S').value = 22;//实发
        row!.getCell('T').value = 0;//实发
        index++;
    }
    // workBook.eachSheet((sheet, id) => {
    //     sheet.eachRow((row, i) => {
    //         if (i <= 1) return;
    //         // const workBook = new Workbook();
    //         // workBook.xlsx.load(files[1][0].arrayBuffer())
    //         // row.eachCell(cell => {
    //         //     console.log(cell.value)
    //         // })

    //     })
    // })

    const buff = await workBook.xlsx.writeBuffer();
    const file = new File([buff], "xx.xlsx", { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    await Bun.write("./生成工资.xlsx", file);
}

//convert()



const con = async (files: FileList[])=>{//prod
    let sourceFile : File;//prod
// const con = async (files: File[])=>{
//     let sourceFile = files[0];
    //prod
    for (const file of files[0]){
        if (file.name.match(/^(\.\/)?工资.*/)) sourceFile = file;
    }
    //prod end
    function convertToExcelColumn(number: number) {
        let result = '';
        while (number > 0) {
          let remainder = (number - 1) % 26; // 计算余数
          result = String.fromCharCode(65 + remainder) + result; // 将余数转换为对应的字母
          number = Math.floor((number - 1) / 26); // 更新数字
        }
        return result;
    }

    const workBook = await new Workbook().xlsx.load(await sourceFile!.arrayBuffer());
    let outputs: File[] = []
    let rows = workBook.getWorksheet("Sheet1")!.rowCount;
    new Promise<File[]>((resolve, reject)=>{
        workBook.getWorksheet("Sheet1")?.eachRow(async (row, rIndex)=>{
            if(rIndex<=1) return;
            const personalXlsx =  new Workbook();
    
            let employeeName = row.getCell("A").value?.toString();//获取员工名字
            const newSheet = personalXlsx.addWorksheet(employeeName);//建表
            let rowObj: CellValue[] = [];//创建表头
            workBook.getWorksheet("Sheet1")!.getRow(1).eachCell((cell)=>{
                rowObj.push(cell.value);
            })
            newSheet.addRow(rowObj);//创建完成表头
            workBook.getWorksheet("Sheet1")!.getRow(1).eachCell((cell, cIndex)=>{
                newSheet.getRow(1).getCell(cIndex).style = cell.style;//拷贝
            })
            row.eachCell((cell)=>{
                newSheet.getRow(2).getCell(cell.col).value = cell.value;
                newSheet.getRow(2).getCell(cell.col).style = cell.style;
            })
            newSheet.columns = workBook.getWorksheet("Sheet1")!.columns.map((col)=>({width: col.width, alignment: col.alignment}));
            let independent = new File([await personalXlsx.xlsx.writeBuffer()], employeeName + ".xlsx", { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            outputs.push(independent);
            postMessage({
                type: "progress",
                payload: {
                    type: "processing",
                    payload: Math.floor((rIndex/rows)*100)
                }
            });
            //console.log(Math.floor((rIndex/rows)*100), outputs);
            if(rIndex===rows) resolve(outputs);
        })
    }).then((out: File[])=>{
        postMessage({
            type: "result",
            payload: out
        })
    })
}
con([Bun.file("./工资生成.xlsx") as unknown as File])