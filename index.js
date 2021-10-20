const fs = require('fs')
const Excel = require('exceljs');
var args = process.argv.slice(2)
var src = args[0]
var dist = args[1]
var type = args[2]
const arbToXlsx = async function (src, dist) {
  const fileStream = await fs.readFileSync(src);
  const arr = JSON.parse(fileStream)
  writeToXlsx(arr, dist)
}

var excelToJson = async function (src) {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(src);
  const sheet = workbook.getWorksheet('Sheet')
  var zh = {}
  var en = {}
  var es = {}
  var pt = {}
  sheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
    if (rowNumber != 1) {
      zh[row.getCell(1).value] = row.getCell(2).value
      en[row.getCell(1).value] = row.getCell(3).value
      es[row.getCell(1).value] = row.getCell(4).value
      pt[row.getCell(1).value] = row.getCell(5).value
    }
  });
  return {
    "zh": zh,
    "en": en,
    "es": es,
    "pt": pt
  }
}


function writeToArb(data, dist) {
  for (let key of Object.keys(data)) {
    fs.writeFileSync(dist + `app_${key}.arb`, formatArbData(data[key]))
  }
}

function formatArbData(data) {
  let str = "{\n"
  for (let id of Object.keys(data)) {
    str += `    \"${id}\":\"${data[id]}\",\n`
  }
  str += "}"
  return str;
}

function writeToXlsx(data, dist) {

  const workbook = new Excel.Workbook()
  const worksheet = workbook.addWorksheet('Sheet');
  worksheet.views = [
    { state: 'frozen', xSplit: 1, ySplit: 0 }
  ];
  worksheet.columns = [
    { header: 'Id', key: 'id', width: 30, style: { alignment: { horizontal: "left", vertical: "middle" }, protection: { locked: true } } },
    { header: 'ZH', key: 'ZH', width: 50, style: { alignment: { horizontal: "center", vertical: "middle" }, protection: { locked: true } } },
    { header: 'EN', key: 'EN', width: 50, style: { alignment: { horizontal: "center", vertical: "middle" }, protection: { locked: true } } },
    { header: 'ES', key: 'ES', width: 50, style: { alignment: { horizontal: "center", vertical: "middle" }, protection: { locked: true } } },
    { header: 'PT.', key: 'PT', width: 50, style: { alignment: { horizontal: "center", vertical: "middle" }, protection: { locked: true } } }
  ];
  for (let item of Object.keys(data)) {
    worksheet.addRow({
      id: item,
      ZH: data[item],
      EN: "",
      ES: "",
      PT: ""
    });
  }
  workbook.xlsx.writeFile(dist)
}

function formatCSharpXmlData(data) {
  let str = "<?xml version=\"1.0\" encoding=\"utf-8\" ?>\n<localizationDictionary culture=\"zh-Hans\"\>\n<texts></texts>\n"
  for (let id of Object.keys(data)) {
    str += `    <text name = "${id}\">\"${data[id]}\"</text>\n`
  }
  str += "      </texts>\n</localizationDictionary>"
  return str;
}

function writeToCsharpXml(data, dist) {
  let name = {
    "zh": "zh-Hans"
  }
  for (let key of Object.keys(data)) {
    fs.writeFileSync(dist + `exchange-${name[key] ?? key}.xml`, formatCSharpXmlData(data[key]))
  }
}

function parseXMLToXlsx(originFile) {
  let data = {};
  let filedata = fs.readFileSync(originFile);
  let tempData = filedata.toString().split("\n")
  let lineString;
  let start = false;
  for (let index = 0; index < tempData.length; index++) {
    lineString = tempData[index].trimStart();
    if (lineString === "<!-- begin -->") {
      start = true;
      continue;
    }
    if (lineString.startsWith("<text ") && start) {
      let first = lineString.split("=\"")
      let second = first[first.length - 1].split("\">")
      let key = second[0]
      let value = second[second.length - 1].replace("</text>", "")
      data[key] = value;
    }
  }
  return data;
}
function parseIOSToJson(originFile) {
  let data = {};
  let filedata = fs.readFileSync(originFile);
  let tempData = filedata.toString().split("\n")
  let lineString;
  let start = false;
  for (let index = 0; index < tempData.length; index++) {
    lineString = tempData[index].trimStart();
    let first = lineString.split("\"=\"")
    let second = first[first.length - 1].split("\"")
    let key = first[0].toString().replace("\"","");
    let value = second[0]
    data[key] = value;
  }
  return data;
}

function parseAndroidXMLToXlsx(originFile) {
  let data = {};
  let filedata = fs.readFileSync(originFile);
  let tempData = filedata.toString().split("\n")
  let lineString;
  for (let index = 0; index < tempData.length; index++) {
    lineString = tempData[index].trimStart();
    if (lineString.startsWith("<string ")) {
      let first = lineString.split("=\"")
      let second = first[first.length - 1].split("\">")
      let key = second[0]
      let value = second[second.length - 1].replace("</string>", "")
      data[key] = value;
    }
  }
  return data;
}


if (type == 0) {
  //Flutter arb to xlsx
  arbToXlsx(src, dist);
} else if (type == 1) {
  //xlsx to Flutter arb
  let parseXlsxToArb = async function () {
    let data = await excelToJson(src)
    console.log(data)
    writeToArb(data, dist)
  }
  parseXlsxToArb()
} else if (type == 2) {
  //C# xml to xlsx
  let data = parseXMLToXlsx(src);
  writeToXlsx(data, dist)
} else if (type == 3) {
  //xlsx to C# xml 
  let parseXlsxToXml = async function () {
    let data = await excelToJson(src)
    console.log(data)
    writeToCsharpXml(data, dist)
  }
  parseXlsxToXml()
}else if(type ==4){
//android xml to xlsx 
let data = parseAndroidXMLToXlsx(src);
  writeToXlsx(data, dist)
}else if(type ==5){
//xlsx to android xml 
}else if(type ==6){
  //IOS to xlsx
  let data = parseIOSToJson(src);
  writeToXlsx(data,dist);
}else{
  //xlsx to IOS
}




