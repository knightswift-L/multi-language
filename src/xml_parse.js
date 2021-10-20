const fs = require("fs")

export function parseXMLToXlsx(originFile,targetFile){
    let data = [];
    let filedata = fs.readFileSync(originFile);
    let tempData = filedata.toString().split("\n")
    let lineString;
    for(let index = 0;index < tempData.length;index++){
        lineString = tempData[index].trimStart();
        if(lineString.startsWith("<text ")){
        let first = lineString.split("=\"")
        let second = strs[strs.length-1].split("\">")
        let key = second[0]
        let value = second[second.length-1].replace("</text>","")
        data.push({key:value})
        }
    }
    console.log(data);
    return data;
}


export default{

}