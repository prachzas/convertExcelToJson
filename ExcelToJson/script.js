const xlsx = require("xlsx");
const fs = require('fs');

var wb = xlsx.readFile("testExcel.xlsx")
var ws = wb.Sheets["Prach"];
var data = xlsx.utils.sheet_to_json(ws);
var result = []
var obj = {}
data.forEach((x => {
    obj.RuleName = "APDown"
    obj.system = "CFMS"
    obj.emsIp = "10.219.209.134"
    obj.manager = "Huawei_WifiOffload"
    obj.domain = "RAN-3G"
    obj.neType = "FBB"
    obj.region = "BKK"
    obj.amoName = "aaaaaaaaaaaaaaa"
    obj.alertName = "AP_Disassociated"
    obj.severity = "critical or major or minor or warning or clear"
    obj.description = "hwApFaultTrap"
    obj.col1 = "wWLCSUK01HW"
    obj.col2 = "AIS_WIFI_NOKIA"
    obj.col3 = "Null value or Shopping mall"
    obj.col4 = "3G2100"
    obj.col5 = "NMC"
    obj.alarmTime = "29/07/2018 10:57:40"
    obj.node = "3RNCBPL2H_KSURM_SR_SAT"
    obj.sitecode = "KSURM"
    obj.vendor = "ZTE"
    obj.serverserial = "POS01_+A4:A222999999999"
    obj.a = x.name
    result.push(obj);
    obj = {}
}));

//please check path 
fs.writeFile("C:/Users/WiN 10 Pro/Desktop/ExcelToJson/json.txt", JSON.stringify(result), 'utf8', function (err) {
    if (err) {
        return console.log(err);
    }
    console.log("The file was saved!");
});

console.log(result)
