var xlsx = require('xlsx');
var fs = require('fs');

//file to process
var file = 'large_file.xlsx'


console.log("start")
var result = run(file)
console.log("done")
fs.writeFile("test.txt", JSON.stringify(result), function(err) {
    if(err) {
        return console.log(err);
    }
    console.log("The file was saved!");
});

function run(file){
  var json =  {}
  json[file] = arrayPrint(file, (file.endsWith(".xls") || file.endsWith(".xlsx")))
  return json
}

function csvRead(file){
  return fs.readFileSync(file).toString().split("\n")
}

function excelsheetRead(workbook, sheet_name){
  var csv = xlsx.utils.sheet_to_csv(workbook.Sheets[sheet_name])
  return csv.split("\n")
}

function rowArrayToArray(rowArray) {
  var splitArray = []
  rowArray.forEach(function(row) {
    if(row != ""){
      var row_s = row.split(",")
      splitArray.push(row_s)
    }
  })
return splitArray
}

function arrayPrint(file, excel) {
  var jsonObj = {}
  if(excel){
    jsonObj['size'] = {}
    var workbook = xlsx.readFile(file)
    workbook.SheetNames.forEach(function(sheet_name) {
      var worksheet = workbook.Sheets[sheet_name]
      if(Object.keys(worksheet).length != 0){
        jsonObj['size'][sheet_name] = worksheet["!ref"];
      }
      jsonObj[sheet_name] = rowArrayToArray(excelsheetRead(workbook, sheet_name))
    })
  }
  else{
    jsonObj['csv'] = rowArrayToArray(csvRead(file))
  }
  return jsonObj
}
