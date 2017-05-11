var xlsx = require('xlsx');
var fs = require('fs');

//file to process
var file = '73acb11627d6702e51c3bfa21a598e76b291a9f2823ff633a4f34f8444165a6a.xls'

//file to output
var output = "test.txt"

console.log("start")
var result = run(file)
console.log("done")
fs.writeFile(output, JSON.stringify(result), function(err) {
    if(err) {
        return console.log(err);
    }
    console.log("The file was saved!");
});

function run(file){
  var json =  {}
  json[file] = arrayPrint(file)
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
  if(file.endsWith(".xls") || file.endsWith(".xlsx")){
    jsonObj['size'] = {}
    var workbook = xlsx.readFile(file)
    workbook.SheetNames.forEach(function(sheet_name) {
      var worksheet = workbook.Sheets[sheet_name]
      if(Object.keys(worksheet).length != 0){
        jsonObj['size'][sheet_name] = worksheet["!ref"]
      }
      jsonObj[sheet_name] = rowArrayToArray(excelsheetRead(workbook, sheet_name))
    })
  }
  else if (file.endsWith(".csv")){
    jsonObj['csv'] = rowArrayToArray(csvRead(file))
  }
  else {
    return "Not a CSV or Excel File"
  }
  return jsonObj
}
