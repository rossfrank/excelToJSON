# excelToJSON
Converts CSV or Excel Files to JSON

## Run
```
npm install xlsx
```
then
```
node excelToJSON.js
```

### Output Format:
For Excel:
```
{FileName:
   {size:
      {sheet1 : A1:C4, sheet2 : B2:E5},
    sheet1: 2D Array of Data,
    sheet2: 2D Array of Data}}   
```
For CSV:
```
{FileName:
   {csv: 2D Array of Data}}
```
