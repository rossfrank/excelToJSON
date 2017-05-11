# excelToJSON
Converts CSV or Excel Files to JSON

## Node
### Install
```
npm install -g xlsx
```
### Run
```
node excelToJSON.js
```

## Python
### Install
```
pip install xlrd
```
### Run
```
python excelToJSON.py
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
