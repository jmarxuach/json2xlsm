
# json2xlsm
Command line tool for update data into xlsm first sheet of a macro excel file. 
Usefull from PHP, python or any language that doesn't have Excel Macro library that keeps visual basic and macro code after updating xlsm data.

## Usage

Usually you create a xlsm with your macros in Vb and you do not need to edit macros, but you need to replace sheet data in XLSM files.

```
java -jar json2xlsm.jar <strFileJSON> <strMacroExcelFileIn> <strMacroExcelFileOut>
```

Where : 

- strFileJSON : Is the data to insert into the first excel sheet.
- strMacroExcelFileIn : Is your report template with macros in and the first sheet empty.
- strMacroExcelFileOut : Is the resulting excel file with json data in the sheet and your vb code intact. I you have an Workbook_Open report will generate on open excel.

## Creating JSON file from Python and executing json2xlsm

```python
import json
import os

data = [
{'field1': 'Value', 'field2': 'Value', 'field3': 'Value'},
{'field1': 'Value', 'field2': 'Value', 'field3': 'Value'},
{'field1': 'Value', 'field2': 'Value', 'field3': 'Value'},
]

with open('jsonFilename.json', 'w') as fout:
    json.dump(data , fout)

os.system('java -jar json2xlsm.jar jsonFilename.json MacroExcelTemplateFile.xlsm MacroExcelFileOut.xlsm')

```

## Creating JSON file from PHP and executing json2xlsm

All values must be in UTF8.

```php

$array = array(
    0 => array("field1" => "Value", "field2" => "Value"),
    1 => array("field1" => "Value", "field2" => "Value"),
    2 => array("field1" => "Value", "field2" => "Value"),
);

$jsonString = json_encode($array);

file_put_contents("jsonFilename.json", $jsonString);

shell_exec("java -jar json2xlsm.jar jsonFilename.json MacroExcelTemplateFile.xlsm MacroExcelFileOut.xlsm");

```
