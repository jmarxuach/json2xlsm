
# json2xlsm
Adds json data file into xlsm first sheet macro excel file. 
Usefull from PHP or any language that doesn't have Excel Macro library. 

## Usage

If you want to use a xlsm excel sheet from PHP, python or any other language without xlms library, json2xlsm will fill first sheet data with JSON file. Usually you create a xlsm with your macros in Vb and you do not need to edit macros, but you need to replace sheet data in XLSM files.

```
java -jar json2xlsm.jar <strFileJSON> <strMacroExcelFileIn> <strMacroExcelFileOut>
```
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
