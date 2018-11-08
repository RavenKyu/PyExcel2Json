# PyExcel2Json
It converts Excel to JSON.

## Features
- Define a specific range (e.g `A1:D1`)
    - If not defined, It'll find the `HEAD` and `DATA` range.

## Install
```shell
pip install pyexcel2json
``` 
or
```shell
git clone https://github.com/RavenKyu/PyExcel2Json.git
```

## Usage
### Help
```shell
pyxl2json --help

usage: C:\Users\deokyu\AppData\Local\Programs\Python\Python36-32\Scripts\pyxl2json.exe
       [-h] [--head HEAD] [--data DATA] [--sheet SHEET] [--verbose] [--asfile]
       excel_filename [excel_filename ...]

PyExcel2Json is easy to convert Excel to Json

positional arguments:
  excel_filename

optional arguments:
  -h, --help            show this help message and exit
  --head HEAD, -t HEAD  A1:F1
  --data DATA, -d DATA  A2:F10
  --sheet SHEET         SHEET NAME
  --verbose, -v         print JSON output converted
  --asfile              save as file(s)

```

### Example

`test.xlsx`

| |A|B|C|D|
|----|-----|---------|-------|----|
|1|NAME |	VALUE	|COLOR	|DATE|
|2|Alan	|12|	blue|	Sep. 25, 2009|
|3|Shan	|13|	green|	blue	Sep. 27, 2009|
|4|John	|45	|orange	|Sep. 29, 2009|
|5|Minna	|27	|teal	|Sep. 30, 2009|

#### Define a specific `HEAD` range
```shell
pyxl2json test.xlsx -v --head A1:D1 
```
result
```json
[
    {
        "NAME": "Alan",
        "VALUE": 12,
        "COLOR": "blue",
        "DATE": "Sep. 25, 2009"
    },
    {
        "NAME": "Shan",
        "VALUE": 13,
        "COLOR": "green\tblue",
        "DATE": "Sep. 27, 2009"
    },
    {
        "NAME": "John",
        "VALUE": 45,
        "COLOR": "orange",
        "DATE": "Sep. 29, 2009"
    },
    {
        "NAME": "Minna",
        "VALUE": 27,
        "COLOR": "teal",
        "DATE": "Sep. 30, 2009"
    }
]
```
It's not necessary to define. you can also use like this. It'll be same as abobe.
```shell
pyxl2json test.xlsx -v  
```

#### Select `Data` range
```shell
pyxl2json test.xlsx -v --data A4:D5 
```
result
```json
[
    {
        "NAME": "John",
        "VALUE": 45,
        "COLOR": "orange",
        "DATE": "Sep. 29, 2009"
    },
    {
        "NAME": "Minna",
        "VALUE": 27,
        "COLOR": "teal",
        "DATE": "Sep. 30, 2009"
    }
]
```

#### Save as JSON file
```shell
pyxl2json test.xlsx --asfile
# it will be saved as test.json
```
or
```shell
pyxl2json test.xlsx -v > test.json
``` 

#### Convert Multiple Excel Files
```shell
pyxl2json --asfile test_1.xlsx test_2.xlsx test_3.xlsx 
``` 

