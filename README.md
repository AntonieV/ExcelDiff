# SpreadSheetDiff

A tool to compare two Excel or OpenOffice files (.xlsx or .ods format)
with annotation of the differences. The files 
to be compared must have the same basic structure. That means the labels and the 
number of their sheets must be the same in both files. 

## Download and Installation for Linux:

### Install or upgrade pip:
````
sudo apt install python3-pip
pip3 install --upgrade pip
````
[//]: # (&#40;requires pip>=22.2.1&#41;)

Clone the repo or download the latest release in .zip or .tar.gz format.

### Download and extract repository in .tar.gz format: 

Check if `wget` is installed:
````
which wget; echo $?
````
If `wget` isn't installed, it's command returns 1.

If necessary install `wget`:
````
sudo apt install wget
````

For .tar.gz format execute the following command in terminal:
````
wget https://github.com/AntonieV/SpreadSheetDiff/archive/refs/tags/latest.tar.gz -O - | tar -xz
````

### Alternatively: download and extract repository in .zip format: 
Check if `unzip` and `wget` are installed:
````
which unzip; echo $?
which wget; echo $?
````
If one of the tools isn't installed, it's command returns 1.
If necessary install `wget` and/or `unzip`:
````
sudo apt install unzip
sudo apt install wget
````

For .zip format execute the following command in terminal:
````
wget https://github.com/AntonieV/SpreadSheetDiff/archive/refs/tags/latest.zip
unzip latest.zip
rm latest.zip
````

### Installation:
````
cd SpreadSheetDiff
pip3 install .
````

## Compare two Excel/OpenOffice files:

### Input files:

As input .xlsx or .ods files are accepted.

### Output files:


The differences are stored in the cells of an .xlsx results file. A difference
is marked in corresponding cell with `>>>` annotation between the different 
values: `[value of first input file] >>> [value of second input file]`

![image info](/home/tarja/Downloads/Screenshot_SpreadSheetDiff.png)


In addition, they are stored with the corresponding cell localization 
(sheet name, row and column) in a SpreadSheetDiff_annotations file in 
.txt format:

![image info](./assets/Screenshot_SpreadSheetDiff_annot.png)

### Help:

````
spreadsheetdiff -h
````
Help output:
````
usage: spreadsheetdiff [-h] -i INPUT_FILES INPUT_FILES -o OUT_DIR [-b] [-c BG_COLOR] [-v] [-q]

A tool to compare two excel files with annotation of the differences.

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT_FILES INPUT_FILES, --input-files INPUT_FILES INPUT_FILES
                        Two paths to the Excel files (.xlsx or .ods format) to be compared with each other.
  -o OUT_DIR, --out-dir OUT_DIR
                        Path to the output directory.
  -b, --bold            Displays differences in resulting spreadsheet in bold text style.
  -c BG_COLOR, --bg-color BG_COLOR
                        Displays differences in resulting spreadsheet in a specific color. The color has to be given in quoted hex color code (e.g.
                        '#ff0000') or color name (e.g. red).
  -v, --verbose         Increase output verbosity to debug level.
  -q, --quiet           Decrease output verbosity to warning level. Ignores -v flag.

````

### Examples package execution:
````
spreadsheetdiff -v -c '#ff0000' -b -i ../test_table_1.ods ../test_table_3.xlsx -o ../diff

spreadsheetdiff -q -c 'yellow' -i ../test_table_3.xlsx ../test_table_4.xlsx -o ../diff
````
Example output:
````
31-07-2022 22:32:25 | INFO | Starting SpreadSheetDiff analysis...
31-07-2022 22:32:25 | INFO | Analysing sheet 'Tabelle1'
31-07-2022 22:32:25 | INFO | In sheet 'Tabelle1' [row: 7, col: Test2]: dfghd >>> asdf
31-07-2022 22:32:25 | INFO | In sheet 'Tabelle1' [row: 13, col: Test1]: 12 >>> 112
31-07-2022 22:32:25 | INFO | Analysing sheet 'Tabelle2'
31-07-2022 22:32:25 | INFO | SpreadSheetDiff analysis finished!
````

### Example for local execution of the main method:

````
python3 spreadsheetdiff/main.py -v -i ../file_3.xlsx ../file_2.ods -o ../diff

````

### Example of module import:

````
from spreadsheetdiff import main as ssd

style = [("bold", True), ("bg_color", "#ff0000")]
ssd.compare_excel_files('../file_1.ods', '../file_2.ods', '../diff', style)

style = [("bg_color", "yellow")]
ssd.compare_excel_files('../file_1.ods', '../file_3.xlsx', '../diff', style)
````

## Uninstall package:
````
pip3 uninstall spreadsheetdiff
````


