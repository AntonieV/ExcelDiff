# ExcelDiff

A tool to compare two excel files with annotation of the differences. 

sudo apt install python3-pip
pip3 install --upgrade pip
(requires pip>=22.2.1)


pip3 uninstall exceldiff
pip3 install .
exceldiff -v -i ../test_table_1 ../test_table_2 -o ../diff
