# XLSX Generator
This is the implementation of a script to merge multiple csv into one xlsx file. This script arranges values according \
to their type.

# Get Started

### Environment
+ python version >= 3.6
+ openpyxl

## Execution Method
### From Terminal

To execute the script from terminal run the following command. Please change the file names according to your file names.

```
python3 ./xlsx_gen.py \
    --output path_to_output.xlsx \
    --files path_to_file_1.csv \
        path_to_file_2.csv \
        path_to_file_3.csv \
        ..
        ..
        path_to_file_n.csv
```


### From shell file
The script can also be run from **run.sh** file. To do so change the file names according to your file names in the **run.sh** file. \
And then run the following command.

```
./run.sh
```

### From another python file
To run this script from another python file first you have to import **MergeFile** class from **xlsx_gen.py**. Then you have to create \
an object of that class and call **read_files** method of the **MergeFile** object. A demo code is provided below.

```
from xlsx_gen import MergeFile
    
file_path_array = [
    'file_path_1.csv',
    'file_path_2.csv',
    ..
    ..
    ..
    'file_path_n.csv'
]

output_file_path = 'path_to_output_file.xlsx'

merge_file_object = MergeFile(file_list = file_path_array, output_file = output_file_path)
merge_file_object.read_files()
```

