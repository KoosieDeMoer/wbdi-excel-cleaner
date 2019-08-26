# wbdi-excel-cleaner
Transforms and filters World Bank Development Index annual data to features &amp; instances

## Usage:
java -jar wbdiExcelCleaner.jar INPUT_FILE_PATH <FILTER_FILE_PATH>

produces a file alongside INPUT_FILE_PATH named clean.xlsx


## Correlation - written in Java for performance comparison with SciPy
### Usage
java -cp wbdiExcelCleaner.jar za.co.ennui.wbdi.Correlate CLEAN_FILE_PATH

produces a file alongside CLEAN_FILE_PATH named correlations.xlsx

