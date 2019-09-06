# wbdi-excel-cleaner
Transforms and filters [World Bank Development Index](https://databank.worldbank.org/source/world-development-indicators) annual data to features &amp; instances

## Usage:
java -jar wbdiExcelCleaner.jar INPUT_FILE_PATH <FILTER_FILE_PATH>


produces a file alongside INPUT_FILE_PATH named clean.xlsx


If you get a "GC overhead limit exceeded" because you have a large input file, try a little of: -Xmx8192m


## Correlation - written in Java for performance comparison with SciPy (which does it 60% faster)
### Usage
java -cp wbdiExcelCleaner.jar za.co.ennui.wbdi.Correlate CLEAN_FILE_PATH

produces a file alongside CLEAN_FILE_PATH named correlations.xlsx

