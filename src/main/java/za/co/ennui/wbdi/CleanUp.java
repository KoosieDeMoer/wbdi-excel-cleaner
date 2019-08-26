package za.co.ennui.wbdi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CleanUp {
	public static void main(String[] args) throws EncryptedDocumentException, IOException {

		File inFile = new File(args[0]);

		// Creating a Workbook from an Excel file (.xls or .xlsx)
		Workbook inputWorkbook = WorkbookFactory.create(inFile, null, true);

		List<String> includedSeries = new ArrayList<>();
		List<String> includedCountries = new ArrayList<>();
		Map<String, Integer> seriesIndexes = new HashMap<>();
		Map<String, Integer> countryIndexes = new HashMap<>();

		File filterFile = null;
		if(args.length > 1) {
			// use the second are as the filter file name
			filterFile = new File(args[1]);
		} else {
			// use the data file as a filter, ie no filter
			filterFile = new File(args[0]);
		}

		Workbook filterWorkbook = WorkbookFactory.create(filterFile, null, true);

		readCriteria(filterWorkbook, "Country", includedCountries, countryIndexes);
		readCriteria(filterWorkbook, "Series", includedSeries, seriesIndexes);


		Workbook outputWorkbook = new XSSFWorkbook();
		Sheet outputDataSheet = outputWorkbook.createSheet("Data");
		Sheet outputSeriesSheet = outputWorkbook.createSheet("Series");
		Sheet outputCountrySheet = outputWorkbook.createSheet("Country");

		// get the series
		Sheet seriesSheet = inputWorkbook.getSheet("Series");

		Row outRow = outputSeriesSheet.createRow(0);
		outRow.createCell(0).setCellValue("Series Code");
		outRow.createCell(1).setCellValue("Topic");
		outRow.createCell(2).setCellValue("Indicator Name");

		// header row
		Row dataSheetHeaderRow = outputDataSheet.createRow(0);
		dataSheetHeaderRow.createCell(0).setCellValue("Country Code\\Series Code");

		for (int seriesNo = 1; seriesNo <= seriesSheet.getLastRowNum(); seriesNo++) {
			Row inRow = seriesSheet.getRow(seriesNo);

			String seriesCode = inRow.getCell(0).getStringCellValue();
			if (includedSeries.contains(seriesCode)) {

				int seriesIndex = seriesIndexes.get(seriesCode) + 1;

				outRow = outputSeriesSheet.createRow(seriesIndex);

				outRow.createCell(0).setCellValue(seriesCode);
				outRow.createCell(1).setCellValue(inRow.getCell(1).getStringCellValue());
				outRow.createCell(2).setCellValue(inRow.getCell(2).getStringCellValue());

				// add the series header
				dataSheetHeaderRow.createCell(seriesIndex).setCellValue(seriesCode);
			}
		}

		// get the countries
		Sheet countriesSheet = inputWorkbook.getSheet("Country");

		outRow = outputCountrySheet.createRow(0);
		outRow.createCell(0).setCellValue("Country Code");
		outRow.createCell(1).setCellValue("Short Name");
		outRow.createCell(2).setCellValue("Long Name");
		outRow.createCell(3).setCellValue("2-alpha code");

		for (int countryNo = 1; countryNo <= countriesSheet.getLastRowNum(); countryNo++) {
			Row inRow = countriesSheet.getRow(countryNo);
			String threeLetterCountryCode = inRow.getCell(0).getStringCellValue();
			if (includedCountries.contains(threeLetterCountryCode)) {
				int countryIndex = countryIndexes.get(threeLetterCountryCode) + 1;

				outRow = outputCountrySheet.createRow(countryIndex);
				outRow.createCell(0).setCellValue(threeLetterCountryCode);
				outRow.createCell(1).setCellValue(inRow.getCell(1).getStringCellValue());
				outRow.createCell(2).setCellValue(inRow.getCell(3).getStringCellValue());
				Cell countryShortCodeCell = inRow.getCell(4);
				if (countryShortCodeCell != null) {
					outRow.createCell(3).setCellValue(countryShortCodeCell.getStringCellValue());
				}

				// put a blank row in for each country
				Row outDataSheetRow = outputDataSheet.createRow(countryIndex);
				outDataSheetRow.createCell(0).setCellValue(threeLetterCountryCode);
			}

		}

		// get the data
		Sheet dataSheet = inputWorkbook.getSheet("Data");

		for (int rowNo = 1; rowNo <= dataSheet.getLastRowNum(); rowNo++) {
			Row row = dataSheet.getRow(rowNo);
			String indicatorCode = row.getCell(3).getStringCellValue();
			String threeLetterCountryCode = row.getCell(1).getStringCellValue();
			if (includedSeries.contains(indicatorCode) && includedCountries.contains(threeLetterCountryCode)) {
				short lastCellNum = row.getLastCellNum();
				if (lastCellNum > 4) {

					double numericCellValue = row.getCell(lastCellNum - 1).getNumericCellValue();
					outputDataSheet.getRow(countryIndexes.get(threeLetterCountryCode) + 1)
							.createCell(seriesIndexes.get(indicatorCode) + 1).setCellValue(numericCellValue);
				}
			}
		}

		FileOutputStream fileOut = new FileOutputStream(inFile.getParent() + File.separator + "clean.xlsx");
		outputWorkbook.write(fileOut);
		fileOut.close();

		// Closing the workbook
		inputWorkbook.close();
		outputWorkbook.close();

	}

	private static void readCriteria(Workbook filterWorkbook, String sheetName, List<String> includedItems,
			Map<String, Integer> itemIndexes) {

		Sheet filterSheet = filterWorkbook.getSheet(sheetName);
		for (int itemNo = 1; itemNo <= filterSheet.getLastRowNum(); itemNo++) {
			Row inRow = filterSheet.getRow(itemNo);
			String code = inRow.getCell(0).getStringCellValue();

			includedItems.add(code);
			itemIndexes.put(code, itemNo - 1);

		}
	}

}
