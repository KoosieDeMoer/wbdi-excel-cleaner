package za.co.ennui.wbdi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.math3.linear.Array2DRowRealMatrix;
import org.apache.commons.math3.linear.RealMatrix;
import org.apache.commons.math3.stat.correlation.PearsonsCorrelation;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Correlate {
	public static void main(String[] args) throws EncryptedDocumentException, IOException {

		File inFile = new File(args[0]);

		// Creating a Workbook from an Excel file (.xls or .xlsx)
		Workbook inputWorkbook = WorkbookFactory.create(inFile, null, true);

		Workbook outputWorkbook = new XSSFWorkbook();
		Sheet outputCorrelationsSheet = outputWorkbook.createSheet("Correlations");

		Sheet dataSheet = inputWorkbook.getSheet("Data");

		int instanceCount = dataSheet.getLastRowNum();

		Row dataSheetHeaderRow = dataSheet.getRow(0);

		int featureCount = dataSheetHeaderRow.getLastCellNum() - 1;
		
		String[] featureNames = new String[featureCount];
		Row headerRow = dataSheet.getRow(0);
		for (int featureNo = 0; featureNo < featureCount; featureNo++) {
			Cell cell = headerRow.getCell(featureNo + 1);
			featureNames[featureNo] = cell.getStringCellValue();
		}		

		RealMatrix data = new Array2DRowRealMatrix(instanceCount, featureCount);

		for (int instanceNo = 0; instanceNo < instanceCount; instanceNo++) {
			Row row = dataSheet.getRow(instanceNo + 1);
			for (int featureNo = 0; featureNo < featureCount; featureNo++) {
				Cell cell = row.getCell(featureNo + 1);
				if(cell != null) {
				data.addToEntry( instanceNo, featureNo, cell.getNumericCellValue());
				}
			}

		}
		
		
		PearsonsCorrelation pearsonsCorrelation = new PearsonsCorrelation();
		long startTime = System.nanoTime();

		RealMatrix correlationMatrix = pearsonsCorrelation.computeCorrelationMatrix(data);
		
		long endTime = System.nanoTime();

		long duration = (endTime - startTime) / 1000000;  
		System.out.println("Pearson correlation duration for " + instanceCount + " instances and " + featureCount + " features: " + duration + "ms");
		
		// header row
		Row resultSheetHeaderRow = outputCorrelationsSheet.createRow(0);
		resultSheetHeaderRow.createCell(0).setCellValue("Series Code");
		for (int featureNo = 0; featureNo < featureCount; featureNo++) {
			resultSheetHeaderRow.createCell(featureNo + 1).setCellValue(featureNames[featureNo]);
		}		
		
		for (int featureNo = 0; featureNo < featureCount; featureNo++) {
			Row resultSheetRow = outputCorrelationsSheet.createRow(featureNo + 1);
			resultSheetRow.createCell(0).setCellValue(featureNames[featureNo]);
			for (int featureNo2 = 0; featureNo2 < featureCount; featureNo2++) {
				resultSheetRow.createCell(featureNo2 + 1).setCellValue(correlationMatrix.getEntry(featureNo, featureNo2));
			}
		}		

		FileOutputStream fileOut = new FileOutputStream(inFile.getParent() + File.separator + "correlations.xlsx");
		outputWorkbook.write(fileOut);
		fileOut.close();

		// Closing the workbook
		inputWorkbook.close();
		outputWorkbook.close();

	}

}
