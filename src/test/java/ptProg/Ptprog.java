package ptProg;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
public class Ptprog {
	// Set path here
	String path = "file/pt_file.xlsx";
	String pathInput = "file/pt_file.xlsx";
	String pathOutput = "file/test2.xlsx";
	double value1;
	double value2;
	double sum;
	double div;
	double subst;
	double multipl;

	public static void main(String[] args) throws IOException {
		Ptprog obj = new Ptprog();
		obj.writeInfile();
		// Iterate in cells

		// CLose workbook
	}

	/**
	 * @throws EncryptedDocumentException
	 * @throws IOException
	 */
//	public static void readInFile() throws IOException {
//		// Read file
//		File f = new File(pathInput);
//		FileInputStream fi = new FileInputStream(f);
//		XSSFWorkbook wb = new XSSFWorkbook(fi);
//		XSSFSheet sheet = wb.getSheetAt(0);
////		DataFormatter dataFormatter = new DataFormatter();
//		// Iterate
//		Iterator it = sheet.iterator();
//		while (it.hasNext()) {
//			XSSFRow row = (XSSFRow) it.next();
//			Iterator cellIteator = row.cellIterator();
//			while (cellIteator.hasNext()) {
//				XSSFCell cell = (XSSFCell) cellIteator.next();
//				switch (cell.getCellType()) {
//				case STRING:
//					System.out.println(cell.getStringCellValue());
//					break;
//				case NUMERIC:
//					System.out.println(cell.getNumericCellValue());
//					break;
//				default:
//					break;
//				}
//				System.out.print("|");
//
//			}
//			System.out.println();
//		}
//	}

	public void writeInfile() throws IOException {
		try {
			// Load existing workbook
			FileInputStream fileInputStream = new FileInputStream(pathInput);
			Workbook workbook = new XSSFWorkbook(fileInputStream);

			// Access the sheet
			Sheet sheet = workbook.getSheetAt(0); // Assuming you want the first sheet, change index accordingly

			// Iterate through rows and apply formula for each row
			int lastRowNum = sheet.getLastRowNum();
			for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
				Row row = sheet.getRow(rowNum);
				if (row != null) {
					Cell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					Cell cell2 = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

					value1 = cell1.getCellType() == CellType.NUMERIC ? cell1.getNumericCellValue() : 0;
					value2 = cell2.getCellType() == CellType.NUMERIC ? cell2.getNumericCellValue() : 0;

					sum = value1 + value2;

					// Create a new cell for the sum in the third column
					Cell sumCell = row.createCell(2, CellType.NUMERIC);
					sumCell.setCellValue(sum);
				}
			}

			// Write the changes back to the workbook
			FileOutputStream fileOutputStream = new FileOutputStream(pathOutput);
			workbook.write(fileOutputStream);

			// Close streams
			fileInputStream.close();
			fileOutputStream.close();

			System.out.println("File Created successfully.");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
