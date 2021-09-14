package MavenProject.MavenPKG;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sun.tools.classfile.Opcode.Set;
import com.sun.tools.javac.api.JavacTaskPool.Worker;

public class XlsReader {
	Worker work_book;
	Set sheet;
	Row row;

	public XlsReader(String wbpath) {
		try {
			work_book = new XSSFWorkbook(wbpath);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public int getRowCount(String sheetName) {
		sheet = work_book.getSheet(sheetName);
		int lastRowNo = sheet.getLastRowNum();
		int firstRowNo = sheet.getFirstRowNum();

		int rcount = (lastRowNo - firstRowNo) + 1;
		return rcount;
	}

	public int getcellCount() {
		row = sheet.getRow(0);
		int lastRowNo = sheet.getLastRowNum();
		int firstRowNo = sheet.getFirstRowNum();

		int cellcount = (lastRowNo - firstRowNo);
		return cellcount;

	}

	public String getCellData(int row, int col) {
		String value = sheet.getRow(row).getCell(col).getStringCellValue();
		return value;
	}

	public static void main(String[] args) {
		String wbpath = System.getProperty("user.dir")+"\\src\\TestData\\test.xlsx";
		XlsReader xlr = new XlsReader(wbpath);
		int rowcount = xlr.getRowCount("Sheet1");
		int colcount = xlr.getcellCount();
		System.out.println(rowcount+" - "+colcount);
		for (int i = 1; i < rowcount; i++) {
			for (int j = 0; j < colcount; j++) {
				String data = xlr.getCellData(i, j);
				System.out.print(data + "  ");

			}
			System.out.println();
		}

	}

}
