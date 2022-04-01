import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	public static void main(String[] args) throws IOException {

		FileInputStream file = new FileInputStream("C:/Users/sudheer.pragnapuram/OneDrive - Datacom/Desktop/DataDriven/excel.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		int sheets = workbook.getNumberOfSheets();

		for(int i=0; i<sheets; i++) {

			if(workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {

				XSSFSheet sheet = workbook.getSheetAt(i);

				Iterator<Row> rows= sheet.iterator();
				Row firstrow =  rows.next();

				Iterator<Cell> cell = firstrow.iterator();
				int k=0;
				int column = 0;
				//To scan column
				while (cell.hasNext())
				{
					Cell cvalue = cell.next();

					if(cvalue.getStringCellValue().equalsIgnoreCase("TestCases"))
					{
						column=k;
					}

					k++;
				}
				System.out.println(column);
				//To scan Rows
				while (rows.hasNext())
				{
					Row rvalue = rows.next();

					if(rvalue.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase"))
					{
						Iterator<Cell> cv = rvalue.cellIterator();
						while(cv.hasNext()) {
							System.out.println(cv.next().getStringCellValue());
						}
					}


				}

			}
		}

	}

}
