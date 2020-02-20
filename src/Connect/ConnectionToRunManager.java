package Connect;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	


public class ConnectionToRunManager {
	public void testCaseName()
	{
		String TestCaseName = null;
		try
		{
		String Projectpath = System.getProperty("user.dir");
		
		String FileName = Projectpath + "\\src\\RunManager.xlsx";
		
		File f = new File(FileName);
		
		FileInputStream fs = new FileInputStream(f);
		
		XSSFWorkbook wb = new XSSFWorkbook(fs);
		
		XSSFSheet ws = wb.getSheet("Regression");
		
		int TotRows = ws.getPhysicalNumberOfRows();
		
		for (int i = 1; i <= TotRows; i++)
		{
			XSSFRow row =  ws.getRow(i);
			String flag = row.getCell(2).getStringCellValue();
			if (flag.toLowerCase().equals("yes"))
			{
			TestCaseName = row.getCell(1).getStringCellValue();
			System.out.println(TestCaseName);
			ConnectionToBusinessFlow.keyWords(TestCaseName);
			}
		}
		
		}
		catch(Exception e)
		{
			System.out.println(e.getMessage());
		}
		
		
	}

	}

