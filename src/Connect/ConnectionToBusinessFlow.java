package Connect;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConnectionToBusinessFlow {
	
	public static String keyWords(String TestCase)
	{
		String Keyword = null;
			try
			{
			String Projectpath = System.getProperty("user.dir");
			
			String FileName = Projectpath + "\\src\\DataTable\\BusinessFlow.xlsx";
			
			File f = new File(FileName);
			
			FileInputStream fs = new FileInputStream(f);
			
			XSSFWorkbook wb = new XSSFWorkbook(fs);
			
			XSSFSheet ws = wb.getSheet("BusinessFlow");
			
			int TotRows = ws.getLastRowNum();
			
			for (int i = 1; i <= TotRows; i++)
			{
				XSSFRow row =  ws.getRow(i);
				String TCName_InBusinessFlow = row.getCell(1).getStringCellValue();
				if (TestCase.equals(TCName_InBusinessFlow))
				{
					
				}
				
			}
			
			}
			catch(Exception e)
			{
				System.out.println(e.getMessage());
			}
		return Keyword;
		}
	}
