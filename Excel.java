package THIRD;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	XSSFWorkbook wb;
	XSSFSheet s;
	Excel(String excelPath,String sheetname) throws IOException
	{
		FileInputStream fis=new FileInputStream(excelPath);
		wb=new XSSFWorkbook(fis);
		s=wb.getSheet(sheetname);
	}
	public void setCellData(int rowindex,int colindex,String data,String excelPath) throws IOException
	{
		s.getRow(rowindex).createCell(colindex).setCellValue(data);
		FileOutputStream fos=new FileOutputStream(excelPath);
		wb.write(fos);
	}
	public String getCellData(int rowindex,int colindex)
	{
		return s.getRow(rowindex).getCell(colindex).getStringCellValue();
	}

}



