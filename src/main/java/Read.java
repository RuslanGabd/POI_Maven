import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;


public class Read
{
    public static void readFromExcel(String file) throws IOException
    {
        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheet("Лист1");
        XSSFRow row = myExcelSheet.getRow(0);
        XSSFCell cell = row.getCell(0);
        DataFormatter fmt = new DataFormatter();
        ArrayList<String> buffer = new ArrayList<String>();

        Workbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Матрица доступа");

        int i;
        int m = 0;
        while  ( row!=null) //пока не пустая строка

        {  i=0;
            buffer.clear();
            while (i < 14)  //Вычитывание строки в массив
            {
                if (row.getCell(i) == null)
                {
                    buffer.add("");
                    i++;
                    continue;
                }
                else if (row.getCell(i).getCellType() == CellType.NUMERIC)
                {
                    buffer.add(fmt.formatCellValue(row.getCell(i)));
                }
                else if (row.getCell(i).getCellType() == CellType.STRING)
                {
                    buffer.add(row.getCell(i).getStringCellValue());
                }
                else if (row.getCell(i).getCellType() == CellType.BLANK)
                {
                    buffer.add("");
                }
                else if (row.getCell(i).getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty())
                {
                    buffer.add("");
                }

                i++;
                cell = row.getCell(i);

            }
            m++;
            System.out.println(buffer);
            Row row_w = sheet.createRow(m);
            Cell name = row_w.createCell(m);
            name.setCellValue("John");
row=myExcelSheet.getRow(m);

           // Cell birthdate = row_w.createCell(1);



// Get current cell value value and overwrite the value


        }


        book.write(new FileOutputStream("Lenta Area.xlsx"));
        book.close();



        myExcelBook.close();
    }

    public static boolean isCellEmpty(final XSSFCell cell)
    {
        if (cell == null)
        { // use row.getCell(x, Row.CREATE_NULL_AS_BLANK) to avoid null cells
            return true;
        }

        if (cell.getCellType() == CellType.BLANK)
        {
            return true;
        }

        if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty())
        {
            return true;
        }

        return false;
    }

}
