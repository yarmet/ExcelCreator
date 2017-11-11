package ruslan;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.IOException;
import java.util.Date;

/**
 * Created by ruslan on 22.05.2017.
 */
public class Main {


    public static void main(String[] args) throws IOException {

        XlsWriter xlsWriter = new XlsWriter(2);


        XSSFCellStyle greenBackground = xlsWriter.createXssfCellStyle();
        greenBackground.setFillForegroundColor(new XSSFColor(new java.awt.Color(50, 150, 50)));
        greenBackground.setFillPattern(CellStyle.SOLID_FOREGROUND);

        XSSFCellStyle redBackground = xlsWriter.createXssfCellStyle();
        redBackground.setFillForegroundColor(new XSSFColor(new java.awt.Color(150, 50, 50)));
        redBackground.setFillPattern(CellStyle.SOLID_FOREGROUND);
        //--------------------------------------------------------------------------------------------------------------

        xlsWriter.changeSheet(0);
        xlsWriter.establishSheetName("лист1");

        String[] columnNames = {"column1", "column2", "column3"};

        xlsWriter.mergeCells(10, 12, 3, 3);
        xlsWriter.createheader(columnNames);


        xlsWriter.createRow();
        xlsWriter.createCell("cell1");
        xlsWriter.createCell(true);

        Cell cell1 = xlsWriter.createCell(new Date());
        cell1.setCellStyle(redBackground);
        xlsWriter.finishRow();


        xlsWriter.createRow();
        xlsWriter.createCell("cell4");
        xlsWriter.createCell(false);

        Cell cell = xlsWriter.createCell(new Date());
        cell.setCellStyle(greenBackground);

        xlsWriter.finishRow();

        xlsWriter.saveInFile("myexcel");
    }
}
