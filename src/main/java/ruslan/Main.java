package ruslan;


import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import java.io.IOException;


/**
 * Created by ruslan on 22.05.2017.
 */
public class Main {


    public static void main(String[] args) throws IOException {

        XlsWriter xlsWriter = new XlsWriter(2);


        XSSFCellStyle greenBackground = xlsWriter.createXSSFCellStyle();
        greenBackground.setFillForegroundColor(new XSSFColor(new java.awt.Color(50, 150, 50)));
        greenBackground.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle redBackground = xlsWriter.createXSSFCellStyle();
        redBackground.setFillForegroundColor(new XSSFColor(new java.awt.Color(150, 50, 50)));
        redBackground.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //--------------------------------------------------------------------------------------------------------------

        xlsWriter.changeSheet(0);
        xlsWriter.establishCurrentSheetName("лист1");
        xlsWriter.createNewRow();


        xlsWriter.createCellAndGet(123);
        xlsWriter.createCellAndGet(23);


        xlsWriter.createCellAndGet(33);

        Double teslDouble = null;
        xlsWriter.createCellAndGet(teslDouble);
        xlsWriter.createCellAndGet(44);

        xlsWriter.createCellAndGet(77);

        Boolean teslBoolean = null;
        xlsWriter.createCellAndGet(teslBoolean);


        xlsWriter.createCellAndGet(33);



        xlsWriter.changeSheet(1);




        xlsWriter.changeSheet(0);

        xlsWriter.createCellAndGet(11);
        xlsWriter.createCellAndGet(11);

        xlsWriter.saveInFile("myexcel");
    }
}
