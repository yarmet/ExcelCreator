package ruslan;

import java.io.IOException;

/**
 * Created by ruslan on 22.05.2017.
 */
public class Main {


    public static void main(String[] args) throws IOException {

        XlsWriter xlsWriter = new XlsWriter(2);

        //--------------------------------------------------------------------------------------------------------------
        xlsWriter.changeSheet(0);
        xlsWriter.establishSheetName("лист1");

        String[] columnNames = {"column1" , "column2" , "column3"};

        xlsWriter.createheader(columnNames);

        xlsWriter.createRow();
        xlsWriter.createCell("cell1");
        xlsWriter.createCell("cell2");
        xlsWriter.createCell("cell3");
        xlsWriter.finishRow();

        xlsWriter.createRow();
        xlsWriter.createCell("cell4");
        xlsWriter.createCell("cell5");
        xlsWriter.createCell("cell6");
        xlsWriter.finishRow();

        //--------------------------------------------------------------------------------------------------------------
        xlsWriter.changeSheet(1);
        xlsWriter.establishSheetName("лист2");

        String[] columnNames1 = {"column1" , "column2" , "column3"};

        xlsWriter.createheader(columnNames1);

        xlsWriter.createRow();
        xlsWriter.createCell("cell111111");
        xlsWriter.createCell("cell222222");
        xlsWriter.createCell("cell333333");
        xlsWriter.finishRow();

        xlsWriter.createRow();
        xlsWriter.createCell("cell444444");
        xlsWriter.createCell("cell555555");
        xlsWriter.createCell("cell666666");
        xlsWriter.finishRow();



        xlsWriter.saveInFile("myexcel");

    }
}
