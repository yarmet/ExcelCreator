package ruslan;

/**
 * Created by ruslan on 22.05.2017.
 */

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;


/**
 * @author ruslan
 */
public class XlsWriter {

    private final DateFormat DATE_FORMAT = new SimpleDateFormat("dd.MM.yyyy");
    private Row row = null;
    private final SXSSFWorkbook workbook;
    private int currSheet = 0;
    private final int[] rowNumber;
    private SXSSFSheet[] sheets = null;
    private int cellNumber = 0;

    public XlsWriter(int sheetCount) {
        rowNumber = new int[sheetCount];
        workbook = new SXSSFWorkbook();
        sheets = new SXSSFSheet[sheetCount];
        for (int i = 0; i < sheetCount; i++) {
            rowNumber[i] = 1;
            sheets[i] = (SXSSFSheet) workbook.createSheet();
        }
    }

    protected void establishSheetName(String name) {
        workbook.setSheetName(currSheet, name);
    }

    protected void createheader(String[] headers) {
        Row rowhead = sheets[currSheet].createRow(0);
        for (int i = 0; i < headers.length; i++) {
            rowhead.createCell(i).setCellValue(headers[i]);
        }
    }

    protected void changeSheet(int sheetNumber) {
        currSheet = sheetNumber;
    }


    protected void createCell(double value) {
        row.createCell(cellNumber++).setCellValue(value);
    }

    protected void createCell(String value) {
        row.createCell(cellNumber++).setCellValue(value);
    }

    protected void createCell(boolean value) {
        row.createCell(cellNumber++).setCellValue(value);
    }

    protected void createCell(Date value) {
        row.createCell(cellNumber++).setCellValue(value);
    }

    protected void createCell(Calendar value) {
        row.createCell(cellNumber++).setCellValue(value);
    }

    protected void createCell(RichTextString value) {
        row.createCell(cellNumber++).setCellValue(value);
    }


    protected void finishRow() {
        ++rowNumber[currSheet];
        cellNumber = 0;
    }

    protected void createRow() {
        row = sheets[currSheet].createRow(rowNumber[currSheet]);
    }

    public void saveInFile(String fileName) throws IOException {
        try (FileOutputStream fileOut = new FileOutputStream(fileName.concat(".xlsx"))) {
            workbook.write(fileOut);
        }
    }

}