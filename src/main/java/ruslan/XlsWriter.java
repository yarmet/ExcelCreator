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
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;


/**
 * @author ruslan
 */
public class XlsWriter {

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

    public void establishSheetName(String name) {
        workbook.setSheetName(currSheet, name);
    }


    public void createheader(String[] headers) {
        Row rowhead = sheets[currSheet].createRow(0);
        for (int i = 0; i < headers.length; i++) {
            rowhead.createCell(i).setCellValue(headers[i]);
        }
    }

    public void changeSheet(int sheetNumber) {
        currSheet = sheetNumber;
    }


    private Cell createCellAndGet(Consumer<Cell> consumer) {
        Cell cell = row.createCell(cellNumber++);
        consumer.accept(cell);
        return cell;
    }


    public Cell createCell(double value) {
        return createCellAndGet((cell) -> cell.setCellValue(value));
    }

    public Cell createCell(String value) {
        return createCellAndGet((cell) -> cell.setCellValue(value));
    }

    public Cell createCell(boolean value) {
        return createCellAndGet((cell) -> cell.setCellValue(value));
    }

    public Cell createCell(Date value) {
        return createCellAndGet((cell) -> cell.setCellValue(value));
    }

    public Cell createCell(Calendar value) {
        return createCellAndGet((cell) -> cell.setCellValue(value));
    }

    public Cell createCell(RichTextString value) {
        return createCellAndGet((cell) -> cell.setCellValue(value));
    }


    public void mergeCells(int firstRow, int lastRow, int firstColumn, int lastColumn) {
        sheets[currSheet].addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn));
    }


    public void finishRow() {
        ++rowNumber[currSheet];
        cellNumber = 0;
    }


    public XSSFCellStyle createXssfCellStyle() {
        return (XSSFCellStyle) workbook.createCellStyle();
    }


    public void createRow() {
        row = sheets[currSheet].createRow(rowNumber[currSheet]);
    }


    public void saveInFile(String fileName) throws IOException {
        try (FileOutputStream fileOut = new FileOutputStream(fileName.concat(".xlsx"))) {
            workbook.write(fileOut);
        }
    }

}