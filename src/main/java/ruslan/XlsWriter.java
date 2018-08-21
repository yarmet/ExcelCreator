package ruslan;

/**
 * Created by ruslan on 22.05.2017.
 */

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;


/**
 * @author ruslan
 */
public class XlsWriter {

    private final SXSSFWorkbook workbook;
    private int currSheetNumber = 0;
    private CustomSheetWrapper[] customSheetWrappers;

    public XlsWriter(int sheetCount) {
        workbook = new SXSSFWorkbook();
        customSheetWrappers = new CustomSheetWrapper[sheetCount];
        for (int i = 0; i < sheetCount; i++) {
            customSheetWrappers[i] = new CustomSheetWrapper(workbook.createSheet());
        }
    }

    public void establishCurrentSheetName(String name) {
        workbook.setSheetName(currSheetNumber, name);
    }

    public void changeSheet(int sheetNumber) {
        if (sheetNumber < 0 || sheetNumber >= customSheetWrappers.length) {
            throw new IllegalArgumentException("wrong sheet number");
        }
        currSheetNumber = sheetNumber;
    }

    public Cell createCellAngGet(double value) {
        return customSheetWrappers[currSheetNumber].createCellAngGet((cell) -> cell.setCellValue(value));
    }

    public Cell createCellAngGet(String value) {
        return customSheetWrappers[currSheetNumber].createCellAngGet((cell) -> cell.setCellValue(value));
    }

    public Cell createCellAngGet(boolean value) {
        return customSheetWrappers[currSheetNumber].createCellAngGet((cell) -> cell.setCellValue(value));
    }

    public Cell createCellAngGet(Date value) {
        return customSheetWrappers[currSheetNumber].createCellAngGet((cell) -> cell.setCellValue(value));
    }

    public Cell createCellAngGet(Calendar value) {
        return customSheetWrappers[currSheetNumber].createCellAngGet((cell) -> cell.setCellValue(value));
    }

    public Cell createCellAngGet(RichTextString value) {
        return customSheetWrappers[currSheetNumber].createCellAngGet((cell) -> cell.setCellValue(value));
    }

    public void mergeCells(int firstRow, int lastRow, int firstColumn, int lastColumn) {
        customSheetWrappers[currSheetNumber].mergeCells(firstRow, lastRow, firstColumn, lastColumn);
    }

    public XSSFCellStyle createXSSFCellStyle() {
        return (XSSFCellStyle) workbook.createCellStyle();
    }

    public void createNewRow() {
        customSheetWrappers[currSheetNumber].createRow();
    }

    public int getRowNumber() {
        return customSheetWrappers[currSheetNumber].getRowNumber();
    }

    public void saveInFile(String fileName) throws IOException {
        try (FileOutputStream fileOut = new FileOutputStream(fileName.concat(".xlsx"))) {
            workbook.write(fileOut);
        }
    }


    private class CustomSheetWrapper {
        private int currentRowNumber = 0;
        private int currentCellNumber = 0;
        private Row row;
        private SXSSFSheet sheet;

        public CustomSheetWrapper(SXSSFSheet sheet) {
            this.sheet = sheet;
        }

        private void createRow() {
            row = sheet.createRow(currentRowNumber++);
            currentCellNumber = 0;
        }

        private Cell createCellAngGet(Consumer<Cell> consumer) {
            Cell cell = row.createCell(currentCellNumber++);
            consumer.accept(cell);
            return cell;
        }

        private void mergeCells(int firstRow, int lastRow, int firstColumn, int lastColumn) {
            sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn));
        }

        private int getRowNumber() {
            return currentRowNumber;
        }

    }

}