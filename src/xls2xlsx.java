

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Partial copy of one of style and data..
 * 
 * @author jcalfee
 */
public class xls2xlsx {

    /**
     * @param args
     * @throws InvalidFormatException
     * @throws IOException
     */
    public static void main(String[] args) throws InvalidFormatException,
            IOException {

        String inpFn = "test1.xls";
        String outFn = "test1.xlsx";

        InputStream in = new BufferedInputStream(new FileInputStream(inpFn));
        try {
            Workbook wbIn = new HSSFWorkbook(in);
            File outF = new File(outFn);
            if (outF.exists())
                outF.delete();

            Workbook wbOut = new SXSSFWorkbook();
            int sheetCnt = wbIn.getNumberOfSheets();
            for (int i = 0; i < sheetCnt; i++) {
                Sheet sIn = wbIn.getSheetAt(0);
                Sheet sOut = wbOut.createSheet(sIn.getSheetName());
                Iterator<Row> rowIt = sIn.rowIterator();
                while (rowIt.hasNext()) {
                    Row rowIn = rowIt.next();
                    Row rowOut = sOut.createRow(rowIn.getRowNum());

                    Iterator<Cell> cellIt = rowIn.cellIterator();
                    while (cellIt.hasNext()) {
                        Cell cellIn = cellIt.next();
                        Cell cellOut = rowOut.createCell(
                                cellIn.getColumnIndex(), cellIn.getCellType());

                        switch (cellIn.getCellType()) {
                        case Cell.CELL_TYPE_BLANK:
                            break;

                        case Cell.CELL_TYPE_BOOLEAN:
                            cellOut.setCellValue(cellIn.getBooleanCellValue());
                            break;

                        case Cell.CELL_TYPE_ERROR:
                            cellOut.setCellValue(cellIn.getErrorCellValue());
                            break;

                        case Cell.CELL_TYPE_FORMULA:
                            cellOut.setCellFormula(cellIn.getCellFormula());
                            break;

                        case Cell.CELL_TYPE_NUMERIC:
                            cellOut.setCellValue(cellIn.getNumericCellValue());
                            break;

                        case Cell.CELL_TYPE_STRING:
                            cellOut.setCellValue(cellIn.getStringCellValue());
                            break;
                        }

                        {
                            CellStyle styleIn = cellIn.getCellStyle();
                            CellStyle styleOut = cellOut.getCellStyle();
                            styleOut.setDataFormat(styleIn.getDataFormat());
                        }
                        cellOut.setCellComment(cellIn.getCellComment());

                        // HSSFCellStyle cannot be cast to XSSFCellStyle
                        // cellOut.setCellStyle(cellIn.getCellStyle());
                    }
                }
            }
            OutputStream out = new BufferedOutputStream(new FileOutputStream(
                    outF));
            try {
                wbOut.write(out);
            } finally {
                out.close();
            }
        } finally {
            in.close();
        }
    }
}

