package br.com.xbrain.app;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.NumberFormat;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author xbrain
 */
public class Converter {

    public static void execute(String inputFilePath, String outputFilePath) {
        if (StringUtils.isNotEmpty(inputFilePath)
                && StringUtils.isNotEmpty(outputFilePath)) {
            File inputFile = new File(inputFilePath);
            File outputFile = new File(outputFilePath);
            xls(inputFile, outputFile);
        }
    }

    private static void xls(File inputFile, File outputFile) {
        StringBuilder data = new StringBuilder();
        try {
            FileOutputStream fos = new FileOutputStream(outputFile);
            Workbook wb = WorkbookFactory.create(inputFile);
            Sheet mySheet = wb.getSheetAt(0);
            Row row;
            Iterator<Row> rowIterator = mySheet.rowIterator();
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                for (int i = 0; i < row.getLastCellNum(); i++) {
                    if (row.getCell(i) != null) {
                        switch (row.getCell(i).getCellType()) {
                            case Cell.CELL_TYPE_BOOLEAN:
                                data.append(row.getCell(i).getBooleanCellValue()).append(";");
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                double numericValue = row.getCell(i).getNumericCellValue();
                                if (((long) numericValue) == numericValue) {
                                    NumberFormat numberFormat = NumberFormat.getIntegerInstance();
                                    numberFormat.setGroupingUsed(false);
                                    data.append(numberFormat.format(numericValue)).append(";");
                                }
                                break;
                            case Cell.CELL_TYPE_STRING:
                                data.append(row.getCell(i).getStringCellValue().replaceAll(";", "")).append(";");
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                data.append(";");
                                break;
                            default:
                                data.append(row.getCell(i)).append(";");
                        }
                    } else {
                        data.append(";");
                    }
                }
                data.append('\n');
            }
            fos.write(data.toString().getBytes());
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Converter.class.getName()).log(Level.SEVERE, ex.getMessage(), ex);
        } catch (IOException ex) {
            Logger.getLogger(Converter.class.getName()).log(Level.SEVERE, ex.getMessage(), ex);
        } catch (EncryptedDocumentException | InvalidFormatException ex) {
            Logger.getLogger(Converter.class.getName()).log(Level.SEVERE, ex.getMessage(), ex);
        }
    }

}
