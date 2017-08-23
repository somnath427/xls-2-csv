package br.com.xbrain.app;

import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

/**
 *
 * @author xbrain
 */
public class ConverterTest {

    private static final String FILE_NAME_XLS = "new_excel.xls";

    private static final String FILE_NAME_CSV = "csv_converted.csv";

    @Test
    public void converterFile() throws IOException {
        createAXlsFile();
        Converter.execute(FILE_NAME_XLS, FILE_NAME_CSV);
        BufferedReader br = new BufferedReader(new FileReader(FILE_NAME_CSV));
        String line = br.readLine();
        Assert.assertEquals("Name;City;Address;Email;", line);
    }

    private void createAXlsFile() throws FileNotFoundException, IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("FirstSheet");

        HSSFRow rowhead = sheet.createRow((short) 0);
        rowhead.createCell(0).setCellValue("Name");
        rowhead.createCell(1).setCellValue("City");
        rowhead.createCell(2).setCellValue("Address");
        rowhead.createCell(3).setCellValue("Email");

        HSSFRow row = sheet.createRow((short) 1);
        row.createCell(0).setCellValue("Thiago Marcello");
        row.createCell(1).setCellValue("Londrina-PR");
        row.createCell(2).setCellValue("Brazil");
        row.createCell(3).setCellValue("thiagocmarcello@gmail.com");

        FileOutputStream fileOut = new FileOutputStream(FILE_NAME_XLS);
        workbook.write(fileOut);
        fileOut.close();
    }

    private void createAXlsFileWithColumnNumeric() throws FileNotFoundException, IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("FirstSheet");

        HSSFRow rowhead = sheet.createRow((short) 0);
        rowhead.createCell(0).setCellValue("Name");
        rowhead.createCell(1).setCellValue("City");
        rowhead.createCell(2).setCellValue("Address");
        rowhead.createCell(3).setCellValue("Email");
        rowhead.createCell(4).setCellValue("Zip");

        HSSFRow row = sheet.createRow((short) 1);
        row.createCell(0).setCellValue("Thiago Marcello");
        row.createCell(1).setCellValue("Londrina-PR");
        row.createCell(2).setCellValue("Brazil");
        row.createCell(3).setCellValue("thiagocmarcello@gmail.com");
        row.createCell(3).setCellValue("86025040");

        FileOutputStream fileOut = new FileOutputStream(FILE_NAME_XLS);
        workbook.write(fileOut);
        fileOut.close();
    }
}
