package helper;

import java.io.File;
import java.io.IOException;

import jxl.write.*;
import jxl.write.Number;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.biff.RowsExceededException;

public class ExcelHelper {

    public void createExcelFile(String path) throws IOException, RowsExceededException, WriteException {

        // 1. Create an Excel file
        WritableWorkbook myFirstWbook = null;
        try {

            myFirstWbook = Workbook.createWorkbook(new File(path));

            // create an Excel sheet
            WritableSheet excelSheet = myFirstWbook.createSheet("Sheet 1", 0);

            // add something into the Excel sheet
            Label label = new Label(0, 0, "Test Count");
            excelSheet.addCell(label);

            Number number = new Number(0, 1, 1);
            excelSheet.addCell(number);

            label = new Label(1, 0, "Result");
            excelSheet.addCell(label);

            label = new Label(1, 1, "Passed");
            excelSheet.addCell(label);

            number = new Number(0, 2, 2);
            excelSheet.addCell(number);

            label = new Label(1, 2, "Passed 2");
            excelSheet.addCell(label);

            myFirstWbook.write();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } finally {

            if (myFirstWbook != null) {
                try {
                    myFirstWbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
            }

        }

    }

    public void readExcelFile(String path) {
        Workbook workbook = null;

        try {
            workbook = Workbook.getWorkbook(new File(path));

            Sheet sheet = workbook.getSheet(0);
            Cell cell = sheet.getCell(0, 0);
            Cell cell2 = sheet.getCell(0, 1);

            System.out.println(cell.getContents() + " : " + cell2.getContents());

        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }

    }

    public static void main(String[] args) {
        try {
            new ExcelHelper().createExcelFile("G:\\temp\\Excel.xls");
            new ExcelHelper().readExcelFile("G:\\temp\\Excel.xls");
        } catch (Exception e) {
            e.printStackTrace();
        }

        System.out.println("-------finish--------");
    }
}
