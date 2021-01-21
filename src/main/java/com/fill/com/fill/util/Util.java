package com.fill.com.fill.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.print.*;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Sides;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

/**
 * Created by Rohit.Bansal on 15-04-2017.
 */
public class Util {

    public static String CITY = "Sri Ganganagar";

    private Row row;

    public Util(Row row) {
        this.row = row;
    }

    public static String getCellValue(Iterator<Cell> cellIterator) {
        Cell cell = cellIterator.next();
        cell.setCellType(CellType.STRING);
        return (cell.getStringCellValue() != null) ? cell.getStringCellValue().toUpperCase() : cell.getStringCellValue();
    }

    public static Util getUtilObject(String sourceExcelFile, int sheetIndex) throws IOException {
        FileInputStream file = new FileInputStream(new File(sourceExcelFile));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
        Iterator<Row> iterator = sheet.iterator();
        Row headingRow = iterator.next();
        Row currentRow;
        if (iterator.hasNext())
            currentRow = iterator.next();
        else
            return null;

        Util util = new Util(currentRow);
        return util;
    }

    public static Iterator<Row> getUtilIteratorObject(String sourceExcelFile, int sheetIndex) throws IOException {
        FileInputStream file = new FileInputStream(new File(sourceExcelFile));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
        Iterator<Row> iterator = sheet.iterator();
        Row headingRow = iterator.next();
        return iterator;
    }

    public String getCellValue(int cellNumber) {
        Cell cell = row.getCell(cellNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        cell.setCellType(CellType.STRING);
        return (cell.getStringCellValue() != null) ? cell.getStringCellValue().toUpperCase() : "";
    }

    public static String getCellValue(Row row, int cellNumber) {
        Cell cell = row.getCell(cellNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        cell.setCellType(CellType.STRING);
        return (cell.getStringCellValue() != null) ? cell.getStringCellValue().toUpperCase() : "";
    }

    public static String getAmountInWords(String amount) {
        int amt;
        if (amount == null || amount.equalsIgnoreCase(""))
            return "";
        try {
            amt = Integer.parseInt(amount);
            if (amt == 0) {
                return "";
            }
        } catch (Exception e) {
            return "";
        }
        return NumberToWord.convertNumberToWords(amt) + " ONLY";
    }

    public static void printPdfOutput(String fileName) throws IOException, PrintException {
        DocFlavor flavor = DocFlavor.SERVICE_FORMATTED.PAGEABLE;
        PrintRequestAttributeSet patts = new HashPrintRequestAttributeSet();
        patts.add(Sides.ONE_SIDED);
        PrintService[] ps = PrintServiceLookup.lookupPrintServices(flavor, patts);
        if (ps.length == 0) {
            throw new IllegalStateException("No Printer found");
        }
        System.out.println("Available printers: " + Arrays.asList(ps));

        PrintService myService = null;
        for (PrintService printService : ps) {
            if (printService.getName().equals("HP LaserJet 1018")) {
                myService = printService;
                break;
            }
        }

        if (myService == null) {
            throw new IllegalStateException("Printer not found");
        }

        FileInputStream fis = new FileInputStream(fileName);
        Doc pdfDoc = new SimpleDoc(fis, DocFlavor.INPUT_STREAM.AUTOSENSE, null);
        DocPrintJob printJob = myService.createPrintJob();
        printJob.print(pdfDoc, new HashPrintRequestAttributeSet());
        fis.close();
    }

    public static String getDirectoryPath() {
        if (System.getProperty("os.name").toLowerCase().contains("mac"))
            return "/Users/rohibans/personal/Mutual Fund Forms/MFU Set";
        else
            return "C:\\Personal\\Mutual Fund Forms\\MFU Set";
    }

    public static String getDestinationDirectoryPath() {
        if (System.getProperty("os.name").toLowerCase().contains("mac"))
            return "/Users/rohibans/personal/Mutual Fund Forms/MFU Set/Filled Forms";
        else
            return "C:\\Personal\\Mutual Fund Forms\\MFU Set\\Filled Forms";
    }

    public static String getTotalAmount(List<String> amounts) {
        Long totalAmount = 0l;
        for (String s : amounts) {
            try {
                totalAmount = totalAmount + Long.parseLong(s);
            } catch (Exception e) {
                System.out.println("Invalid amount");
            }

        }
        if (totalAmount == 0) {
            return "";
        }
        return totalAmount.toString();
    }

    public String getDateCellValue(int cellNumber) {
        Cell cell = row.getCell(cellNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        return (cell.getDateCellValue() != null) ? sdf.format(cell.getDateCellValue()) : "";
    }

    public static String maskString(String strText, int start, int end, char maskChar)
            throws Exception {

        if (strText == null || strText.equals(""))
            return "";

        if (start < 0)
            start = 0;

        if (end > strText.length())
            end = strText.length();

        if (start > end)
            throw new Exception("End index cannot be greater than start index");

        int maskLength = end - start;

        if (maskLength == 0)
            return strText;

        StringBuilder sbMaskString = new StringBuilder(maskLength);

        for (int i = 0; i < maskLength; i++) {
            sbMaskString.append(maskChar);
        }

        return strText.substring(0, start)
                + sbMaskString.toString()
                + strText.substring(start + maskLength);
    }
}
