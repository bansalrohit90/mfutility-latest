package com.fill.com.fill.util;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPageable;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.print.*;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Sides;
import java.awt.print.PageFormat;
import java.awt.print.Paper;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
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

    public static String CITY = "NEW DELHI";

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

    public String getCellValue(int cellNumber) {
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

    static public void printPdfOutput2(String fileName, String jobName) throws PrinterException, IOException {
        PrinterJob job = PrinterJob.getPrinterJob();
        PageFormat pf = job.defaultPage();
        Paper paper = new Paper();
        paper.setSize(595, 842.0);
        double margin = 10;
        paper.setImageableArea(margin, margin, paper.getWidth() - margin, paper.getHeight() - margin);
        pf.setPaper(paper);
        pf.setOrientation(PageFormat.LANDSCAPE);

        // PDFBox
        PDDocument document = PDDocument.load(fileName);
        job.setPageable(new PDPageable(document, job));

        job.setJobName(jobName);
        try {
            job.print();
        } catch (PrinterException e) {
            System.out.println(e);
        }
    }

    public static String getDirectoryPath() {
        return "C:\\Personal\\Mutual Fund Forms\\MFU Set";
    }

    public static String getDestinationDirectoryPath() {
        //return getDirectoryPath();
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
}
