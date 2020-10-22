package com.fill;

import com.fill.com.fill.util.Util;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

public final class FillKYCRegistrationForm {

    private static String PAN;
    private static String PREFIX;
    private static String FIRST_NAME;
    private static String MIDDLE_NAME;
    private static String SURNAME;
    private static String FATHER_FIRST_NAME;
    private static String FATHER_MIDDLE_NAME;
    private static String FATHER_LAST_NAME;
    private static String MOTHER_FIRST_NAME;
    private static String MOTHER_MIDDLE_NAME;
    private static String MOTHER_LAST_NAME;
    private static String ADDRESS_LINE1;
    private static String ADDRESS_LINE2;
    private static String ADDRESS_LINE3;
    private static String STATE;
    private static String COUNTRY;
    private static String EMAIL;
    private static String MOBILE;
    private static String printFile;

    private static final String sourcePayEezzFile = Util.getDirectoryPath() + "/CKYC-Individual-Fillable.pdf";
    private static String destinationFile;
    private static final String sourceExcelFile = Util.getDirectoryPath() + "/fill-can-registration.xlsm";

    public static void main(String[] args) throws Exception {

        Util util = Util.getUtilObject(sourceExcelFile, 4);
        if (util == null)
            return;
        fillFromExcel(util);
        destinationFile = Util.getDestinationDirectoryPath() + "/" + FIRST_NAME + "_" + PAN + "_" + "KYC_Old" + ".pdf";
        editPdfDocument();

    }

    private static void fillFromExcel(Util util) throws Exception {
        PAN = util.getCellValue(0);
        PREFIX = util.getCellValue(1);
        FIRST_NAME = util.getCellValue(2);
        MIDDLE_NAME = util.getCellValue(3);
        SURNAME = util.getCellValue(4);
        FATHER_FIRST_NAME = util.getCellValue(5);
        FATHER_MIDDLE_NAME = util.getCellValue(6);
        FATHER_LAST_NAME = util.getCellValue(7);
        MOTHER_FIRST_NAME = util.getCellValue(8);
        MOTHER_MIDDLE_NAME = util.getCellValue(9);
        MOTHER_LAST_NAME = util.getCellValue(10);
        ADDRESS_LINE1 = util.getCellValue(14);
        ADDRESS_LINE2 = util.getCellValue(15);
        ADDRESS_LINE3 = util.getCellValue(16);
        STATE = util.getCellValue(18);
        COUNTRY = util.getCellValue(20);
        MOBILE = util.getCellValue(21);
        EMAIL = util.getCellValue(22);

    }

    private static void editPdfDocument() throws Exception {
        Date date = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        String strDate = sdf.format(date);


        PdfReader reader = null;
        PdfStamper stamper = null;
        try {
            System.out.println("Reading from following file :-" + sourcePayEezzFile + " \n and excel file :- " + sourceExcelFile);
            reader = new PdfReader(sourcePayEezzFile);
            stamper = new PdfStamper(reader, new FileOutputStream(destinationFile));
            AcroFields form = stamper.getAcroFields();
            Map<String, AcroFields.Item> fields = form.getFields();
            System.out.println(fields);
            form.setField("PAN", PAN);
            form.setField("Name same as ID proof", PREFIX);
            form.setField("First Name", FIRST_NAME);
            form.setField("Middle Name", MIDDLE_NAME);
            form.setField("Last Name", SURNAME);
            form.setField("Father Spouse Name", "MR");
            form.setField("undefined_5", FATHER_FIRST_NAME);
            form.setField("undefined_6", FATHER_MIDDLE_NAME);
            form.setField("undefined_7", FATHER_LAST_NAME);
            form.setField("undefined_8", "MRS");
            form.setField("undefined_9", MOTHER_FIRST_NAME);
            form.setField("undefined_10", MOTHER_MIDDLE_NAME);
            form.setField("undefined_11", MOTHER_LAST_NAME);
            form.setField("Line 1", ADDRESS_LINE1);
            form.setField("Line 2", ADDRESS_LINE2);
            form.setField("Line 3", ADDRESS_LINE3);
            form.setField("StateUT", STATE);
            form.setField("Country", COUNTRY);
            form.setField("Email ID", EMAIL);
            form.setField("undefined_36", MOBILE);

            stamper.setFormFlattening(true);

        } catch (Exception dex) {
            if (stamper != null)
                stamper.close();
            if (reader != null)
                reader.close();
            throw dex;
        }
        stamper.close();
        reader.close();
        System.out.println("File saved at following location :- " + destinationFile);
    }

}
