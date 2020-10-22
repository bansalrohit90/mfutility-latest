package com.fill;

import com.fill.com.fill.util.Util;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.Map;

import static com.fill.com.fill.util.Util.CITY;

/**
 * Created by Rohit.Bansal on 15-04-2017.
 */
public class FillPurchaseTransactionForm {

    private static String CAN;
    private static String AMOUNT;
    private static String PAN;
    private static String NAME;
    private static String ACCOUNT_NUMBER;
    private static String MICR;
    private static String IFSC;
    private static String BANK_NAME;
    private static String BRANCH_NAME;
    private static String CHQ_NO;
    private static String CHQ_DT;
    private static String CHQ_MONTH;
    private static String CHQ_YEAR;
    private static String LUMPSUM1_AMC;
    private static String LUMPSUM1_FOLIO;
    private static String LUMPSUM1_SCHEME;
    private static String LUMPSUM1_AMOUNT;
    private static String LUMPSUM2_AMC;
    private static String LUMPSUM2_FOLIO;
    private static String LUMPSUM2_SCHEME;
    private static String LUMPSUM2_AMOUNT;
    private static String LUMPSUM3_AMC;
    private static String LUMPSUM3_FOLIO;
    private static String LUMPSUM3_SCHEME;
    private static String LUMPSUM3_AMOUNT;
    private static String LUMPSUM4_AMC;
    private static String LUMPSUM4_FOLIO;
    private static String LUMPSUM4_SCHEME;
    private static String LUMPSUM4_AMOUNT;
    private static String LUMPSUM5_AMC;
    private static String LUMPSUM5_FOLIO;
    private static String LUMPSUM5_SCHEME;
    private static String LUMPSUM5_AMOUNT;
    private static String printFile;

    private static final String sourcePayEezzFile = Util.getDirectoryPath() + "/TF-Purchase-Fillable.pdf";
    private static String destinationFile;
    private static final String sourceExcelFile = Util.getDirectoryPath() + "/fill-can-registration1.xlsm";

    public static void main(String[] args) throws Exception {

        Util util = Util.getUtilObject(sourceExcelFile, 3);
        if (util == null)
            return;
        fillFromExcel(util);
        destinationFile = Util.getDestinationDirectoryPath() + "/" + NAME + "_" + PAN + "_" + "PurchaseTransaction" + ".pdf";
        editPdfDocument();
        /*if (printFile.equalsIgnoreCase("yes")) {
//            Util.printPdfOutput2(destinationFile, NAME + "LUMPSUM_REG");
        }*/
    }

    private static void fillFromExcel(Util util) throws Exception {
        CAN = util.getCellValue(0);
        PAN = util.getCellValue(1);
        NAME = util.getCellValue(2);
        ACCOUNT_NUMBER = util.getCellValue(3);
        MICR = util.getCellValue(4);
        IFSC = util.getCellValue(5);
        BANK_NAME = util.getCellValue(6);
        BRANCH_NAME = util.getCellValue(7);
        CHQ_NO = util.getCellValue(8);
        CHQ_DT = util.getCellValue(9);
        CHQ_MONTH = util.getCellValue(10);
        CHQ_YEAR = util.getCellValue(11);
        LUMPSUM1_AMC = util.getCellValue(12);
        LUMPSUM1_FOLIO = util.getCellValue(13);
        LUMPSUM1_SCHEME = util.getCellValue(14);
        LUMPSUM1_AMOUNT = util.getCellValue(15);
        LUMPSUM2_AMC = util.getCellValue(16);
        LUMPSUM2_FOLIO = util.getCellValue(17);
        LUMPSUM2_SCHEME = util.getCellValue(18);
        LUMPSUM2_AMOUNT = util.getCellValue(19);
        LUMPSUM3_AMC = util.getCellValue(20);
        LUMPSUM3_FOLIO = util.getCellValue(21);
        LUMPSUM3_SCHEME = util.getCellValue(22);
        LUMPSUM3_AMOUNT = util.getCellValue(23);
        LUMPSUM4_AMC = util.getCellValue(24);
        LUMPSUM4_FOLIO = util.getCellValue(25);
        LUMPSUM4_SCHEME = util.getCellValue(26);
        LUMPSUM4_AMOUNT = util.getCellValue(27);
        LUMPSUM5_AMC = util.getCellValue(28);
        LUMPSUM5_FOLIO = util.getCellValue(29);
        LUMPSUM5_SCHEME = util.getCellValue(30);
        LUMPSUM5_AMOUNT = util.getCellValue(31);
        AMOUNT = Util.getTotalAmount(Arrays.asList(LUMPSUM1_AMOUNT, LUMPSUM2_AMOUNT, LUMPSUM3_AMOUNT, LUMPSUM4_AMOUNT, LUMPSUM5_AMOUNT));
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
            form.setField("CANNO", CAN);
            form.setField("PAN_PEKRN", PAN);
            form.setField("NAME", NAME);
            form.setField("ARNNO", "10911");
            form.setField("ARNNAME", "SANTOSH BANSAL");
            form.setField("EUIN", "E036366");
            form.setField("PAY_MODE1", "Yes");
            form.setField("PAY_MODE12", "Yes");
            form.setField("PAY_REF_NO", CHQ_NO);
            form.setField("PAY_DD", CHQ_DT);
            form.setField("PAY_MM", CHQ_MONTH);
            form.setField("PAY_YEAR", CHQ_YEAR);
            form.setField("INSTR_AMT", AMOUNT);
            form.setField("DD_CHRG", "--");
            form.setField("TOTAL_AMT", AMOUNT);
            form.setField("AMT_IN_WORDS", Util.getAmountInWords(AMOUNT));
            form.setField("BK_ACC_NO", ACCOUNT_NUMBER);
            form.setField("MICR", MICR);
            form.setField("IFSC", IFSC);
            form.setField("AC_TYPE1", "Yes");
            form.setField("BANK_NAME", BANK_NAME);
            form.setField("BR_NAME", BRANCH_NAME);

            form.setField("AMC1", LUMPSUM1_AMC);
            form.setField("FOLIO1", LUMPSUM1_FOLIO);
            form.setField("SCHEME_PLAN1", LUMPSUM1_SCHEME);
            form.setField("INVST_AMT1", LUMPSUM1_AMOUNT);
            form.setField("INVST_AMT1_IN_WORDS", Util.getAmountInWords(LUMPSUM1_AMOUNT));

            form.setField("AMC2", LUMPSUM2_AMC);
            form.setField("FOLIO2", LUMPSUM2_FOLIO);
            form.setField("SCHEME_PLAN2", LUMPSUM2_SCHEME);
            form.setField("INVST_AMT2", LUMPSUM2_AMOUNT);
            form.setField("INVST_AMT2_IN_WORDS", Util.getAmountInWords(LUMPSUM2_AMOUNT));

            form.setField("AMC3", LUMPSUM3_AMC);
            form.setField("FOLIO3", LUMPSUM3_FOLIO);
            form.setField("SCHEME_PLAN3", LUMPSUM3_SCHEME);
            form.setField("INVST_AMT3", LUMPSUM3_AMOUNT);
            form.setField("INVST_AMT3_IN_WORDS", Util.getAmountInWords(LUMPSUM3_AMOUNT));

            form.setField("AMC4", LUMPSUM4_AMC);
            form.setField("FOLIO4", LUMPSUM4_FOLIO);
            form.setField("SCHEME_PLAN4", LUMPSUM4_SCHEME);
            form.setField("INVST_AMT4", LUMPSUM4_AMOUNT);
            form.setField("INVST_AMT4_IN_WORDS", Util.getAmountInWords(LUMPSUM4_AMOUNT));

            form.setField("AMC5", LUMPSUM5_AMC);
            form.setField("FOLIO5", LUMPSUM5_FOLIO);
            form.setField("SCHEME_PLAN5", LUMPSUM5_SCHEME);
            form.setField("INVST_AMT5", LUMPSUM5_AMOUNT);
            form.setField("INVST_AMT5_IN_WORDS", Util.getAmountInWords(LUMPSUM5_AMOUNT));

            form.setField("SUB_DATE", strDate);
            form.setField("PLACE", CITY);

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
