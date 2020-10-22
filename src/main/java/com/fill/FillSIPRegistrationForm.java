/*
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
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
import static com.fill.com.fill.util.Util.getAmountInWords;

/**
 * Example to show filling form fields.
 */
public final class FillSIPRegistrationForm {

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
    private static String SIP1_AMC;
    private static String SIP1_FOLIO;
    private static String SIP1_SCHEME;
    private static String SIP1_DATE;
    private static String SIP1_START_MONTH;
    private static String SIP1_START_YEAR;
    private static String SIP1_AMOUNT;
    private static String SIP2_AMC;
    private static String SIP2_FOLIO;
    private static String SIP2_SCHEME;
    private static String SIP2_DATE;
    private static String SIP2_START_MONTH;
    private static String SIP2_START_YEAR;
    private static String SIP2_AMOUNT;
    private static String SIP3_AMC;
    private static String SIP3_FOLIO;
    private static String SIP3_SCHEME;
    private static String SIP3_DATE;
    private static String SIP3_START_MONTH;
    private static String SIP3_START_YEAR;
    private static String SIP3_AMOUNT;
    private static String SIP4_AMC;
    private static String SIP4_FOLIO;
    private static String SIP4_SCHEME;
    private static String SIP4_DATE;
    private static String SIP4_START_MONTH;
    private static String SIP4_START_YEAR;
    private static String SIP4_AMOUNT;
    private static String SIP5_AMC;
    private static String SIP5_FOLIO;
    private static String SIP5_SCHEME;
    private static String SIP5_DATE;
    private static String SIP5_START_MONTH;
    private static String SIP5_START_YEAR;
    private static String SIP5_AMOUNT;
    private static String printFile;


    //Fillable1 - checked growth & monthly; Fillable - unchecked growth & monthly
    private static final String sourcePayEezzFile = Util.getDirectoryPath() + "/CTF-SIP-Fillable1.pdf";
    private static String destinationFile;
    private static final String sourceExcelFile = Util.getDirectoryPath() + "/fill-can-registration1.xlsm";

    public static void main(String[] args) throws Exception {

        Util util = Util.getUtilObject(sourceExcelFile, 2);
        if (util == null)
            return;
        fillFromExcel(util);
        destinationFile = Util.getDestinationDirectoryPath() + "/" + NAME + "_" + PAN + "_" + "SIP" + ".pdf";
        editPdfDocument();
        /*if (printFile.equalsIgnoreCase("yes")) {
//            Util.printPdfOutput2(destinationFile,NAME+"SIP_REG");
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
        SIP1_AMC = util.getCellValue(12);
        SIP1_FOLIO = util.getCellValue(13);
        SIP1_SCHEME = util.getCellValue(14);
        SIP1_DATE = util.getCellValue(15);
        SIP1_START_MONTH = util.getCellValue(16);
        SIP1_START_YEAR = util.getCellValue(17);
        SIP1_AMOUNT = util.getCellValue(18);
        SIP2_AMC = util.getCellValue(19);
        SIP2_FOLIO = util.getCellValue(20);
        SIP2_SCHEME = util.getCellValue(21);
        SIP2_DATE = util.getCellValue(22);
        SIP2_START_MONTH = util.getCellValue(23);
        SIP2_START_YEAR = util.getCellValue(24);
        SIP2_AMOUNT = util.getCellValue(25);
        SIP3_AMC = util.getCellValue(26);
        SIP3_FOLIO = util.getCellValue(27);
        SIP3_SCHEME = util.getCellValue(28);
        SIP3_DATE = util.getCellValue(29);
        SIP3_START_MONTH = util.getCellValue(30);
        SIP3_START_YEAR = util.getCellValue(31);
        SIP3_AMOUNT = util.getCellValue(32);
        SIP4_AMC = util.getCellValue(33);
        SIP4_FOLIO = util.getCellValue(34);
        SIP4_SCHEME = util.getCellValue(35);
        SIP4_DATE = util.getCellValue(36);
        SIP4_START_MONTH = util.getCellValue(37);
        SIP4_START_YEAR = util.getCellValue(38);
        SIP4_AMOUNT = util.getCellValue(39);
        SIP5_AMC = util.getCellValue(40);
        SIP5_FOLIO = util.getCellValue(41);
        SIP5_SCHEME = util.getCellValue(42);
        SIP5_DATE = util.getCellValue(43);
        SIP5_START_MONTH = util.getCellValue(44);
        SIP5_START_YEAR = util.getCellValue(45);
        SIP5_AMOUNT = util.getCellValue(46);
        AMOUNT = Util.getTotalAmount(Arrays.asList(SIP1_AMOUNT, SIP2_AMOUNT, SIP3_AMOUNT, SIP4_AMOUNT, SIP5_AMOUNT));
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
            form.setField("CANNO", CAN);
            form.setField("PAN_PEKRN", PAN);
            form.setField("NAME", NAME);
            form.setField("ARN", "10911");
            form.setField("ARN NAME", "SANTOSH BANSAL");
            form.setField("EUIN", "E036366");
            form.setField("PAY_REFNO", CHQ_NO);
            form.setField("PAY_DD", CHQ_DT);
            form.setField("PAY_MM", CHQ_MONTH);
            form.setField("PAY_YEAR", CHQ_YEAR);
            form.setField("NET_AMT", AMOUNT);
            form.setField("DD_CHRG", "--");
            form.setField("TOT_AMT", AMOUNT);
            form.setField("AMT_IN_WORDS", getAmountInWords(AMOUNT).toUpperCase());
            form.setField("BK_AC_NO", ACCOUNT_NUMBER);
            form.setField("BK_MICR", MICR);
            form.setField("BK_IFSC", IFSC);
            form.setField("BANK_NAME", BANK_NAME);
            form.setField("BR_NAME", BRANCH_NAME);

            form.setField("AMC1", SIP1_AMC);
            form.setField("FOLIO1", SIP1_FOLIO);
            form.setField("SCHEME_AMC1", SIP1_SCHEME);
            form.setField("FREQ1_DD", SIP1_DATE);
            form.setField("FREQ1_START_MM", SIP1_START_MONTH);
            form.setField("FREQ1_START_YY", SIP1_START_YEAR);
            form.setField("AMT_AMC1", SIP1_AMOUNT);
            if (SIP1_AMC != "") form.setField("FREQ1_D", "Yes");

            form.setField("AMC2", SIP2_AMC);
            form.setField("FOLIO2", SIP2_FOLIO);
            form.setField("SCHEME_AMC2", SIP2_SCHEME);
            form.setField("FREQ2_DD", SIP2_DATE);
            form.setField("FREQ2_START_MM", SIP2_START_MONTH);
            form.setField("FREQ2_START_YY", SIP2_START_YEAR);
            form.setField("AMT_AMC2", SIP2_AMOUNT);
            if (SIP2_AMC != "") form.setField("FREQ2_D", "Yes");

            form.setField("AMC3", SIP3_AMC);
            form.setField("FOLIO3", SIP3_FOLIO);
            form.setField("SCHEME_AMC3", SIP3_SCHEME);
            form.setField("FREQ3_DD", SIP3_DATE);
            form.setField("FREQ3_START_MM", SIP3_START_MONTH);
            form.setField("FREQ3_START_YEAR", SIP3_START_YEAR);
            form.setField("AMT_AMC3", SIP3_AMOUNT);
            if (SIP3_AMC != "") form.setField("FREQ3_D", "Yes");

            form.setField("AMC4", SIP4_AMC);
            form.setField("FOLIO4", SIP4_FOLIO);
            form.setField("SCHEME_AMC4", SIP4_SCHEME);
            form.setField("FREQ4_DD", SIP4_DATE);
            form.setField("FREQ4_START_MM", SIP4_START_MONTH);
            form.setField("FREQ4_START_YEAR", SIP4_START_YEAR);
            form.setField("AMT_AMC4", SIP4_AMOUNT);
            if (SIP4_AMC != "") form.setField("FREQ4_D", "Yes");

            form.setField("AMC5", SIP5_AMC);
            form.setField("FOLIO5", SIP5_FOLIO);
            form.setField("SCHEME_AMC5", SIP5_SCHEME);
            form.setField("FREQ5_DD", SIP5_DATE);
            form.setField("FREQ5_START_MM", SIP5_START_MONTH);
            form.setField("FREQ5_START_YEAR", SIP5_START_YEAR);
            form.setField("AMT_AMC5", SIP5_AMOUNT);
            if (SIP5_AMC != "") form.setField("FREQ5_D", "Yes");

            form.setField("TXN_DATE", strDate);
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
