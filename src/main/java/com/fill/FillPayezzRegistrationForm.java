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
import java.util.Date;
import java.util.Map;

import static com.fill.com.fill.util.Util.CITY;
import static com.fill.com.fill.util.Util.getAmountInWords;

/**
 * Example to show filling form fields.
 */
public final class FillPayezzRegistrationForm {

    private static String CAN;
    private static String AMOUNT;
    private static String PAN;
    private static String NAME;
    private static String MOBILE;
    private static String ACCOUNT_NUMBER;
    private static String MICR;
    private static String IFSC;
    private static String BANK_NAME;
    private static String EMAIL;
    private static String printFile;

    private static final String sourcePayEezzFile = Util.getDirectoryPath() + "\\PayEezz-Mandate-Fillable.pdf";
    private static String destinationFile;
    private static final String sourceExcelFile = Util.getDirectoryPath() + "\\fill-can-registration.xlsm";

    public static void main(String[] args) throws Exception {
        Util util = Util.getUtilObject(sourceExcelFile, 1);
        if (util == null)
            return;
        fillFromExcel(util);
        destinationFile = Util.getDestinationDirectoryPath() + "\\" + NAME + "_" + PAN + "_" + "PAYEZZ" + ".pdf";
        editPdfDocument();
        /*if (printFile.equalsIgnoreCase("yes")) {
//            Util.printPdfOutput2(destinationFile,NAME+"PAYEZZ_REG");
        }*/
    }

    private static void fillFromExcel(Util util) throws Exception {
        CAN = util.getCellValue(0);
        PAN = util.getCellValue(1);
        NAME = util.getCellValue(2);
        MOBILE = util.getCellValue(3);
        EMAIL = util.getCellValue(4);
        ACCOUNT_NUMBER = util.getCellValue(5);
        MICR = util.getCellValue(6);
        IFSC = util.getCellValue(7);
        BANK_NAME = util.getCellValue(8);
        AMOUNT = util.getCellValue(9);
        printFile = util.getCellValue(10);
    }

    private static void editPdfDocument() throws Exception {
        Date date = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        String strDate = sdf.format(date);


        PdfReader reader = null;
        PdfStamper stamper = null;
        try {
            reader = new PdfReader(sourcePayEezzFile);
            System.out.println("Reading from following file :-" + sourcePayEezzFile + " \n and excel file :- " + sourceExcelFile);
            stamper = new PdfStamper(reader, new FileOutputStream(destinationFile));
            AcroFields form = stamper.getAcroFields();
            Map<String, AcroFields.Item> fields = form.getFields();
            form.setField("ARN_CODE", "ARN-10911");
            form.setField("EUIN_CODE", "E036366");
            form.setField("CANID", CAN);
            form.setField("CANID1", CAN);
            form.setField("PANID", PAN);
            form.setField("NAME", NAME);
            form.setField("BK_NAME1", NAME);
            form.setField("PHONENO", MOBILE);
            form.setField("EMAILID", EMAIL);
            form.setField("BANKACNO", ACCOUNT_NUMBER);
            form.setField("BANKACNO1", ACCOUNT_NUMBER);
            form.setField("MICR", MICR);
            form.setField("IFSC", IFSC);
            form.setField("BANKNAME", BANK_NAME);
            form.setField("AMT_FIG", AMOUNT);
            form.setField("AMOUNT_WORD", getAmountInWords(AMOUNT));
            form.setField("PRD_CHK", "Yes");
            form.setField("FROM_DATE", strDate);
            form.setField("DATE1", strDate);
            form.setField("DATE", strDate);
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