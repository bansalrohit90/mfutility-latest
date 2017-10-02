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
import java.util.Calendar;
import java.util.Date;
import java.util.Map;

/**
 * Example to show filling form fields.
 */
public final class FillKYC2RegistrationForm {

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
    private static String DOB_DATE;
    private static String DOB_MONTH;
    private static String DOB_YEAR;
    private static String ADDRESS_LINE1;
    private static String ADDRESS_LINE2;
    private static String ADDRESS_LINE3;
    private static String PIN_CODE;
    private static String STATE;
    private static String STATE_CODE;
    private static String COUNTRY;
    private static String EMAIL;
    private static String MOBILE;
    private static String PROOF_TYPE;
    private static String PROOF_VALUE;
    private static String EXPIRY_DATE;
    private static String EXPIRY_MONTH;
    private static String EXPIRY_YEAR;
    private static String APPLICATION_YEAR;
    private static String APPLICATION_MONTH;
    private static String APPLICATION_DATE;
    private static String printFile;

    private static final String sourcePayEezzFile = Util.getDirectoryPath() + "\\ckyc-application-form-individual.pdf";
    private static String destinationFile;
    private static final String sourceExcelFile = Util.getDirectoryPath() + "\\fill-can-registration.xlsm";

    public static void main(String[] args) throws Exception {

        Util util = Util.getUtilObject(sourceExcelFile, 4);
        if (util == null)
            return;
        fillFromExcel(util);
        destinationFile = Util.getDestinationDirectoryPath() + "\\" + FIRST_NAME + "_" + PAN + "_" + "KYC" + ".pdf";
        editPdfDocument();
        /*if (printFile.equalsIgnoreCase("yes")) {
//            Util.printPdfOutput2(destinationFile, FIRST_NAME + "KYC_REG");
        }*/
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
        DOB_DATE = util.getCellValue(11);
        DOB_MONTH = util.getCellValue(12);
        DOB_YEAR = util.getCellValue(13);
        ADDRESS_LINE1 = util.getCellValue(14);
        ADDRESS_LINE2 = util.getCellValue(15);
        ADDRESS_LINE3 = util.getCellValue(16);
        PIN_CODE = util.getCellValue(17);
        STATE = util.getCellValue(18);
        STATE_CODE = util.getCellValue(19);
        COUNTRY = util.getCellValue(20);
        MOBILE = util.getCellValue(21);
        EMAIL = util.getCellValue(22);
        PROOF_TYPE = util.getCellValue(23);
        PROOF_VALUE = util.getCellValue(24);
        EXPIRY_DATE = util.getCellValue(25);
        EXPIRY_MONTH = util.getCellValue(26);
        EXPIRY_YEAR = util.getCellValue(27);
        APPLICATION_YEAR = String.valueOf(Calendar.getInstance().get(Calendar.YEAR));
        int month = Calendar.getInstance().get(Calendar.MONTH);
        APPLICATION_MONTH = (month < 9) ? "0" + String.valueOf(month + 1) : String.valueOf(month + 1);
        int date = Calendar.getInstance().get(Calendar.DATE);
        APPLICATION_DATE = (date <= 9) ? "0" + String.valueOf(date) : String.valueOf(date);
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
            form.setField("Father First Name", FATHER_FIRST_NAME);
            form.setField("Date", DOB_DATE);
            form.setField("Month", DOB_MONTH);
            form.setField("Year", DOB_YEAR);
            form.setField("Father Middle Name", FATHER_MIDDLE_NAME);
            form.setField("Group9", "IN");
            form.setField("Post code", PIN_CODE);
            form.setField("Mother Last Name", MOTHER_LAST_NAME);
            form.setField("Code6", STATE_CODE);
            form.setField("Zip/Post code 1", "IN");
            form.setField("Text6", PAN);
            form.setField("Mother First Name", MOTHER_FIRST_NAME);
            if (PROOF_TYPE.equals("PASSPORT")) {
                form.setField("Passport Number 1", PROOF_VALUE);
                form.setField("PED Date1", EXPIRY_DATE);
                form.setField("PED Month1", EXPIRY_MONTH);
                form.setField("PED Year1", EXPIRY_YEAR);
            } else if (PROOF_TYPE.equals("VOTER_ID")) {
                form.setField("Voter ID1", PROOF_VALUE);
            } else if (PROOF_TYPE.equals("DRIVING LICENSE")) {
                form.setField("Driving Licence1", PROOF_VALUE);
                form.setField("DLE Date1", EXPIRY_DATE);
                form.setField("DLE Month1", EXPIRY_MONTH);
                form.setField("DLE Year1", EXPIRY_YEAR);
            } else if (PROOF_TYPE.equals("AADHAR")) {
                form.setField("Aadhaar Card1", PROOF_VALUE);
            }

            form.setField("State", STATE);
            form.setField("Mother Middle Name", MOTHER_MIDDLE_NAME);
            form.setField("Mobile Number", MOBILE);
            form.setField("Address Line 1", ADDRESS_LINE1);
            form.setField("Address Line 2", ADDRESS_LINE2);
            form.setField("Address Line 3", ADDRESS_LINE3);
            form.setField("Email ID", EMAIL);
            form.setField("Country", "INDIA");
            form.setField("Prefix 3", "MR");
            form.setField("Prefix 4", "MRS");
            form.setField("Prefix 1", PREFIX);
            form.setField("First Name", FIRST_NAME);
            form.setField("Middle Name", MIDDLE_NAME);
            form.setField("Last Name", SURNAME);
            form.setField("Father Last Name", FATHER_LAST_NAME);
            form.setField("Place2", "GANGANAGAR ");
            form.setField("Application Year", APPLICATION_YEAR);
            form.setField("Application Month", APPLICATION_MONTH);
            form.setField("Application Date", APPLICATION_DATE);
            form.setField("Year5", APPLICATION_YEAR);
            form.setField("Year6", APPLICATION_YEAR);
            form.setField("Month5", APPLICATION_MONTH);
            form.setField("Month6", APPLICATION_MONTH);
            form.setField("Date6", APPLICATION_DATE);
            form.setField("Date5", APPLICATION_DATE);

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