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

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

import static com.fill.com.fill.util.Util.CITY;

public final class FillCanRegistrationForm {

    private static String PAN;
    private static String NAME;
    private static String DOB;
    private static String AADHAR;
    private static String MOBILE;
    private static String ACC1_NUMBER;
    private static String ACC1_MICR;
    private static String ACC1_IFSC;
    private static String ACC1_BANK_NAME;
    private static String ACC1_BRANCH;
    private static String ACC1_CITY;
    private static String ACC2_NUMBER;
    private static String ACC2_MICR;
    private static String ACC2_IFSC;
    private static String ACC2_BANK_NAME;
    private static String ACC2_BRANCH;
    private static String ACC2_CITY;
    private static String EMAIL;
    private static String NOMINEE;
    private static String NOMINEE_RELATION;
    private static String PLACE_OF_BIRTH;
    private static String printFile;

    private static final String sourceCanEezzFile = Util.getDirectoryPath() + "\\CANEezz-Fillable.pdf";
    private static String destinationFile;
    private static final String sourceExcelFile = Util.getDirectoryPath() + "\\fill-can-registration.xlsm";

    public static void main(String[] args) throws Exception {
        Util util = Util.getUtilObject(sourceExcelFile, 0);
        if (util == null)
            return;
        fillFromExcel(util);
        destinationFile = Util.getDestinationDirectoryPath() + "\\" + NAME + "_" + PAN + "_" + "CAN_REGISTRATION" + ".pdf";
        editPdfDocument();
        /*if (printFile.equalsIgnoreCase("yes")) {
//            Util.printPdfOutput2(destinationFile,NAME+"CAN_REG");
        }*/
    }

    private static void fillFromExcel(Util util) throws Exception {

        PAN = util.getCellValue(0);
        NAME = util.getCellValue(1);
        DOB = util.getDateCellValue(2);
        AADHAR = util.getCellValue(3);
        AADHAR = Util.maskString(AADHAR,0,8,'*');
        MOBILE = util.getCellValue(4);
        EMAIL = util.getCellValue(5);
        ACC1_NUMBER = util.getCellValue(6);
        ACC1_MICR = util.getCellValue(7);
        ACC1_IFSC = util.getCellValue(8);
        ACC1_BANK_NAME = util.getCellValue(9);
        ACC1_BRANCH = util.getCellValue(10);
        ACC1_CITY = util.getCellValue(11);
        ACC2_NUMBER = util.getCellValue(12);
        ACC2_MICR = util.getCellValue(13);
        ACC2_IFSC = util.getCellValue(14);
        ACC2_BANK_NAME = util.getCellValue(15);
        ACC2_BRANCH = util.getCellValue(16);
        ACC2_CITY = util.getCellValue(17);
        NOMINEE = util.getCellValue(18);
        NOMINEE_RELATION = util.getCellValue(19);
        PLACE_OF_BIRTH = util.getCellValue(20);
        printFile = util.getCellValue(21);
    }

    private static void editPdfDocument() throws Exception {
        Date date = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        String strDate = sdf.format(date);


        PdfReader reader = null;
        PdfStamper stamper = null;
        ByteArrayOutputStream baosPDF = new ByteArrayOutputStream();
        try {
            System.out.println("Reading from following file :-" + sourceCanEezzFile + " \n and excel file :- " + sourceExcelFile);
            reader = new PdfReader(sourceCanEezzFile);
            stamper = new PdfStamper(reader, new FileOutputStream(destinationFile));
            AcroFields form = stamper.getAcroFields();
            Map<String, AcroFields.Item> fields = form.getFields();
            System.out.println(fields);
            form.setField("ARN", "ARN-10911");
            form.setField("EUIN", "E036366");
            form.setField("MOH1", "Yes");
            form.setField("CATG1", "Yes");
            form.setField("RI_ST1", "Yes");
            form.setField("PAN1", PAN);
            form.setField("NAME1", NAME);
            form.setField("DOB1", DOB);
            form.setField("AADHAR1", AADHAR);
            form.setField("MOB1", MOBILE);
            form.setField("EMAIL1", EMAIL);
            form.setField("BK_AC1", ACC1_NUMBER);
            form.setField("BK_MICR1", ACC1_MICR);
            form.setField("BK_IFSC1", ACC1_IFSC);
            form.setField("BK_NAME1", ACC1_BANK_NAME);
            form.setField("BK_BR1", ACC1_BRANCH);
            form.setField("BK_CITY1", ACC1_CITY);
            form.setField("BK_AC2", ACC2_NUMBER);
            form.setField("BK_MICR2", ACC2_MICR);
            form.setField("BK_IFSC2", ACC2_IFSC);
            form.setField("BK_NAME2", ACC2_BANK_NAME);
            form.setField("BK_BR2", ACC2_BRANCH);
            form.setField("BK_CITY2", ACC2_CITY);
            form.setField("NOMINEE1", NOMINEE);
            form.setField("NOM_REL1", NOMINEE_RELATION);
            form.setField("NOM1_PER", "100%");
            form.setField("AC1_TYPE1", "Yes");
            form.setField("AC2_TYPE1", "Yes");
            form.setField("Nom_Reg_Y", "Yes");
            form.setField("NAME1_POB", PLACE_OF_BIRTH);
            form.setField("NAME1_CON", "INDIA");
            form.setField("NAME1_COC", "INDIA");
            form.setField("NAME1_COB", "INDIA");
            form.setField("IND1", "Yes");
            form.setField("DATE_SUB", strDate);
            form.setField("SUB_PLACE", CITY);
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