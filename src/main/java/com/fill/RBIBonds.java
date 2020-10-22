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
import org.apache.poi.ss.usermodel.Row;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;

/**
 * Example to show filling form fields.
 */
public final class RBIBonds {

    private static String subBrokerCode;
    private static String applicantName;
    private static String applicantDOB;
    private static String applicantGender;
    private static String applicantPAN;
    private static String applicantMotherName;
    private static String applicantAdd1;
    private static String applicantAdd2;
    private static String applicantCity;
    private static String applicantMobile;
    private static String applicantEmail;
    private static String chqNumber;
    private static String chqDate;
    private static String chqBankBranch;
    private static String amount;
    private static String nameinBank;
    private static String bankName;
    private static String branch;
    private static String accountNumber;
    private static String iFSC;
    private static String mICR;
    private static String place;
    private static String nomineeName;
    private static String nomineeDOB;
    private static String nomineeRelation;

    private static Iterator<Row> rowIterator;
    private static Row firstRow;

    //Fillable1 - checked growth & monthly; Fillable - unchecked growth & monthly
    private static final String sourceBankMandateFile = Util.getDirectoryPath() + "/RBI_BondFillable.pdf";
    private static String destinationFile;
    private static final String sourceExcelFile = Util.getDirectoryPath() + "/fill-can-registration1.xlsm";

    public static void main(String[] args) throws Exception {

        Iterator<Row> iterator = Util.getUtilIteratorObject(sourceExcelFile, 6);
        fillFromExcel(iterator);
        destinationFile = Util.getDestinationDirectoryPath() + "/" + applicantName + ".pdf";
        editPdfDocument();
    }

    private static void fillFromExcel(Iterator<Row> iterator) {
        if (iterator.hasNext()) {
            firstRow = iterator.next();
        }
        rowIterator = iterator;
        subBrokerCode = Util.getCellValue(firstRow, 1);
        firstRow = rowIterator.next();
        applicantName = Util.getCellValue(firstRow, 1);
        firstRow = rowIterator.next();
    }

    private static void editPdfDocument() throws Exception {
        Date date = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        String strDate = sdf.format(date);


        PdfReader reader = null;
        PdfStamper stamper = null;
        try {
            System.out.println("Reading from following file :-" + sourceBankMandateFile + " \n and excel file :- " + sourceExcelFile);
            reader = new PdfReader(sourceBankMandateFile);
            stamper = new PdfStamper(reader, new FileOutputStream(destinationFile));
            AcroFields form = stamper.getAcroFields();
            Map<String, AcroFields.Item> fields = form.getFields();
            System.out.println(fields);
            Row row = firstRow;


            applicantDOB = (Util.getCellValue(row, 1));
            row = rowIterator.next();
            applicantPAN = Util.getCellValue(row, 1);
            row = rowIterator.next();
            applicantMotherName = Util.getCellValue(row, 1);
            row = rowIterator.next();
            applicantAdd1 = Util.getCellValue(row, 1);
            row = rowIterator.next();
            applicantAdd2 = Util.getCellValue(row, 1);
            row = rowIterator.next();
            applicantCity = Util.getCellValue(row, 1);
            row = rowIterator.next();
            applicantMobile = Util.getCellValue(row, 1);
            row = rowIterator.next();
            applicantEmail = Util.getCellValue(row, 1);
            row = rowIterator.next();
            chqNumber = Util.getCellValue(row, 1);
            row = rowIterator.next();
            chqDate = Util.getCellValue(row, 1);
            row = rowIterator.next();
            chqBankBranch = Util.getCellValue(row, 1);
            row = rowIterator.next();
            amount = Util.getCellValue(row, 1);
            row = rowIterator.next();
            nameinBank = Util.getCellValue(row, 1);
            row = rowIterator.next();
            bankName = Util.getCellValue(row, 1);
            row = rowIterator.next();
            branch = Util.getCellValue(row, 1);
            row = rowIterator.next();
            accountNumber = Util.getCellValue(row, 1);
            row = rowIterator.next();
            iFSC = Util.getCellValue(row, 1);
            row = rowIterator.next();
            mICR = Util.getCellValue(row, 1);
            row = rowIterator.next();
            place = Util.getCellValue(row, 1);
            row = rowIterator.next();
            nomineeName = Util.getCellValue(row, 1);
            row = rowIterator.next();
            nomineeDOB = (Util.getCellValue(row, 1));
            row = rowIterator.next();
            nomineeRelation = (Util.getCellValue(row, 1));
            row = rowIterator.next();
            String nomineeStatus = (Util.getCellValue(row, 1));
            row = rowIterator.next();
            String witness1Name = Util.getCellValue(row, 1);
            row = rowIterator.next();
            String witness2Name = Util.getCellValue(row, 1);
            row = rowIterator.next();
            String witness1Add = Util.getCellValue(row, 1);
            row = rowIterator.next();
            String witness2Add = Util.getCellValue(row, 1);
//            row = rowIterator.next();

            form.setField("SubBroker", subBrokerCode);
            form.setField("AppName", applicantName);
            form.setField("AppDOB", applicantDOB);
            form.setField("AppMotherName", applicantMotherName);
            form.setField("AppPAN", applicantPAN);
            form.setField("AppAdd1", applicantAdd1);
            form.setField("AppAdd2", applicantAdd2);
            form.setField("AppCity", applicantCity);
            form.setField("AppMobile", applicantMobile);
            form.setField("AppEmail", applicantEmail);
            form.setField("Number", chqNumber);
            form.setField("Dated", chqDate);
            form.setField("DrawnOnBankBranch", chqBankBranch);
            form.setField("AmtFig", amount);
            form.setField("AmtWords", Util.getAmountInWords(amount));
            form.setField("AppBankAccHolderName", nameinBank);
            form.setField("AppBankName", bankName);
            form.setField("AppBranchName", branch);
            form.setField("AppMICR", mICR);
            form.setField("AppAccNo", accountNumber);
            form.setField("AppIFSC", iFSC);
            form.setField("Date", strDate);
            form.setField("Place", place);
            form.setField("NomAppName", applicantName);
            form.setField("NomNameAdd", nomineeName);

            form.setField("NomDate", strDate);
            form.setField("NomAmt", amount);
            form.setField("NomDate_2", strDate);
            form.setField("NomDOB", nomineeDOB);
            form.setField("NomRelation", nomineeRelation);
            form.setField("NomStatus", nomineeStatus);

            form.setField("Witness1Name", witness1Name);
            form.setField("Witness1Add", witness1Add);
            form.setField("Witness2Name", witness2Name);
            form.setField("Witness2Add", witness2Add);
            form.setField("ChqNo", chqNumber);
            form.setField("ChqDate", chqDate);
            form.setField("BankBranch", chqBankBranch);


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
