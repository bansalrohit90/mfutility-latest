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
public final class FillBankChangeRegistrationForm {

    private static String CAN;
    private static String NAME;
    private static Iterator<Row> rowIterator;
    private static Row firstRow;

    //Fillable1 - checked growth & monthly; Fillable - unchecked growth & monthly
    private static final String sourceBankMandateFile = Util.getDirectoryPath() + "/NCT-Mandate-Fillable.pdf";
    private static String destinationFile;
    private static final String sourceExcelFile = Util.getDirectoryPath() + "/fill-can-registration1.xlsm";

    public static void main(String[] args) throws Exception {

        Iterator<Row> iterator = Util.getUtilIteratorObject(sourceExcelFile, 5);
        fillFromExcel(iterator);
        destinationFile = Util.getDestinationDirectoryPath() + "/" + NAME + "_" + CAN + "_" + "BankChange" + ".pdf";
        editPdfDocument();
        /*if (printFile.equalsIgnoreCase("yes")) {
//            Util.printPdfOutput2(destinationFile,NAME+"SIP_REG");
        }*/
    }

    private static void fillFromExcel(Iterator<Row> iterator) {
        if (iterator.hasNext()) {
            firstRow = iterator.next();
            CAN = Util.getCellValue(firstRow, 0);
            NAME = Util.getCellValue(firstRow, 1);
        }
        rowIterator = iterator;
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
            form.setField("CAN", CAN);
            form.setField("FHName", NAME);
            form.setField("Date", strDate);
            Row row = firstRow;
            int i = 1;
            while (row != null) {
                form.setField("B" + i + "Act", Util.getCellValue(row, 2));
                form.setField("B" + i + "MICR", Util.getCellValue(row, 3));
                form.setField("B" + i + "IFSC", Util.getCellValue(row, 4));
                form.setField("B" + i + "BANK", Util.getCellValue(row, 5));
                form.setField("B" + i + "BR", Util.getCellValue(row, 6));
                form.setField("B" + i + "CITY", Util.getCellValue(row, 7));
                if (rowIterator.hasNext())
                    row = rowIterator.next();
                else
                    row = null;
                i++;
            }


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
