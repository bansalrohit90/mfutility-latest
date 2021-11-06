package com.returns;

import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.io.RandomAccessRead;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;

import java.io.File;
import java.io.FileInputStream;

public class PDFReader {

    public static void main(String[] args) {
        PDFParser parser = null;
        PDDocument pdDoc = null;
        COSDocument cosDoc = null;
        PDFTextStripper pdfStripper;

        String parsedText;
        String fileName = "E:\\Files\\Small Files\\PDF\\JDBC.pdf";
        File file = new File(fileName);
        try {
            parser = new PDFParser((RandomAccessRead) new FileInputStream(file));
            parser.parse();
            cosDoc = parser.getDocument();
            pdfStripper = new PDFTextStripper();
            pdDoc = new PDDocument(cosDoc);
            parsedText = pdfStripper.getText(pdDoc);
            System.out.println(parsedText.replaceAll("[^A-Za-z0-9. ]+", ""));
        } catch (Exception e) {
            e.printStackTrace();
            try {
                if (cosDoc != null)
                    cosDoc.close();
                if (pdDoc != null)
                    pdDoc.close();
            } catch (Exception e1) {
                e1.printStackTrace();
            }

        }


        try {
            PDDocument document = null;
            document = PDDocument.load(new File("/Users/rohibans/Downloads/komal.pdf"), "123456", MemoryUsageSetting.setupMainMemoryOnly());
            document.getClass();
            if (document.isEncrypted()) {
                PDFTextStripperByArea stripper = new PDFTextStripperByArea();
                stripper.setSortByPosition(true);
                PDFTextStripper Tstripper = new PDFTextStripper();
                String st = Tstripper.getText(document);
                System.out.println("Text:" + st);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
