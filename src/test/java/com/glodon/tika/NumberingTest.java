package com.glodon.tika;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNum;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.Iterator;
import java.util.List;

/**
 *
 * @author Mark Beardsley
 */
public class NumberingTest {

    private final static String filename = "C:\\Users\\liuzk\\Desktop\\lzhk.docx";

    public static void main(String[] args) throws IOException {
        File file = null;
        FileInputStream fis = null;
        XWPFDocument document = null;
        XWPFNumbering numbering = null;
        XWPFParagraph para = null;
        XWPFNum num = null;
        List<XWPFParagraph> paraList = null;
        Iterator<XWPFParagraph> paraIter = null;
        BigInteger numID = null;
        int numberingID = -1;
        try {
            file = new File(filename);
            fis = new FileInputStream(file);
            document = new XWPFDocument(fis);

            fis.close();
            fis = null;

            numbering = document.getNumbering();

            paraList = document.getParagraphs();
            paraIter = paraList.iterator();
            while(paraIter.hasNext()) {
                para = paraIter.next();
                numID = para.getNumID();
                if(numID != null) {
                    if(numID.intValue() != numberingID) {
                        num = numbering.getNum(numID);
                        numberingID = numID.intValue();
                        System.out.println("Getting details of the new numbering system " + numberingID);
                        System.out.println("It's abstract numID is " + num.getCTNum().getAbstractNumId().getVal().intValue());
                    }
                    else {
                        System.out.println("Iterating through the numbers.");
                    }
                }
                else {
                    System.out.print("Null numID ");
                }
                System.out.println("Text " + para.getParagraphText());
            }
        }
        finally {
            if(fis != null) {
                fis.close();
                fis = null;
            }
        }
    }

}
