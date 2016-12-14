package com.glodon.tika;


import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author liuzk
 * @create 2016-12-14 14:04.
 */
public class XWPFTests {
//    public static final String FILE_NAME = "C:\\Users\\liuzk\\Desktop\\建筑安全产品\\广西建工三建安全管理系统联合开发方案.docx";
    public static final String FILE_NAME = "C:\\Users\\zhongkai\\Desktop\\sentry调研文档.docx";
//    public static final String FILE_NAME = "E:\\毕业论文\\计算机专业张亚涛-OA-办公自动化系统的设计与实现.docx";
    private static Pattern NUM_PATTERN = Pattern.compile("\\d+");
    public static void main(String[] args) {
        XWPFTests test = new XWPFTests();
        try {
            test.doGenerateSysOut(FILE_NAME);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void doGenerateSysOut(String fileName) throws IOException{
        long startTime = System.currentTimeMillis();
        InputStream is = null;
        try {
            is = new FileInputStream(fileName);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        XWPFDocument doc = new XWPFDocument(is);
//        CTDocument1 ctDocument1 = doc.getDocument();
        List<XWPFParagraph> paragraphList = doc.getParagraphs();
        String style = null;
        for(XWPFParagraph paragraph : paragraphList){
//            if(paragraph.isEmpty()) continue;
            style = paragraph.getStyleID();
            if(null != style){
                Matcher styleMatcher = NUM_PATTERN.matcher(style);
                if(styleMatcher.matches()){
                    /*XWPFNumbering numbering = paragraph.getDocument().getNumbering();
                    XWPFNum xwpfNum = numbering.getNum(paragraph.getNumID());
                    System.out.println(xwpfNum.getCTNum().getNumId());*/
                    System.out.println("NumberId = "+ paragraph.getNumID() +"    style = " + style + "     content：" + paragraph.getParagraphText());
                }
            }

        }

    }
}
