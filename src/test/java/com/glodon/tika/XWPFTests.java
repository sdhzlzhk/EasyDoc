package com.glodon.tika;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author liuzk
 * @create 2016-12-14 14:04.
 */
public class XWPFTests {
//    public static final String FILE_NAME = "C:\\Users\\liuzk\\Desktop\\建筑安全产品\\广西建工三建安全管理系统联合开发方案.docx";
//    public static final String FILE_NAME = "C:\\Users\\zhongkai\\Desktop\\sentry调研文档.docx";
//    public static final String FILE_NAME = "E:\\毕业论文\\计算机专业张亚涛-OA-办公自动化系统的设计与实现.docx";
    public static final String FILE_NAME = "C:\\Users\\zhongkai\\Desktop\\sentry调研文档.docx";
//    public static final String FILE_NAME = "D:\\workspace\\help-document\\header.docx";
    private static Pattern NUM_PATTERN = Pattern.compile("\\d+");
    private static final Pattern CODE_EXTRACT_PATTERN = Pattern.compile("^(\\d+\\.*)+");
    private static final Pattern CODE_PATTERN = Pattern.compile("^(\\d+\\.?)+(?<=\\d)$");
    public static void main(String[] args) throws Exception{
        XWPFTests test = new XWPFTests();
        try {
            test.doGenerateSysOut(FILE_NAME);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void doGenerateSysOut(String fileName) throws IOException{
        long startTime = System.currentTimeMillis();
        FileInputStream is = null;
        try {
            is = new FileInputStream(fileName);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            XWPFDocument doc = new XWPFDocument(is);
            is.close();
            is = null;
//            List<XWPFParagraph> paragraphList = doc.getParagraphs();
            Iterator<XWPFParagraph> paragraphList = doc.getParagraphsIterator();
            String style = null;
            XWPFParagraph paragraph = null;
            while(paragraphList.hasNext()){
                paragraph = paragraphList.next();
                style = paragraph.getStyleID();
                if(null != style) {
                    Matcher styleMatcher = NUM_PATTERN.matcher(style);
                    if(styleMatcher.matches()){
                        String paragraphText = paragraph.getParagraphText();
                        System.out.println("NumberId = "+ paragraph.getNumID() + "   NumLevelText = " + paragraph.getNumLevelText() + "   NumFmt = " + paragraph.getNumIlvl() + "    style = " + style + "     content：" + paragraphText);
                        Matcher catalogMatcher = CODE_EXTRACT_PATTERN.matcher(paragraphText);
                        if(catalogMatcher.find()){
                            String catalogCode = catalogMatcher.group();
                            catalogMatcher = CODE_PATTERN.matcher(catalogCode);
                            System.out.println("code = " + catalogCode + "   是否符合格式 ：" + catalogMatcher.matches());
                        }
                    }
                }
            }
        } finally {
            if(null != is){
                is.close();
                is = null;
            }
        }


    }
}
