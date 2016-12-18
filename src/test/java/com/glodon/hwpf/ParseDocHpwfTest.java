package com.glodon.hwpf;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by zhongkai on 2016/12/17.
 */
public class ParseDocHpwfTest {
    private static final String FILE_PATH = "C:\\Users\\zhongkai\\Desktop\\sentry调研文档.doc";
    public static final String NEXT_PAGE = "\f";
    public static void main(String[] args) throws IOException {
        //TODO first create a HWPFDocument
        FileInputStream inputStream = null;
        HWPFDocument document = null;
        try {
            inputStream = new FileInputStream(FILE_PATH);
            document = new HWPFDocument(HWPFDocument.verifyAndBuildPOIFS(inputStream));
        } finally {
            if(null != inputStream){
                inputStream.close();
            }
        }
        //TODO fetch the range with getRange()
        Range docRange = document.getRange();
        int paragraphs = docRange.numParagraphs();
        int sections = docRange.numSections();
        int characterRuns = docRange.numCharacterRuns();
        System.out.println("paragraphs = " + paragraphs + "  sections = " + sections +"  characterRuns = " + characterRuns);
        for(int i = 0; i < sections; i++){
            Section curSection = docRange.getSection(i);
            System.out.println("###############Section 第" + (i+1) + " 章 has " + curSection.numSections() + " sub sections.#############");
            int sectionPara = curSection.numParagraphs();
            for(int j = 0; j < sectionPara; j++){
                Paragraph paragraph = curSection.getParagraph(j);
                //TODO: 判断空行
                if(paragraph.numCharacterRuns() > 1){
                    String text = paragraph.text();
                    if (text.contains("HYPERLINK") && text.indexOf('\u0013') > -1 && text.indexOf('\u0015') > -1){
                        System.out.println(" StartOffset "+ paragraph.getStartOffset() + "  EndOffset = " + paragraph.getEndOffset() + "目录：" + Paragraph.stripFields(text));
                    } else {
                        System.out.println(paragraph.numCharacterRuns() + " StartOffset "+ paragraph.getStartOffset() + "  EndOffset = " + paragraph.getEndOffset() +"   内容：" + Paragraph.stripFields(text));
                    }

                }
            }
            System.out.println("##################第 "+(i+1)+" 章结束##################");
            if(i > 1){
                break;
            }

        }
    }

}