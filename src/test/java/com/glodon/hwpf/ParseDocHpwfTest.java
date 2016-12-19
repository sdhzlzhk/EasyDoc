package com.glodon.hwpf;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.UUID;

/**
 * Created by zhongkai on 2016/12/17.
 */
public class ParseDocHpwfTest {
    private static final String FILE_PATH = "C:\\Users\\liuzk\\Desktop\\sentry调研文档.doc";
//    private static final String FILE_PATH = "C:\\Users\\zhongkai\\Desktop\\sentry调研文档.doc";
    private static final String IMAGE_DIR = "\\sentry图片";
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
        PicturesTable picturesTable = document.getPicturesTable();
        List<Picture> pictures = picturesTable.getAllPictures();
        for(int i = 0; i < sections; i++){
            Section curSection = docRange.getSection(i);
            System.out.println("###############Section 第" + (i+1) + " 章 has " + curSection.numSections() + " sub sections.#############");
//            picturesTable.extractPicture()
            for(int j = 0; j < curSection.numParagraphs(); j++){
                Paragraph paragraph = curSection.getParagraph(j);
                //TODO: 判断空行
                if(paragraph.numCharacterRuns() > 1) {
                    String text = paragraph.text();
                    if (text.contains("HYPERLINK") && text.indexOf('\u0013') > -1 && text.indexOf('\u0015') > -1){
                        text = Paragraph.stripFields(text);
                        while (text.indexOf("\u0013") > -1 || text.indexOf("\t") > -1){
                            int exists13 = text.indexOf("\u0013");
                            int tab = text.indexOf("\t");
                            if(exists13 > -1){
                                text = text.substring(0, exists13);
                            }
                            if(tab > -1){
                                text = text.substring(0, tab);
                            }
                        }
                        System.out.println("styleIndex "+ paragraph.getStyleIndex() + " " + paragraph.numCharacterRuns()  + "  目录：" + text);
                    } else {
                        if(text.indexOf("\u0001") > -1){
                            String imgPath = null;
                            for(int k = 0; k < paragraph.numCharacterRuns(); k++){
                                CharacterRun characterRun = paragraph.getCharacterRun(k);
                                if(picturesTable.hasPicture(characterRun)){
                                    Picture picture = picturesTable.extractPicture(characterRun,false);
                                    imgPath = IMAGE_DIR + UUID.randomUUID().toString() + ".jpg";
                                    picture.writeImageContent(new FileOutputStream(imgPath));
                                }
                            }
                            System.out.println(paragraph.isWordWrapped() + " 有图片" + imgPath);
                        } else {
                            if(paragraph.isInTable()) {
                                Table table = curSection.getTable(paragraph);
                                System.out.println("表格:"+table.numRows()+"行");
                            } else {
                                System.out.println("styleIndex "+ paragraph.getStyleIndex() + " #" + paragraph.numCharacterRuns() + " StartOffset "+ paragraph.getStartOffset() + "  EndOffset = " + paragraph.getEndOffset() +"   内容：" + Paragraph.stripFields(text));
                            }
                        }
                    }
                }
            }
            System.out.println("##################第 "+(i+1)+" 章结束##################");
            /*if(i > 1){
                break;
            }*/
        }
    }

}
