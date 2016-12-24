package com.glodon.hwpf;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by zhongkai on 2016/12/17.
 */
public class ParseDocHpwfTest {
//    private static final String FILE_PATH = "C:\\Users\\liuzk\\Desktop\\sentry调研文档.doc";
    private static final String FILE_PATH = "C:\\Users\\liuzk\\Desktop\\header.doc";
//    private static final String FILE_PATH = "C:\\Users\\zhongkai\\Desktop\\sentry调研文档.doc";
    private static final String IMAGE_DIR = "\\sentry图片\\";
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
        for(int i = 0; i < sections; i++) {
            Section curSection = docRange.getSection(i);
            System.out.println("###############Section 第" + (i+1) + " 章 has " + curSection.numSections() + " sub sections.#############");
//            picturesTable.extractPicture()
            processSection(document,curSection,i);
            /*for(int j = 0; j < curSection.numParagraphs(); j++){
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
                                    imgPath = System.getProperty("user.dir") + IMAGE_DIR;
                                    File imgOut = new File(imgPath);
                                    if(!imgOut.exists()){
                                        imgOut.mkdirs();
                                    }
                                    imgPath = imgPath + picture.suggestFullFileName();
                                    picture.writeImageContent(new FileOutputStream(imgPath));
                                }
                            }
                            System.out.println(paragraph.isWordWrapped() + " 有图片" + imgPath);
                        } else {
                            if(paragraph.isInTable()) {
                                System.out.println("tableLevel = " + paragraph.getTableLevel());
                                Table table = curSection.getTable(paragraph);
                                TableRow tableRow = null;
                                TableCell tableCell = null;
                                for(int t = 0; t < table.numRows(); t++){
                                    System.out.println("第" + (t + 1) + "行");
                                    tableRow = table.getRow(t);
                                    int cellNums = tableRow.numCells();
                                    for(int col = 0; col < cellNums; col++) {
                                        tableCell = tableRow.getCell(col);
                                        System.out.print(tableCell.text() + " # ");
                                    }
                                    System.out.println();
                                }
                                j += table.numParagraphs();
                                j--;
                                continue;
                            } else {
                                System.out.println("styleIndex "+ paragraph.getStyleIndex() + " #" + paragraph.numCharacterRuns() + " StartOffset "+ paragraph.getStartOffset() + "  EndOffset = " + paragraph.getEndOffset() +"   内容：" + Paragraph.stripFields(text));
                            }
                        }
                    }
                }
            }*/
            System.out.println("##################第 "+(i+1)+" 章结束##################");
        }
    }

    /***
     * 处理目录章节，我们约定目录章节为第二章节
     * @param section
     * @param catalogIndex
     * @param comfirm
     */
    private static void processCatalogSection(HWPFDocumentCore hwpfDoc,final Section section,int catalogIndex,boolean comfirm){
        Section catalogSection = null;
        if(comfirm){
            catalogSection = section.getSection(catalogIndex);
        } else {
            catalogSection = section.getSection(2);
        }
        if(null != catalogSection){
            for(int p = 0;p < catalogSection.numParagraphs(); p++){
                processParagraph(hwpfDoc,catalogSection.getParagraph(p));
            }
        }
    }

    /***
     * 处理普通章节
     * @param section
     * @param rangeIndex
     */
    private static void processSection(HWPFDocumentCore hwpfDoc,final Section section ,int rangeIndex){
            processParagraphs(hwpfDoc,section,Integer.MIN_VALUE);
    }

    /***
     * 处理段落
     * @param hwpfDoc
     * @param paragraph
     */
    private static void processParagraph(HWPFDocumentCore hwpfDoc,Paragraph paragraph) {
        final int charRuns = paragraph.numCharacterRuns();
        if(charRuns == 0){
            return;
        }
        processCharacters(hwpfDoc,paragraph);
    }

    /**
     * 处理段落中字符
     * @param range
     */
    private static void processCharacters(HWPFDocumentCore hwpfDoc,Range range) {
        if(hwpfDoc instanceof HWPFDocument){
            HWPFDocument wordDoc = (HWPFDocument)hwpfDoc;
            String resut = "";
            for (int c = 0; c < range.numCharacterRuns(); c++) {
                CharacterRun chRun = range.getCharacterRun(c);
                if(wordDoc.getPicturesTable().hasPicture(chRun)){
                    Picture picture = wordDoc.getPicturesTable().extractPicture(chRun,false);
                    System.out.println("图片路径：" + processImage(picture));
                    continue;
                }
                //处理一般文本数据
                if(chRun.isSpecialCharacter() || chRun.isObj() || chRun.isOle2()) {
                    continue;
                }
                String text = chRun.text();

                if(null == text || text.isEmpty()) {
                    continue;
                }
                if ( text.charAt(0) == 20 ) {
                    // shall not appear without FIELD_BEGIN_MARK
                    continue;
                }
                if ( text.charAt(0) == 21 )
                {
                    // shall not appear without FIELD_BEGIN_MARK
                    continue;
                }
                if(text.equals("\f")){
                    continue;
                }
                if (text.endsWith( "\r") || (text.charAt(text.length() - 1) == 7)) {
                    text = text.substring( 0, text.length() - 1 );
                }
                /*StringBuilder stringBuilder = new StringBuilder();
                for ( char charChar : text.toCharArray() ) {
                    if ( charChar >= 0x20 || charChar == 0x09
                            || charChar == 0x0A || charChar == 0x0D )
                    {
                        stringBuilder.append(charChar);
                    }
                }
                if (stringBuilder.length() > 0 ) {
                    stringBuilder.toString();
                    stringBuilder.setLength( 0 );
                }*/
                resut = resut + text;
            }
            System.out.println("文本：" + resut);
        } else {
            throw new RuntimeException("仅支持HWPF处理");
        }
    }

    /***
     * 处理图片
     * @param picture
     */
    private static String processImage(Picture picture) {
        DocPictureManager picManger = new DocPictureManager();
        try {
            return picManger.savePicture(picture);
        } catch (Exception ex) {
            ex.printStackTrace();
            return "";
        }
    }

    /***
     * 处理段落
     * @param range
     */
    private static void processParagraphs(HWPFDocumentCore hwpfDoc,Range range,int currentTableLevel) {
        for(int p = 0;p < range.numParagraphs();p++){
            Paragraph para = range.getParagraph(p);
            if(para.isInTable() && para.getTableLevel() != currentTableLevel){
                Table table = range.getTable(para);
                processTable(hwpfDoc,table);
                p += table.numParagraphs();
                p--;
                continue;
            }
            if(para.isInList()){
                HWPFList hwpfList = para.getList();
                String numberText = hwpfList.getNumberText((char) para.getIlvl());
                System.out.println(numberText);
            }
            processParagraph(hwpfDoc,para);
        }
    }

    /**
     * 处理表格
     * @param table
     */
    private static void processTable(HWPFDocumentCore hwpfDoc,Table table) {
        for(int r = 0; r < table.numRows(); r++) {
            TableRow tableRow = table.getRow(r);
            for(int col = 0; col < tableRow.numCells(); col++) {
                TableCell tableCell = tableRow.getCell(col);
//                System.out.print(tableCell.text() + " # ");
                processParagraphs(hwpfDoc,tableCell,table.getTableLevel());
            }
        }
    }

}
