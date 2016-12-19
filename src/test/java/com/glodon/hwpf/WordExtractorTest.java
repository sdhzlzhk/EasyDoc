package com.glodon.hwpf;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * @author liuzk
 * @create 2016-12-19 17:08.
 */
public class WordExtractorTest {
    private static final String FILE_PATH = "C:\\Users\\liuzk\\Desktop\\sentry调研文档.doc";
    public static void main(String[] args) throws IOException {
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
        WordExtractor wordExtractor = new WordExtractor(document);
        try {

            String[] paragraphText = wordExtractor.getParagraphText();//"\u000c"
            for(String str : paragraphText){
                System.out.println(str);
            }
        } finally {
            wordExtractor.close();
        }
    }
}
