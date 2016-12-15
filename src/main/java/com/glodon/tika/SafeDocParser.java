package com.glodon.tika;

import com.glodon.tika.vo.DocCatalog;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author liuzk
 * @create 2016-12-15 19:02.
 */
public class SafeDocParser {
    private static final Pattern NUM_PATTERN = Pattern.compile("\\d+");

    public List<DocCatalog> parseDocForCatalog(XWPFDocument document) {
        Iterator<XWPFParagraph> paragraphs = document.getParagraphsIterator();
        XWPFParagraph currentPara = null;
        String style = null;
        while (paragraphs.hasNext()){
            currentPara = paragraphs.next();
            style = currentPara.getStyleID();
            if(null != style) {
                Matcher styleMatcher = NUM_PATTERN.matcher(style);
                if(styleMatcher.matches()){
//                    System.out.println("NumberId = "+ paragraph.getNumID() + "   NumLevelText = " + paragraph.getNumLevelText() + "   NumFmt = " + paragraph.getNumIlvl() + "    style = " + style + "     content：" + paragraph.getParagraphText());
                    currentPara.getElementType().name();
                }
            }
        }
        return null;
    }
}
