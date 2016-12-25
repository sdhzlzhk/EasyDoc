package com.glodon.tika;

import com.glodon.tika.vo.DocCatalog;
import com.glodon.tika.vo.SafeDocument;
import org.apache.poi.hwpf.converter.AbstractWordUtils;
import org.apache.poi.hwpf.converter.NumberFormatter;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.HWPFList;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author liuzk
 * @create 2016-12-15 19:02.
 */
public class SafeDocParser {
    private static final Pattern NUM_PATTERN = Pattern.compile("\\d+");
    private static final Pattern CODE_EXTRACT_PATTERN = Pattern.compile("^(\\d+\\.*)+");
    private static final Pattern CODE_VAL_PATTERN = Pattern.compile("^(\\d+\\.?)+(?<=\\d)$");
    private final XWPFDocument xwpfDocument;

    public SafeDocParser(XWPFDocument xwpfDocument) {
        this.xwpfDocument = xwpfDocument;
    }

    public List<DocCatalog> parseDocForCatalog(XWPFDocument document) {
        List<DocCatalog> retList = new ArrayList<>();
        Iterator<XWPFParagraph> paragraphs = document.getParagraphsIterator();
        Map<String,DocCatalog> catalogMap = new HashMap<>();
        XWPFParagraph currentPara = null;
        String style = null;
        String catalogTopCode = null;
        while (paragraphs.hasNext()){
            currentPara = paragraphs.next();
            style = currentPara.getStyleID();
            if(null != style) {
                Matcher styleMatcher = NUM_PATTERN.matcher(style);
                if(styleMatcher.matches()){
//                    System.out.println("NumberId = "+ paragraph.getNumID() + "   NumLevelText = " + paragraph.getNumLevelText() + "   NumFmt = " + paragraph.getNumIlvl() + "    style = " + style + "     content：" + paragraph.getParagraphText());
                    String paragraphText = currentPara.getParagraphText();
                    Matcher catalogMatcher = CODE_EXTRACT_PATTERN.matcher(paragraphText);
                    if(catalogMatcher.find()){
                        String catalogCode = catalogMatcher.group();
                        if(CODE_VAL_PATTERN.matcher(catalogCode).matches()){
                            throw new RuntimeException("该目录编号【"+catalogCode+"】不正确");
                        }
                        if(isTopLevel(catalogCode)){
                            catalogTopCode = catalogCode;
                            if(catalogMap.containsKey(catalogCode)){
                                throw new RuntimeException("该节点【" + catalogCode + "】已经存在");
                            } else {
                                DocCatalog docCatalog = new DocCatalog(paragraphText);
                                catalogMap.put(catalogCode,docCatalog);
                                retList.add(docCatalog);
                            }
                        } else {
                            //子级节点
                            if(isFamilyCode(catalogCode,catalogTopCode)){
                                if (catalogMap.containsKey(catalogCode)) {
                                    throw new RuntimeException("该节点【" + catalogCode + "】已经存在");
                                }
                                if(catalogMap.containsKey(getParentCode(catalogCode))){
                                    //存在其父节点
                                    DocCatalog parentCatalog = catalogMap.get(getParentCode(catalogCode));
                                    DocCatalog docCatalog = new DocCatalog(paragraphText);
                                    parentCatalog.addChildCatalog(docCatalog);
                                    catalogMap.put(catalogCode,docCatalog);
                                } else {
                                    throw new RuntimeException("当前目录【"+catalogCode+"】找不到父节点目录");
                                }
                            } else {
                                throw new RuntimeException("该目录编号不正确，当前目录为【" + catalogTopCode +"】");
                            }
                        }
                    }
                }
            }
        }
        return retList;
    }

    /**
     * 是否同一家族
     * @param code
     * @param topCode
     * @return
     */
    private boolean isFamilyCode(String code,String topCode){
        return code.startsWith(topCode);
    }
    /***
     * 是否顶级节点
     * @param code
     * @return
     */
    private boolean isTopLevel(String code){
        return !code.contains(".");
    }

    /***
     * 父节点编码
     * @param code
     * @return
     */
    private String getParentCode(String code){
        int lastIndex = code.lastIndexOf(".") == -1 ? code.length() : code.lastIndexOf(".");
        return code.substring(0,lastIndex);
    }

    public SafeDocument getSafeDocument(){
        SafeDocument document = new SafeDocument();
        DocCatalog catalog = new DocCatalog("文档目录");
        catalog.setChildCatalog(parseDocForCatalog(this.xwpfDocument));
        document.setDocCatalog(catalog);
        return document;
    }


    public static void main(String[] args) {
        try {
            WordToHtmlConverter.main(new String[]{"C:\\Users\\zhongkai\\Desktop\\header.doc","C:\\Users\\zhongkai\\Desktop\\header.html"});
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
