package com.glodon.tika;

import com.glodon.tika.vo.DocCatalog;
import com.glodon.tika.vo.SafeDocument;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.AbstractWordConverter;
import org.apache.poi.hwpf.converter.HtmlDocumentFacade;
import org.apache.poi.hwpf.usermodel.*;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.util.List;

/**
 * @author liuzk
 * @create 2016-12-16 16:53.
 */
public class SafeHtmlConvert extends AbstractWordConverter {
    private final HtmlDocumentFacade htmlDocumentFacade;

    public SafeHtmlConvert(Document document) {
        this.htmlDocumentFacade = new HtmlDocumentFacade(document);
    }

    @Override
    public Document getDocument() {
        return this.htmlDocumentFacade.getDocument();
    }

    public void processDocument(SafeDocument safeDocument){
        DocCatalog docCatalog = safeDocument.getDocCatalog();
        this.processCatalog(docCatalog);
    }

    private void processCatalog(DocCatalog docCatalog) {
        if(null != docCatalog){
            List<DocCatalog> childCatalog = docCatalog.getChildCatalog();
            if(null != childCatalog){
                Element div = this.htmlDocumentFacade.createBlock();
                this.htmlDocumentFacade.addStyleClass(div,"d",getDefaultStyle());
                this.htmlDocumentFacade.getBody().appendChild(div);
                DocCatalog parentCatalog = null;
                for(DocCatalog catalog : childCatalog){
                    parentCatalog = catalog.getParentCatalog();
                    if (null != parentCatalog) {
                        this.htmlDocumentFacade.createHyperlink(parentCatalog.getCatalogName());
                    }

                }
            }
        }
    }

    private String getDefaultStyle() {
        return  "margin: " + 10 + "in " + 10 + "in " + 10 + "in " + 10 + "in;";
    }

    @Override
    protected void outputCharacters(Element element, CharacterRun characterRun, String s) {

    }

    @Override
    protected void processBookmarks(HWPFDocumentCore hwpfDocumentCore, Element element, Range range, int i, List<Bookmark> list) {

    }

    @Override
    protected void processDocumentInformation(SummaryInformation summaryInformation) {

    }

    @Override
    protected void processDrawnObject(HWPFDocument hwpfDocument, CharacterRun characterRun, OfficeDrawing officeDrawing, String s, Element element) {

    }

    @Override
    protected void processEndnoteAutonumbered(HWPFDocument hwpfDocument, int i, Element element, Range range) {

    }

    @Override
    protected void processFootnoteAutonumbered(HWPFDocument hwpfDocument, int i, Element element, Range range) {

    }

    @Override
    protected void processHyperlink(HWPFDocumentCore hwpfDocumentCore, Element element, Range range, int i, String s) {

    }

    @Override
    protected void processImage(Element element, boolean b, Picture picture, String s) {

    }

    @Override
    protected void processImageWithoutPicturesManager(Element element, boolean b, Picture picture) {

    }

    @Override
    protected void processLineBreak(Element element, CharacterRun characterRun) {

    }

    @Override
    protected void processPageBreak(HWPFDocumentCore hwpfDocumentCore, Element element) {

    }

    @Override
    protected void processPageref(HWPFDocumentCore hwpfDocumentCore, Element element, Range range, int i, String s) {

    }

    @Override
    protected void processParagraph(HWPFDocumentCore hwpfDocumentCore, Element element, int i, Paragraph paragraph, String s) {

    }

    @Override
    protected void processSection(HWPFDocumentCore hwpfDocumentCore, Section section, int i) {

    }

    @Override
    protected void processTable(HWPFDocumentCore hwpfDocumentCore, Element element, Table table) {

    }
}
