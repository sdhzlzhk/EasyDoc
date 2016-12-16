package com.glodon.tika.vo;

/**
 * @author liuzk
 * @create 2016-12-16 18:46.
 */
public class SafeDocument {
    /**文档名*/
    private String fileName;
    /**文档路径*/
    private String filePath;
    /**文档作者*/
    private String author;
    /**文档创建日期*/
    private String createDate;
    /**文档目录*/
    private DocCatalog docCatalog;

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public String getCreateDate() {
        return createDate;
    }

    public void setCreateDate(String createDate) {
        this.createDate = createDate;
    }

    public DocCatalog getDocCatalog() {
        return docCatalog;
    }

    public void setDocCatalog(DocCatalog docCatalog) {
        this.docCatalog = docCatalog;
    }
}
