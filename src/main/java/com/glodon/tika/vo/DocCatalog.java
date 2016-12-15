package com.glodon.tika.vo;

import java.io.Serializable;
import java.util.List;

/**
 * @author liuzk
 * @create 2016-12-15 19:10.
 * 目录对象
 */
public class DocCatalog implements Serializable {
    /**父目录*/
    private DocCatalog parentCatalog;
    /**目录名称*/
    private String catalogName;
    /**子目录*/
    private List<DocCatalog> childCatalog;

    public DocCatalog(String catalogName) {
        this.catalogName = catalogName;
    }

    public DocCatalog getParentCatalog() {
        return parentCatalog;
    }

    public void setParentCatalog(DocCatalog parentCatalog) {
        this.parentCatalog = parentCatalog;
    }

    public List<DocCatalog> getChildCatalog() {
        return childCatalog;
    }

    public void setChildCatalog(List<DocCatalog> childCatalog) {
        this.childCatalog = childCatalog;
    }

    public String getCatalogName() {
        return catalogName;
    }

    public void setCatalogName(String catalogName) {
        this.catalogName = catalogName;
    }
}
