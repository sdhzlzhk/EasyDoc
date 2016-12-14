package com.glodon.tika;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Bookmark;
import org.apache.poi.hwpf.usermodel.Bookmarks;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.InputStream;

/**
 * @author liuzk
 * @create 2016-12-14 9:20.
 */
public class HWPFTests {

    @Test
    public void testReadDoc() throws Exception{
        InputStream is = new FileInputStream("C:\\Users\\liuzk\\Desktop\\建筑安全产品\\广西建工三建安全管理系统联合开发方案.docx");
        HWPFDocument doc = new HWPFDocument(is);
        //输出标签信息
        this.printInfo(doc.getBookmarks());
    }

    private void printInfo(Bookmarks bookmarks) {
        int count = bookmarks.getBookmarksCount();
        System.out.println("书签数量：" + count);
        Bookmark bookmark;
        for (int i = 0; i < count; i++){
            bookmark = bookmarks.getBookmark(i);
            System.out.println("书签：" + (i + 1) + "的名称是：" + bookmark.getName());
            System.out.println("开始位置：" + bookmark.getStart());
            System.out.println("结束位置：" + bookmark.getEnd());
        }
    }
}
